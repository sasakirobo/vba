using System.Diagnostics.CodeAnalysis;
using System.Numerics;
using System.Text;
using DuckDB.NET.Data;
using DuckDB.NET.Native;
using ExcelDna.Integration;
using System.Diagnostics;

namespace xlDuckDb;

public static class DuckDbHelper
{
    /// <summary>
    /// DuckDB クエリを Prepared Statement で実行して結果を 2D 配列で返す。
    ///
    /// パラメータバインド:
    ///   ranges の 2D 配列を行優先で展開し $1, $2, $3, ... にバインドする。
    ///   例) A1:C2 → $1=A1, $2=B1, $3=C1, $4=A2, $5=B2, $6=C2
    ///
    /// キャッシュ戦略:
    ///   1. xlRange を使わないクエリ → 結果キャッシュを確認し、ヒットで即返却
    ///   2. Prepared Statement をキャッシュから取得または新規 Prepare
    ///   3. パラメータをバインドして実行
    ///   4. xlRange を使わない場合のみ結果をキャッシュに登録
    ///
    /// 出力最大行数制限: xlDuckDbConfig.MaxOutputRows で制御。超過時はエラー行を返す。
    /// </summary>
    [Experimental("DuckDBNET001")]
    public static object[,] ExecuteQuery(
        string query,
        string dataSource = "",
        object[]? parameters = null,          // ranges を行優先展開したスカラー配列
        string[]? rangeAddresses = null,      // xlRange テーブル関数用アドレス
        IReadOnlyDictionary<string, object[,]>? preloadedRanges = null)
    {
        if (string.IsNullOrEmpty(query))
            throw new ArgumentException("Value cannot be null or empty.", nameof(query));

        var normalizedSource = DuckDbConnectionManager.NormalizeDataSource(dataSource);
        var hasXlRange = rangeAddresses is { Length: > 0 } && query.Contains("xlRange");

        // ── 1. 結果キャッシュ確認（xlRange クエリは対象外）──────────────────
        if (!hasXlRange)
        {
            if (DuckDbConnectionManager.Instance.ResultCache
                    .TryGet(normalizedSource, query, parameters, out var cached))
                return cached;
        }

        using var pooled = DuckDbConnectionManager.Instance.Acquire(dataSource);
        var conn = pooled.Connection;

        // ── 2. xlRange テーブル関数の登録（接続ごとに 1 度）─────────────────
        var resolvedQuery = query;
        if (hasXlRange)
        {
            if (!pooled.XlRangeRegistered)
            {
                var capturedRanges = preloadedRanges
                    ?? throw new InvalidOperationException(
                        "preloadedRanges must be provided when xlRange is used with IsThreadSafe = true.");

                conn.RegisterTableFunction<string>(
                    "xlRange",
                    p => ExcelRangeTableFunctions.ResultCallback(p, capturedRanges),
                    ExcelRangeTableFunctions.MapperCallback);

                pooled.MarkXlRangeRegistered();
            }

            resolvedQuery = SubstituteXlRange(query, rangeAddresses!);
        }

        // ── 3. Prepared Statement をキャッシュから取得または新規 Prepare ─────
        var cmd = pooled.GetOrPrepare(resolvedQuery);

        // ── 4. パラメータバインド（$1, $2, $3, ... に順番に設定）────────────
        cmd.Parameters.Clear();
        if (parameters is { Length: > 0 })
        {
            for (var i = 0; i < parameters.Length; i++)
            {
                var p = cmd.CreateParameter();
                p.ParameterName = "$" + (i + 1);
                p.Value = parameters[i] ?? DBNull.Value;
                cmd.Parameters.Add(p);
            }
        }

        // ── 5. 実行 ──────────────────────────────────────────────────────────
        using var reader = cmd.ExecuteReader();
        if (!reader.HasRows || reader.IsClosed)
            return new object[,] { { ExcelError.ExcelErrorNA } };

        var result = ReadResultsWithLimit(reader, xlDuckDbConfig.MaxOutputRows, out bool limitReached);

        // ── 6. 結果キャッシュへ登録（xlRange なしのみ）──────────────────────
        if (!hasXlRange)
        {
            DuckDbConnectionManager.Instance.ResultCache
                .Set(normalizedSource, query, parameters, result);
        }

        // ── 7. 行数制限到達時はエラー行を追加 ──────────────────────────────
        if (limitReached)
        {
            Debug.WriteLine($"[xlDuckDb] 出力最大行数({xlDuckDbConfig.MaxOutputRows})に到達しました。以降の行は切り捨てられました。");
            var rows = result.GetLength(0);
            var cols = result.GetLength(1);
            var errorRow = new object[cols];
            for (int i = 0; i < cols; i++) errorRow[i] = "#ERR: 出力最大行数に到達";
            var newResult = new object[rows + 1, cols];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    newResult[r, c] = result[r, c];
            for (int c = 0; c < cols; c++)
                newResult[rows, c] = errorRow[c];
            return newResult;
        }

        return result;
    }

    // ── ヘルパー: xlRange 置換 ────────────────────────────────────────────

    private static string SubstituteXlRange(string query, string[] addresses)
    {
        if (addresses.Length == 1)
        {
            var enc = Convert.ToBase64String(Encoding.Unicode.GetBytes(addresses[0]));
            return query.Replace("xlRange", $"xlRange($${enc}$$)");
        }
        for (var i = 0; i < addresses.Length; i++)
        {
            var enc = Convert.ToBase64String(Encoding.Unicode.GetBytes(addresses[i]));
            query = query.Replace($"xlRange[{i + 1}]", $"xlRange($${enc}$$)");
        }
        return query;
    }

    // ── ヘルパー: DataReader → object[,]（最大行数制限付き） ──────────────
    private static object[,] ReadResultsWithLimit(DuckDBDataReader reader, int maxRows, out bool limitReached)
    {
        var rows = new List<object[]>();

        // フィールド型の分類
        var fc = reader.FieldCount;
        var bigInt   = new bool[fc];
        var blob     = new bool[fc];
        var dec      = new bool[fc];
        var flt      = new bool[fc];
        var tspan    = new bool[fc];
        var timeTz   = new bool[fc];
        var dateOnly = new bool[fc];
        var uuid     = new bool[fc];

        var header = new object[fc];
        for (var i = 0; i < fc; i++)
        {
            header[i] = reader.GetName(i);
            var t = reader.GetFieldType(i);
            bigInt[i]   = t == typeof(BigInteger) || t == typeof(ulong);
            blob[i]     = t == typeof(Stream);
            dec[i]      = t == typeof(decimal);
            flt[i]      = t == typeof(float);
            tspan[i]    = t == typeof(TimeSpan);
            timeTz[i]   = t == typeof(DateTimeOffset);
            dateOnly[i] = t == typeof(DateOnly);
            uuid[i]     = t == typeof(Guid);
        }
        rows.Add(header);

        int rowCount = 0;
        foreach (var _ in reader)
        {
            if (rowCount >= maxRows)
            {
                limitReached = true;
                break;
            }
            var row = new object[fc];
            for (var i = 0; i < fc; i++)
            {
                if      (bigInt[i])   row[i] = reader.GetInt64(i);
                else if (blob[i])   { using var sr = new StreamReader(reader.GetStream(i), Encoding.UTF8); row[i] = sr.ReadToEnd(); }
                else if (dec[i])      row[i] = decimal.ToDouble(reader.GetDecimal(i));
                else if (flt[i])      row[i] = (double)reader.GetFloat(i);
                else if (tspan[i])
                {
                    var v = reader.GetValue(i);
                    row[i] = v switch
                    {
                        DuckDBTimeOnly ddb => new DateTime(1899, 12, 30) + TimeSpan.FromTicks(ddb.Ticks),
                        TimeSpan ts        => new DateTime(1899, 12, 30) + ts,
                        _                  => v
                    };
                }
                else if (timeTz[i])   row[i] = new DateTime(1899, 12, 30) + TimeSpan.FromTicks(((DateTimeOffset)reader.GetValue(i)).Ticks);
                else if (dateOnly[i]) row[i] = reader.GetDateTime(i);
                else if (uuid[i])     row[i] = reader.GetGuid(i).ToString();
                else                  row[i] = reader.GetValue(i);
            }
            // DBNull / NaN / Inf → Excel エラー値に変換
            for (var i = 0; i < fc; i++)
                row[i] = row[i] switch
                {
                    DBNull                                          => ExcelError.ExcelErrorNA,
                    double d when double.IsNaN(d) || double.IsInfinity(d) => ExcelError.ExcelErrorNum,
                    float  f when float.IsNaN(f)  || float.IsInfinity(f)  => ExcelError.ExcelErrorNum,
                    _                                               => row[i]
                };
            rows.Add(row);
            rowCount++;
        }
        limitReached = rowCount >= maxRows;
        return rows.AsMultiDimensionalArray();
    }
}
