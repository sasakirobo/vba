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
    /// DuckDB クエリをシンプルに実行して結果を 2D 配列で返す。
    ///
    /// キャッシュ戦略:
    ///   1. 結果キャッシュを確認し、ヒットで即返却
    ///   2. Prepared Statement をキャッシュから取得または新規 Prepare
    ///   3. クエリを実行
    ///   4. 結果をキャッシュに登録
    ///
    /// ⚠️ 注意: DuckDB.NET v1.5.0 は $1, $2, $3 ... スタイルのパラメータバインディングをサポートしていません。
    /// パラメータが必要な場合は、以下の推奨パターンを使用してください:
    ///   • xlRange テーブル関数を使用する（動的にセル範囲をクエリする）
    ///   • SQL に値を直接埋め込む（Excel 関数で SQL 文字列を動的構築）
    ///   • DuckDbNormalize, DuckDbReplace, DuckDbRemove などの UDF 関数を組み合わせる
    ///
    /// 出力最大行数制限: xlDuckDbConfig.MaxOutputRows で制御。超過時はエラー行を返す。
    /// </summary>
    [Experimental("DuckDBNET001")]
    public static object[,] ExecuteQuery(
        string query,
        string dataSource = "")
    {
        if (string.IsNullOrEmpty(query))
            throw new ArgumentException("Value cannot be null or empty.", nameof(query));

        var normalizedSource = DuckDbConnectionManager.NormalizeDataSource(dataSource);

        // ── 1. 結果キャッシュ確認 ─────────────────────────────────────────
        if (DuckDbConnectionManager.Instance.ResultCache
                .TryGet(normalizedSource, query, parameters: null, out var cached))
            return cached;

        using var pooled = DuckDbConnectionManager.Instance.Acquire(normalizedSource);
        var conn = pooled.Connection;

        // ── 2. Prepared Statement をキャッシュから取得 ────────────────────
        var cmd = pooled.GetOrPrepare(query);

        // ── 3. 実行 ──────────────────────────────────────────────────────────
        using var reader = cmd.ExecuteReader();
        if (!reader.HasRows || reader.IsClosed)
            return new object[,] { { ExcelError.ExcelErrorNA } };

        var result = ReadResultsWithLimit(reader, xlDuckDbConfig.MaxOutputRows, out bool limitReached);

        // ── 4. 結果キャッシュへ登録 ──────────────────────────────────────
        DuckDbConnectionManager.Instance.ResultCache
            .Set(normalizedSource, query, parameters: null, result);

        // ── 5. 行数制限到達時はエラー行を追加 ────────────────────────────
        if (limitReached)
        {
            Debug.WriteLine($"[xlDuckDb] 出力最大行数({xlDuckDbConfig.MaxOutputRows})に到達しました。以降の行は切り捨てられました。");
            var rows = result.GetLength(0);
            var cols = result.GetLength(1);
            var newResult = new object[rows + 1, cols];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    newResult[r, c] = result[r, c];
            for (int c = 0; c < cols; c++)
                newResult[rows, c] = "#ERR: 出力最大行数に到達";
            return newResult;
        }

        return result;
    }

    // ════════════════════════════════════════════════════════════════════
    // 列の型分類
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// DuckDB.NET が返す CLR 型を列ごとに分類し、読み取り戦略を決定する。
    /// </summary>
    private enum ColKind
    {
        /// <summary>GetValue() で得た値を ToExcelValue() に通す汎用パス</summary>
        Generic,
        /// <summary>BigInteger / ulong → long (オーバーフロー時は文字列)</summary>
        BigInt,
        /// <summary>Stream (BLOB) → UTF-8 文字列</summary>
        Blob,
        /// <summary>decimal → double</summary>
        Decimal,
        /// <summary>float → double</summary>
        Float,
        /// <summary>TimeSpan / DuckDBTimeOnly → Excel シリアル値 (TIME)</summary>
        TimeSpan,
        /// <summary>DateTimeOffset → Excel シリアル値 (TIMETZ)</summary>
        TimeTz,
        /// <summary>DateOnly → DateTime (Excel DATE)</summary>
        DateOnly,
        /// <summary>Guid → string</summary>
        Uuid,
    }

    private static ColKind ClassifyColumn(Type t)
    {
        if (t == typeof(BigInteger) || t == typeof(ulong))   return ColKind.BigInt;
        if (t == typeof(Stream))                              return ColKind.Blob;
        if (t == typeof(decimal))                             return ColKind.Decimal;
        if (t == typeof(float))                               return ColKind.Float;
        if (t == typeof(TimeSpan))                            return ColKind.TimeSpan;
        if (t == typeof(DateTimeOffset))                      return ColKind.TimeTz;
        if (t == typeof(DateOnly))                            return ColKind.DateOnly;
        if (t == typeof(Guid))                                return ColKind.Uuid;
        return ColKind.Generic;
    }

    // Excel の DATE シリアル値の起点（1899-12-30）
    private static readonly DateTime ExcelEpoch = new(1899, 12, 30);

    // ════════════════════════════════════════════════════════════════════
    // 汎用値変換（Generic パス・および後処理共通）
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// DuckDB から取得した任意の CLR 値を Excel が受け取れる型に変換する。
    /// Excel が受け取れる型: double / string / bool / DateTime / ExcelError
    ///
    /// 変換規則:
    ///   DBNull / null                  → ExcelError.ExcelErrorNA
    ///   double (NaN / Inf)             → ExcelError.ExcelErrorNum
    ///   float  (NaN / Inf)             → ExcelError.ExcelErrorNum
    ///   float  (正常値)                → (double)
    ///   decimal                        → (double)
    ///   sbyte / short / int / long     → (double)  ※Excel は数値を double で保持
    ///   byte / ushort / uint / ulong   → (double)
    ///   BigInteger                     → long に収まれば double、超過時は string
    ///   bool                           → bool (そのまま)
    ///   DateTime                       → DateTime (そのまま)
    ///   DateOnly                       → DateTime
    ///   DateTimeOffset                 → Excel シリアル値 (double)
    ///   TimeSpan / DuckDBTimeOnly      → Excel TIME シリアル値 (DateTime)
    ///   Guid                           → string
    ///   Stream (BLOB)                  → UTF-8 string
    ///   Int128 / UInt128               → long に収まれば double、超過時は string
    ///   その他                          → ToString()（null/空 → ExcelErrorNA）
    /// </summary>
    private static object ToExcelValue(object? raw)
    {
        switch (raw)
        {
            case null:
            case DBNull:
                return ExcelError.ExcelErrorNA;

            // ── 浮動小数点 ───────────────────────────────────────────────
            case double d:
                return double.IsNaN(d) || double.IsInfinity(d)
                    ? (object)ExcelError.ExcelErrorNum
                    : d;
            case float f:
                return float.IsNaN(f) || float.IsInfinity(f)
                    ? (object)ExcelError.ExcelErrorNum
                    : (double)f;

            // ── 固定小数点 ───────────────────────────────────────────────
            case decimal dm:
                return decimal.ToDouble(dm);

            // ── 整数（符号あり）────────────────────────────────────────
            case long l:    return (double)l;
            case int iv:    return (double)iv;
            case short s:   return (double)s;
            case sbyte sb:  return (double)sb;

            // ── 整数（符号なし）────────────────────────────────────────
            case ulong ul:
                // ulong は long 範囲を超える可能性があるため checked
                try   { return (double)(long)ul; }
                catch { return ul.ToString(); }
            case uint ui:   return (double)ui;
            case ushort us: return (double)us;
            case byte b:    return (double)b;

            // ── 大整数 ──────────────────────────────────────────────────
            case BigInteger bi:
                if (bi >= long.MinValue && bi <= long.MaxValue)
                    return (double)(long)bi;
                return bi.ToString();

            // ── 論理値 ──────────────────────────────────────────────────
            case bool bv:
                return bv;

            // ── 日時 ────────────────────────────────────────────────────
            case DateTime dt:
                return dt;

            case DateOnly donly:
                return donly.ToDateTime(TimeOnly.MinValue);

            case DateTimeOffset dto:
                // DateTimeOffset → Excel シリアル値
                // UTC に変換してから ExcelEpoch との差分を日数 double で返す
                return (dto.UtcDateTime - ExcelEpoch).TotalDays;

            case TimeSpan ts:
                // TIME 型: ExcelEpoch + TimeSpan で DateTime として返す
                // 負値や 24 時間超もそのまま受け入れる
                return ExcelEpoch + ts;

            case DuckDBTimeOnly ddbTime:
                return ExcelEpoch + TimeSpan.FromTicks(ddbTime.Ticks);

            // ── GUID ────────────────────────────────────────────────────
            case Guid g:
                return g.ToString();

            // ── Stream (BLOB) ────────────────────────────────────────────
            case Stream st:
                try
                {
                    using var sr = new StreamReader(st, Encoding.UTF8, leaveOpen: false);
                    return (object)sr.ReadToEnd();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[xlDuckDb] BLOB 読み取りエラー: {ex.Message}");
                    return ExcelError.ExcelErrorValue;
                }

            // ── Int128 / UInt128（.NET 7+、DuckDB HUGEINT）────────────
            case Int128 i128:
                if (i128 >= long.MinValue && i128 <= long.MaxValue)
                    return (double)(long)i128;
                return i128.ToString();

            case UInt128 u128:
                if (u128 <= (UInt128)long.MaxValue)
                    return (double)(long)u128;
                return u128.ToString();

            // ── その他（フォールバック）──────────────────────────────────
            default:
                var str = raw.ToString();
                return string.IsNullOrEmpty(str) ? (object)ExcelError.ExcelErrorNA : str;
        }
    }

    // ════════════════════════════════════════════════════════════════════
    // DataReader → object[,]（最大行数制限付き）
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// DataReader の全行を読み取り、Excel 互換の 2D 配列に変換する。
    ///
    /// 各セルの変換は <see cref="ToExcelValue"/> に集約しており、
    /// 列の型分類（ColKind）によって最適な読み取りメソッドを使いつつ、
    /// 例外が発生した場合はすべて <see cref="ToExcelValue"/> の汎用パスへフォールバックする。
    /// </summary>
    private static object[,] ReadResultsWithLimit(DuckDBDataReader reader, int maxRows, out bool limitReached)
    {
        var fc = reader.FieldCount;

        // ── ヘッダー行 ──────────────────────────────────────────────────
        var header = new object[fc];
        var kinds  = new ColKind[fc];
        var types  = new Type[fc];
        for (var i = 0; i < fc; i++)
        {
            header[i] = reader.GetName(i);
            kinds[i]  = ClassifyColumn(reader.GetFieldType(i));
            types[i]  = reader.GetFieldType(i);
        }

        var rows = new List<object[]> { header };

        int rowCount = 0;
        limitReached = false;

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
                if (reader.IsDBNull(i))
                {
                    row[i] = ExcelError.ExcelErrorNA;
                    continue;
                }
                try
                {
                    // すべてToExcelValue経由でExcel互換型に統一
                    row[i] = kinds[i] switch
                    {
                        ColKind.BigInt => ToExcelValue(ReadBigInt(reader, i)),
                        ColKind.Blob => ToExcelValue(ReadBlob(reader, i)),
                        ColKind.Decimal => ToExcelValue(reader.GetDecimal(i)),
                        ColKind.Float => ToExcelValue((double)reader.GetFloat(i)),
                        ColKind.TimeSpan => ToExcelValue(ReadTimeSpan(reader, i)),
                        ColKind.TimeTz => ToExcelValue(ReadTimeTz(reader, i)),
                        ColKind.DateOnly => ToExcelValue(reader.GetDateTime(i)),
                        ColKind.Uuid => ToExcelValue(reader.GetGuid(i).ToString()),
                        _ => ToExcelValue(reader.GetValue(i)),
                    };
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[xlDuckDb] 列[{i}] ({reader.GetName(i)}) 変換エラー: {ex.GetType().Name}: {ex.Message}");
                    row[i] = ExcelError.ExcelErrorValue;
                }
                row[i] = row[i] switch
                {
                    double d when double.IsNaN(d) || double.IsInfinity(d) => ExcelError.ExcelErrorNum,
                    float  f when float.IsNaN(f)  || float.IsInfinity(f)  => ExcelError.ExcelErrorNum,
                    _ => row[i],
                };
            }
            rows.Add(row);
            rowCount++;
        }
        if (!limitReached)
            limitReached = rowCount >= maxRows;
        return rows.AsMultiDimensionalArray();
    }

    // ── 個別読み取りヘルパー ──────────────────────────────────────────────

    /// <summary>
    /// BigInteger / ulong を Excel 互換の double または string に変換する。
    /// long 範囲内であれば double、超過時は文字列。
    /// </summary>
    private static object ReadBigInt(DuckDBDataReader reader, int i)
    {
        var raw = reader.GetValue(i);
        return raw switch
        {
            BigInteger bi =>
                bi >= long.MinValue && bi <= long.MaxValue
                    ? (object)(double)(long)bi
                    : bi.ToString(),
            ulong ul =>
                ul <= (ulong)long.MaxValue
                    ? (object)(double)(long)ul
                    : ul.ToString(),
            long l  => (double)l,
            // DuckDB.NET が将来 Int128 を返す場合も対応
            Int128  i128 =>
                i128 >= long.MinValue && i128 <= long.MaxValue
                    ? (object)(double)(long)i128
                    : i128.ToString(),
            UInt128 u128 =>
                u128 <= (UInt128)long.MaxValue
                    ? (object)(double)(long)u128
                    : u128.ToString(),
            _ => ToExcelValue(raw),
        };
    }

    /// <summary>
    /// BLOB (Stream) を UTF-8 文字列に読み取る。
    /// NULL や空ストリームは ExcelErrorNA を返す。
    /// </summary>
    private static object ReadBlob(DuckDBDataReader reader, int i)
    {
        var st = reader.GetStream(i);
        if (st is null) return ExcelError.ExcelErrorNA;
        using var sr = new StreamReader(st, Encoding.UTF8, leaveOpen: false);
        return (object)sr.ReadToEnd();
    }

    /// <summary>
    /// TIME 型（TimeSpan / DuckDBTimeOnly）を Excel TIME シリアル値（DateTime）に変換する。
    /// </summary>
    private static object ReadTimeSpan(DuckDBDataReader reader, int i)
    {
        var v = reader.GetValue(i);
        return v switch
        {
            DuckDBTimeOnly ddb => ExcelEpoch + TimeSpan.FromTicks(ddb.Ticks),
            TimeSpan ts        => ExcelEpoch + ts,
            DBNull             => ExcelError.ExcelErrorNA,
            null               => ExcelError.ExcelErrorNA,
            _                  => ToExcelValue(v),
        };
    }

    /// <summary>
    /// TIMETZ 型（DateTimeOffset）を Excel シリアル値（double）に変換する。
    /// UTC 変換後に ExcelEpoch からの経過日数を返す。
    /// </summary>
    private static object ReadTimeTz(DuckDBDataReader reader, int i)
    {
        var v = reader.GetValue(i);
        return v switch
        {
            DateTimeOffset dto => (object)(dto.UtcDateTime - ExcelEpoch).TotalDays,
            DBNull             => ExcelError.ExcelErrorNA,
            null               => ExcelError.ExcelErrorNA,
            _                  => ToExcelValue(v),
        };
    }
}
