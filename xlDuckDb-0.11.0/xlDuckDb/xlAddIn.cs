using ExcelDna.Integration;
using System.Diagnostics.CodeAnalysis;

namespace xlDuckDb;

public class xlAddIn : IExcelAddIn
{
    private const string ErrorPrefix = "#ERR";

    public void AutoOpen()
    {
        ExcelIntegration.RegisterUnhandledExceptionHandler(
            ex =>
            {
                if (ex is Exception e) return $"{ErrorPrefix} - {e.Message}";
                return $"{ErrorPrefix} - {ex}";
            });
    }

    public void AutoClose()
    {
        DuckDbConnectionManager.Instance.Dispose();
    }

    // ════════════════════════════════════════════════════════════════════
    // DuckDbQuery — 単一 SQL クエリ
    // ════════════════════════════════════════════════════════════════════

    [Experimental("DuckDBNET001")]
    [ExcelFunction(
        Description = "Executes a single DuckDB SQL query with Prepared Statement.\n" +
                      "SQL: DuckDB SQL statement. Use $1, $2, ... for parameters.\n" +
                      "DatabaseFile: Path to .duckdb file (default: READ ONLY). Add ';RW' to open as writable. Empty for in-memory DB.\n" +
                      "Params/Ranges: Values or ranges for $1, $2, ... (row-first expansion).",
        Category = "xlDuckDb",
        HelpTopic = "https://duckdb.org/docs/index",
        IsMacroType = true,
        IsThreadSafe = true,
        Name = "DuckDbQuery")]
    public static object DuckDbQuery(
        [ExcelArgument(
            Name = "SQL",
            Description = "DuckDB SQL statement. Use $1, $2, ... for parameters. Example: 'SELECT * FROM my_table WHERE id = $1'",
            AllowReference = true)]
        object query,
        [ExcelArgument(
            Name = "DatabaseFile",
            Description = "Path to .duckdb file (default: READ ONLY). Add ';RW' to open as writable. Empty for in-memory DB.")]
        string dataSource,
        [ExcelArgument(
            Name = "Params/Ranges",
            Description = "Values or ranges for $1, $2, ... (row-first expansion). String: '@NORM:...'. Range: row-first expansion.",
            AllowReference = true)]
        params object[] ranges)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNull;

        var queryString = ResolveQueryString(query);
        if (queryString is null) return ExcelError.ExcelErrorValue;

        var (parameters, rangeAddrs, preloaded) = ResolveRanges(queryString, ranges);

        var result = DuckDbHelper.ExecuteQuery(
            queryString, dataSource, parameters, rangeAddrs, preloaded);

        return result;
    }

    // ════════════════════════════════════════════════════════════════════
    // DuckDbQueryMulti — 複数 SQL クエリを UNION ALL
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// SQL 式をセル範囲で受け取り、各 SQL を順番に実行して UNION ALL した結果を返す。
    ///
    /// ── SQL 範囲 ──────────────────────────────────────────────────────────
    ///   縦1列（A1:A5）または横1行（A1:E1）を自動判定。空セルはスキップ。
    ///   行数 ≥ 列数なら縦、行数 &lt; 列数なら横と判定する。
    ///
    /// ── 引数範囲（省略可）────────────────────────────────────────────────
    ///   行インデックス = SQL 番号（SQL 範囲上の順番に対応）
    ///   列インデックス = $1, $2, ... の引数番号
    ///
    ///   例: SQL 範囲が縦3行の場合
    ///     argsRange[0, 0] → SQL[0] の $1
    ///     argsRange[0, 1] → SQL[0] の $2
    ///     argsRange[1, 0] → SQL[1] の $1
    ///     （空セルを含む末尾の引数は自動トリム）
    ///
    ///   横 SQL でも引数範囲の構造は同じ（行=SQL番号、列=引数番号）。
    ///   SQL 範囲の方向に関わらず引数範囲の構造は一定。
    ///
    ///   文字列セルには @NORM: プレフィックスで正規化オプションを指定可能。
    ///
    /// ── 結果 ─────────────────────────────────────────────────────────────
    ///   1番目の SQL の結果（ヘッダー行含む）を先頭に、
    ///   2番目以降はデータ行のみ（ヘッダー行なし）を連結。
    ///   列数が不一致の場合は #VALUE! を返す。
    /// </summary>
    [Experimental("DuckDBNET001")]
    [ExcelFunction(
        Description = "Executes multiple DuckDB SQL queries from a cell range and returns UNION ALL results (VSTACK).\n" +
                      "SQLRange: 1-column or 1-row range of SQLs.\n" +
                      "DatabaseFile: Path to .duckdb file (default: READ ONLY). Add ';RW' to open as writable. Empty for in-memory DB.\n" +
                      "ArgsRange: Optional. Parameters for each SQL ($1, $2, ...). Row=SQL index, Col=param index.",
        Category = "xlDuckDb",
        HelpTopic = "https://duckdb.org/docs/index",
        IsMacroType = true,
        IsThreadSafe = true,
        Name = "DuckDbQueryMulti")]
    public static object DuckDbQueryMulti(
        [ExcelArgument(
            Name = "SQLRange",
            Description = "1-column or 1-row range of SQLs. Empty cells are skipped. Vertical/horizontal auto-detected.",
            AllowReference = true)]
        object sqlRange,
        [ExcelArgument(
            Name = "DatabaseFile",
            Description = "Path to .duckdb file (default: READ ONLY). Add ';RW' to open as writable. Empty for in-memory DB.")]
        string dataSource,
        [ExcelArgument(
            Name = "ArgsRange",
            Description = "Optional. Parameters for each SQL ($1, $2, ...). Row=SQL index, Col=param index.",
            AllowReference = true)]
        object argsRange)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNull;

        var sqlData = ResolveRangeData(sqlRange);
        if (sqlData is null) return ExcelError.ExcelErrorValue;

        var sqlRows = sqlData.GetLength(0);
        var sqlCols = sqlData.GetLength(1);
        var isVertical = sqlRows >= sqlCols; // 縦横の自動判定

        var sqlList = ExtractSqlList(sqlData, isVertical);
        if (sqlList.Count == 0) return new object[,] { { ExcelError.ExcelErrorNA } };

        object[,]? argsData = null;
        if (argsRange is ExcelReference argsRef)
        {
            var argsAddress = XlCall.Excel(XlCall.xlfReftext, argsRef, true).ToString();
            if (argsAddress is not null)
                argsData = ExcelHelper.GetRangeValues(argsAddress);
        }

        List<object[]>? unionRows   = null;
        int?            headerCount = null;
        int             totalRows = 0;
        bool            limitReached = false;
        int             maxRows = xlDuckDbConfig.MaxOutputRows;

        for (var i = 0; i < sqlList.Count && !limitReached; i++)
        {
            var (sql, _) = sqlList[i];
            if (string.IsNullOrWhiteSpace(sql)) continue;

            var parameters = ExtractArgsRow(argsData, i);

            object[,] queryResult;
            try
            {
                queryResult = DuckDbHelper.ExecuteQuery(sql, dataSource, parameters);
            }
            catch (Exception ex)
            {
                return $"{ErrorPrefix} [SQL#{i + 1}] {ex.Message}";
            }

            if (queryResult.Length == 1) continue; // 0件はスキップ

            var qRows = queryResult.GetLength(0);
            var qCols = queryResult.GetLength(1);

            if (headerCount is null)
            {
                headerCount = qCols;
                unionRows   = new List<object[]>(qRows);
                for (var r = 0; r < qRows; r++)
                {
                    if (totalRows >= maxRows) { limitReached = true; break; }
                    var row = new object[qCols];
                    for (var c = 0; c < qCols; c++) row[c] = queryResult[r, c];
                    unionRows.Add(row);
                    totalRows++;
                }
            }
            else
            {
                if (qCols != headerCount) return ExcelError.ExcelErrorValue;

                unionRows ??= [];
                for (var r = 1; r < qRows; r++)
                {
                    if (totalRows >= maxRows) { limitReached = true; break; }
                    var row = new object[qCols];
                    for (var c = 0; c < qCols; c++) row[c] = queryResult[r, c];
                    unionRows.Add(row);
                    totalRows++;
                }
            }
        }

        if (unionRows is null || unionRows.Count == 0)
            return new object[,] { { ExcelError.ExcelErrorNA } };

        // 行数制限に到達した場合、エラー行を追加
        if (limitReached)
        {
            var cols = unionRows[0].Length;
            var errorRow = new object[cols];
            for (int i = 0; i < cols; i++) errorRow[i] = "#ERR: 出力最大行数に到達";
            unionRows.Add(errorRow);
        }

        return unionRows.AsMultiDimensionalArray();
    }

    // ════════════════════════════════════════════════════════════════════
    // DuckDbNormalize — 正規化内容のステップ別確認
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 正規化の各ステップ後の中間値を行として返す。
    ///
    /// 戻り値（2列）:
    ///   列1: ステップ名（"Input" / "After NFKC+CaseFold" / ... / "Final"）
    ///   列2: そのステップ後の文字列値
    ///
    ///   オプションが有効かつ適用前後で値が変化したステップのみ中間行を出力する。
    ///   "Input" と "Final" は常に出力される。
    ///
    /// セル範囲を渡した場合:
    ///   各セルに対して同じオプションでステップ別結果を生成し、
    ///   セルごとにブロックを縦に連結して返す。
    ///   各ブロックの先頭に "Input" 行、末尾に "Final" 行が入る。
    ///   ブロック間は空行（空白セル）で区切られる。
    /// </summary>
    [ExcelFunction(
        Description = "Applies text normalization step by step and returns each intermediate result as a row.\n" +
                      "Input: Input string or cell range.\n" +
                      "NormalizeOptions: Bit-flag (default: 511=All).\n" +
                      "AdditionalRemove: Characters to remove.\n" +
                      "AdditionalReplace: Replacement mappings (from,to per line).",
        Category = "xlDuckDb",
        IsMacroType = true,
        IsThreadSafe = true,
        Name = "DuckDbNormalize")]
    public static object DuckDbNormalize(
        [ExcelArgument(
            Name = "Input",
            Description = "Input string or cell range to normalize step by step.",
            AllowReference = true)]
        object input,
        [ExcelArgument(
            Name = "NormalizeOptions",
            Description = "Bit-flag (default: 511 = All). 1=NFKC+CaseFold 2=LongVowel 4=EmojiRemove 8=FullwidthAlnum 16=HalfAlnum 32=HalfKana 64=FullSymbol 128=Whitespace 256=Bracket 511=All")]
        double normalizeOptions,
        [ExcelArgument(
            Name = "AdditionalRemove",
            Description = "Characters to remove. E.g. '①②③'")]
        string additionalRemove,
        [ExcelArgument(
            Name = "AdditionalReplace",
            Description = "Replacement mappings. 'from,to' per line. E.g. '㈱,株式会社'")]
        string additionalReplace)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNull;

        var options     = (NormalizeOptions)(int)normalizeOptions;
        var removeSpec  = string.IsNullOrWhiteSpace(additionalRemove)  ? null : additionalRemove;
        var replaceSpec = string.IsNullOrWhiteSpace(additionalReplace) ? null : additionalReplace;

        // ── 入力を文字列リストとして収集 ─────────────────────────────────
        var inputs = new List<string>();

        if (input is ExcelReference reference)
        {
            var data = ExcelHelper.GetRangeValues(
                XlCall.Excel(XlCall.xlfReftext, reference, true).ToString()!);
            var rows = data.GetLength(0);
            var cols = data.GetLength(1);
            for (var r = 0; r < rows; r++)
                for (var c = 0; c < cols; c++)
                    inputs.Add(data[r, c]?.ToString() ?? string.Empty);
        }
        else
        {
            inputs.Add(input is string s ? s : input?.ToString() ?? string.Empty);
        }

        // ── 各入力のステップ結果をまとめて 2D 配列に積み上げる ─────────
        var allRows = new List<object[]>();

        for (var idx = 0; idx < inputs.Count; idx++)
        {
            // 複数入力の場合、2番目以降の前に空行を挿入して視認性を向上
            if (idx > 0)
                allRows.Add(["", ""]);

            var steps = TextNormalizer.NormalizeSteps(
                inputs[idx], options, removeSpec, replaceSpec);

            foreach (var (label, value) in steps)
                allRows.Add([label, value]);
        }

        if (allRows.Count == 0)
            return new object[,] { { string.Empty, string.Empty } };

        return allRows.AsMultiDimensionalArray();
    }

    // ════════════════════════════════════════════════════════════════════
    // DuckDbCascadeSearch — 複数 SQL クエリを順次実行し、累計出力行数が maxRows に到達するまで VSTACK で返す
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 複数SQLを順次実行し、累計出力行数がmaxRowsに到達するまでVSTACK(縦連結)で返す。
    /// maxRowsに到達した時点で以降のSQLは実行しない。
    /// </summary>
    [Experimental("DuckDBNET001")]
    [ExcelFunction(
        Description = "Executes multiple DuckDB SQL queries from a cell range and returns VSTACK results up to MaxRows. Stops when the cumulative output row count reaches MaxRows.\n" +
                      "SQLRange: 1-column or 1-row range of SQLs.\n" +
                      "DatabaseFile: Path to .duckdb file (default: READ ONLY). Add ';RW' to open as writable. Empty for in-memory DB.\n" +
                      "ArgsRange: Optional. Parameters for each SQL ($1, $2, ...). Row=SQL index, Col=param index.\n" +
                      "MaxRows: Maximum cumulative output row count.",
        Category = "xlDuckDb",
        HelpTopic = "https://duckdb.org/docs/index",
        IsMacroType = true,
        IsThreadSafe = true,
        Name = "DuckDbCascadeSearch")]
    public static object DuckDbCascadeSearch(
        [ExcelArgument(
            Name = "SQLRange",
            Description = "1-column or 1-row range of SQLs. Empty cells are skipped. Vertical/horizontal auto-detected.",
            AllowReference = true)]
        object sqlRange,
        [ExcelArgument(
            Name = "DatabaseFile",
            Description = "Path to .duckdb file (default: READ ONLY). Add ';RW' to open as writable. Empty for in-memory DB.")]
        string dataSource,
        [ExcelArgument(
            Name = "ArgsRange",
            Description = "Optional. Parameters for each SQL ($1, $2, ...). Row=SQL index, Col=param index.",
            AllowReference = true)]
        object argsRange,
        [ExcelArgument(
            Name = "MaxRows",
            Description = "Maximum cumulative output row count.")]
        int maxRows)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNull;

        var sqlData = ResolveRangeData(sqlRange);
        if (sqlData is null) return ExcelError.ExcelErrorValue;

        var sqlRows = sqlData.GetLength(0);
        var sqlCols = sqlData.GetLength(1);
        var isVertical = sqlRows >= sqlCols;
        var sqlList = ExtractSqlList(sqlData, isVertical);
        if (sqlList.Count == 0) return new object[,] { { ExcelError.ExcelErrorNA } };

        object[,]? argsData = null;
        if (argsRange is ExcelReference argsRef)
        {
            var argsAddress = XlCall.Excel(XlCall.xlfReftext, argsRef, true).ToString();
            if (argsAddress is not null)
                argsData = ExcelHelper.GetRangeValues(argsAddress);
        }

        List<object[]>? unionRows = null;
        int? headerCount = null;
        int totalRows = 0;
        bool limitReached = false;

        for (var i = 0; i < sqlList.Count && !limitReached; i++)
        {
            var (sql, _) = sqlList[i];
            if (string.IsNullOrWhiteSpace(sql)) continue;
            var parameters = ExtractArgsRow(argsData, i);
            object[,] queryResult;
            try
            {
                queryResult = DuckDbHelper.ExecuteQuery(sql, dataSource, parameters);
            }
            catch (Exception ex)
            {
                return $"{ErrorPrefix} [SQL#{i + 1}] {ex.Message}";
            }
            if (queryResult.Length == 1) continue;
            var qRows = queryResult.GetLength(0);
            var qCols = queryResult.GetLength(1);
            if (headerCount is null)
            {
                headerCount = qCols;
                unionRows = new List<object[]>(qRows);
                for (var r = 0; r < qRows; r++)
                {
                    if (totalRows >= maxRows) { limitReached = true; break; }
                    var row = new object[qCols];
                    for (var c = 0; c < qCols; c++) row[c] = queryResult[r, c];
                    unionRows.Add(row);
                    totalRows++;
                }
            }
            else
            {
                if (qCols != headerCount) return ExcelError.ExcelErrorValue;
                for (var r = 1; r < qRows; r++)
                {
                    if (totalRows >= maxRows) { limitReached = true; break; }
                    var row = new object[qCols];
                    for (var c = 0; c < qCols; c++) row[c] = queryResult[r, c];
                    unionRows.Add(row);
                    totalRows++;
                }
            }
        }
        if (unionRows is null || unionRows.Count == 0)
            return new object[,] { { ExcelError.ExcelErrorNA } };
        // 行数制限に到達した場合、エラー行を追加
        if (limitReached)
        {
            var cols = unionRows[0].Length;
            var errorRow = new object[cols];
            for (int i = 0; i < cols; i++) errorRow[i] = "#ERR: 累計出力最大行数に到達";
            unionRows.Add(errorRow);
        }
        return unionRows.AsMultiDimensionalArray();
    }

    // ════════════════════════════════════════════════════════════════════
    // 内部ヘルパー
    // ════════════════════════════════════════════════════════════════════

    /// <summary>SQL 引数（query）を文字列に解決する。</summary>
    internal static string? ResolveQueryString(object query) => query switch
    {
        string s when !string.IsNullOrWhiteSpace(s) => s,
        ExcelReference r => r.GetValue() switch
        {
            string s when !string.IsNullOrWhiteSpace(s) => s,
            double d => d.ToString(),
            _ => null
        },
        _ => null
    };

    /// <summary>
    /// ranges を走査してスカラーパラメータ・xlRange アドレス・先読みデータに仕分ける。
    /// </summary>
    internal static (object[]? parameters, string[]? rangeAddrs,
        IReadOnlyDictionary<string, object[,]>? preloaded)
        ResolveRanges(string queryString, IEnumerable<object> ranges)
    {
        var scalarParams     = new List<object>();
        var xlRangeAddresses = new List<string>();
        var preloadedRanges  = new Dictionary<string, object[,]>();
        var usesXlRange      = queryString.Contains("xlRange");

        foreach (var arg in ranges)
        {
            if (arg is ExcelReference reference)
            {
                var address = XlCall.Excel(XlCall.xlfReftext, reference, true).ToString()
                    ?? throw new InvalidOperationException(
                        "Failed to determine the address of the supplied range.");
                var isSingleCell = reference.RowFirst == reference.RowLast
                                && reference.ColumnFirst == reference.ColumnLast;

                if (!isSingleCell && usesXlRange
                    && xlRangeAddresses.Count < CountXlRangePlaceholders(queryString))
                {
                    xlRangeAddresses.Add(address);
                    if (!preloadedRanges.ContainsKey(address))
                        preloadedRanges[address] = ExcelHelper.GetRangeValues(address);
                }
                else
                {
                    var data = isSingleCell
                        ? new object[,] { { NormalizeScalar(reference.GetValue()) } }
                        : ExcelHelper.GetRangeValues(address);
                    FlattenRowFirst(data, scalarParams);
                }
            }
            else if (arg is string s)
                scalarParams.Add(NormParam.Parse(s));
            else
                scalarParams.Add(NormalizeScalar(arg));
        }

        var parameters = scalarParams.Count > 0 ? scalarParams.ToArray() : null;
        var addrs      = xlRangeAddresses.Count > 0 ? xlRangeAddresses.ToArray() : null;
        IReadOnlyDictionary<string, object[,]>? preloaded =
            preloadedRanges.Count > 0 ? preloadedRanges : null;

        return (parameters, addrs, preloaded);
    }

    /// <summary>ExcelReference または文字列から 2D データを取得する。</summary>
    private static object[,]? ResolveRangeData(object arg)
    {
        if (arg is ExcelReference reference)
        {
            var address = XlCall.Excel(XlCall.xlfReftext, reference, true).ToString();
            if (address is null) return null;
            return ExcelHelper.GetRangeValues(address);
        }
        if (arg is string s && !string.IsNullOrWhiteSpace(s))
            return new object[,] { { s } };
        return null;
    }

    /// <summary>
    /// 2D データから SQL 文字列リストを抽出する。
    /// isVertical=true  → 各行の列0 が SQL
    /// isVertical=false → 行0 の各列が SQL
    /// </summary>
    private static List<(string Sql, int OriginalIndex)> ExtractSqlList(
        object[,] data, bool isVertical)
    {
        var list = new List<(string, int)>();
        if (isVertical)
        {
            for (var r = 0; r < data.GetLength(0); r++)
            {
                var cell = data[r, 0]?.ToString();
                if (!string.IsNullOrWhiteSpace(cell))
                    list.Add((cell!, r));
            }
        }
        else
        {
            for (var c = 0; c < data.GetLength(1); c++)
            {
                var cell = data[0, c]?.ToString();
                if (!string.IsNullOrWhiteSpace(cell))
                    list.Add((cell!, c));
            }
        }
        return list;
    }

    /// <summary>
    /// 引数範囲の sqlIndex 行目を取得して パラメータ配列を返す。
    ///
    /// argsData 構造: 行=SQL番号、列=$1,$2,...
    ///   argsData[sqlIndex, 0] → $1
    ///   argsData[sqlIndex, 1] → $2
    ///   末尾の DBNull は自動トリム。
    /// </summary>
    private static object[]? ExtractArgsRow(object[,]? argsData, int sqlIndex)
    {
        if (argsData is null) return null;
        if (sqlIndex >= argsData.GetLength(0)) return null;

        var args = new List<object>();
        for (var c = 0; c < argsData.GetLength(1); c++)
        {
            var v = argsData[sqlIndex, c];
            args.Add(v is string s ? NormParam.Parse(s) : NormalizeScalar(v));
        }

        // 末尾の DBNull をトリム（固定幅範囲の余白を除去）
        while (args.Count > 0 && args[args.Count - 1] is DBNull)
            args.RemoveAt(args.Count - 1);

        return args.Count > 0 ? args.ToArray() : null;
    }

    internal static object NormalizeScalar(object? value) => value switch
    {
        double d     => d,
        string s     => s,
        bool b       => b,
        ExcelError   => DBNull.Value,
        ExcelEmpty   => DBNull.Value,
        ExcelMissing => DBNull.Value,
        null         => DBNull.Value,
        _            => value
    };

    private static void FlattenRowFirst(object[,] data, List<object> target)
    {
        var rows = data.GetLength(0);
        var cols = data.GetLength(1);
        for (var r = 0; r < rows; r++)
            for (var c = 0; c < cols; c++)
                target.Add(NormalizeScalar(data[r, c]));
    }

    private static int CountXlRangePlaceholders(string sql)
    {
        var count = 0;
        var idx   = 0;
        while ((idx = sql.IndexOf("xlRange", idx, StringComparison.Ordinal)) >= 0)
        {
            count++;
            idx += "xlRange".Length;
        }
        return count;
    }
}
