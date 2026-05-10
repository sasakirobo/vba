using ExcelDna.Integration;
using System.Diagnostics.CodeAnalysis;
using System.Diagnostics;

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

        // ── Ribbon UI は .dna ファイルで自動登録 ──────────────────────
        // RibbonUICallbacks の public メソッドが自動的に .dna の コールバックとして
        // 登録される（ExcelRibbon を継承することで）
        Debug.WriteLine("[xlDuckDb] xlDuckDb Add-in opened. Ribbon UI ready.");
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
        Description = "Executes a single DuckDB SQL query and returns results as a 2D array.\n" +
                      "SQL: DuckDB SQL statement. Use xlRange(address) for dynamic ranges or embed values directly.\n" +
                      "DatabaseFile: Path to .duckdb file (default: READ ONLY). Add ';RW' to open as writable. Empty for in-memory DB.",
        Category = "xlDuckDb",
        HelpTopic = "https://duckdb.org/docs/index",
        IsMacroType = true,
        IsThreadSafe = true,
        Name = "DuckDbQuery")]
    public static object DuckDbQuery(
        [ExcelArgument(
            Name = "SQL",
            Description = "DuckDB SQL statement. Use xlRange(address) for dynamic ranges or embed values directly in SQL.",
            AllowReference = true)]
        object query,
        [ExcelArgument(
            Name = "DatabaseFile",
            Description = "Path to .duckdb file (default: READ ONLY). Add ';RW' to open as writable. Empty for in-memory DB.")]
        string dataSource)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNull;

        var queryString = ResolveQueryString(query);
        if (queryString is null) return ExcelError.ExcelErrorValue;

        var result = DuckDbHelper.ExecuteQuery(queryString, dataSource);

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
    /// ── SQL の値埋め込み ──────────────────────────────────────────────────
    ///   DuckDB.NET v1.5.0 はパラメータバインディング（$1, $2, ...）をサポートしていません。
    ///   SQL 内で値を直接埋め込むか、xlRange() テーブル関数を使用してください。
    ///   例：
    ///     SELECT * FROM table WHERE id = CAST(123 AS INTEGER)
    ///     SELECT * FROM xlRange('$A$1:$B$10') AS t(id INT, name VARCHAR)
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
                      "DatabaseFile: Path to .duckdb file (default: READ ONLY). Add ';RW' to open as writable. Empty for in-memory DB.",
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
        string dataSource)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNull;

        var sqlData = ResolveRangeData(sqlRange);
        if (sqlData is null) return ExcelError.ExcelErrorValue;

        var sqlRows = sqlData.GetLength(0);
        var sqlCols = sqlData.GetLength(1);
        var isVertical = sqlRows >= sqlCols; // 縦横の自動判定

        var sqlList = ExtractSqlList(sqlData, isVertical);
        if (sqlList.Count == 0) return new object[,] { { ExcelError.ExcelErrorNA } };

        List<object[]>? unionRows   = null;
        int?            headerCount = null;
        int             totalRows = 0;
        bool            limitReached = false;
        int             maxRows = xlDuckDbConfig.MaxOutputRows;

        for (var i = 0; i < sqlList.Count && !limitReached; i++)
        {
            var (sql, _) = sqlList[i];
            if (string.IsNullOrWhiteSpace(sql)) continue;

            object[,] queryResult;
            try
            {
                queryResult = DuckDbHelper.ExecuteQuery(sql, dataSource);
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
    /// テキスト正規化を適用し、最終結果のみを返す。
    /// 
    /// 単一入力の場合:
    ///   正規化後の文字列を直接返す。
    /// 
    /// セル範囲入力の場合:
    ///   各セルに正規化を適用し、セル配列を返す。
    /// </summary>
    [ExcelFunction(
        Description = "Applies text normalization and returns the final result only.\n" +
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
            Description = "Input string or cell range to normalize.",
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

        // ── セル範囲の場合: 各セルに正規化を適用して配列で返す ─────────────────
        if (input is ExcelReference reference)
        {
            var address = XlCall.Excel(XlCall.xlfReftext, reference, true).ToString()!;
            var data = ExcelHelper.GetRangeValues(address);
            var rows = data.GetLength(0);
            var cols = data.GetLength(1);
            
            var results = new object[rows, cols];
            for (var r = 0; r < rows; r++)
                for (var c = 0; c < cols; c++)
                {
                    var cellValue = data[r, c]?.ToString() ?? string.Empty;
                    try
                    {
                        results[r, c] = TextNormalizer.Normalize(cellValue, options, removeSpec, replaceSpec);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[xlDuckDb] DuckDbNormalize セル[{r},{c}] 正規化エラー: {ex.Message}");
                        results[r, c] = ExcelError.ExcelErrorValue;
                    }
                }
            
            return results;
        }

        // ── 単一入力の場合: 正規化結果を直接返す ──────────────────────────────
        var inputStr = input is string s ? s : input?.ToString() ?? string.Empty;
        try
        {
            return TextNormalizer.Normalize(inputStr, options, removeSpec, replaceSpec);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[xlDuckDb] DuckDbNormalize 正規化エラー: {ex.Message}");
            return ExcelError.ExcelErrorValue;
        }
    }

    // ════════════════════════════════════════════════════════════════════
    // DuckDbCascadeSearch — 複数 SQL クエリを順次実行し、累計出力行数が maxRows に到達するまで VSTACK で返す
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 複数SQLを順次実行し、累計出力行数がmaxRowsに到達するまでVSTACK(縦連結)で返す。
    /// maxRowsに到達した時点で以降のSQLは実行しない。
    ///
    /// ⚠️ 注意: DuckDB.NET v1.5.0 はパラメータバインディング（$1, $2, ...）をサポートしていません。
    /// SQL 内で値を直接埋め込むか、xlRange() テーブル関数を使用してください。
    /// </summary>
    [Experimental("DuckDBNET001")]
    [ExcelFunction(
        Description = "Executes multiple DuckDB SQL queries from a cell range and returns VSTACK results up to MaxRows. Stops when the cumulative output row count reaches MaxRows.\n" +
                      "SQLRange: 1-column or 1-row range of SQLs.\n" +
                      "DatabaseFile: Path to .duckdb file (default: READ ONLY). Add ';RW' to open as writable. Empty for in-memory DB.\n" +
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
            Name = "Reserved",
            Description = "Reserved. Pass null or empty.",
            AllowReference = true)]
        object reserved,
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

        List<object[]>? unionRows = null;
        int? headerCount = null;
        int totalRows = 0;
        bool limitReached = false;

        for (var i = 0; i < sqlList.Count && !limitReached; i++)
        {
            var (sql, _) = sqlList[i];
            if (string.IsNullOrWhiteSpace(sql)) continue;

            object[,] queryResult;
            try
            {
                queryResult = DuckDbHelper.ExecuteQuery(sql, dataSource);
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
    // DuckDbResultCacheList — クエリ結果キャッシュの一覧を表示
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// DuckDBクエリ結果キャッシュの一覧を表示します。
    /// </summary>
    [ExcelFunction(
        Description = "DuckDBクエリ結果キャッシュの一覧を表示します。",
        Category = "xlDuckDb",
        IsMacroType = true,
        IsThreadSafe = true,
        Name = "DuckDbResultCacheList")]
    public static object[,] DuckDbResultCacheList()
    {
        var cache = DuckDbConnectionManager.Instance.ResultCache;
        var entries = cache.GetAllEntries().ToList();
        var result = new object[entries.Count + 1, 3];
        result[0, 0] = "DataSource";
        result[0, 1] = "SQL";
        result[0, 2] = "RowCount";
        for (int i = 0; i < entries.Count; i++)
        {
            result[i + 1, 0] = entries[i].Key.DataSource;
            result[i + 1, 1] = entries[i].Key.Sql;
            result[i + 1, 2] = entries[i].Value.GetLength(0) - 1; // ヘッダー除くデータ行数
        }
        return result;
    }

    // ════════════════════════════════════════════════════════════════════
    // 非同期クエリ関数
    // ════════════════════════════════════════════════════════════════════

    // ════════════════════════════════════════════════════════════════════
    // DuckDbQueryAsync — 非同期 SQL クエリ（Excel-DNA AsyncUtil対応）
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// DuckDB SQL クエリを非同期で実行し、結果を返す。
    ///
    /// ── 非同期処理の利点 ──────────────────────────────────────────────────
    ///   • Excel UI をブロッキングしない（計算中もExcel操作可能）
    ///   • 複数シートの並列クエリ実行
    ///   • ネットワーク遅延（HTTPS）の影響を軽減
    ///   • ロングランニングクエリの進捗表示対応
    ///
    /// ── 使用上の注意 ──────────────────────────────────────────────────────
    ///   • IsThreadSafe=false（Excel-DNA非同期API仕様）
    ///   • タイムアウト: 5分
    ///   • 計算途中の値: #N/A が表示される（完了時に結果に置換）
    ///   • 同一キーの重複実行: 先行タスク再利用
    /// </summary>
    [Experimental("DuckDBNET001")]
    [ExcelFunction(
        Description = "Executes a DuckDB SQL query asynchronously, avoiding Excel UI blocking.\n" +
                      "Useful for long-running queries and parallel execution across multiple sheets.\n" +
                      "Note: IsThreadSafe=false (Excel-DNA requirement). Timeout: 5 minutes.",
        Category = "xlDuckDb",
        HelpTopic = "https://duckdb.org/docs/index",
        IsMacroType = true,
        IsThreadSafe = false, // ExcelAsyncUtil requirement
        Name = "DuckDbQueryAsync")]
    public static object DuckDbQueryAsync(
        [ExcelArgument(
            Name = "SQL",
            Description = "DuckDB SQL statement.",
            AllowReference = true)]
        object query,
        [ExcelArgument(
            Name = "DatabaseFile",
            Description = "Path to .duckdb file (default: READ ONLY). Empty for in-memory DB.")]
        string dataSource)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNull;

        var queryString = ResolveQueryString(query);
        if (queryString is null) return ExcelError.ExcelErrorValue;

        var asyncHandle = AsyncDuckDbCache.ExecuteQueryAsync(queryString, dataSource);
        return asyncHandle;
    }

    /// <summary>
    /// 複数の DuckDB SQL クエリを並列非同期で実行し、結果を VSTACK で返す。
    ///
    /// ── 並列実行の効果 ────────────────────────────────────────────────────
    ///   • N個のクエリを理論上N分の1の時間で実行
    ///   • HTTP/HTTPS リモートデータの同時取得
    ///   • ネットワーク遅延を並列化で吸収
    ///
    /// ── 結果の集約 ────────────────────────────────────────────────────────
    ///   • 1番目のクエリのヘッダー行を採用
    ///   • 2番目以降のデータ行のみを追加（VSTACK）
    ///   • 列数不一致時は #VALUE! を返す
    /// </summary>
    [Experimental("DuckDBNET001")]
    [ExcelFunction(
        Description = "Executes multiple DuckDB SQL queries in parallel and returns VSTACK results.\n" +
                      "Useful for querying multiple data sources simultaneously.",
        Category = "xlDuckDb",
        HelpTopic = "https://duckdb.org/docs/index",
        IsMacroType = true,
        IsThreadSafe = false,
        Name = "DuckDbQueriesAsync")]
    public static object DuckDbQueriesAsync(
        [ExcelArgument(
            Name = "SQLArray",
            Description = "Array of SQL statements to execute in parallel.",
            AllowReference = true)]
        object sqlArray,
        [ExcelArgument(
            Name = "DatabaseFile",
            Description = "Path to .duckdb file (default: READ ONLY). Empty for in-memory DB.")]
        string dataSource)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNull;

        var sqlData = ResolveRangeData(sqlArray);
        if (sqlData is null) return ExcelError.ExcelErrorValue;

        var sqlList = ExtractSqlList(sqlData, isVertical: true);
        if (sqlList.Count == 0) return ExcelError.ExcelErrorNA;

        var queries = sqlList.Select(x => x.Sql).ToArray();
        var asyncHandle = AsyncDuckDbCache.ExecuteQueriesAsync(queries, dataSource);
        return asyncHandle;
    }

    // ════════════════════════════════════════════════════════════════════
    // DuckDbAsyncCacheClear — 非同期キャッシュ管理
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 非同期クエリのタスクキャッシュを完全にクリア（メモリ解放）。
    /// </summary>
    [Experimental("DuckDBNET001")]
    [ExcelFunction(
        Description = "Clears the async query task cache, freeing memory.",
        Category = "xlDuckDb",
        Name = "DuckDbAsyncCacheClear")]
    public static string DuckDbAsyncCacheClear()
    {
        AsyncDuckDbCache.ClearTaskCache();
        return "Async task cache cleared.";
    }

    /// <summary>
    /// 完了済みで古い非同期タスクを削除。
    /// </summary>
    [Experimental("DuckDBNET001")]
    [ExcelFunction(
        Description = "Removes completed async tasks older than specified minutes (default: 10).",
        Category = "xlDuckDb",
        Name = "DuckDbAsyncCacheCleanup")]
    public static string DuckDbAsyncCacheCleanup(
        [ExcelArgument(
            Name = "MinutesOld",
            Description = "Remove tasks older than N minutes (default: 10).")]
        double minutesOld)
    {
        var threshold = TimeSpan.FromMinutes(minutesOld > 0 ? minutesOld : 10);
        AsyncDuckDbCache.RemoveCompletedTasks(threshold);
        return $"Async cache cleanup: removed tasks older than {threshold.TotalMinutes} minutes.";
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
