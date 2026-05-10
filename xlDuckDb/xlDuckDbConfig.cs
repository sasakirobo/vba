namespace xlDuckDb;

/// <summary>
/// xlDuckDb 全体設定の管理クラス。
/// Ribbon UI から設定可能な項目を提供。
/// </summary>
public static class xlDuckDbConfig
{
    // ──────────────────────────────────────────────────────────────────
    // 出力設定
    // ──────────────────────────────────────────────────────────────────

    /// <summary>クエリ出力の最大行数（デフォルト: 1,048,570 = Excel最大行数）。</summary>
    public static int MaxOutputRows { get; set; } = 1048570;

    // ──────────────────────────────────────────────────────────────────
    // キャッシュ設定
    // ──────────────────────────────────────────────────────────────────

    /// <summary>結果キャッシュ有効化フラグ。</summary>
    public static bool EnableResultCache { get; set; } = true;

    /// <summary>非同期タスクキャッシュの最大エントリ数（LRU）。</summary>
    public static int AsyncTaskCacheCapacity { get; set; } = 256;

    /// <summary>結果キャッシュの最大メモリ（MB）。0=無制限。</summary>
    public static long ResultCacheMaxMemoryMb { get; set; } = 512;

    // ──────────────────────────────────────────────────────────────────
    // タイムアウト設定
    // ──────────────────────────────────────────────────────────────────

    /// <summary>非同期クエリのタイムアウト（ミリ秒）。デフォルト: 300,000ms (5分)。</summary>
    public static int AsyncTimeoutMs { get; set; } = 300_000;

    // ──────────────────────────────────────────────────────────────────
    // 接続設定
    // ──────────────────────────────────────────────────────────────────

    /// <summary>DuckDB 接続プール最大サイズ。0=自動（CPUコア数×2）。</summary>
    public static int MaxPoolSize { get; set; } = 0;

    // ──────────────────────────────────────────────────────────────────
    // Extension設定
    // ──────────────────────────────────────────────────────────────────

    /// <summary>Built-in Extension の自動ロード有効化（Excel, Parquet, HTTPS, Lance）。</summary>
    public static bool EnableBuiltInExtensions { get; set; } = true;

    // ──────────────────────────────────────────────────────────────────
    // デバッグ設定
    // ──────────────────────────────────────────────────────────────────

    /// <summary>詳細ログ出力有効化。</summary>
    public static bool EnableDetailedLogging { get; set; } = false;

    /// <summary>キャッシュ統計情報の表示有効化。</summary>
    public static bool ShowCacheStatistics { get; set; } = false;

    // ──────────────────────────────────────────────────────────────────
    // リセット
    // ──────────────────────────────────────────────────────────────────

    /// <summary>すべての設定をデフォルト値にリセット。</summary>
    public static void ResetToDefaults()
    {
        MaxOutputRows = 1048570;
        EnableResultCache = true;
        AsyncTaskCacheCapacity = 256;
        ResultCacheMaxMemoryMb = 512;
        AsyncTimeoutMs = 300_000;
        MaxPoolSize = 0;
        EnableBuiltInExtensions = true;
        EnableDetailedLogging = false;
        ShowCacheStatistics = false;
    }
}
