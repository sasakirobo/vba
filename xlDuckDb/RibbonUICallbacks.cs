using ExcelDna.Integration.CustomUI;
using System.Diagnostics;
using System.Reflection;

namespace xlDuckDb;

/// <summary>
/// Ribbon UI コールバックハンドラー。
/// ExcelDNA Ribbon フレームワークと連携してRibbon UIイベントを処理します。
/// 
/// 登録方法:
///   • .dna ファイル内に &lt;CustomUI&gt; セクションで Ribbon XML を定義
///   • このクラスの public メソッドが自動的にコールバックとして認識される
///   • メソッド名は Ribbon XML の onAction, onChange, getLabel などと一致する必要あり
/// </summary>
public class RibbonUICallbacks : ExcelRibbon
{
    private static IRibbonUI _ribbonUI;

    /// <summary>
    /// Ribbon ロード時に呼び出される。
    /// </summary>
    public void OnRibbonLoad(IRibbonUI ribbonUI)
    {
        _ribbonUI = ribbonUI;
        Debug.WriteLine("[xlDuckDb] Ribbon UI loaded and initialized.");
    }

    // ──────────────────────────────────────────────────────────────────
    // Label コールバック（表示文字列の動的生成）
    // ──────────────────────────────────────────────────────────────────

    public string GetCacheMemoryLabel(IRibbonControl control)
    {
        return $"最大メモリ: {xlDuckDbConfig.ResultCacheMaxMemoryMb} MB";
    }

    public string GetMaxRowsLabel(IRibbonControl control)
    {
        return $"最大出力行: {xlDuckDbConfig.MaxOutputRows:N0}";
    }

    public string GetAsyncTimeoutLabel(IRibbonControl control)
    {
        var seconds = xlDuckDbConfig.AsyncTimeoutMs / 1000;
        return $"タイムアウト: {seconds} 秒";
    }

    // ──────────────────────────────────────────────────────────────────
    // キャッシュ設定コールバック
    // ──────────────────────────────────────────────────────────────────

    public void OnEnableResultCacheChanged(IRibbonControl control, bool pressed)
    {
        xlDuckDbConfig.EnableResultCache = pressed;
        Debug.WriteLine($"[xlDuckDb] Result cache: {(pressed ? "Enabled" : "Disabled")}");
        InvalidateRibbon();
    }

    public void OnResultCacheMaxMemoryChanged(IRibbonControl control, string text)
    {
        if (long.TryParse(text, out var mb) && mb >= 0)
        {
            xlDuckDbConfig.ResultCacheMaxMemoryMb = mb;
            Debug.WriteLine($"[xlDuckDb] Result cache max memory: {mb} MB");
            InvalidateRibbon();
        }
    }

    public void OnClearCache(IRibbonControl control)
    {
        try
        {
            DuckDbConnectionManager.Instance.ResultCache.Clear();
            AsyncDuckDbCache.ClearTaskCache();
            Debug.WriteLine("[xlDuckDb] Cache cleared.");
            InvalidateRibbon();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[xlDuckDb] Error clearing cache: {ex.Message}");
        }
    }

    // ──────────────────────────────────────────────────────────────────
    // 出力設定コールバック
    // ──────────────────────────────────────────────────────────────────

    public void OnMaxOutputRowsChanged(IRibbonControl control, string text)
    {
        if (int.TryParse(text, out var rows) && rows > 0 && rows <= 1048570)
        {
            xlDuckDbConfig.MaxOutputRows = rows;
            Debug.WriteLine($"[xlDuckDb] Max output rows: {rows:N0}");
            InvalidateRibbon();
        }
    }

    public void OnPresetMaxRows(IRibbonControl control)
    {
        if (int.TryParse(control.Tag, out var rows))
        {
            xlDuckDbConfig.MaxOutputRows = rows;
            Debug.WriteLine($"[xlDuckDb] Max output rows preset: {rows:N0}");
            InvalidateRibbon();
        }
    }

    // ──────────────────────────────────────────────────────────────────
    // タイムアウト設定コールバック
    // ──────────────────────────────────────────────────────────────────

    public void OnAsyncTimeoutChanged(IRibbonControl control, string text)
    {
        if (int.TryParse(text, out var seconds) && seconds >= 0)
        {
            xlDuckDbConfig.AsyncTimeoutMs = seconds * 1000;
            Debug.WriteLine($"[xlDuckDb] Async timeout: {seconds} seconds");
            InvalidateRibbon();
        }
    }

    public void OnPresetAsyncTimeout(IRibbonControl control)
    {
        if (int.TryParse(control.Tag, out var seconds))
        {
            xlDuckDbConfig.AsyncTimeoutMs = seconds * 1000;
            Debug.WriteLine($"[xlDuckDb] Async timeout preset: {seconds} seconds");
            InvalidateRibbon();
        }
    }

    // ──────────────────────────────────────────────────────────────────
    // 接続設定コールバック
    // ──────────────────────────────────────────────────────────────────

    public void OnEnableExtensionsChanged(IRibbonControl control, bool pressed)
    {
        xlDuckDbConfig.EnableBuiltInExtensions = pressed;
        Debug.WriteLine($"[xlDuckDb] Built-in extensions: {(pressed ? "Enabled" : "Disabled")}");
        InvalidateRibbon();
    }

    // ──────────────────────────────────────────────────────────────────
    // デバッグ &amp; 情報コールバック
    // ──────────────────────────────────────────────────────────────────

    public void OnEnableDetailedLoggingChanged(IRibbonControl control, bool pressed)
    {
        xlDuckDbConfig.EnableDetailedLogging = pressed;
        Debug.WriteLine($"[xlDuckDb] Detailed logging: {(pressed ? "Enabled" : "Disabled")}");
        InvalidateRibbon();
    }

    public void OnResetConfig(IRibbonControl control)
    {
        xlDuckDbConfig.ResetToDefaults();
        Debug.WriteLine("[xlDuckDb] Configuration reset to defaults.");
        InvalidateRibbon();
    }

    public void OnShowAbout(IRibbonControl control)
    {
        var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "Unknown";
        var about = $@"xlDuckDb - DuckDB Excel Add-in v{version}

【現在の設定値】
  • 最大出力行: {xlDuckDbConfig.MaxOutputRows:N0}
  • キャッシュ有効: {(xlDuckDbConfig.EnableResultCache ? "✓" : "✗")}
  • キャッシュメモリ: {xlDuckDbConfig.ResultCacheMaxMemoryMb} MB
  • タスク上限: {xlDuckDbConfig.AsyncTaskCacheCapacity} 件
  • タイムアウト: {(xlDuckDbConfig.AsyncTimeoutMs / 1000)} 秒
  • Extensions: {(xlDuckDbConfig.EnableBuiltInExtensions ? "✓" : "✗")}

【サポート】
  https://github.com/sasakirobo/xlDuck";

        Debug.WriteLine(about);
    }

    // ──────────────────────────────────────────────────────────────────
    // 内部ヘルパー
    // ──────────────────────────────────────────────────────────────────

    private void InvalidateRibbon()
    {
        try
        {
            _ribbonUI?.Invalidate();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[xlDuckDb] Error invalidating ribbon: {ex.Message}");
        }
    }
}
