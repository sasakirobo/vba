#pragma warning disable DuckDBNET001 // Experimental API usage

using System.Diagnostics.CodeAnalysis;
using DuckDB.NET.Data;

namespace xlDuckDb;

/// <summary>
/// DuckDB拡張機能（Extension）の初期化・管理モジュール。
/// 
/// Built-in Extensionの自動ロード：
///   • excel - ExcelファイルのDirect読み取り対応
///   • parquet - Parquetファイル形式対応
///   • https - HTTPS/TLS通信対応
///   • lance - Lance形式（ベクトル検索）対応
/// 
/// 用途：
///   1. DuckDB接続ごとのExtension自動初期化
///   2. Extension依存の機能テスト・検証
///   3. Runtime時のExtensionメタデータ取得
/// </summary>
public static class DuckDbExtensionInitializer
{
    /// <summary>
    /// BuiltIn拡張機能の名前リスト。
    /// これらは接続確立時に自動的にINSTALL/LOADされる。
    /// </summary>
    public static readonly string[] BuiltInExtensions = { "excel", "parquet", "https", "lance" };

    /// <summary>
    /// 指定のDuckDB接続にBuiltIn拡張機能をロードする。
    /// 
    /// このメソッドは以下の処理を実行：
    ///   1. 各拡張機能を INSTALL（DuckDBプロセスに追加）
    ///   2. 各拡張機能を LOAD（メモリにロード）
    ///   3. エラーハンドリング（失敗しても処理継続）
    /// 
    /// ⚠️ 注意: この処理は自動的にDuckDbConnectionManager内で実行されるため、
    ///   通常は明示的に呼び出す必要はありません。
    ///   デバッグやテスト目的でのみ使用してください。
    /// </summary>
    [Experimental("DuckDBNET001")]
    public static void InitializeExtensions(DuckDBConnection conn)
    {
        if (conn == null)
            throw new ArgumentNullException(nameof(conn));

        if (conn.State != System.Data.ConnectionState.Open)
            throw new InvalidOperationException("DuckDB connection must be open before initializing extensions.");

        using var cmd = conn.CreateCommand();

        foreach (var ext in BuiltInExtensions)
        {
            try
            {
                // INSTALL: 拡張機能をDuckDBプロセスにインストール
                cmd.CommandText = $"INSTALL {ext};";
                cmd.ExecuteNonQuery();

                // LOAD: メモリにロード
                cmd.CommandText = $"LOAD {ext};";
                cmd.ExecuteNonQuery();

                System.Diagnostics.Debug.WriteLine(
                    $"[xlDuckDb] Extension '{ext}' successfully loaded.");
            }
            catch (Exception ex)
            {
                // 拡張機能ロード失敗は警告のみ（処理継続）
                // これにより、特定の拡張機能が利用不可でも他の機能は動作する
                System.Diagnostics.Debug.WriteLine(
                    $"[xlDuckDb] Warning: Failed to load extension '{ext}': {ex.Message}");
            }
        }
    }

    /// <summary>
    /// DuckDB接続から利用可能なExtensionの一覧を取得する。
    /// </summary>
    [Experimental("DuckDBNET001")]
    public static List<ExtensionInfo> GetLoadedExtensions(DuckDBConnection conn)
    {
        if (conn == null)
            throw new ArgumentNullException(nameof(conn));

        if (conn.State != System.Data.ConnectionState.Open)
            throw new InvalidOperationException("DuckDB connection must be open.");

        var extensions = new List<ExtensionInfo>();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT extension_name, loaded, install_repository FROM duckdb_extensions() WHERE installed;";

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            extensions.Add(new ExtensionInfo
            {
                Name = reader.GetString(0),
                IsLoaded = reader.GetBoolean(1),
                Repository = reader.IsDBNull(2) ? null : reader.GetString(2)
            });
        }

        return extensions;
    }

    /// <summary>
    /// 指定のExtensionが利用可能かどうかをチェックする。
    /// </summary>
    [Experimental("DuckDBNET001")]
    public static bool IsExtensionAvailable(DuckDBConnection conn, string extensionName)
    {
        if (conn == null)
            throw new ArgumentNullException(nameof(conn));

        if (string.IsNullOrWhiteSpace(extensionName))
            throw new ArgumentException("Extension name cannot be null or empty.", nameof(extensionName));

        if (conn.State != System.Data.ConnectionState.Open)
            throw new InvalidOperationException("DuckDB connection must be open.");

        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT installed FROM duckdb_extensions() WHERE extension_name = '{extensionName.ToLowerInvariant()}' LIMIT 1;";
            var result = cmd.ExecuteScalar();
            return result != null && (bool)result;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// DuckDB Extensionのメタ情報。
    /// </summary>
    public class ExtensionInfo
    {
        /// <summary>拡張機能名（例: "parquet"、"excel"）</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>拡張機能がメモリにロードされているか</summary>
        public bool IsLoaded { get; set; }

        /// <summary>拡張機能の取得元（通常は "core" または URL）</summary>
        public string? Repository { get; set; }
    }
}

#pragma warning restore DuckDBNET001
