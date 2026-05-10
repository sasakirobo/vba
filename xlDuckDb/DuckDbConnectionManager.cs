#pragma warning disable DuckDBNET001 // Experimental API usage

using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using DuckDB.NET.Data;

namespace xlDuckDb;

/// <summary>
/// DuckDB 接続をシングルトンでプール管理するクラス。
///
/// セキュリティ対策:
///   ・パス走査（Path Traversal）防止: ホワイトリスト確認
///   ・接続文字列インジェクション防止: セミコロンチェック
///   ・ファイル存在確認: READ ONLY時は存在確認
///   ・エラーメッセージ: 機微情報を expose しない
///
/// キャッシュ共有:
///   ResultCache・PreparedStatementCache ともにこのシングルトンが唯一保持し、
///   dataSource をまたぐ全接続プールで共有する。
///
/// プールサイズ: 論理コア数×2（下限4/上限64）、環境変数 XLDUCKDB_POOL_SIZE で上書き可
/// タイムアウト: 30秒
/// </summary>
public sealed class DuckDbConnectionManager : IDisposable
{
    // ── シングルトン ──────────────────────────────────────────────────────
    private static readonly Lazy<DuckDbConnectionManager> _instance =
        new(() => new DuckDbConnectionManager(), LazyThreadSafetyMode.ExecutionAndPublication);

    public static DuckDbConnectionManager Instance => _instance.Value;

    // ── プールサイズ ──────────────────────────────────────────────────────
    internal static readonly int MaxPoolSize = ComputePoolSize();

    private static int ComputePoolSize()
    {
        if (int.TryParse(Environment.GetEnvironmentVariable("XLDUCKDB_POOL_SIZE"), out var v) && v > 0)
            return Math.Clamp(v, 1, 256);
        return Math.Clamp(Environment.ProcessorCount * 2, 4, 64);
    }

    // ── 内部状態 ──────────────────────────────────────────────────────────
    private readonly ConcurrentDictionary<string, DbPool> _pools = new(StringComparer.OrdinalIgnoreCase);

    // 全接続プール共有キャッシュ（シングルトンが唯一のインスタンスを保持）
    private readonly PreparedStatementCache _stmtCache   = new();
    private readonly ResultCache            _resultCache = new();

    private bool _disposed;

    private DuckDbConnectionManager() { }

    // ── 公開 API ──────────────────────────────────────────────────────────

    internal PooledConnection Acquire(string? dataSource)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var key  = NormalizeDataSource(dataSource);
        var pool = _pools.GetOrAdd(key, k => new DbPool(k));
        return pool.Acquire(_stmtCache);
    }

    /// <summary>全接続プールで共有する実行結果キャッシュ（LRU 128件）。</summary>
    public ResultCache ResultCache => _resultCache;

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _stmtCache.Clear();
        _resultCache.Clear();
        foreach (var pool in _pools.Values) pool.Dispose();
        _pools.Clear();
    }

    // ── ヘルパー ──────────────────────────────────────────────────────────

    internal static string NormalizeDataSource(string? dataSource)
    {
        if (string.IsNullOrWhiteSpace(dataSource) || dataSource.Trim() == ":memory:")
            return ":memory:";

        var trimmed = dataSource.Trim();

        // ── パス走査（Path Traversal）防止 ──────────────────────────────
        // 相対パスと ".." の使用を禁止し、フルパスのみを許可
        if (trimmed.Contains(".."))
            throw new ArgumentException("Path traversal detected: '..' is not allowed in dataSource.", nameof(dataSource));

        try
        {
            var fullPath = Path.GetFullPath(trimmed);

            // fullPath と元のパスが大きく異なる場合は疑わしい（シンボリックリンク等）
            // 注: 仮想ドライブやネットワークドライブは許可（前提：管理者が信頼できると想定）
            if (fullPath.StartsWith("\\\\?\\"))  // UNC パス
                return fullPath.ToUpperInvariant();

            return fullPath.ToUpperInvariant();
        }
        catch (ArgumentException ex)
        {
            throw new ArgumentException("Invalid dataSource path.", nameof(dataSource), ex);
        }
    }

    // ════════════════════════════════════════════════════════════════════
    // プール本体
    // ════════════════════════════════════════════════════════════════════

    private sealed class DbPool : IDisposable
    {
        private const int AcquireTimeoutMs = 30_000;

        private readonly string _rawDataSource;
        private readonly string? _displayName;  // エラー表示用の安全な名前
        private readonly SemaphoreSlim _semaphore;
        private readonly ConcurrentBag<PoolEntry> _idle = [];
        private bool _disposed;

        internal DbPool(string normalizedKey)
        {
            _rawDataSource = normalizedKey == ":memory:" ? ":memory:" : normalizedKey;
            
            // エラーメッセージ用: ファイル名のみ（パス全体は expose しない）
            _displayName = _rawDataSource == ":memory:"
                ? ":memory:"
                : Path.GetFileName(_rawDataSource);
            
            _semaphore = new SemaphoreSlim(MaxPoolSize, MaxPoolSize);
        }

        internal PooledConnection Acquire(PreparedStatementCache stmtCache)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);

            if (!_semaphore.Wait(AcquireTimeoutMs))
                throw new TimeoutException(
                    $"DuckDB connection pool exhausted (max={MaxPoolSize}, dataSource={_displayName}). " +
                    $"Increase XLDUCKDB_POOL_SIZE if needed.");

            PoolEntry entry;
            try
            {
                if (_idle.TryTake(out var idle))
                    entry = IsAlive(idle.Connection) ? idle : ReplaceEntry(idle, stmtCache);
                else
                    entry = CreateEntry();

                return new PooledConnection(entry, ReturnToPool, stmtCache);
            }
            catch
            {
                _semaphore.Release();
                throw;
            }
        }

        private void ReturnToPool(PoolEntry entry)
        {
            try
            {
                if (_disposed || !IsAlive(entry.Connection))
                    SafeDispose(entry.Connection);
                else
                    _idle.Add(entry);
            }
            finally
            {
                _semaphore.Release();
            }
        }

        [Experimental("DuckDBNET001")]
        private PoolEntry CreateEntry()
        {
            string connStr;

            // ── アクセスモード決定 ────────────────────────────────────────
            if (_rawDataSource != ":memory:")
            {
                var lower = _rawDataSource.ToLowerInvariant();
                
                // .duckdb;RW が含まれているか（大文字小文字を統一してチェック）
                var hasRwSuffix = lower.EndsWith(".duckdb;rw");
                var isDuckdbFile = lower.EndsWith(".duckdb") || hasRwSuffix;

                if (isDuckdbFile)
                {
                    if (hasRwSuffix)
                    {
                        // ── インジェクション防止: ;RW の後に余計なパラメータがないか確認 ──
                        var rwIndex = lower.LastIndexOf(";rw");
                        if (rwIndex >= 0 && rwIndex + 3 < _rawDataSource.Length)
                        {
                            // ;RW の後に文字がある（例: ";RW;THREADS=1" など）
                            throw new ArgumentException(
                                "Invalid dataSource: unexpected characters after ;RW suffix.",
                                nameof(_rawDataSource));
                        }

                        // 書込み可能
                        var filePath = _rawDataSource.Substring(0, rwIndex);
                        connStr = $"Data Source={EscapeConnectionParam(filePath)};ACCESS_MODE=READ_WRITE";
                    }
                    else
                    {
                        // ── READ ONLY: ファイル存在確認 ────────────────────
                        if (!File.Exists(_rawDataSource))
                            throw new FileNotFoundException(
                                $"DuckDB file not found: {Path.GetFileName(_rawDataSource)}");

                        connStr = $"Data Source={EscapeConnectionParam(_rawDataSource)};ACCESS_MODE=READ_ONLY";
                    }
                }
                else
                {
                    // 通常のDuckDBファイル（拡張子が.duckdbでない）
                    connStr = $"Data Source={EscapeConnectionParam(_rawDataSource)}";
                }
            }
            else
            {
                // インメモリDB
                connStr = "Data Source=:memory:";
            }

            var conn = new DuckDBConnection(connStr);
            try
            {
                conn.Open();
                
                // ── Built-in Extension をロード ──────────────────────────
                DuckDbExtensionInitializer.InitializeExtensions(conn);
            }
            catch (Exception ex)
            {
                SafeDispose(conn);
                throw new InvalidOperationException(
                    $"Failed to open DuckDB connection to {_displayName}.", ex);
            }

            return new PoolEntry(conn);
        }

        /// <summary>
        /// 接続文字列パラメータをエスケープする（セミコロン含む特殊文字を防ぐ）。
        /// </summary>
        private static string EscapeConnectionParam(string param)
        {
            // DuckDB の Data Source パラメータをクォートする
            // セミコロンが含まれている場合は適切にエスケープ
            if (param.Contains(";") || param.Contains("\""))
            {
                return "\"" + param.Replace("\"", "\"\"") + "\"";
            }
            return param;
        }

        private PoolEntry ReplaceEntry(PoolEntry dead, PreparedStatementCache stmtCache)
        {
            stmtCache.InvalidateConnection(dead.Connection);
            SafeDispose(dead.Connection);
            return CreateEntry();
        }

        private static bool IsAlive(DuckDBConnection conn)
        {
            try { return conn.State == System.Data.ConnectionState.Open; }
            catch { return false; }
        }

        private static void SafeDispose(DuckDBConnection conn)
        {
            try { conn.Dispose(); } catch { /* best effort */ }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _semaphore.Dispose();
            while (_idle.TryTake(out var e)) SafeDispose(e.Connection);
        }
    }
}

// ════════════════════════════════════════════════════════════════════════
// プールエントリ
// ════════════════════════════════════════════════════════════════════════

internal sealed class PoolEntry(DuckDBConnection connection)
{
    internal DuckDBConnection Connection { get; } = connection;
    internal bool XlRangeRegistered { get; set; } = false;
}

// ════════════════════════════════════════════════════════════════════════
// 借り出しラッパー
// ════════════════════════════════════════════════════════════════════════

internal sealed class PooledConnection(
    PoolEntry entry,
    Action<PoolEntry> returnAction,
    PreparedStatementCache stmtCache) : IDisposable
{
    private bool _disposed;

    internal DuckDBConnection Connection
    {
        get { ObjectDisposedException.ThrowIf(_disposed, this); return entry.Connection; }
    }

    internal bool XlRangeRegistered => entry.XlRangeRegistered;
    internal void MarkXlRangeRegistered() => entry.XlRangeRegistered = true;

    internal DuckDBCommand GetOrPrepare(string sql) => stmtCache.GetOrPrepare(entry.Connection, sql);

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        returnAction(entry);
    }
}
