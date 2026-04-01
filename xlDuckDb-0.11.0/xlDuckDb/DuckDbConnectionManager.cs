#pragma warning disable DuckDBNET001 // Experimental API usage

using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using DuckDB.NET.Data;

namespace xlDuckDb;

/// <summary>
/// DuckDB 接続をシングルトンでプール管理するクラス。
///
/// キャッシュ共有:
///   ResultCache・PreparedStatementCache ともにこのシングルトンが唯一保持し、
///   dataSource をまたぐ全接続プールで共有する。
///   ResultCache のキーは dataSource を含むため、DB をまたいだ結果の混在はない。
///
/// プールサイズ: 論理コア数×2（下限4/上限64）、環境変数 XLDUCKDB_POOL_SIZE で上書き可
/// タイムアウト: 30秒
/// </summary>
internal sealed class DuckDbConnectionManager : IDisposable
{
    // ── シングルトン ──────────────────────────────────────────────────────
    private static readonly Lazy<DuckDbConnectionManager> _instance =
        new(() => new DuckDbConnectionManager(), LazyThreadSafetyMode.ExecutionAndPublication);

    internal static DuckDbConnectionManager Instance => _instance.Value;

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
    internal ResultCache ResultCache => _resultCache;

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
        return Path.GetFullPath(dataSource.Trim()).ToUpperInvariant();
    }

    // ════════════════════════════════════════════════════════════════════
    // プール本体
    // ════════════════════════════════════════════════════════════════════

    private sealed class DbPool : IDisposable
    {
        private const int AcquireTimeoutMs = 30_000;

        private readonly string _rawDataSource;
        private readonly SemaphoreSlim _semaphore;
        private readonly ConcurrentBag<PoolEntry> _idle = [];
        private bool _disposed;

        internal DbPool(string normalizedKey)
        {
            _rawDataSource = normalizedKey == ":memory:" ? ":memory:" : normalizedKey;
            _semaphore     = new SemaphoreSlim(MaxPoolSize, MaxPoolSize);
        }

        internal PooledConnection Acquire(PreparedStatementCache stmtCache)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);

            if (!_semaphore.Wait(AcquireTimeoutMs))
                throw new TimeoutException(
                    $"DuckDB pool exhausted (max={MaxPoolSize}, dataSource={_rawDataSource}). " +
                    $"Set XLDUCKDB_POOL_SIZE to increase.");

            PoolEntry entry;
            if (_idle.TryTake(out var idle))
                entry = IsAlive(idle.Connection) ? idle : ReplaceEntry(idle, stmtCache);
            else
                entry = CreateEntry();

            return new PooledConnection(entry, ReturnToPool, stmtCache);
        }

        private void ReturnToPool(PoolEntry entry)
        {
            if (_disposed || !IsAlive(entry.Connection))
                SafeDispose(entry.Connection);
            else
                _idle.Add(entry);
            _semaphore.Release();
        }

        [Experimental("DuckDBNET001")]
        private PoolEntry CreateEntry()
        {
            // データソースが. duckdbファイルの場合、デフォルトでREAD ONLYでATTACH
            // 書込み可能で開きたい場合は末尾に ";RW" を付加
            string connStr;
            if (_rawDataSource != ":memory:" && _rawDataSource.EndsWith(".DUCKDB", StringComparison.OrdinalIgnoreCase))
            {
                if (_rawDataSource.EndsWith(".DUCKDB;RW", StringComparison.OrdinalIgnoreCase))
                {
                    // 書込み可能
                    var filePath = _rawDataSource.Substring(0, _rawDataSource.Length - 3); // ;RW除去
                    connStr = $"Data Source={filePath};ACCESS_MODE=READ_WRITE";
                }
                else
                {
                    // 既定はREAD ONLY
                    connStr = $"Data Source={_rawDataSource};ACCESS_MODE=READ_ONLY";
                }
            }
            else
            {
                connStr = $"Data Source={_rawDataSource}";
            }
            var conn = new DuckDBConnection(connStr);
            conn.Open();
            return new PoolEntry(conn);
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
