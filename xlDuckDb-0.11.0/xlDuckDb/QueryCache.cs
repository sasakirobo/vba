using DuckDB.NET.Data;

namespace xlDuckDb;

// ════════════════════════════════════════════════════════════════════════
// Prepared Statement キャッシュ
// ════════════════════════════════════════════════════════════════════════

/// <summary>
/// Prepared Statement（DuckDBCommand）を接続ごとに LRU キャッシュする。
///
/// キー: (接続インスタンスの識別子, SQL文字列)
///   接続をキーに含めることで、接続が再作成されたときに
///   古い Statement が別接続に誤って使われないようにする。
///
/// キャッシュ上限: 1024件（固定）
///   追い出し時は LruCache が DuckDBCommand を自動 Dispose する。
/// </summary>
internal sealed class PreparedStatementCache
{
    private const int Capacity = 1024;

    // キー: (接続の InstanceId, SQL文字列)
    private readonly LruCache<(int ConnId, string Sql), CachedCommand> _cache = new(Capacity);

    /// <summary>
    /// キャッシュから Prepared Statement を取得する。
    /// なければ新規に Prepare して登録する。
    /// </summary>
    internal DuckDBCommand GetOrPrepare(DuckDBConnection connection, string sql)
    {
        var key = (RuntimeHelpers.GetHashCode(connection), sql);

        if (_cache.TryGet(key, out var cached))
            return cached.Command;

        var cmd = connection.CreateCommand();
        cmd.CommandText = sql;
        cmd.Prepare();

        _cache.Set(key, new CachedCommand(cmd));
        return cmd;
    }

    /// <summary>
    /// 指定接続に紐付く Statement を無効化する（接続再作成時に呼ぶ）。
    /// LruCache は全件スキャン不可のため全クリアで対応。
    /// </summary>
    internal void InvalidateConnection(DuckDBConnection _) => _cache.Clear();

    internal void Clear() => _cache.Clear();
}

/// <summary>
/// キャッシュ内の Prepared Statement ラッパー。
/// LruCache の追い出し時に IDisposable として Dispose される。
/// </summary>
internal sealed class CachedCommand(DuckDBCommand command) : IDisposable
{
    internal DuckDBCommand Command { get; } = command;
    public void Dispose() { try { Command.Dispose(); } catch { /* best effort */ } }
}

// ════════════════════════════════════════════════════════════════════════
// 実行結果キャッシュ
// ════════════════════════════════════════════════════════════════════════

/// <summary>
/// クエリ実行結果（object[,]）を LRU でキャッシュする。
///
/// キー: (dataSource, SQL, パラメータ値の配列)
///   パラメータ値が変わると別エントリになる。
///
/// キャッシュ上限: 10000件（固定）
///
/// xlRange を使うクエリはセルの値に依存するためキャッシュ対象外。
/// </summary>
internal sealed class ResultCache
{
    private const int Capacity = 10000;

    private readonly LruCache<ResultCacheKey, object[,]> _cache = new(Capacity);

    internal bool TryGet(string dataSource, string sql, object[]? parameters, out object[,] result)
        => _cache.TryGet(new ResultCacheKey(dataSource, sql, parameters), out result!);

    internal void Set(string dataSource, string sql, object[]? parameters, object[,] result)
        => _cache.Set(new ResultCacheKey(dataSource, sql, parameters), result);

    internal void Clear() => _cache.Clear();
}

/// <summary>
/// 結果キャッシュのキー。
/// dataSource（大文字小文字無視） + SQL（完全一致） + パラメータ値 で同一性を判定。
/// </summary>
internal readonly struct ResultCacheKey(string dataSource, string sql, object[]? parameters)
    : IEquatable<ResultCacheKey>
{
    private readonly string _dataSource = dataSource;
    private readonly string _sql = sql;
    private readonly object[]? _parameters = parameters;
    private readonly int _hashCode = ComputeHash(dataSource, sql, parameters);

    private static int ComputeHash(string dataSource, string sql, object[]? parameters)
    {
        var h = new HashCode();
        h.Add(dataSource, StringComparer.OrdinalIgnoreCase);
        h.Add(sql, StringComparer.Ordinal);
        if (parameters != null)
            foreach (var p in parameters) h.Add(p);
        return h.ToHashCode();
    }

    public bool Equals(ResultCacheKey other)
    {
        if (!string.Equals(_dataSource, other._dataSource, StringComparison.OrdinalIgnoreCase)) return false;
        if (!string.Equals(_sql, other._sql, StringComparison.Ordinal)) return false;
        if (_parameters == null && other._parameters == null) return true;
        if (_parameters == null || other._parameters == null) return false;
        if (_parameters.Length != other._parameters.Length) return false;
        for (var i = 0; i < _parameters.Length; i++)
            if (!Equals(_parameters[i], other._parameters[i])) return false;
        return true;
    }

    public override bool Equals(object? obj) => obj is ResultCacheKey k && Equals(k);
    public override int GetHashCode() => _hashCode;
}

// ════════════════════════════════════════════════════════════════════════
// using static のための using
// ════════════════════════════════════════════════════════════════════════
file static class RuntimeHelpers
{
    internal static int GetHashCode(object obj)
        => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
}
