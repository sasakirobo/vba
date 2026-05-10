namespace xlDuckDb;

/// <summary>
/// スレッドセーフな汎用 LRU キャッシュ。
///
/// 実装:
///   LinkedList + Dictionary の古典的な LRU。
///   先頭 = 最近使用、末尾 = 最も古い（次に追い出される）。
///   追い出し時に値が IDisposable であれば自動 Dispose する。
/// </summary>
internal sealed class LruCache<TKey, TValue>(int capacity) where TKey : notnull
{
    private readonly int _capacity = capacity > 0
        ? capacity
        : throw new ArgumentOutOfRangeException(nameof(capacity), "Capacity must be > 0.");

    private readonly LinkedList<(TKey Key, TValue Value)> _list = new();
    private readonly Dictionary<TKey, LinkedListNode<(TKey Key, TValue Value)>> _map = new();
    private readonly object _lock = new();

    internal int Count { get { lock (_lock) return _map.Count; } }

    /// <summary>
    /// キーに対応する値を取得する。ヒット時は先頭に移動（最近使用としてマーク）。
    /// </summary>
    internal bool TryGet(TKey key, out TValue value)
    {
        lock (_lock)
        {
            if (_map.TryGetValue(key, out var node))
            {
                _list.Remove(node);
                _list.AddFirst(node);
                value = node.Value.Value;
                return true;
            }
            value = default!;
            return false;
        }
    }

    /// <summary>
    /// キーと値を追加または更新する。
    /// キャパシティを超える場合は末尾（最古）を自動削除する。
    /// </summary>
    internal void Set(TKey key, TValue value)
    {
        lock (_lock)
        {
            if (_map.TryGetValue(key, out var existing))
            {
                _list.Remove(existing);
                _map.Remove(key);
            }
            else if (_map.Count >= _capacity)
            {
                var lru = _list.Last!;
                _map.Remove(lru.Value.Key);
                _list.RemoveLast();
                if (lru.Value.Value is IDisposable d) DisposeQuietly(d);
            }

            var node = new LinkedListNode<(TKey, TValue)>((key, value));
            _list.AddFirst(node);
            _map[key] = node;
        }
    }

    /// <summary>全エントリを削除する。IDisposable な値は自動 Dispose する。</summary>
    internal void Clear()
    {
        lock (_lock)
        {
            foreach (var (_, value) in _list)
                if (value is IDisposable d) DisposeQuietly(d);
            _list.Clear();
            _map.Clear();
        }
    }

    private static void DisposeQuietly(IDisposable d)
    {
        try { d.Dispose(); } catch { /* best effort */ }
    }
}
