using DuckDB.NET.Data;
using ExcelDna.Integration;
using System.IO.Compression;
using System.Runtime.CompilerServices;

namespace xlDuckDb;

// ════════════════════════════════════════════════════════════════════════
// 結果キャッシュキー
// ════════════════════════════════════════════════════════════════════════

public readonly struct ResultCacheKey : IEquatable<ResultCacheKey>
{
    public string DataSource { get; }
    public string Sql { get; }
    public object[]? Parameters { get; }
    private readonly int _hashCode;

    public ResultCacheKey(string dataSource, string sql, object[]? parameters)
    {
        DataSource = dataSource;
        Sql = sql;
        Parameters = parameters;
        _hashCode = ComputeHash(dataSource, sql, parameters);
    }

    public bool Equals(ResultCacheKey other)
    {
        if (!string.Equals(DataSource, other.DataSource, StringComparison.OrdinalIgnoreCase)) return false;
        if (!string.Equals(Sql, other.Sql, StringComparison.Ordinal)) return false;
        if (Parameters == null && other.Parameters == null) return true;
        if (Parameters == null || other.Parameters == null) return false;
        if (Parameters.Length != other.Parameters.Length) return false;
        for (var i = 0; i < Parameters.Length; i++)
            if (!Equals(Parameters[i], other.Parameters[i])) return false;
        return true;
    }
    public override bool Equals(object? obj) => obj is ResultCacheKey k && Equals(k);
    public override int GetHashCode() => _hashCode;

    private static int ComputeHash(string dataSource, string sql, object[]? parameters)
    {
        var h = new HashCode();
        h.Add(dataSource, StringComparer.OrdinalIgnoreCase);
        h.Add(sql, StringComparer.Ordinal);
        if (parameters != null)
            foreach (var p in parameters) h.Add(p);
        return h.ToHashCode();
    }
}

// ════════════════════════════════════════════════════════════════════════
// Brotli シリアライザ
// ════════════════════════════════════════════════════════════════════════

public static class ResultCacheSerializer
{
    public static byte[] Compress(object[,] data)
    {
        // ここに圧縮ロジックを実装
        var raw = SerializeToBytes(data);
        var maxLen = System.IO.Compression.BrotliEncoder.GetMaxCompressedLength(raw.Length);
        if (maxLen < 0 || maxLen > 500L * 1024 * 1024)
            throw new InvalidOperationException($"Compression size {maxLen} exceeds limit.");
        var buf = new byte[maxLen];
        if (!System.IO.Compression.BrotliEncoder.TryCompress(raw, buf, out int written, 4, window: 22))
            throw new InvalidOperationException("Brotli compression failed.");
        return buf[..written];
    }
    public static object[,] Decompress(byte[] compressed)
    {
        // ここに展開ロジックを実装
        using var src = new MemoryStream(compressed);
        using var dst = new MemoryStream((int)Math.Min(compressed.Length * 10L, 500L * 1024 * 1024));
        try
        {
            using (var brotli = new System.IO.Compression.BrotliStream(src, CompressionMode.Decompress, leaveOpen: true))
            {
                var buffer = new byte[4096];
                int read;
                while ((read = brotli.Read(buffer, 0, buffer.Length)) > 0)
                {
                    if (dst.Length + read > 500L * 1024 * 1024)
                        throw new InvalidOperationException($"Decompressed data exceeds 500MB limit.");
                    dst.Write(buffer, 0, read);
                }
            }
        }
        catch (Exception ex) when (!(ex is InvalidOperationException))
        {
            throw new InvalidOperationException("Brotli decompression failed.", ex);
        }
        return DeserializeFromBytes(dst.ToArray());
    }

    private static byte[] SerializeToBytes(object[,] data)
    {
        var rows = data.GetLength(0);
        var cols = data.GetLength(1);

        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms, System.Text.Encoding.UTF8, leaveOpen: true);

        writer.Write(rows);
        writer.Write(cols);

        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                WriteCell(writer, data[r, c]);
            }
        }

        writer.Flush();
        return ms.ToArray();
    }

    private static void WriteCell(BinaryWriter w, object? value)
    {
        switch (value)
        {
            case null:
            case DBNull:
                w.Write((byte)0x00);
                break;
            case double d:
                w.Write((byte)0x01);
                w.Write(d);
                break;
            case string s:
                w.Write((byte)0x02);
                var bytes = System.Text.Encoding.UTF8.GetBytes(s);
                w.Write(bytes.Length);
                w.Write(bytes);
                break;
            case bool b:
                w.Write((byte)0x03);
                w.Write((byte)(b ? 1 : 0));
                break;
            case DateTime dt:
                w.Write((byte)0x04);
                w.Write(dt.Ticks);
                w.Write((byte)dt.Kind);
                break;
            case ExcelError err:
                w.Write((byte)0x05);
                w.Write((int)err);
                break;
            default:
                w.Write((byte)0x02);
                var fb = System.Text.Encoding.UTF8.GetBytes(value.ToString() ?? string.Empty);
                w.Write(fb.Length);
                w.Write(fb);
                break;
        }
    }

    private static object[,] DeserializeFromBytes(byte[] raw)
    {
        using var ms = new MemoryStream(raw);
        using var reader = new BinaryReader(ms, System.Text.Encoding.UTF8, leaveOpen: true);

        var rows = reader.ReadInt32();
        var cols = reader.ReadInt32();

        if (rows < 0 || cols < 0)
            throw new InvalidDataException($"Invalid array dimensions: rows={rows}, cols={cols}");
        if (rows > 1_000_000 || cols > 1_000_000)
            throw new InvalidDataException($"Array dimensions exceed limit: {rows}x{cols} (max: 1000000x1000000)");
        if (rows > 0 && cols > long.MaxValue / rows)
            throw new InvalidDataException("Array size would overflow.");

        var data = new object[rows, cols];

        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                data[r, c] = ReadCell(reader);
            }
        }
        return data;
    }

    private static object ReadCell(BinaryReader r)
    {
        var tag = r.ReadByte();
        return tag switch
        {
            0x00 => ExcelError.ExcelErrorNA,
            0x01 => r.ReadDouble(),
            0x02 => ReadStringValue(r),
            0x03 => r.ReadByte() != 0,
            0x04 => new DateTime(r.ReadInt64(), (DateTimeKind)r.ReadByte()),
            0x05 => (ExcelError)r.ReadInt32(),
            _    => throw new InvalidDataException($"不明な TypeTag: 0x{tag:X2}")
        };
    }

    private static string ReadStringValue(BinaryReader r)
    {
        var len = r.ReadInt32();
        if (len < 0 || len > 10 * 1024 * 1024)
            throw new InvalidDataException($"Invalid string length: {len}");
        try
        {
            return System.Text.Encoding.UTF8.GetString(r.ReadBytes(len));
        }
        catch (ArgumentException ex)
        {
            throw new InvalidDataException("Failed to read string data.", ex);
        }
    }
}

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
/// クエリ実行結果（object[,]）を Brotli 圧縮してメモリ上にキャッシュする。
///
/// ── キャッシュ上限の考え方 ──────────────────────────────────────────
///   固定件数上限を廃止し、「圧縮済み合計バイト数」でメモリ圧迫を制御する。
///
///   許容バイト上限 = min(物理メモリ × MaxMemoryFraction, AbsoluteMaxBytes)
///     MaxMemoryFraction : 0.10（デフォルト。物理メモリの 10 %）
///     AbsoluteMaxBytes  : 512 MB（上記割合が巨大マシンで膨らみすぎないよう絶対上限）
///
///   Set() のたびに合計圧縮バイト数を確認し、上限を超えたら LRU 末尾から
///   合計が上限の 80 % を下回るまで追い出す（ヒステリシス付き）。
///
///   さらに Set() の 100 回に 1 回、GCMemoryInfo.MemoryLoadBytes を確認し、
///   物理メモリ使用率が HighPressureThreshold（既定 0.90）を超えている場合は
///   キャッシュを半減させる緊急縮退を行う。
///
/// ── 圧縮 ──────────────────────────────────────────────────────────────
///   object[,] → 独自バイナリ形式 → Brotli 圧縮（Quality=4, Window=22）
///   キャッシュ登録時に圧縮、ヒット時に自動展開して object[,] を返す。
///
///   シリアライズ対応型（DuckDbHelper の出力型に準拠）:
///     double / string / bool / DateTime / ExcelError / DBNull(→ExcelErrorNA)
///
/// ── xlRange クエリ ──────────────────────────────────────────────────
///   セルの値に依存するためキャッシュ対象外（呼び出し側で制御）。
/// </summary>
public sealed class ResultCache : IDisposable
{
    // ── 設定定数 ──────────────────────────────────────────────────────
    /// <summary>物理メモリに対してキャッシュが占める最大割合（デフォルト 10 %）。</summary>
    private const double MaxMemoryFraction = 0.10;

    /// <summary>割合計算の結果に関わらず超えない絶対上限（512 MB）。</summary>
    private const long AbsoluteMaxBytes = 512L * 1024 * 1024;

    /// <summary>追い出し後の目標使用率（ヒステリシス: 上限の 80 %）。</summary>
    private const double EvictTargetFraction = 0.80;

    /// <summary>GC メモリ負荷チェックを行う Set() 呼び出し間隔。</summary>
    private const int GcCheckInterval = 100;

    /// <summary>
    /// 物理メモリ使用率がこの値を超えたら緊急縮退（キャッシュを半減）する。
    /// GCMemoryInfo.HighMemoryLoadThresholdBytes との比較で使う割合。
    /// </summary>
    private const double EmergencyLoadFraction = 0.90;

    // ── 内部状態 ──────────────────────────────────────────────────────
    private readonly object _lock = new();

    /// <summary>LRU 順序管理: 先頭 = 最近使用、末尾 = 最古。</summary>
    private readonly LinkedList<CacheEntry> _list = new();

    /// <summary>高速アクセス用辞書。</summary>
    private readonly Dictionary<ResultCacheKey, LinkedListNode<CacheEntry>> _map = new();

    /// <summary>全エントリの圧縮済みバイト数合計。</summary>
    private long _totalCompressedBytes;

    /// <summary>Set() 呼び出しカウンタ（GC チェック用）。</summary>
    private int _setCallCount;

    private bool _disposed;

    // ── 動的上限計算 ──────────────────────────────────────────────────

    /// <summary>
    /// 現在の物理メモリ量から許容キャッシュバイト上限を算出する。
    /// 算出が失敗した場合は絶対上限（512 MB）を返す。
    /// </summary>
    private static long ComputeMaxBytes()
    {
        try
        {
            var info = GC.GetGCMemoryInfo();
            var fraction = (long)(info.TotalAvailableMemoryBytes * MaxMemoryFraction);
            return Math.Min(fraction, AbsoluteMaxBytes);
        }
        catch
        {
            return AbsoluteMaxBytes;
        }
    }

    // ── 公開 API ──────────────────────────────────────────────────────

    /// <summary>現在のキャッシュエントリ数。</summary>
    public int Count { get { lock (_lock) return _map.Count; } }

    /// <summary>現在の圧縮済み合計バイト数。</summary>
    public long TotalCompressedBytes { get { lock (_lock) return _totalCompressedBytes; } }

    /// <summary>
    /// キャッシュを検索する。ヒット時はデータを Brotli 展開して返す。
    /// </summary>
    public bool TryGet(string dataSource, string sql, object[]? parameters, out object[,] result)
    {
        var key = new ResultCacheKey(dataSource, sql, parameters);
        lock (_lock)
        {
            if (!_map.TryGetValue(key, out var node))
            {
                result = null!;
                return false;
            }
            // LRU: 先頭へ移動
            _list.Remove(node);
            _list.AddFirst(node);

            result = ResultCacheSerializer.Decompress(node.Value.CompressedData);
            return true;
        }
    }

    /// <summary>
    /// クエリ結果を Brotli 圧縮してキャッシュに登録する。
    /// 登録後にメモリ上限チェックと必要に応じた LRU 追い出しを行う。
    /// </summary>
    internal void Set(string dataSource, string sql, object[]? parameters, object[,] result)
    {
        var key            = new ResultCacheKey(dataSource, sql, parameters);
        var compressed     = ResultCacheSerializer.Compress(result);
        var compressedSize = compressed.Length;

        lock (_lock)
        {
            if (_disposed) return;

            // 既存エントリの更新
            if (_map.TryGetValue(key, out var existing))
            {
                _totalCompressedBytes -= existing.Value.CompressedData.Length;
                _list.Remove(existing);
                _map.Remove(key);
            }

            // 新エントリを先頭に追加
            var entry = new CacheEntry(key, compressed);
            var node  = new LinkedListNode<CacheEntry>(entry);
            _list.AddFirst(node);
            _map[key]              = node;
            _totalCompressedBytes += compressedSize;

            // GC メモリ負荷チェック（一定間隔）
            _setCallCount++;
            if (_setCallCount % GcCheckInterval == 0)
                ApplyEmergencyEvictionIfNeeded();

            // 通常の上限チェックと LRU 追い出し
            EvictIfOverLimit();
        }
    }

    /// <summary>全エントリを削除する。</summary>
    public void Clear()
    {
        lock (_lock)
        {
            _list.Clear();
            _map.Clear();
            _totalCompressedBytes = 0;
        }
    }

    public void Dispose()
    {
        lock (_lock)
        {
            if (_disposed) return;
            _disposed = true;
            Clear();
        }
    }

    // ── 追い出しロジック ──────────────────────────────────────────────

    /// <summary>
    /// 合計圧縮バイト数が上限を超えていれば、目標値まで LRU 末尾から追い出す。
    /// ロック内から呼ぶこと。
    /// </summary>
    private void EvictIfOverLimit()
    {
        var maxBytes    = ComputeMaxBytes();
        var targetBytes = (long)(maxBytes * EvictTargetFraction);

        while (_totalCompressedBytes > maxBytes && _list.Count > 0)
        {
            RemoveLruTail();
            // ヒステリシス: 目標まで追い出す
            if (_totalCompressedBytes <= targetBytes) break;
        }
    }

    /// <summary>
    /// GC のメモリ負荷情報を参照し、物理メモリが逼迫している場合はキャッシュを半減する。
    /// ロック内から呼ぶこと。
    /// </summary>
    private void ApplyEmergencyEvictionIfNeeded()
    {
        try
        {
            var info = GC.GetGCMemoryInfo();
            if (info.HighMemoryLoadThresholdBytes <= 0) return;

            var loadRatio = (double)info.MemoryLoadBytes / info.HighMemoryLoadThresholdBytes;
            if (loadRatio < EmergencyLoadFraction) return;

            // キャッシュを半減
            var target = _list.Count / 2;
            while (_list.Count > target)
                RemoveLruTail();
        }
        catch
        {
            // GC API が失敗しても処理を継続する
        }
    }

    /// <summary>
    /// LRU 末尾（最古エントリ）を 1 件削除する。
    /// ロック内から呼ぶこと。
    /// </summary>
    private void RemoveLruTail()
    {
        var tail = _list.Last;
        if (tail is null) return;
        _totalCompressedBytes -= tail.Value.CompressedData.Length;
        _map.Remove(tail.Value.Key);
        _list.RemoveLast();
    }

    // ── 内部レコード ──────────────────────────────────────────────────

    private sealed record CacheEntry(ResultCacheKey Key, byte[] CompressedData);

    /// <summary>
    /// キャッシュ内のすべてのエントリを列挙する。
    /// </summary>
    public IEnumerable<(ResultCacheKey Key, object[,] Value)> GetAllEntries()
    {
        lock (_lock)
        {
            foreach (var node in _map.Values)
            {
                var key = node.Value.Key;
                var value = ResultCacheSerializer.Decompress(node.Value.CompressedData);
                yield return (key, value);
            }
        }
    }
}