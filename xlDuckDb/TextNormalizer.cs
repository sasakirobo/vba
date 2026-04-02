using System;
using System.Buffers;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;

namespace xlDuckDb;

// ════════════════════════════════════════════════════════════════════════
// 正規化オプション（ビットフラグ）
// ════════════════════════════════════════════════════════════════════════

/// <summary>
/// テキスト正規化オプション。複数を OR で組み合わせて使用する。
/// </summary>
[Flags]
public enum NormalizeOptions : int
{
    None = 0,
    NfkcCaseFold = 1 << 0,   //   1  .NET 標準: FormKC + 簡易 CaseFold
    LongVowelNormalize = 1 << 1,   //   2  長音「ー」正規化
    EmojiRemove = 1 << 2,   //   4  絵文字削除（簡易）
    FullwidthAlphanumToHalf = 1 << 3,   //   8  全角英数字 → 半角（+必要なら大文字）
    HalfwidthAlphanumToUpper = 1 << 4,   //  16  半角英数字 → 大文字
    HalfKanaToFullKana = 1 << 5,   //  32  半角カナ → 全角カナ（自前変換）
    FullwidthSymbolToHalf = 1 << 6,   //  64  全角記号 → 半角記号
    WhitespaceNormalize = 1 << 7,   // 128  空白正規化・Trim
    BracketNormalize = 1 << 8,   // 256  括弧類 → ( )、( 前空白削除

    All = NfkcCaseFold | LongVowelNormalize | EmojiRemove
        | FullwidthAlphanumToHalf | HalfwidthAlphanumToUpper
        | HalfKanaToFullKana | FullwidthSymbolToHalf
        | WhitespaceNormalize | BracketNormalize,   // 511
}

// ════════════════════════════════════════════════════════════════════════
// テキスト正規化エンジン（.NET 8 標準ライブラリのみ）
// ════════════════════════════════════════════════════════════════════════

/// <summary>
/// .NET 8 標準ライブラリのみを使ったテキスト正規化クラス。
///
/// ── 方針（ICU4N 不使用）────────────────────────────────────────────
///   ・NfkcCaseFold : string.Normalize(FormKC) + ToLowerInvariant()
///                   + 最低限の追加（ß/ẞ → ss）
///                   ※ Unicode CaseFold の完全互換ではありません
///   ・HalfKanaToFullKana : 半角カナ区間（U+FF61〜U+FF9F）のみ自前変換
///                          英数字・記号は全角化しない
///   ・EmojiRemove : 既定の簡易範囲判定（必要なら強化可能）
///   ・他の処理は元ロジックを維持（1パス、ArrayPool、LRU キャッシュ）
/// </summary>
public static class TextNormalizer
{
    // ── 正規化結果キャッシュ（LRU）───────────────────────────────────────
    // キー: (入力文字列, options, additionalRemove, additionalReplace)
    // デフォルト上限: 1024件（環境変数 XLDUCKDB_NORM_CACHE_SIZE で上書き可）
    public static readonly int NormCacheCapacity = ResolveNormCacheCapacity();

    private static int ResolveNormCacheCapacity()
    {
        if (int.TryParse(Environment.GetEnvironmentVariable("XLDUCKDB_NORM_CACHE_SIZE"), out var v) && v > 0)
            return Math.Clamp(v, 64, 65536);
        return 1024;
    }

    private static readonly LruCache<NormCacheKey, string> NormCache = new(NormCacheCapacity);

    // ── 置換辞書キャッシュ（mappingSpec 文字列 → 解析済み辞書）──────────
    private static readonly LruCache<string, List<(string From, string To)>> ReplaceSpecCache = new(64);

    // ════════════════════════════════════════════════════════════════════
    // 公開 API
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// テキストに正規化オプションを適用して返す。
    /// 同じ (input, options, remove, replace) の組み合わせは LRU キャッシュから返す。
    /// </summary>
    public static string Normalize(
        string input,
        NormalizeOptions options,
        string? additionalRemove = null,
        string? additionalReplace = null)
    {
        if (string.IsNullOrEmpty(input)) return input;
        if (options == NormalizeOptions.None
            && string.IsNullOrEmpty(additionalRemove)
            && string.IsNullOrEmpty(additionalReplace))
            return input;

        var cacheKey = new NormCacheKey(input, options, additionalRemove, additionalReplace);
        if (NormCache.TryGet(cacheKey, out var cached))
            return cached;

        var result = NormalizeCore(input, options, additionalRemove, additionalReplace);
        NormCache.Set(cacheKey, result);
        return result;
    }

    /// <summary>
    /// 正規化の各ステップ後の中間値を順番に返す。
    /// DuckDbNormalize 関数の "ステップ別表示" モード用。
    ///
    /// 戻り値: (stepLabel, value) のリスト
    ///   - 最初の行は常に ("Input", input)
    ///   - オプションが有効かつ適用前後で値が変化したステップのみ行を追加
    ///   - 最後の行は常に ("Final", 最終値)（最終値が Input と同じでも出力する）
    /// </summary>
    public static List<(string Label, string Value)> NormalizeSteps(
        string input,
        NormalizeOptions options,
        string? additionalRemove = null,
        string? additionalReplace = null)
    {
        var steps = new List<(string, string)> { ("Input", input) };
        if (string.IsNullOrEmpty(input)) return steps;

        var s = input;

        void AddStep(string label, string next)
        {
            if (next != s) steps.Add((label, next));
            s = next;
        }

        if (options.HasFlag(NormalizeOptions.NfkcCaseFold))
            AddStep("After NFKC+CaseFold", NfkcCaseFoldPass(s));

        if (options.HasFlag(NormalizeOptions.HalfKanaToFullKana))
            AddStep("After HalfKana→Full", HalfKanaToFullKanaPass(s));

        if (options.HasFlag(NormalizeOptions.EmojiRemove))
            AddStep("After EmojiRemove", CorePass(s, NormalizeOptions.EmojiRemove));

        if (options.HasFlag(NormalizeOptions.FullwidthSymbolToHalf))
            AddStep("After FullSymbol→Half", CorePass(s, NormalizeOptions.FullwidthSymbolToHalf));

        if (options.HasFlag(NormalizeOptions.FullwidthAlphanumToHalf))
        {
            AddStep("After FullwidthAlnum→Half", CorePass(s,
                NormalizeOptions.FullwidthAlphanumToHalf
                | (options & NormalizeOptions.HalfwidthAlphanumToUpper)));
        }
        else if (options.HasFlag(NormalizeOptions.HalfwidthAlphanumToUpper))
        {
            AddStep("After HalfAlnum→Upper", CorePass(s, NormalizeOptions.HalfwidthAlphanumToUpper));
        }

        if (options.HasFlag(NormalizeOptions.BracketNormalize))
        {
            AddStep("After Bracket→()", CorePass(s,
                NormalizeOptions.BracketNormalize
                | (options & NormalizeOptions.WhitespaceNormalize)));
        }
        else if (options.HasFlag(NormalizeOptions.WhitespaceNormalize))
        {
            AddStep("After Whitespace", CorePass(s, NormalizeOptions.WhitespaceNormalize));
        }

        if (options.HasFlag(NormalizeOptions.LongVowelNormalize))
            AddStep("After LongVowel", LongVowelPass(s));

        if (!string.IsNullOrEmpty(additionalReplace))
            AddStep("After AdditionalReplace", ApplyAdditionalReplace(s, additionalReplace!));

        if (!string.IsNullOrEmpty(additionalRemove))
            AddStep("After AdditionalRemove", ApplyAdditionalRemove(s, additionalRemove!));

        steps.Add(("Final", s));
        return steps;
    }

    /// <summary>
    /// params[] の各文字列要素に正規化を適用する。
    /// </summary>
    public static object[] NormalizeParams(
        object[] parameters,
        NormalizeOptions options,
        string? additionalRemove = null,
        string? additionalReplace = null)
    {
        if (options == NormalizeOptions.None
            && string.IsNullOrEmpty(additionalRemove)
            && string.IsNullOrEmpty(additionalReplace))
            return parameters;

        var result = new object[parameters.Length];
        for (var i = 0; i < parameters.Length; i++)
        {
            result[i] = parameters[i] is string s
                ? Normalize(s, options, additionalRemove, additionalReplace)
                : parameters[i];
        }
        return result;
    }

    // ════════════════════════════════════════════════════════════════════
    // 正規化本体
    // ════════════════════════════════════════════════════════════════════

    private static string NormalizeCore(
        string input,
        NormalizeOptions options,
        string? additionalRemove,
        string? additionalReplace)
    {
        var s = input;

        // ① NFKC + 簡易 CaseFold（.NET 標準）
        if (options.HasFlag(NormalizeOptions.NfkcCaseFold))
            s = NfkcCaseFoldPass(s);

        // ② 半角カナ → 全角カナ（半角カナ区間のみ自前変換）
        if (options.HasFlag(NormalizeOptions.HalfKanaToFullKana))
            s = HalfKanaToFullKanaPass(s);

        // ③ メイン 1 パス処理
        var needCore =
            options.HasFlag(NormalizeOptions.EmojiRemove) ||
            options.HasFlag(NormalizeOptions.WhitespaceNormalize) ||
            options.HasFlag(NormalizeOptions.BracketNormalize) ||
            options.HasFlag(NormalizeOptions.FullwidthSymbolToHalf) ||
            options.HasFlag(NormalizeOptions.FullwidthAlphanumToHalf) ||
            options.HasFlag(NormalizeOptions.HalfwidthAlphanumToUpper);

        if (needCore)
            s = CorePass(s, options);

        // ④ 長音符正規化
        if (options.HasFlag(NormalizeOptions.LongVowelNormalize))
            s = LongVowelPass(s);

        // ⑤ 追加置換
        if (!string.IsNullOrEmpty(additionalReplace))
            s = ApplyAdditionalReplace(s, additionalReplace);

        // ⑥ 追加削除
        if (!string.IsNullOrEmpty(additionalRemove))
            s = ApplyAdditionalRemove(s, additionalRemove);

        return s;
    }

    // ════════════════════════════════════════════════════════════════════
    // ① NFKC + 簡易 CaseFold（.NET 標準）
    // ════════════════════════════════════════════════════════════════════

    private static string NfkcCaseFoldPass(string s)
    {
        // NFKC
        var t = s.Normalize(NormalizationForm.FormKC);
        // 簡易 case-fold（Unicode の full case-fold と完全一致ではない）
        t = t.ToLowerInvariant();

        // 最低限の補正（一般的に検索で問題になりやすい ß を展開）
        // 例: "straße" → "strasse"
        if (t.IndexOf('ß') >= 0)
            t = t.Replace("ß", "ss", StringComparison.Ordinal);
        if (t.IndexOf('ẞ') >= 0)
            t = t.Replace("ẞ", "ss", StringComparison.Ordinal);

        return t;
    }

    // ════════════════════════════════════════════════════════════════════
    // ② 半角カナ → 全角カナ（自前変換）
    // ════════════════════════════════════════════════════════════════════

    // 半角カナ（U+FF61〜U+FF9F）のみを対象
    private static bool IsHalfKana(char ch) => ch >= '\uFF61' && ch <= '\uFF9F';

    // U+FF9E: dakuten, U+FF9F: handakuten
    private const char HalfDakuten = '\uFF9E';
    private const char HalfHandakuten = '\uFF9F';

    private static readonly Dictionary<char, char> HalfKanaBaseMap = new()
    {
        // punctuation
        ['\uFF61'] = '。', // ｡
        ['\uFF62'] = '「', // ｢
        ['\uFF63'] = '」', // ｣
        ['\uFF64'] = '、', // ､
        ['\uFF65'] = '・', // ･

        // small kana
        ['\uFF67'] = 'ァ', // ｧ
        ['\uFF68'] = 'ィ', // ｨ
        ['\uFF69'] = 'ゥ', // ｩ
        ['\uFF6A'] = 'ェ', // ｪ
        ['\uFF6B'] = 'ォ', // ｫ
        ['\uFF6C'] = 'ャ', // ｬ
        ['\uFF6D'] = 'ュ', // ｭ
        ['\uFF6E'] = 'ョ', // ｮ
        ['\uFF6F'] = 'ッ', // ｯ

        // prolonged sound mark
        ['\uFF70'] = 'ー', // ｰ

        // vowels
        ['\uFF71'] = 'ア',
        ['\uFF72'] = 'イ',
        ['\uFF73'] = 'ウ',
        ['\uFF74'] = 'エ',
        ['\uFF75'] = 'オ',

        // K
        ['\uFF76'] = 'カ',
        ['\uFF77'] = 'キ',
        ['\uFF78'] = 'ク',
        ['\uFF79'] = 'ケ',
        ['\uFF7A'] = 'コ',

        // S
        ['\uFF7B'] = 'サ',
        ['\uFF7C'] = 'シ',
        ['\uFF7D'] = 'ス',
        ['\uFF7E'] = 'セ',
        ['\uFF7F'] = 'ソ',

        // T
        ['\uFF80'] = 'タ',
        ['\uFF81'] = 'チ',
        ['\uFF82'] = 'ツ',
        ['\uFF83'] = 'テ',
        ['\uFF84'] = 'ト',

        // N
        ['\uFF85'] = 'ナ',
        ['\uFF86'] = 'ニ',
        ['\uFF87'] = 'ヌ',
        ['\uFF88'] = 'ネ',
        ['\uFF89'] = 'ノ',

        // H
        ['\uFF8A'] = 'ハ',
        ['\uFF8B'] = 'ヒ',
        ['\uFF8C'] = 'フ',
        ['\uFF8D'] = 'ヘ',
        ['\uFF8E'] = 'ホ',

        // M
        ['\uFF8F'] = 'マ',
        ['\uFF90'] = 'ミ',
        ['\uFF91'] = 'ム',
        ['\uFF92'] = 'メ',
        ['\uFF93'] = 'モ',

        // Y
        ['\uFF94'] = 'ヤ',
        ['\uFF95'] = 'ユ',
        ['\uFF96'] = 'ヨ',

        // R
        ['\uFF97'] = 'ラ',
        ['\uFF98'] = 'リ',
        ['\uFF99'] = 'ル',
        ['\uFF9A'] = 'レ',
        ['\uFF9B'] = 'ロ',

        // W
        ['\uFF9C'] = 'ワ',
        ['\uFF9D'] = 'ン',

        // dakuten/handakuten itself are handled separately
    };

    private static readonly Dictionary<(char Base, char Mark), char> VoicedMap = new()
    {
        // K -> G
        [('カ', HalfDakuten)] = 'ガ',
        [('キ', HalfDakuten)] = 'ギ',
        [('ク', HalfDakuten)] = 'グ',
        [('ケ', HalfDakuten)] = 'ゲ',
        [('コ', HalfDakuten)] = 'ゴ',

        // S -> Z
        [('サ', HalfDakuten)] = 'ザ',
        [('シ', HalfDakuten)] = 'ジ',
        [('ス', HalfDakuten)] = 'ズ',
        [('セ', HalfDakuten)] = 'ゼ',
        [('ソ', HalfDakuten)] = 'ゾ',

        // T -> D
        [('タ', HalfDakuten)] = 'ダ',
        [('チ', HalfDakuten)] = 'ヂ',
        [('ツ', HalfDakuten)] = 'ヅ',
        [('テ', HalfDakuten)] = 'デ',
        [('ト', HalfDakuten)] = 'ド',

        // H -> B / P
        [('ハ', HalfDakuten)] = 'バ',
        [('ヒ', HalfDakuten)] = 'ビ',
        [('フ', HalfDakuten)] = 'ブ',
        [('ヘ', HalfDakuten)] = 'ベ',
        [('ホ', HalfDakuten)] = 'ボ',

        [('ハ', HalfHandakuten)] = 'パ',
        [('ヒ', HalfHandakuten)] = 'ピ',
        [('フ', HalfHandakuten)] = 'プ',
        [('ヘ', HalfHandakuten)] = 'ペ',
        [('ホ', HalfHandakuten)] = 'ポ',

        // U -> V
        [('ウ', HalfDakuten)] = 'ヴ',

        // small kana (optional, but common)
        [('ァ', HalfDakuten)] = 'ヴ', // "ｧﾞ" は稀だが "ヴ" 寄せ
    };

    /// <summary>
    /// 半角カナ区間（U+FF61〜U+FF9F）のみを検出し、連続区間を自前で全角カナへ変換する。
    /// 英数字・記号は一切変換しない。
    /// </summary>
    private static string HalfKanaToFullKanaPass(string s)
    {
        // 半角カナが無ければ即返却
        var has = false;
        foreach (var ch in s)
        {
            if (IsHalfKana(ch)) { has = true; break; }
        }
        if (!has) return s;

        var sb = new StringBuilder(s.Length);
        var i = 0;

        while (i < s.Length)
        {
            var ch = s[i];
            if (!IsHalfKana(ch))
            {
                sb.Append(ch);
                i++;
                continue;
            }

            // 半角カナ連続区間をまとめて処理
            var start = i;
            while (i < s.Length && IsHalfKana(s[i])) i++;
            sb.Append(ConvertHalfKanaSegment(s.AsSpan(start, i - start)));
        }

        return sb.ToString();
    }

    private static string ConvertHalfKanaSegment(ReadOnlySpan<char> seg)
    {
        // seg はすべて半角カナ範囲
        var sb = new StringBuilder(seg.Length);

        for (int i = 0; i < seg.Length; i++)
        {
            var ch = seg[i];

            if (ch == HalfDakuten || ch == HalfHandakuten)
            {
                // 直前文字に付く濁点/半濁点
                if (sb.Length > 0)
                {
                    var prev = sb[sb.Length - 1];
                    if (VoicedMap.TryGetValue((prev, ch), out var voiced))
                    {
                        sb[sb.Length - 1] = voiced;
                        continue;
                    }
                }

                // 付与先が無い/変換不可 → 記号として残す
                sb.Append(ch == HalfDakuten ? '゛' : '゜');
                continue;
            }

            if (HalfKanaBaseMap.TryGetValue(ch, out var mapped))
            {
                sb.Append(mapped);
            }
            else
            {
                // 未定義はそのまま（安全策）
                sb.Append(ch);
            }
        }

        return sb.ToString();
    }

    // ════════════════════════════════════════════════════════════════════
    // ③ メイン 1 パス（ArrayPool バッファ）
    // ════════════════════════════════════════════════════════════════════

    private static string CorePass(string s, NormalizeOptions options)
    {
        var emoji = options.HasFlag(NormalizeOptions.EmojiRemove);
        var wsNorm = options.HasFlag(NormalizeOptions.WhitespaceNormalize);
        var bracket = options.HasFlag(NormalizeOptions.BracketNormalize);
        var fullSym = options.HasFlag(NormalizeOptions.FullwidthSymbolToHalf);
        var fullAlnum = options.HasFlag(NormalizeOptions.FullwidthAlphanumToHalf);
        var halfUpper = options.HasFlag(NormalizeOptions.HalfwidthAlphanumToUpper);

        var buf = ArrayPool<char>.Shared.Rent(s.Length);
        var pos = 0;

        var prevSpace = wsNorm;
        var pendingSpaces = 0;

        foreach (var rune in s.EnumerateRunes())
        {
            if (emoji && IsEmoji(rune)) continue;

            if (rune.Value > 0xFFFF)
            {
                if (wsNorm) prevSpace = false;
                if (pendingSpaces > 0)
                {
                    pos = FlushPendingSpaces(buf, pos, pendingSpaces);
                    pendingSpaces = 0;
                }
                pos += rune.EncodeToUtf16(buf.AsSpan(pos));
                continue;
            }

            var ch = (char)rune.Value;

            if (wsNorm && IsNormSpace(ch))
            {
                if (!prevSpace)
                {
                    if (bracket)
                        pendingSpaces++;
                    else
                        buf[pos++] = ' ';
                    prevSpace = true;
                }
                continue;
            }

            char mapped = bracket ? MapBracket(ch) : ch;

            if (bracket && mapped == '(')
            {
                pendingSpaces = 0;
                if (wsNorm) prevSpace = false;
                buf[pos++] = '(';
                continue;
            }

            if (pendingSpaces > 0)
                pos = FlushPendingSpaces(buf, pos, pendingSpaces);
            pendingSpaces = 0;

            ch = mapped;
            if (wsNorm) prevSpace = false;

            if (fullSym && IsFullwidthSymbol(ch))
                ch = ToHalfwidthSymbol(ch);

            if (fullAlnum)
            {
                if (ch >= '\uFF10' && ch <= '\uFF19')
                    ch = (char)(ch - 0xFEE0);
                else if (ch >= '\uFF21' && ch <= '\uFF3A')
                    ch = (char)(ch - 0xFEE0);
                else if (ch >= '\uFF41' && ch <= '\uFF5A')
                    ch = halfUpper
                        ? (char)(ch - 0xFEE0 - 0x20)
                        : (char)(ch - 0xFEE0);
            }

            if (halfUpper && ch >= 'a' && ch <= 'z')
                ch = (char)(ch - 0x20);

            buf[pos++] = ch;
        }

        if (pendingSpaces > 0 && !wsNorm)
            pos = FlushPendingSpaces(buf, pos, pendingSpaces);

        if (wsNorm)
            while (pos > 0 && buf[pos - 1] == ' ') pos--;

        var result = new string(buf, 0, pos);
        ArrayPool<char>.Shared.Return(buf);
        return result;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static int FlushPendingSpaces(char[] buf, int pos, int count)
    {
        for (var k = 0; k < count; k++) buf[pos + k] = ' ';
        return pos + count;
    }

    // ── 絵文字判定（簡易）───────────────────────────────────────────────
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool IsEmoji(Rune rune)
    {
        var v = rune.Value;
        if (v >= 0x1F000) return true;
        if (v >= 0x2600 && v <= 0x27BF) return true;
        if (v >= 0x2300 && v <= 0x23FF) return true;
        if (v >= 0xFE00 && v <= 0xFE0F) return true;   // Variation Selectors
        if (v == 0x200D) return true;                  // ZWJ
        return false;
    }

    // ── 空白判定 ─────────────────────────────────────────────────────────
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool IsNormSpace(char ch) =>
        ch == ' ' || ch == '\u3000' || ch == '\t' ||
        ch == '\r' || ch == '\n' || ch == '\u00A0' ||
        ch == '\u2002' || ch == '\u2003' || ch == '\u200B' ||
        char.IsControl(ch);

    // ── 括弧変換 ─────────────────────────────────────────────────────────
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static char MapBracket(char ch) => ch switch
    {
        '（' or '「' or '『' or '【' or '〔' or '〈' or '《' or '〖'
        or '[' or '［' or '{' or '｛' or '<' or '＜' => '(',

        '）' or '」' or '』' or '】' or '〕' or '〉' or '》' or '〗'
        or ']' or '］' or '}' or '｝' or '>' or '＞' => ')',

        _ => ch
    };

    // ── 全角記号判定・変換 ───────────────────────────────────────────────
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool IsFullwidthSymbol(char ch) =>
        (ch >= '\uFF01' && ch <= '\uFF0F') ||   // ！〜／
        (ch >= '\uFF1A' && ch <= '\uFF20') ||   // ：〜＠
        (ch >= '\uFF3B' && ch <= '\uFF40') ||   // ［〜｀
        (ch >= '\uFF5B' && ch <= '\uFF60') ||   // ｛〜｠
        (ch >= '\uFFE0' && ch <= '\uFFE6');     // ￠〜￦

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static char ToHalfwidthSymbol(char ch) => ch switch
    {
        '\uFFE0' => '\u00A2',   // ￠ → ¢
        '\uFFE1' => '\u00A3',   // ￡ → £
        '\uFFE2' => '\u00AC',   // ￢ → ¬
        '\uFFE3' => '\u00AF',   // ￣ → ¯
        '\uFFE4' => '\u00A6',   // ￤ → ¦
        '\uFFE5' => '\u00A5',   // ￥ → ¥
        '\uFFE6' => '\u20A9',   // ￦ → ₩
        _ => (char)(ch - 0xFEE0)
    };

    // ════════════════════════════════════════════════════════════════════
    // ④ 長音符正規化（1 パス ステートマシン）
    // ════════════════════════════════════════════════════════════════════

    private static readonly HashSet<char> UDan = new(
    [
        'ク','ス','ツ','ヌ','フ','ム','ユ','ル',
        'グ','ズ','ヅ','ブ','プ'
    ]);

    private static readonly HashSet<char> EDan = new(
    [
        'ケ','セ','テ','ネ','ヘ','メ','レ',
        'ゲ','ゼ','デ','ベ','ペ'
    ]);

    private static string LongVowelPass(string s)
    {
        StringBuilder? sb = null;

        for (var i = 0; i < s.Length; i++)
        {
            var ch = s[i];

            if (ch == 'ー' && i > 0)
            {
                var prev = (sb is { Length: > 0 }) ? sb[sb.Length - 1] : s[i - 1];
                if (UDan.Contains(prev) || EDan.Contains(prev))
                {
                    sb ??= new StringBuilder(s.Length).Append(s, 0, i);
                    continue;
                }
            }

            sb?.Append(ch);
        }

        return sb?.ToString() ?? s;
    }

    // ════════════════════════════════════════════════════════════════════
    // ⑤ 追加置換（辞書キャッシュ + 1 パスマッチ）
    // ════════════════════════════════════════════════════════════════════

    private static string ApplyAdditionalReplace(string s, string mappingSpec)
    {
        if (!ReplaceSpecCache.TryGet(mappingSpec, out var pairs))
        {
            pairs = ParseReplacePairs(mappingSpec);
            ReplaceSpecCache.Set(mappingSpec, pairs);
        }
        if (pairs.Count == 0) return s;

        var sb = new StringBuilder(s.Length);
        var i = 0;

        while (i < s.Length)
        {
            var matched = false;
            foreach (var (from, to) in pairs)
            {
                if (from.Length == 0) continue;

                if (i + from.Length <= s.Length &&
                    s.AsSpan(i, from.Length).SequenceEqual(from.AsSpan()))
                {
                    sb.Append(to);
                    i += from.Length;
                    matched = true;
                    break;
                }
            }

            if (!matched) sb.Append(s[i++]);
        }

        return sb.ToString();
    }

    private static List<(string From, string To)> ParseReplacePairs(string mappingSpec)
    {
        var pairs = new List<(string, string)>();

        foreach (var line in mappingSpec.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var idx = line.IndexOf(',');
            if (idx < 0) continue;

            var from = line[..idx].Trim();
            var to = line[(idx + 1)..].Trim();

            if (from.Length > 0) pairs.Add((from, to));
        }

        return pairs;
    }

    // ════════════════════════════════════════════════════════════════════
    // ⑥ 追加削除（HashSet で 1 パス）
    // ════════════════════════════════════════════════════════════════════

    private static string ApplyAdditionalRemove(string s, string removeChars)
    {
        var set = new HashSet<char>(removeChars);
        if (set.Count == 0) return s;

        StringBuilder? sb = null;
        for (var i = 0; i < s.Length; i++)
        {
            if (set.Contains(s[i]))
            {
                sb ??= new StringBuilder(s.Length).Append(s, 0, i);
                continue;
            }

            sb?.Append(s[i]);
        }

        return sb?.ToString() ?? s;
    }
}

// ════════════════════════════════════════════════════════════════════════
// 正規化結果キャッシュのキー
// ════════════════════════════════════════════════════════════════════════

/// <summary>
/// (input, options, additionalRemove, additionalReplace) の組み合わせをキーとする。
/// </summary>
internal readonly struct NormCacheKey(
    string input,
    NormalizeOptions options,
    string? additionalRemove,
    string? additionalReplace) : IEquatable<NormCacheKey>
{
    private readonly string _input = input;
    private readonly NormalizeOptions _opts = options;
    private readonly string? _remove = additionalRemove;
    private readonly string? _replace = additionalReplace;
    private readonly int _hash = ComputeHash(input, options, additionalRemove, additionalReplace);

    private static int ComputeHash(string input, NormalizeOptions opts, string? remove, string? replace)
    {
        var h = new HashCode();
        h.Add(input, StringComparer.Ordinal);
        h.Add((int)opts);
        h.Add(remove, StringComparer.Ordinal);
        h.Add(replace, StringComparer.Ordinal);
        return h.ToHashCode();
    }

    public bool Equals(NormCacheKey other) =>
        _opts == other._opts &&
        string.Equals(_input, other._input, StringComparison.Ordinal) &&
        string.Equals(_remove, other._remove, StringComparison.Ordinal) &&
        string.Equals(_replace, other._replace, StringComparison.Ordinal);

    public override bool Equals(object? obj) => obj is NormCacheKey k && Equals(k);
    public override int GetHashCode() => _hash;
}