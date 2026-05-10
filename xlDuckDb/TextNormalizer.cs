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

    // ── 置換辞書キャッシュ（mappingSpec 文字列 → 解析済み辞典）──────────
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
    /// 正規化した最終文字列を返す。
    /// </summary>
    public static string NormalizeSteps(
        string input,
        NormalizeOptions options,
        string? additionalRemove = null,
        string? additionalReplace = null) =>
        Normalize(input, options, additionalRemove, additionalReplace);

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

    /// <summary>
    /// 正規化のみを行う（追加削除・追加置換なし）。
    /// </summary>
    public static string NormalizeOnly(string input, NormalizeOptions options)
        => Normalize(input, options, null, null);

    /// <summary>
    /// 追加削除のみを行う（正規化・追加置換なし）。
    /// </summary>
    public static string RemoveOnly(string input, string removeSpec)
    {
        if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(removeSpec)) return input;
        return ApplyAdditionalRemove(input, removeSpec);
    }

    /// <summary>
    /// 追加置換のみを行う（正規化・追加削除なし）。
    /// </summary>
    public static string ReplaceOnly(string input, string replaceSpec)
    {
        if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(replaceSpec)) return input;
        return ApplyAdditionalReplace(input, replaceSpec);
    }

    /// <summary>
    /// 正規化・追加削除・追加置換をすべて適用した文字列を返す。
    /// </summary>
    public static string NormalizeAll(string input, NormalizeOptions options, string? additionalRemove = null, string? additionalReplace = null)
    {
        return Normalize(input, options, additionalRemove, additionalReplace);
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
        ['\uFF81'] = 'イ',
        ['\uFF82'] = 'ウ',
        ['\uFF83'] = 'エ',
        ['\uFF84'] = 'オ',

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

        // R
        ['\uFF94'] = 'ラ',
        ['\uFF95'] = 'リ',
        ['\uFF96'] = 'ル',
        ['\uFF97'] = 'レ',
        ['\uFF98'] = 'ロ',

        // W
        ['\uFF99'] = 'ワ',
        ['\uFF9A'] = 'ヰ',
        ['\uFF9B'] = 'ヱ',
        ['\uFF9C'] = 'ヲ', // ﾜ ｸﾞ ﾊﾟ は除外
    };

    private static string HalfKanaToFullKanaPass(string s)
    {
        var sb = new StringBuilder(s.Length);
        foreach (var c in s)
        {
            if (IsHalfKana(c) && HalfKanaBaseMap.TryGetValue(c, out var fc))
            {
                sb.Append(fc);
            }
            else
            {
                sb.Append(c);
            }
        }
        return sb.ToString();
    }

    // ════════════════════════════════════════════════════════════════════
    // ③ メイン処理（1パス）
    // ════════════════════════════════════════════════════════════════════

    // 全角英数字 ＆ 半角英数字 ＆ 半角カナ ＆ 空白 ＆ 一部記号
    // ※ 詳細は実装部のコメントを参照
    private static readonly HashSet<char> CoreNormChar = new()
    {
        // whitespace
        '\u0009', '\u000A', '\u000B', '\u000C', '\u000D', // Control
        ' ',  // SP
        '　', // IDEOGRAPHIC SPACE
        // 全角英数字
        '０', '１', '２', '３', '４', '５', '６', '７', '８', '９', // 0xFF10-0xFF19
        'Ａ', 'Ｂ', 'Ｃ', 'Ｄ', 'Ｅ', 'Ｆ', // 0xFF21-0xFF26
        'Ｇ', 'Ｈ', 'Ｉ', 'Ｊ', 'Ｋ', 'Ｌ', 'Ｍ', 'Ｎ', 'Ｏ', 'Ｐ', 'Ｑ', 'Ｒ', 'Ｓ', 'Ｔ', 'Ｕ', 'Ｖ', 'Ｗ', 'Ｘ', 'Ｙ', 'Ｚ', // 0xFF27-0xFF38
        // 半角英数字
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', // 0x30-0x39
        'A', 'B', 'C', 'D', 'E', 'F', // 0x41-0x46
        'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', // 0x47-0x5A
        // 半角カナ
        'ｱ', 'ｲ', 'ｳ', 'ｴ', 'ｵ',
        'ｶ', 'ｷ', 'ｸ', 'ｹ', 'ｺ',
        'ｻ', 'ｼ', 'ｽ', 'ｾ', 'ｿ',
        'ﾀ', 'ﾁ', 'ﾂ', 'ﾃ', 'ﾄ',
        'ﾅ', 'ﾆ', 'ﾇ', 'ﾈ', 'ﾉ',
        'ﾊ', 'ﾋ', 'ﾌ', 'ﾍ', 'ﾎ',
        'ﾏ', 'ﾐ', 'ﾑ', 'ﾒ', 'ﾓ',
        'ﾗ', 'ﾘ', 'ﾙ', 'ﾚ', 'ﾛ',
        'ﾜ', 'ｦ',
        // 括弧
        '（', '）', '｛', '｝', '［', '］', // U+FF08 〜 U+FF1B
        // 一部記号
        '､', '．', 'ー', // 0x3010-0x3013
        '・', // ･
        '「', '」', '〈', '〉', '《', '》', // 0x300C-0x300F, 0x3018-0x301B
    };

    private static readonly Dictionary<char, char> CoreNormMap = new()
    {
        // 数字
        ['０'] = '0', ['１'] = '1', ['２'] = '2', ['３'] = '3', ['４'] = '4',
        ['５'] = '5', ['６'] = '6', ['７'] = '7', ['８'] = '8', ['９'] = '9',

        // 英大文字
        ['Ａ'] = 'A', ['Ｂ'] = 'B', ['Ｃ'] = 'C', ['Ｄ'] = 'D', ['Ｅ'] = 'E',
        ['Ｆ'] = 'F', ['Ｇ'] = 'G', ['Ｈ'] = 'H', ['Ｉ'] = 'I', ['Ｊ'] = 'J',
        ['Ｋ'] = 'K', ['Ｌ'] = 'L', ['Ｍ'] = 'M', ['Ｎ'] = 'N', ['Ｏ'] = 'O',
        ['Ｐ'] = 'P', ['Ｑ'] = 'Q', ['Ｒ'] = 'R', ['Ｓ'] = 'S', ['Ｔ'] = 'T',
        ['Ｕ'] = 'U', ['Ｖ'] = 'V', ['Ｗ'] = 'W', ['Ｘ'] = 'X', ['Ｙ'] = 'Y',
        ['Ｚ'] = 'Z',

        // 英小文字
        ['ａ'] = 'a', ['ｂ'] = 'b', ['ｃ'] = 'c', ['ｄ'] = 'd', ['ｅ'] = 'e',
        ['ｆ'] = 'f', ['ｇ'] = 'g', ['ｈ'] = 'h', ['ｉ'] = 'i', ['ｊ'] = 'j',
        ['ｋ'] = 'k', ['ｍ'] = 'm', ['ｎ'] = 'n', ['ｏ'] = 'o', ['ｐ'] = 'p',
        ['ｑ'] = 'q', ['ｒ'] = 'r', ['ｓ'] = 's', ['ｔ'] = 't', ['ｕ'] = 'u',
        ['ｖ'] = 'v', ['ｗ'] = 'w', ['ｘ'] = 'x', ['ｙ'] = 'y', ['ｚ'] = 'z',

        // 半角カナ
        ['ｱ'] = 'ア', ['ｲ'] = 'イ', ['ｳ'] = 'ウ', ['ｴ'] = 'エ', ['ｵ'] = 'オ',
        ['ｶ'] = 'カ', ['ｷ'] = 'キ', ['ｸ'] = 'ク', ['ｹ'] = 'ケ', ['ｺ'] = 'コ',
        ['ｻ'] = 'サ', ['ｼ'] = 'シ', ['ｽ'] = 'ス', ['ｾ'] = 'セ', ['ｿ'] = 'ソ',
        ['ﾀ'] = 'タ', ['ﾁ'] = 'チ', ['ﾂ'] = 'ツ', ['ﾃ'] = 'テ', ['ﾄ'] = 'ト',
        ['ﾅ'] = 'ナ', ['ﾆ'] = 'ニ', ['ﾇ'] = 'ヌ', ['ﾈ'] = 'ネ', ['ﾉ'] = 'ノ',
        ['ﾊ'] = 'ハ', ['ﾋ'] = 'ヒ', ['ﾌ'] = 'フ', ['ﾍ'] = 'ヘ', ['ﾎ'] = 'ホ',
        ['ﾏ'] = 'マ', ['ﾐ'] = 'ミ', ['ﾑ'] = 'ム', ['ﾒ'] = 'メ', ['ﾓ'] = 'モ',
        ['ﾗ'] = 'ラ', ['ﾘ'] = 'リ', ['ﾙ'] = 'ル', ['ﾚ'] = 'レ', ['ﾛ'] = 'ロ',
        ['ﾜ'] = 'ワ', ['ｦ'] = 'ヲ',
    };

    private static string CorePass(string s, NormalizeOptions options)
    {
        Span<char> src = stackalloc char[256];
        Span<char> dst = stackalloc char[256];
        var n = 0;
        foreach (var c in s)
        {
            if (CoreNormChar.Contains(c))
            {
                src[n] = c;
                n++;
            }
        }

        // 基本変換
        for (var i = 0; i < n; i++)
        {
            if (CoreNormMap.TryGetValue(src[i], out var fc))
                dst[i] = fc;
            else
                dst[i] = src[i];
        }

        // スペース・グリフ幅調整
        // （半角スペース→全角スペース、連続スペース→単一スペース、前後空白削除）
        //   参考: https://github.com/dotnet/corefx/issues/27395
        //   「　」: U+3000 IDEOGRAPHIC SPACE
        //   「 」: U+0020 SPACE
        //   注意: U+00A0 NO-BREAK SPACE は除外
        //   注意: U+2007 FIGURE SPACE は通常の空白扱い
        if (options.HasFlag(NormalizeOptions.WhitespaceNormalize))
        {
            // 全角スペースは半角に変換
            for (var i = 0; i < n; i++)
            {
                if (dst[i] == '　')
                    dst[i] = ' ';
            }

            // 連続スペース → 単一スペース
            // 前後空白削除
            var l = 0;
            var skip = false;
            for (var i = 0; i < n; i++)
            {
                if (dst[i] == ' ')
                {
                    if (skip) continue;
                    skip = true;
                }
                else
                {
                    skip = false;
                }
                dst[l] = dst[i];
                l++;
            }
            if (l > 0 && dst[l - 1] == ' ')
                l--;

            n = l;
        }

        // 括弧類 → ( )、( 前空白削除
        // 例: 「　（abc）　」→「(abc)」
        if (options.HasFlag(NormalizeOptions.BracketNormalize))
        {
            // ([]{})
            // 前空白削除は CorePass 内で実施
            for (var i = 0; i < n; i++)
            {
                dst[i] = dst[i] switch
                {
                    '（' or '）' or '｛' or '｝' => '(', // U+FF08 〜 U+FF1B
                    '［' or '］' => '[', // U+3010, U+3011
                    _ => dst[i],
                };
            }
        }

        // 絵文字削除（現在はスクリーニングのみ、必要なら強化可）
        // ※ 非BMP 文字（U+10000 以降）は対象外
        if (options.HasFlag(NormalizeOptions.EmojiRemove))
        {
            var l = 0;
            for (var i = 0; i < n; i++)
            {
                var c = dst[i];
                if (c >= 0xE000 && c <= 0xF8FF) // Emoji (Basic Multilingual Plane)
                {
                    // 削除
                }
                else
                {
                    dst[l] = c;
                    l++;
                }
            }
            n = l;
        }

        return new string(dst[..n]);
    }

    // ════════════════════════════════════════════════════════════════════
    // ④ 長音符正規化
    // ════════════════════════════════════════════════════════════════════

    // 一部仮名の長音符（ー）を除去
    // 例: 「こんにちわー」→「こんにちは」
    private static readonly HashSet<string> LongVowelNormSet = new()
    {
        "あ", "い", "う", "え", "お",
        "か", "き", "く", "け", "こ",
        "さ", "し", "す", "せ", "そ",
        "た", "ち", "つ", "て", "と",
        "な", "に", "ぬ", "ね", "の",
        "は", "ひ", "ふ", "へ", "ほ",
        "ま", "み", "む", "め", "も",
        "や", "ゆ", "よ",
        "ら", "り", "る", "れ", "ろ",
        "わ", "を",
    };

    private static string LongVowelPass(string s)
    {
        var sb = new StringBuilder(s.Length);
        for (var i = 0; i < s.Length; i++)
        {
            var c = s[i];
            // 「ー」が前の文字の直後にある場合、前の文字が仮名なら「ー」をスキップ
            if (c == 'ー' && i > 0)
            {
                var prev = s[i - 1];
                if (LongVowelNormSet.Contains(char.ConvertFromUtf32(prev)))
                {
                    continue;
                }
            }
            sb.Append(c);
        }
        return sb.ToString();
    }

    // ════════════════════════════════════════════════════════════════════
    // ⑤ 追加置換
    // ════════════════════════════════════════════════════════════════════

    // シンプルな CSV スタイルの置換仕様（, で分割された値のペア）
    //   例: "abc,ABC,123,¬¨"`など
    private static readonly char[] AdditionalReplaceSep = new[] { ',', '，' };

    private static string ApplyAdditionalReplace(string s, string replaceSpec)
    {
        if (!ReplaceSpecCache.TryGet(replaceSpec, out var list))
        {
            var parts = replaceSpec.Split(AdditionalReplaceSep, StringSplitOptions.RemoveEmptyEntries);
            list = new List<(string From, string To)>(parts.Length / 2);
            for (var i = 0; i + 1 < parts.Length; i += 2)
            {
                list.Add((parts[i], parts[i + 1]));
            }
            ReplaceSpecCache.Set(replaceSpec, list);
        }
        var dst = s;
        foreach ((string from, string to) in list)
        {
            dst = dst.Replace(from, to, StringComparison.Ordinal);
        }
        return dst;
    }

    // ════════════════════════════════════════════════════════════════════
    // ⑥ 追加削除
    // ════════════════════════════════════════════════════════════════════

    private static string ApplyAdditionalRemove(string s, string removeSpec)
    {
        var dst = s;
        foreach (var r in removeSpec)
        {
            dst = dst.Replace(r.ToString(), "", StringComparison.Ordinal);
        }
        return dst;
    }

    // ════════════════════════════════════════════════════════════════════
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
}