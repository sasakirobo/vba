using System.Buffers;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using Icu;
using Icu.Normalization;

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
    None                     = 0,
    NfkcCaseFold             = 1 << 0,   //   1  NFKC + Unicode CaseFold
    LongVowelNormalize       = 1 << 1,   //   2  長音「ー」正規化
    EmojiRemove              = 1 << 2,   //   4  絵文字削除
    FullwidthAlphanumToHalf  = 1 << 3,   //   8  全角英数字 → 半角大文字
    HalfwidthAlphanumToUpper = 1 << 4,   //  16  半角英数字 → 大文字
    HalfKanaToFullKana       = 1 << 5,   //  32  半角カナ → 全角カナ
    FullwidthSymbolToHalf    = 1 << 6,   //  64  全角記号 → 半角記号
    WhitespaceNormalize      = 1 << 7,   // 128  空白正規化・Trim
    BracketNormalize         = 1 << 8,   // 256  括弧類 → ( )、( 前空白削除
    All = NfkcCaseFold | LongVowelNormalize | EmojiRemove
        | FullwidthAlphanumToHalf | HalfwidthAlphanumToUpper
        | HalfKanaToFullKana | FullwidthSymbolToHalf
        | WhitespaceNormalize | BracketNormalize,   // 511
}

// ════════════════════════════════════════════════════════════════════════
// テキスト正規化エンジン
// ════════════════════════════════════════════════════════════════════════

/// <summary>
/// ICU.net を使ったテキスト正規化クラス。
///
/// ── 正確性の保証 ──────────────────────────────────────────────────────
///   ・NfkcCaseFold  : ICU UChar.FoldCase（ToUpperInvariant ではない）
///   ・HalfKanaToFullKana : 半角カナ区間のみ ICU Transliterate
///                          英数字・記号は全角化しない
///   ・IsEmoji       : BMP サロゲートペア + Emoji プレゼンテーション の2段判定
///   ・BracketNormalize: ( 前空白削除を同一パスで実施
///   ・LongVowelNormalize: 1 パスのステートマシン（Regex なし）
///
/// ── 高速化のポイント ──────────────────────────────────────────────────
///   ・NormalizeCore  : ArrayPool バッファで 1 パス処理
///   ・HalfKanaToFullKana : 連続区間をまとめて ICU に渡す
///   ・additionalRemove   : HashSet&lt;char&gt; で O(n) 1 パス
///   ・additionalReplace  : 辞書キャッシュ（LRU）+ 左から 1 パスマッチ
///   ・Normalize 結果 : LRU キャッシュ（上限設定可）
/// </summary>
public static class TextNormalizer
{
    // ── ICU.NET Transliterator（半角カナ → 全角カナ専用）────────────────────
    // "Halfwidth-Fullwidth" を使うが、適用範囲を半角カナ区間のみに限定する
    private static readonly Transliterator HalfKanaTranslit =
        Transliterator.CreateInstance("Halfwidth-Fullwidth");

    // ── 正規化結果キャッシュ（LRU）───────────────────────────────────────
    // キー: (入力文字列, options, additionalRemove, additionalReplace)
    // デフォルト上限: 1024件（環境変数 XLDUCKDB_NORM_CACHE_SIZE で上書き可）
    public static readonly int NormCacheCapacity = ResolveNormCacheCapacity();

    private static int ResolveNormCacheCapacity()
    {
        if (int.TryParse(
                Environment.GetEnvironmentVariable("XLDUCKDB_NORM_CACHE_SIZE"),
                out var v) && v > 0)
            return Math.Clamp(v, 64, 65536);
        return 1024;
    }

    private static readonly LruCache<NormCacheKey, string> NormCache =
        new(NormCacheCapacity);

    // ── 置換辞書キャッシュ（mappingSpec 文字列 → 解析済み辞書）──────────
    // スレッドセーフに LruCache で管理（サイズは小さく十分）
    private static readonly LruCache<string, List<(string From, string To)>> ReplaceSpecCache =
        new(64);

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

        // ── 結果キャッシュ確認 ─────────────────────────────────────────
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
        var steps = new List<(string, string)>();
        steps.Add(("Input", input));

        if (string.IsNullOrEmpty(input)) return steps;

        var s = input;

        void AddStep(string label, string next)
        {
            if (next != s) steps.Add((label, next));
            s = next;
        }

        if (options.HasFlag(NormalizeOptions.NfkcCaseFold))
        {
            // .NET 標準機能で NFKC + CaseFold をシミュレート
            // これにより ICU.net のバージョン依存から解放されます
            var normalized = s.Normalize(NormalizationForm.FormKC).ToLowerInvariant();
            AddStep("After NFKC+CaseFold", normalized);
        }

        if (options.HasFlag(NormalizeOptions.HalfKanaToFullKana))
            AddStep("After HalfKana→Full", HalfKanaToFullKanaPass(s));

        // CorePass を個別ステップに分解して適用
        if (options.HasFlag(NormalizeOptions.EmojiRemove))
            AddStep("After EmojiRemove", CorePass(s,
                NormalizeOptions.EmojiRemove));

        if (options.HasFlag(NormalizeOptions.FullwidthSymbolToHalf))
            AddStep("After FullSymbol→Half", CorePass(s,
                NormalizeOptions.FullwidthSymbolToHalf));

        if (options.HasFlag(NormalizeOptions.FullwidthAlphanumToHalf))
            AddStep("After FullwidthAlnum→Half", CorePass(s,
                NormalizeOptions.FullwidthAlphanumToHalf
                | (options & NormalizeOptions.HalfwidthAlphanumToUpper)));

        else if (options.HasFlag(NormalizeOptions.HalfwidthAlphanumToUpper))
            AddStep("After HalfAlnum→Upper", CorePass(s,
                NormalizeOptions.HalfwidthAlphanumToUpper));

        if (options.HasFlag(NormalizeOptions.BracketNormalize))
            AddStep("After Bracket→()", CorePass(s,
                NormalizeOptions.BracketNormalize
                | (options & NormalizeOptions.WhitespaceNormalize)));

        else if (options.HasFlag(NormalizeOptions.WhitespaceNormalize))
            AddStep("After Whitespace", CorePass(s,
                NormalizeOptions.WhitespaceNormalize));

        if (options.HasFlag(NormalizeOptions.LongVowelNormalize))
            AddStep("After LongVowel", LongVowelPass(s));

        if (!string.IsNullOrEmpty(additionalReplace))
            AddStep("After AdditionalReplace", ApplyAdditionalReplace(s, additionalReplace!));

        if (!string.IsNullOrEmpty(additionalRemove))
            AddStep("After AdditionalRemove", ApplyAdditionalRemove(s, additionalRemove!));

        // Final は Input と同一でも必ず出力
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
            result[i] = parameters[i] is string s
                ? Normalize(s, options, additionalRemove, additionalReplace)
                : parameters[i];
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

        // ① NFKC + Unicode CaseFold（ICU）
        //    ※ UChar.FoldCase は ToUpperInvariant とは異なる
        //      例: ß → ss、ﬁ → fi などの合字展開も行われる
        if (options.HasFlag(NormalizeOptions.NfkcCaseFold))
        {
            // ICU.net を使わず、.NET 標準の Normalization を使用
            // 合字（ﬁ -> fi）や特殊文字（ß -> ss）の展開も .NET 5+ であれば高い精度で行われます
            s = s.Normalize(NormalizationForm.FormKC).ToLowerInvariant();
        }

        // ② 半角カナ → 全角カナ
        //    文字列内の半角カナ連続区間のみ ICU Transliterate に渡す。
        //    英数字・半角記号には一切触れない。
        if (options.HasFlag(NormalizeOptions.HalfKanaToFullKana))
            s = HalfKanaToFullKanaPass(s);

        // ③ メインの 1 パス処理（ArrayPool バッファ使用）
        //    EmojiRemove / WhitespaceNormalize / BracketNormalize /
        //    FullwidthSymbolToHalf / FullwidthAlphanumToHalf /
        //    HalfwidthAlphanumToUpper を一括処理
        var needCore =
            options.HasFlag(NormalizeOptions.EmojiRemove)              ||
            options.HasFlag(NormalizeOptions.WhitespaceNormalize)      ||
            options.HasFlag(NormalizeOptions.BracketNormalize)         ||
            options.HasFlag(NormalizeOptions.FullwidthSymbolToHalf)    ||
            options.HasFlag(NormalizeOptions.FullwidthAlphanumToHalf)  ||
            options.HasFlag(NormalizeOptions.HalfwidthAlphanumToUpper);

        if (needCore)
            s = CorePass(s, options);

        // ④ 長音符正規化（1 パスのステートマシン）
        if (options.HasFlag(NormalizeOptions.LongVowelNormalize))
            s = LongVowelPass(s);

        // ⑤ 追加置換（辞書キャッシュ + 1 パスマッチ）
        if (!string.IsNullOrEmpty(additionalReplace))
            s = ApplyAdditionalReplace(s, additionalReplace);

        // ⑥ 追加削除（HashSet で 1 パス）
        if (!string.IsNullOrEmpty(additionalRemove))
            s = ApplyAdditionalRemove(s, additionalRemove);

        return s;
    }

    // ════════════════════════════════════════════════════════════════════
    // ② 半角カナ → 全角カナ（連続区間まとめて ICU）
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 半角カナ区間（U+FF65〜U+FF9F）のみを検出し、連続区間を 1 回の
    /// ICU Transliterate 呼び出しでまとめて変換する。
    /// 英数字・記号・ひらがな・漢字は一切変換しない。
    /// </summary>
    private static string HalfKanaToFullKanaPass(string s)
    {
        // 半角カナが含まれなければ即返却
        var hasHalfKana = false;
        foreach (var ch in s)
        {
            if (ch >= '\uFF65' && ch <= '\uFF9F') { hasHalfKana = true; break; }
        }
        if (!hasHalfKana) return s;

        var sb = new StringBuilder(s.Length);
        var i  = 0;

        while (i < s.Length)
        {
            var ch = s[i];
            if (ch >= '\uFF65' && ch <= '\uFF9F')
            {
                // 連続する半角カナ区間を探す
                var start = i;
                while (i < s.Length && s[i] >= '\uFF65' && s[i] <= '\uFF9F') i++;
                // 区間をまとめて ICU Transliterate（呼び出し 1 回）
                sb.Append(HalfKanaTranslit.Transliterate(s[start..i]));
            }
            else
            {
                sb.Append(ch);
                i++;
            }
        }

        return sb.ToString();
    }

    // ════════════════════════════════════════════════════════════════════
    // ③ メイン 1 パス（ArrayPool バッファ）
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// EmojiRemove / WhitespaceNormalize / BracketNormalize /
    /// FullwidthSymbolToHalf / FullwidthAlphanumToHalf / HalfwidthAlphanumToUpper
    /// を 1 パスで処理する。
    ///
    /// ArrayPool でバッファを確保し、文字列アロケーションを最小化する。
    ///
    /// BracketNormalize では「(」の直前にある空白を削除する。
    ///   → 空白文字を即バッファに書かず "保留カウンタ" に積み、
    ///     次の文字が「(」なら捨て、それ以外なら確定して書く。
    ///
    /// WhitespaceNormalize では Trim を最後に 1 回だけ行う。
    ///   → Trim 後のコピーを避けるため、先頭空白は skipLeadingSpace フラグで管理、
    ///     末尾空白は最終的に pos を巻き戻す。
    /// </summary>
    private static string CorePass(string s, NormalizeOptions options)
    {
        var emoji      = options.HasFlag(NormalizeOptions.EmojiRemove);
        var wsNorm     = options.HasFlag(NormalizeOptions.WhitespaceNormalize);
        var bracket    = options.HasFlag(NormalizeOptions.BracketNormalize);
        var fullSym    = options.HasFlag(NormalizeOptions.FullwidthSymbolToHalf);
        var fullAlnum  = options.HasFlag(NormalizeOptions.FullwidthAlphanumToHalf);
        var halfUpper  = options.HasFlag(NormalizeOptions.HalfwidthAlphanumToUpper);

        // バッファは入力長で十分（変換後も文字数は増えない）
        var buf = ArrayPool<char>.Shared.Rent(s.Length);
        var pos = 0;

        // WhitespaceNormalize: 先頭 Trim のため最初は "前が空白" 扱い
        var prevSpace        = wsNorm;
        // BracketNormalize: '(' 前空白を保留するカウンタ
        var pendingSpaces    = 0;

        // サロゲートペアを正しく扱うため EnumerateRunes を使用
        foreach (var rune in s.EnumerateRunes())
        {
            // ── 絵文字削除 ──────────────────────────────────────────────
            if (emoji && IsEmoji(rune)) continue;

            // サロゲートペア文字（BMP外）は以下の char 変換処理に含まれないまま書く
            if (rune.Value > 0xFFFF)
            {
                // BMP 外の非絵文字文字はそのまま書く（サロゲートペア 2 char）
                if (wsNorm) prevSpace = false;
                pendingSpaces = FlushPendingSpaces(buf, pos, pendingSpaces);
                pos += rune.EncodeToUtf16(buf.AsSpan(pos));
                continue;
            }

            var ch = (char)rune.Value;

            // ── 空白正規化 ───────────────────────────────────────────────
            if (wsNorm && IsNormSpace(ch))
            {
                if (!prevSpace)
                {
                    if (bracket)
                        pendingSpaces++;        // ( の前かもしれないので保留
                    else
                        buf[pos++] = ' ';
                    prevSpace = true;
                }
                continue;
            }

            // ── 括弧変換 ────────────────────────────────────────────────
            char mapped = bracket ? MapBracket(ch) : ch;

            if (bracket && mapped == '(')
            {
                // '(' → 保留空白を捨てて ( を出力
                pendingSpaces = 0;
                if (wsNorm) prevSpace = false;
                buf[pos++] = '(';
                continue;
            }

            // '(' でない → 保留空白を確定
            if (pendingSpaces > 0)
                pos = FlushPendingSpaces(buf, pos, pendingSpaces);
            pendingSpaces = 0;

            ch = mapped; // 括弧変換済みの文字を使う

            if (wsNorm) prevSpace = false;

            // ── 全角記号 → 半角 ─────────────────────────────────────────
            // 対象: U+FF01〜U+FF0F / U+FF1A〜U+FF20 / U+FF3B〜U+FF40 /
            //       U+FF5B〜U+FF60 / U+FFE0〜U+FFE6
            // ※ U+FF10〜U+FF19（全角数字）・U+FF21〜U+FF5A（全角英字）は
            //   FullwidthAlphanumToHalf で扱うため、ここでは除外
            if (fullSym && IsFullwidthSymbol(ch))
                ch = ToHalfwidthSymbol(ch);

            // ── 全角英数字 → 半角（+ 大文字化）────────────────────────
            if (fullAlnum)
            {
                if (ch >= '\uFF10' && ch <= '\uFF19')          // ０〜９
                    ch = (char)(ch - 0xFEE0);
                else if (ch >= '\uFF21' && ch <= '\uFF3A')     // Ａ〜Ｚ
                    ch = (char)(ch - 0xFEE0);
                else if (ch >= '\uFF41' && ch <= '\uFF5A')     // ａ〜ｚ → 半角
                    ch = halfUpper
                        ? (char)(ch - 0xFEE0 - 0x20)           // → 半角大文字
                        : (char)(ch - 0xFEE0);                 // → 半角小文字
            }

            // ── 半角小文字 → 大文字 ─────────────────────────────────────
            if (halfUpper && ch >= 'a' && ch <= 'z')
                ch = (char)(ch - 0x20);

            buf[pos++] = ch;
        }

        // 末尾に保留空白が残っていれば確定（WhitespaceNormalize の末尾 Trim で除去される）
        if (pendingSpaces > 0 && !wsNorm)
            pos = FlushPendingSpaces(buf, pos, pendingSpaces);

        // WhitespaceNormalize: 末尾 Trim（空白を pos を戻して除去）
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

    // ── 絵文字判定 ───────────────────────────────────────────────────────
    // 判定基準:
    //   1. BMP 外の文字（サロゲートペア）のうち絵文字ブロックに属するもの
    //   2. BMP 内の記号・絵文字ブロック（U+2600〜U+27BF, U+2300〜U+23FF 等）
    //      ただし OtherSymbol 全体ではなく、絵文字として定義される範囲のみ
    //   ※ UnicodeCategory.OtherSymbol だけで判定すると ①②③ 等まで削除されるため不採用
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool IsEmoji(Rune rune)
    {
        var v = rune.Value;
        // BMP 外（U+1F000〜）: 絵文字専用面
        if (v >= 0x1F000) return true;
        // BMP 内の絵文字・記号ブロック（①②③ 等の enclosed alphanumeric は含まない）
        if (v >= 0x2600 && v <= 0x27BF) return true;   // 各種記号
        if (v >= 0x2300 && v <= 0x23FF) return true;   // 技術記号
        if (v >= 0xFE00 && v <= 0xFE0F) return true;   // 異体字セレクタ
        if (v == 0x200D)                return true;   // ゼロ幅接合子
        return false;
    }

    // ── 空白判定 ─────────────────────────────────────────────────────────
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool IsNormSpace(char ch) =>
        ch == ' '      || ch == '\u3000' || ch == '\t'     ||
        ch == '\r'     || ch == '\n'     || ch == '\u00A0' ||
        ch == '\u2002' || ch == '\u2003' || ch == '\u200B' ||
        char.IsControl(ch);

    // ── 括弧変換（switch 式）────────────────────────────────────────────
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
        (ch >= '\uFF01' && ch <= '\uFF0F') ||   // ！〜／ (英数字の前)
        (ch >= '\uFF1A' && ch <= '\uFF20') ||   // ：〜＠ (英数字の間)
        (ch >= '\uFF3B' && ch <= '\uFF40') ||   // ［〜｀ (英大文字の後)
        (ch >= '\uFF5B' && ch <= '\uFF60') ||   // ｛〜｠ (英小文字の後)
        (ch >= '\uFFE0' && ch <= '\uFFE6');     // ￠〜￦

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static char ToHalfwidthSymbol(char ch) => ch switch
    {
        '\uFFE0' => '\u00A2',   // ￠ → ¢
        '\uFFE1' => '\u00A3',   // ￡ → £
        '\uFFE2' => '\u00AC',   // ￢ → ¬
        '\uFFE3' => '\u00AF',   // ￣ → ¯（マクロン）
        '\uFFE4' => '\u00A6',   // ￤ → ¦
        '\uFFE5' => '\u00A5',   // ￥ → ¥
        '\uFFE6' => '\u20A9',   // ￦ → ₩
        _        => (char)(ch - 0xFEE0)
    };

    // ════════════════════════════════════════════════════════════════════
    // ④ 長音符正規化（1 パス ステートマシン）
    // ════════════════════════════════════════════════════════════════════

    // ウ段カナ直後の「ー」を削除
    private static readonly HashSet<char> UDan = new(
    [
        'ク','ス','ツ','ヌ','フ','ム','ユ','ル',
        'グ','ズ','ヅ','ブ','プ'
    ]);
    // エ段カナ直後の「ー」を削除
    private static readonly HashSet<char> EDan = new(
    [
        'ケ','セ','テ','ネ','ヘ','メ','レ',
        'ゲ','ゼ','デ','ベ','ペ'
    ]);

    /// <summary>
    /// ウ段・エ段カナの直後にある「ー」を削除する 1 パスのステートマシン。
    /// Regex を使わないため高速。
    /// </summary>
    private static string LongVowelPass(string s)
    {
        StringBuilder? sb = null;

        for (var i = 0; i < s.Length; i++)
        {
            var ch = s[i];

            if (ch == 'ー' && i > 0)
            {
                // 直前の有効文字（sb がある場合はそこから取得）
                var prev = (sb is { Length: > 0 }) ? sb[sb.Length - 1] : s[i - 1];

                if (UDan.Contains(prev) || EDan.Contains(prev))
                {
                    // 「ー」を削除 → sb に書かない
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

    /// <summary>
    /// 置換ペア辞書を LRU キャッシュに保持し、文字列を左から 1 パスで走査。
    /// 各位置で全パターンの先頭文字を比較し、マッチすれば置換して進む。
    /// </summary>
    private static string ApplyAdditionalReplace(string s, string mappingSpec)
    {
        if (!ReplaceSpecCache.TryGet(mappingSpec, out var pairs))
        {
            pairs = ParseReplacePairs(mappingSpec);
            ReplaceSpecCache.Set(mappingSpec, pairs);
        }
        if (pairs.Count == 0) return s;

        var sb = new StringBuilder(s.Length);
        var i  = 0;

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
            var to   = line[(idx + 1)..].Trim();
            if (from.Length > 0) pairs.Add((from, to));
        }
        return pairs;
    }

    // ════════════════════════════════════════════════════════════════════
    // ⑥ 追加削除（HashSet で 1 パス）
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// HashSet&lt;char&gt; で O(1) 判定し、1 パスで削除文字をスキップする。
    /// </summary>
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
                // 削除: sb に追加しない
            }
            else
            {
                sb?.Append(s[i]);
            }
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
    private readonly string _input           = input;
    private readonly NormalizeOptions _opts  = options;
    private readonly string? _remove         = additionalRemove;
    private readonly string? _replace        = additionalReplace;
    private readonly int _hash               = ComputeHash(input, options, additionalRemove, additionalReplace);

    private static int ComputeHash(
        string input, NormalizeOptions opts, string? remove, string? replace)
    {
        var h = new HashCode();
        h.Add(input,   StringComparer.Ordinal);
        h.Add((int)opts);
        h.Add(remove,  StringComparer.Ordinal);
        h.Add(replace, StringComparer.Ordinal);
        return h.ToHashCode();
    }

    public bool Equals(NormCacheKey other) =>
        _opts == other._opts &&
        string.Equals(_input,   other._input,   StringComparison.Ordinal) &&
        string.Equals(_remove,  other._remove,  StringComparison.Ordinal) &&
        string.Equals(_replace, other._replace, StringComparison.Ordinal);

    public override bool Equals(object? obj) => obj is NormCacheKey k && Equals(k);
    public override int GetHashCode() => _hash;
}
