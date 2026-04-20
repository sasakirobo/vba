namespace xlDuckDb;

/// <summary>
/// 個別パラメータに付けられた正規化プレフィックスを解析する。
///
/// ━━ 構文 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
///
///   @NORM:&lt;flags&gt;[;REMOVE:&lt;削除文字&gt;][;REPLACE:&lt;元&gt;,&lt;先&gt;|&lt;元&gt;,&lt;先&gt;]:&lt;値&gt;
///
///   &lt;flags&gt;
///     NormalizeOptions のビットフラグ整数（0〜511）。
///
///   ;REMOVE:&lt;削除文字&gt;  （省略可）
///     このパラメータから削除したい文字を列挙した文字列。
///     例: ;REMOVE:①②③
///
///   ;REPLACE:&lt;元&gt;,&lt;先&gt;|&lt;元&gt;,&lt;先&gt;  （省略可）
///     "|" 区切りの置換ペア。値に "|" や "," を含む場合は \| \, でエスケープ。
///     例: ;REPLACE:㈱,株式会社|㈲,有限会社
///
///   :&lt;値&gt;
///     バインドする文字列。セクション群の後の最初の ":" 以降すべてが値。
///     → 値に ";" や "|" が含まれてもパース済みセクション後なので安全。
///     → 値に ":" を含めたい場合はセル参照を使うこと（リテラルでは制限あり）。
///
/// ━━ 例 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
///
///   "@NORM:511:山田　太郎"
///     → flags=511, 値="山田　太郎" に全正規化を適用
///
///   "@NORM:511;REMOVE:①②③:株式会社①"
///     → flags=511, 削除="①②③", 値="株式会社①" → "株式会社"
///
///   "@NORM:511;REPLACE:㈱,株式会社|㈲,有限会社:㈱山田商事"
///     → flags=511, 置換あり, 値="㈱山田商事" → "株式会社山田商事"
///
///   "@NORM:511;REMOVE:①;REPLACE:㈱,株式会社:㈱テスト①"
///     → flags=511, 削除="①", 置換あり, 値="㈱テスト①" → "株式会社テスト"
///
///   "@NORM:0:そのまま"
///     → flags=0（正規化なし）, 値="そのまま" をそのままバインド
///
///   "@NORM:511:"
///     → 値が空 → DBNull.Value をバインド
///
/// ━━ 注意 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
///   - プレフィックスなしの通常文字列はそのまま通過する。
///   - セル参照（ExcelReference）はこのパーサーを通らない。
///   - 値に ":" を含めたい場合はリテラルでは表現できないため、
///     セル参照（例: A1 に "21:30" を入力）を使うこと。
/// </summary>
internal static class NormParam
{
    private const string Prefix = "@NORM:";

    // ── 置換エスケープ用プレースホルダー ────────────────────────────────
    private const string PipePlaceholder  = "\x00PIPE\x00";
    private const string CommaPlaceholder = "\x00COMMA\x00";

    /// <summary>
    /// 文字列引数を検査し、@NORM: プレフィックスがあれば正規化を適用した値を返す。
    /// プレフィックスなしはそのまま返す。
    /// </summary>
    internal static object Parse(string raw)
    {
        if (!raw.StartsWith(Prefix, StringComparison.Ordinal))
            return raw;

        // "@NORM:" を除いた残り: "<flags>[;REMOVE:...][;REPLACE:...]:<値>"
        var body = raw[Prefix.Length..];

        // ── 値の切り出し ─────────────────────────────────────────────────
        // セクション群（flags と ;XXX: セクション）は ';' で区切られる。
        // 値は最後のセクションの ":" 以降に来る。
        // アルゴリズム:
        //   1. ":" を左から探す
        //   2. その ":" の前が ";REMOVE:" や ";REPLACE:" のセクション終端でなければ
        //      そこが値の開始点
        //   3. セクションを解析した後に残った ":" 以降が値
        //
        // 実装: セクション部分（;で始まる部分）を認識しながら最初の値用 ":" を探す

        var (sectionPart, value) = SplitSectionsAndValue(body);

        // ── セクション解析 ───────────────────────────────────────────────
        var sections = sectionPart.Split(';', StringSplitOptions.RemoveEmptyEntries);

        NormalizeOptions options = NormalizeOptions.None;
        string? removeSpec  = null;
        string? replaceSpec = null;

        for (var i = 0; i < sections.Length; i++)
        {
            var sec = sections[i].Trim();

            if (i == 0)
            {
                // 最初のセクション = flags
                if (int.TryParse(sec, out var flagVal))
                    options = (NormalizeOptions)flagVal;
            }
            else if (sec.StartsWith("REMOVE:", StringComparison.OrdinalIgnoreCase))
            {
                removeSpec = sec["REMOVE:".Length..];
            }
            else if (sec.StartsWith("REPLACE:", StringComparison.OrdinalIgnoreCase))
            {
                replaceSpec = ParseReplaceSection(sec["REPLACE:".Length..]);
            }
        }

        // ── 値の正規化とバインド ─────────────────────────────────────────
        if (value.Length == 0)
            return DBNull.Value; // 空値 → NULL バインド

        return options == NormalizeOptions.None && removeSpec is null && replaceSpec is null
            ? value
            : TextNormalizer.Normalize(value, options, removeSpec, replaceSpec);
    }

    // ── 内部ヘルパー ──────────────────────────────────────────────────────

    /// <summary>
    /// body を「セクション部分」と「値」に分割する。
    ///
    /// body = "&lt;flags&gt;[;REMOVE:...][;REPLACE:...]:&lt;値&gt;"
    ///
    /// ":" を左から走査し、直前が既知のセクションキーワードで終わっていない
    /// 最初の ":" を値の開始点とする。
    ///
    /// 具体的には:
    ///   - ";REMOVE:" や ";REPLACE:" の中の ":" はセクションヘッダの一部
    ///   - それ以外の ":" が値の区切り
    /// </summary>
    private static (string SectionPart, string Value) SplitSectionsAndValue(string body)
    {
        // セクションキーワード（大文字で保持）
        ReadOnlySpan<string> keywords = ["REMOVE:", "REPLACE:"];

        var i = 0;
        while (i < body.Length)
        {
            var colon = body.IndexOf(':', i);
            if (colon < 0)
            {
                // ":" がない → 値なし（不正または "@NORM:511" 形式）
                return (body, string.Empty);
            }

            // この ":" がセクションヘッダの一部かチェック
            // 直前のテキストを見て "REMOVE" や "REPLACE" で終わっているか確認
            var isKeywordColon = false;
            foreach (var kw in keywords)
            {
                // kw は "REMOVE:" や "REPLACE:" → kw[..^1] = "REMOVE" / "REPLACE"
                var kwNoColon = kw[..^1];
                if (colon >= kwNoColon.Length)
                {
                    var candidate = body[(colon - kwNoColon.Length)..colon];
                    if (string.Equals(candidate, kwNoColon, StringComparison.OrdinalIgnoreCase))
                    {
                        isKeywordColon = true;
                        break;
                    }
                }
            }

            if (!isKeywordColon)
            {
                // この ":" が値の区切り
                return (body[..colon], body[(colon + 1)..]);
            }

            i = colon + 1;
        }

        return (body, string.Empty);
    }

    /// <summary>
    /// REPLACE セクションの値 ("元,先|元,先") を
    /// TextNormalizer が受け取る "\n" 区切り形式に変換する。
    /// \| \, でエスケープされた文字を処理する。
    /// </summary>
    private static string ParseReplaceSection(string raw)
    {
        var escaped = raw
            .Replace(@"\|", PipePlaceholder)
            .Replace(@"\,", CommaPlaceholder);

        var pairs = escaped.Split('|');
        var sb    = new System.Text.StringBuilder();

        foreach (var pair in pairs)
        {
            var restored = pair
                .Replace(PipePlaceholder,  "|")
                .Replace(CommaPlaceholder, ",")
                .Trim();

            if (restored.Length == 0) continue;
            if (sb.Length > 0) sb.Append('\n');
            sb.Append(restored);
        }

        return sb.ToString();
    }
}
