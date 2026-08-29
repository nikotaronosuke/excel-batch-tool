namespace ExcelBatchTool.Core.Ocr;

/// <summary>
/// 項目の種類ごとに、自動確定してよいかを最後に判定する。
///
/// なぜ要るか: 二重読みの「両方が一致した」という根拠は、
/// **2 つのモデルが同じ間違いをしない**ことを前提にしている。
/// ところが採用した 2 モデルはどちらも PP-OCR v4 系で、字形の取り違えは共通して起きる。
/// つまり 0 と O のような取り違えに対しては、一致は証拠にならない。
///
/// 実測(Phase 2F-B2 / PdfScanBench form)。誤って自動確定した 4 件は
/// **すべて同じ形**だった:
///
/// | ページ | 項目 | 読み | 正 | 自信 |
/// |---|---|---|---|---|
/// | p1 | 店舗コード | SO01-24 | S001-24 | 98.4% |
/// | p5 | 店舗コード | SO05-91 | S005-91 | 98.8% |
/// | p1(傾き) | 店舗コード | SO01-24 | S001-24 | 98.4% |
/// | p2(傾き) | 店舗コード | SO02-35 | S002-35 | 98.5% |
///
/// 自信は 98% を超えており、閾値を上げても消えない(上げると正しい読みも
/// まとめて人へ回るだけで、誤確定は残る)。閾値ではなく**形**で判断する。
///
/// 決まり:
/// - コード: 英字と数字が混ざりうるので、取り違えやすい字が 1 つでも入っていたら
///   自動確定しない。人が元のページと見比べて決める
/// - 数量・金額: 数字と区切りだけのはずなので、そこへ英字が現れたら形が壊れている。
///   自動確定しない(数字だけの読みはそのまま自動確定してよい)
/// - 文章・選択肢: そのまま(2F-B1.1 の実測で誤確定 0)
///
/// 自動確定の割合は下がる。それは意図した取引で、
/// 「自動確定を増やすために誤確定を許容しない」という決まりに従っている。
/// </summary>
public static class FieldAutoAcceptPolicy
{
    /// <summary>
    /// 相手と取り違えやすい字。片方だけ見ても、もう片方だったのか判別できない。
    /// 0/O・1/l/I・5/S・8/B・2/Z・6/G・9/q のペアをまとめて挙げてある。
    /// </summary>
    private const string Confusable = "0O1lI5S8B2Z6G9qg";

    /// <summary>桁区切りを上下逆に読んだときに出る記号(アポストロフィの類)。</summary>
    private static readonly char[] RotatedSeparators =
        ['\u0027', '`', '\u2019', '\u02BC'];

    /// <summary>数量・金額として自然な字。これ以外が出たら形が壊れている。</summary>
    private const string NumberShape = "0123456789,.-+()%¥$ 　";

    public static bool CanAutoAccept(FormFieldKind kind, string text) => kind switch
    {
        FormFieldKind.Code => !string.IsNullOrEmpty(text) && !HasConfusable(text),
        FormFieldKind.NumberLike => !string.IsNullOrEmpty(text) && IsNumberShaped(text),
        _ => true,
    };

    public static bool HasConfusable(string text)
    {
        foreach (var c in text)
        {
            if (Confusable.Contains(c))
            {
                return true;
            }
        }

        return false;
    }

    public static bool IsNumberShaped(string text)
    {
        foreach (var c in text)
        {
            if (!NumberShape.Contains(c))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// 上下逆に読んでも別の正しい数に見えてしまう字だけでできているか。
    ///
    /// 0・1・8 は上下逆でも同じ形、6 と 9 は互いに入れ替わる。
    /// つまり「90」を上下逆に読むと「06」になり、**どちらも数として正しい**。
    /// 実測で残った誤確定はすべてこの形だった:
    ///   「90」→「06」(自信 99.1〜99.8%)、「99」→「66」(自信 99.2〜99.7%)。
    /// 切り出しが上下逆のまま両モデルへ渡るので、一致も自信も根拠にならず、
    /// 列の中の文字の種類も同じ(どちらも数字)なので他の安全弁にも掛からない。
    ///
    /// もう 1 つ、**桁区切りが上下逆になった跡**も見る。カンマを上下逆に読むと
    /// アポストロフィになるため、「99,660」が「099'66」として出てくる
    /// (実測。自信 99.0% で自動確定していた)。数字の並びだけを見ると
    /// 列の中で浮かないので、記号のほうで気づく。
    /// </summary>
    public static bool IsUpsideDownAmbiguous(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return false;
        }

        // 桁区切りが上下逆になった跡。数量・金額にこの記号は出てこない。
        if (text.IndexOfAny(RotatedSeparators) >= 0)
        {
            return true;
        }

        foreach (var c in text)
        {
            if (c is not ('0' or '1' or '6' or '8' or '9'))
            {
                return false;
            }
        }

        return true;
    }

    public const string UpsideDownReason =
        "上下逆に読んでも数として成り立つ値です(6 と 9 など)。"
        + "元のページと見比べてください。";

    /// <summary>自動確定を見送った理由。確認画面にそのまま出す。</summary>
    public static string ReasonFor(FormFieldKind kind) => kind switch
    {
        FormFieldKind.Code =>
            "コードに、取り違えやすい字(0 と O など)が含まれます。元のページと見比べてください。",
        FormFieldKind.NumberLike =>
            "数量・金額に、数字以外の字が含まれます。元のページと見比べてください。",
        _ => string.Empty,
    };
}
