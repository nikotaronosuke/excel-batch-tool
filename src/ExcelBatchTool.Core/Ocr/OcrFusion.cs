using ExcelBatchTool.Core.Pdf;

namespace ExcelBatchTool.Core.Ocr;

/// <summary>統合の結果。どちらのモデルを採ったか、なぜその状態かも持つ。</summary>
public sealed record OcrFusionResult(
    string Text, double Confidence, OcrItemStatus Status, string Reason);

/// <summary>
/// 2 つの認識モデルの結果を 1 つにまとめる。
///
/// この規則は「良さそうだから」ではなく、架空 fixture と Ground Truth で 11 方式 ×
/// 3 閾値を機械照合して選んだ(research/PdfFeasibility/PdfFusionBench)。
/// 最重要の指標は完全一致率ではなく、**自動確定にしたのに間違っていた件数**。
///
/// 実測(帳票 120 ページ / 813 項目):
///
/// | 項目 | 多言語のみ | 日本語のみ | この規則 |
/// |---|---|---|---|
/// | 店舗コード(英数字) | 98.3% | 30.8% | 98.3% |
/// | 備考(かな漢字) | 0.0% | 72.0% | 72.0% |
/// | 担当者(氏名) | 49.2% | 79.2% | 79.2% |
/// | 全体 | 74.9% | 78.6% | **87.6%** |
///
/// どちらか一方では成立しない(多言語はかな漢字を静かに落とし、
/// 日本語は英数字コードを崩す)。項目ごとに良いほうへ寄る。
/// </summary>
public static class OcrFusion
{
    public const string MultiEngineName = "多言語";

    public const string JapanEngineName = "日本語";

    /// <summary>
    /// 自動確定にしてよい自信の下限。0.90 / 0.95 / 0.98 を実測し、
    /// 0.98 だけが誤確定をほぼ 0(813 項目中 1 件)に抑えられた。
    /// </summary>
    public const double AutoAcceptThreshold = 0.98;

    /// <summary>これを下回るものは、文字として読めていないものとして扱う。</summary>
    public const double UnreadableThreshold = 0.30;

    /// <summary>
    /// 認識器は NaN / Infinity を返すことがある(空に近い切り出しで実測)。
    /// そのまま比べると「自信 = 無限大」で閾値を通ってしまうので、必ず 0 に倒す。
    /// </summary>
    public static double Finite(double score) => double.IsFinite(score) ? score : 0;

    /// <summary>
    /// 空白を除いてすべて ASCII か。
    ///
    /// 多言語モデルを使うのは「数字・英字・記号だけの領域」に限る。
    /// 少しでも日本語が混ざる領域は日本語モデルに任せる。
    /// 割合を 5 割で切る版も測ったが、「金額:4,917,087円」のように
    /// ASCII が多数でも末尾に日本語が付く領域で多言語モデルを選んでしまい、
    /// 日本語の文章での完全一致が 100% → 75% に落ちた。
    ///
    /// 判定には両モデルが読んだ文字を両方使う
    /// (片方だけを見ると、その片方が壊れたときに文字種の判定ごと外れる)。
    /// </summary>
    internal static bool IsAllAscii(string multi, string japan)
    {
        var counted = 0;

        foreach (var character in multi + japan)
        {
            if (char.IsWhiteSpace(character))
            {
                continue;
            }

            counted++;
            if (character >= 0x80)
            {
                return false;
            }
        }

        return counted > 0;
    }

    public static OcrFusionResult Fuse(OcrRawLine line)
    {
        var multi = PdfTextNormalization.Normalize(line.MultiText);
        var japan = PdfTextNormalization.Normalize(line.JapanText);
        var multiScore = Finite(line.MultiScore);
        var japanScore = Finite(line.JapanScore);

        var agree = multi.Length > 0 && string.Equals(multi, japan, StringComparison.Ordinal);
        var useMulti = IsAllAscii(multi, japan);
        var text = useMulti ? multi : japan;
        var owner = useMulti ? multiScore : japanScore;
        var ownerName = useMulti ? MultiEngineName : JapanEngineName;

        if (text.Length == 0)
        {
            return new OcrFusionResult(
                string.Empty, Math.Max(multiScore, japanScore), OcrItemStatus.Unreadable,
                "文字として読み取れませんでした");
        }

        if (agree)
        {
            // 2 つのモデルが同じ文字を出したときだけ、自動確定の候補にする。
            var confidence = Math.Min(multiScore, japanScore);
            return new OcrFusionResult(
                multi,
                confidence,
                confidence >= AutoAcceptThreshold
                    ? OcrItemStatus.AutoAccepted
                    : OcrItemStatus.NeedsReview,
                confidence >= AutoAcceptThreshold
                    ? "2 つのモデルが一致"
                    : "2 つのモデルは一致したが自信が低い");
        }

        if (Math.Max(multiScore, japanScore) < UnreadableThreshold)
        {
            return new OcrFusionResult(
                text, owner, OcrItemStatus.Unreadable, "どちらのモデルも読み取れていません");
        }

        // 割れた場合は、得意な側の文字を採るが**必ず人へ回す**。
        return new OcrFusionResult(
            text, owner, OcrItemStatus.NeedsReview, $"2 つのモデルで違う({ownerName}を表示)");
    }
}
