namespace ExcelBatchTool.Core.Ocr;

/// <summary>
/// 選択肢 1 つ分の画素の測り方。
///
/// 箱の中(<see cref="BoxInk"/>)と、ラベルを囲む線が通るところ
/// (<see cref="RingInk"/>)を分けて測る。チェックや塗りつぶしは箱の中に出るが、
/// 丸囲みはラベルの周りに出るため、1 つの数値では区別できない。
/// </summary>
public readonly record struct MarkSample(string Label, double BoxInk, double RingInk);

public enum MarkDecision
{
    /// <summary>どれか 1 つが選ばれている。</summary>
    Selected = 0,

    /// <summary>どれも選ばれていない。</summary>
    None,

    /// <summary>選ばれているか判断できない。</summary>
    Unclear,
}

/// <summary>印の判定結果。</summary>
public sealed record MarkResult(
    MarkDecision Decision, string Label, double Confidence, string Reason)
{
    public string Text => Decision == MarkDecision.Selected ? Label : string.Empty;
}

/// <summary>
/// チェック / 塗りつぶし / 丸 / ばつ を、文字としてではなく画素で判定する。
///
/// 2F-R では、きれいなスキャンで 100%、劣化版で 66.7% だった。
/// 迷うものを無理に決めないで人へ回すのが安全側なので、
/// **1 位と 2 位の差が小さいときは「判断できない」**にする。
/// </summary>
public static class MarkClassifier
{
    /// <summary>印が付いているとみなす、いちばん低い濃さ。</summary>
    public const double MarkedThreshold = 0.12;

    /// <summary>どれも付いていないとみなす、いちばん高い濃さ。</summary>
    public const double EmptyThreshold = 0.05;

    /// <summary>自動確定してよい、1 位と 2 位の差。</summary>
    public const double ConfidentMargin = 0.08;

    public static MarkResult Classify(IReadOnlyList<MarkSample> samples)
    {
        if (samples.Count == 0)
        {
            return new MarkResult(MarkDecision.Unclear, string.Empty, 0, "選択肢がありません");
        }

        // 箱の中と丸囲みのうち、濃いほうをその選択肢の強さにする。
        var scored = samples
            .Select(sample => (sample.Label, Score: Math.Max(sample.BoxInk, sample.RingInk)))
            .OrderByDescending(entry => entry.Score)
            .ToList();

        var top = scored[0];
        var second = scored.Count > 1 ? scored[1].Score : 0;

        if (top.Score < EmptyThreshold)
        {
            return new MarkResult(
                MarkDecision.None, string.Empty, 1 - top.Score, "どれにも印がありません");
        }

        if (top.Score < MarkedThreshold)
        {
            return new MarkResult(
                MarkDecision.Unclear, top.Label, top.Score,
                "印が薄く、付いているか判断できません");
        }

        var margin = top.Score - second;
        if (margin < ConfidentMargin)
        {
            return new MarkResult(
                MarkDecision.Unclear, top.Label, margin,
                "複数の選択肢に印があるように見えます");
        }

        // 差が大きいほど確か。0.08 の差で 0.5、0.30 の差で 1.0 くらいになる目安。
        var confidence = Math.Clamp(margin / 0.3, 0, 1);
        return new MarkResult(
            MarkDecision.Selected, top.Label, confidence, $"「{top.Label}」に印があります");
    }

    /// <summary>判定結果を、確認の状態へ落とす。</summary>
    public static OcrItemStatus ToStatus(MarkResult result) => result.Decision switch
    {
        // 印は「文字が一致したか」で確かめられないので、
        // はっきり差がついたものだけを自動確定にする。
        MarkDecision.Selected when result.Confidence >= 0.9 => OcrItemStatus.AutoAccepted,
        MarkDecision.None => OcrItemStatus.NeedsReview,
        MarkDecision.Unclear => OcrItemStatus.NeedsReview,
        _ => OcrItemStatus.NeedsReview,
    };
}
