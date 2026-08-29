namespace PdfBench;

public enum FusedStatus
{
    AutoAccepted,
    NeedsReview,
}

public sealed record FusedRegion(int Page, string Text, double Score, FusedStatus Status, string Reason);

/// <summary>統合方式の候補。どれが良いかは GT の完全一致率で決める(見た目の説得力では決めない)。</summary>
public sealed record FusionStrategy(string Name, Func<DualRegion, double, FusedRegion> Fuse);

public static class Fusion
{
    /// <summary>
    /// 文字種の割合。英数字・記号中心なら多言語 rec、かな漢字中心なら日本語 rec が
    /// 得意、というのが 2F-R の所見。判定には「両モデルが読んだ文字」を両方使う
    /// (片方だけを見ると、その片方が壊れたときに文字種判定ごと外れる)。
    /// </summary>
    public static double AsciiRatio(string a, string b)
    {
        var text = TextMetrics.Strip(a) + TextMetrics.Strip(b);
        if (text.Length == 0)
        {
            return 0;
        }

        var ascii = text.Count(c => c < 0x80);
        return (double)ascii / text.Length;
    }

    public static bool Agree(DualRegion region)
        => TextMetrics.Strip(region.MultiText) == TextMetrics.Strip(region.JapanText)
            && TextMetrics.Strip(region.MultiText).Length > 0;

    public static IReadOnlyList<FusionStrategy> All { get; } =
    [
        new("multi-only", (r, t) => new FusedRegion(
            r.Page, r.MultiText, r.MultiScore,
            r.MultiScore >= t ? FusedStatus.AutoAccepted : FusedStatus.NeedsReview, "multi")),

        new("japan-only", (r, t) => new FusedRegion(
            r.Page, r.JapanText, r.JapanScore,
            r.JapanScore >= t ? FusedStatus.AutoAccepted : FusedStatus.NeedsReview, "japan")),

        new("higher-score", (r, t) =>
        {
            var useMulti = r.MultiScore >= r.JapanScore;
            var text = useMulti ? r.MultiText : r.JapanText;
            var score = Math.Max(r.MultiScore, r.JapanScore);
            return new FusedRegion(
                r.Page, text, score,
                score >= t ? FusedStatus.AutoAccepted : FusedStatus.NeedsReview,
                useMulti ? "multi" : "japan");
        }),

        // 一致したときだけ自動確定し、割れたら必ず人間へ回す(最も保守的)。
        new("agree-only", (r, t) =>
        {
            if (Agree(r))
            {
                var score = Math.Min(r.MultiScore, r.JapanScore);
                return new FusedRegion(
                    r.Page, r.MultiText, score,
                    score >= t ? FusedStatus.AutoAccepted : FusedStatus.NeedsReview, "agree");
            }

            return new FusedRegion(r.Page, r.MultiText, 0, FusedStatus.NeedsReview, "disagree");
        }),

        // 文字種だけで選ぶ(一致は見ない)。
        new("charclass", (r, t) =>
        {
            var useMulti = AsciiRatio(r.MultiText, r.JapanText) >= 0.5;
            var text = useMulti ? r.MultiText : r.JapanText;
            var score = useMulti ? r.MultiScore : r.JapanScore;
            return new FusedRegion(
                r.Page, text, score,
                score >= t ? FusedStatus.AutoAccepted : FusedStatus.NeedsReview,
                useMulti ? "multi" : "japan");
        }),

        // 一致 → 自動確定候補。割れた → 文字種で選ぶが、必ず人間へ回す。
        new("agree-then-charclass", (r, t) =>
        {
            if (Agree(r))
            {
                var score = Math.Min(r.MultiScore, r.JapanScore);
                return new FusedRegion(
                    r.Page, r.MultiText, score,
                    score >= t ? FusedStatus.AutoAccepted : FusedStatus.NeedsReview, "agree");
            }

            var useMulti = AsciiRatio(r.MultiText, r.JapanText) >= 0.5;
            return new FusedRegion(
                r.Page,
                useMulti ? r.MultiText : r.JapanText,
                useMulti ? r.MultiScore : r.JapanScore,
                FusedStatus.NeedsReview,
                useMulti ? "disagree-multi" : "disagree-japan");
        }),

        // 一致 → 自動確定候補。割れた → 文字種で選び、その担当モデルが十分自信を
        // 持っていれば自動確定してよい、という強気版。誤確定が増えないかを見る。
        new("agree-or-confident-owner", (r, t) =>
        {
            if (Agree(r))
            {
                var score = Math.Min(r.MultiScore, r.JapanScore);
                return new FusedRegion(
                    r.Page, r.MultiText, score,
                    score >= t ? FusedStatus.AutoAccepted : FusedStatus.NeedsReview, "agree");
            }

            var useMulti = AsciiRatio(r.MultiText, r.JapanText) >= 0.5;
            var ownerText = useMulti ? r.MultiText : r.JapanText;
            var ownerScore = useMulti ? r.MultiScore : r.JapanScore;
            return new FusedRegion(
                r.Page, ownerText, ownerScore,
                ownerScore >= t ? FusedStatus.AutoAccepted : FusedStatus.NeedsReview,
                useMulti ? "disagree-multi" : "disagree-japan");
        }),

        // 実測で分かったこと: 多言語 rec の失敗は「日本語の文字を静かに落とす」形が多い。
        // ならば「割れたときは長いほうが正しい」はずで、それを規則にできるか試す。
        new("agree-then-longer", (r, t) =>
        {
            if (Agree(r))
            {
                var score = Math.Min(r.MultiScore, r.JapanScore);
                return new FusedRegion(
                    r.Page, r.MultiText, score,
                    score >= t ? FusedStatus.AutoAccepted : FusedStatus.NeedsReview, "agree");
            }

            var multi = TextMetrics.Strip(r.MultiText);
            var japan = TextMetrics.Strip(r.JapanText);
            var useMulti = multi.Length > japan.Length;
            return new FusedRegion(
                r.Page,
                useMulti ? r.MultiText : r.JapanText,
                useMulti ? r.MultiScore : r.JapanScore,
                FusedStatus.NeedsReview,
                useMulti ? "disagree-multi" : "disagree-japan");
        }),

        // 数字・記号だけの領域(コード・金額・日付)は多言語 rec、
        // 日本語が 1 文字でも混ざる領域は日本語 rec。閾値を厳しくした版。
        new("agree-then-charclass-strict", (r, t) =>
        {
            if (Agree(r))
            {
                var score = Math.Min(r.MultiScore, r.JapanScore);
                return new FusedRegion(
                    r.Page, r.MultiText, score,
                    score >= t ? FusedStatus.AutoAccepted : FusedStatus.NeedsReview, "agree");
            }

            var useMulti = AsciiRatio(r.MultiText, r.JapanText) >= 0.999;
            return new FusedRegion(
                r.Page,
                useMulti ? r.MultiText : r.JapanText,
                useMulti ? r.MultiScore : r.JapanScore,
                FusedStatus.NeedsReview,
                useMulti ? "disagree-multi" : "disagree-japan");
        }),

        // 割れたら日本語 rec を既定にする(日本語の帳票が対象なので)。
        new("agree-then-japan", (r, t) =>
        {
            if (Agree(r))
            {
                var score = Math.Min(r.MultiScore, r.JapanScore);
                return new FusedRegion(
                    r.Page, r.MultiText, score,
                    score >= t ? FusedStatus.AutoAccepted : FusedStatus.NeedsReview, "agree");
            }

            return new FusedRegion(
                r.Page, r.JapanText, r.JapanScore, FusedStatus.NeedsReview, "disagree-japan");
        }),

        // 片方の読みがもう片方を丸ごと含むなら、短いほうが文字を落としたとみなす。
        // 含まない(＝別物を読んでいる)ときだけ文字種で選ぶ。
        new("agree-then-contains-then-charclass", (r, t) =>
        {
            if (Agree(r))
            {
                var score = Math.Min(r.MultiScore, r.JapanScore);
                return new FusedRegion(
                    r.Page, r.MultiText, score,
                    score >= t ? FusedStatus.AutoAccepted : FusedStatus.NeedsReview, "agree");
            }

            var multi = TextMetrics.Strip(r.MultiText);
            var japan = TextMetrics.Strip(r.JapanText);
            bool useMulti;
            if (multi.Length > 0 && japan.Contains(multi, StringComparison.Ordinal))
            {
                useMulti = false;
            }
            else if (japan.Length > 0 && multi.Contains(japan, StringComparison.Ordinal))
            {
                useMulti = true;
            }
            else
            {
                useMulti = AsciiRatio(r.MultiText, r.JapanText) >= 0.999;
            }

            return new FusedRegion(
                r.Page,
                useMulti ? r.MultiText : r.JapanText,
                useMulti ? r.MultiScore : r.JapanScore,
                FusedStatus.NeedsReview,
                useMulti ? "disagree-multi" : "disagree-japan");
        }),

        // 担当モデルで読み、もう一方の自信も足切りに使う(両方の較正を使う)。
        new("charclass-both-gated", (r, t) =>
        {
            var useMulti = AsciiRatio(r.MultiText, r.JapanText) >= 0.5;
            var text = useMulti ? r.MultiText : r.JapanText;
            var owner = useMulti ? r.MultiScore : r.JapanScore;
            var both = Math.Min(r.MultiScore, r.JapanScore);
            var status = owner >= t && (Agree(r) || both >= t)
                ? FusedStatus.AutoAccepted
                : FusedStatus.NeedsReview;
            return new FusedRegion(r.Page, text, owner, status, useMulti ? "multi" : "japan");
        }),
    ];
}
