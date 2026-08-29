namespace ExcelBatchTool.Core.Ocr;

/// <summary>
/// 同じ様式の帳票を何ページも読むとき、項目ごとに「その項目らしい形」を
/// 全ページの読み取りから学び、そこから外れたページだけを人へ回す。
///
/// なぜ要るか: Phase 2F-B2 では、コードの項目に取り違えやすい字
/// (0 と O など)が 1 つでもあれば自動確定しない、という粗い決まりにした。
/// 誤確定は 0 になったが、**数字を含むコードはほぼ全部が人へ回る**。
/// 120 ページの帳票で自動確定が 46% から 32% まで落ちていた。
///
/// 店舗コードが 120 ページとも `S001-24` のような形なら、
/// その形どおりに読めたページは自動確定してよいはずで、
/// 形が違うページ(`SO01-24`)だけを人が見ればよい。
///
/// **値は決して書き換えない。** 形から推測して O を 0 に直すようなことはしない。
/// あくまで「自動確定してよいか」の判断にだけ使う。
/// 直してしまうと、人が原文と見比べる前に間違いが確定しうる。
/// </summary>
public static class FieldShapePattern
{
    /// <summary>形を決めるのに要る、読み取れたページの数。</summary>
    public const int MinimumSamples = 8;

    /// <summary>いちばん多い形がこの割合に満たなければ、形を決めない。</summary>
    public const double MajorityRatio = 0.6;

    /// <summary>
    /// 1 文字を種類へ潰した並び。桁数の違いは残す(コードは桁数も含めて様式)。
    ///   英字 = A / 数字 = 9 / 日本語 = 字 / それ以外はその文字のまま。
    /// </summary>
    public static string Of(string text)
    {
        var shape = new System.Text.StringBuilder(text.Length);
        foreach (var c in text)
        {
            shape.Append(
                char.IsAsciiDigit(c) ? '9'
                : char.IsAsciiLetter(c) ? 'A'
                : c > (char)0x7F && char.IsLetter(c) ? '字'
                : c);
        }

        return shape.ToString();
    }

    /// <summary>
    /// 全ページぶんの読みから、その項目の形を決める。決まらなければ null。
    /// </summary>
    public static string? Learn(IEnumerable<string> readings)
    {
        var shapes = readings
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .Select(Of)
            .ToList();

        if (shapes.Count < MinimumSamples)
        {
            return null;
        }

        var best = shapes.GroupBy(shape => shape)
            .OrderByDescending(group => group.Count())
            .First();

        return (double)best.Count() / shapes.Count >= MajorityRatio ? best.Key : null;
    }

    /// <summary>この読みが、学んだ形どおりか。</summary>
    public static bool Matches(string? pattern, string text)
        => pattern is not null && Of(text) == pattern;

    public const string Reason =
        "同じ項目の他のページと形が違います。元のページと見比べてください。";
}
