namespace ExcelBatchTool.Core.Ocr;

/// <summary>ページの位置ずれを吸収する平行移動。</summary>
public readonly record struct FormOffset(double X, double Y)
{
    public static readonly FormOffset None = new(0, 0);

    public OcrBox Apply(OcrBox box) => new(box.X + X, box.Y + Y, box.Width, box.Height);
}

/// <summary>1 ページ分の 1 項目の読み取り結果。</summary>
public sealed record FormFieldReading(
    string Name,
    FormFieldKind Kind,
    string Text,
    double Confidence,
    OcrBox Area,
    bool WasFound,
    string Reason);

/// <summary>
/// 帳票の項目を、全ページ OCR の結果から選び出す。
///
/// **領域だけを切り出して OCR にかけることはしない。** 2F-R で、切り出して読む方式は
/// 全ページ OCR より悪かった(clean 98.6% → 87.3%、劣化版 87.7% → 54.2%)。
/// ここでの領域は「全ページの読み取り結果のうち、この項目に属するものを選ぶため」に使う。
///
/// いちばん大事なのは、**読めなかった項目を消さないこと**。該当する読み取りが
/// 1 つも無くても、その項目は「見つからなかった」として必ず 1 件返す。
/// </summary>
public static class FormFieldExtractor
{
    /// <summary>領域に入っているとみなす、重なりの割合。</summary>
    private const double MinimumOverlap = 0.35;

    /// <summary>位置合わせで動かしてよい幅(領域の高さに対する割合)。</summary>
    private const double MaximumAnchorShift = 3.0;

    /// <summary>
    /// 位置合わせ。指定した手がかりの文字が実際にどこにあったかを見て、
    /// ページ全体のずれを平行移動として求める。
    /// 手がかりが見つからなければ動かさない(勝手に大きく動かさない)。
    /// </summary>
    public static FormOffset FindOffset(
        FormTemplate template,
        IReadOnlyList<OcrRawLine> lines,
        Func<OcrRawLine, string> read)
    {
        if (template.Anchors.Count == 0 || lines.Count == 0)
        {
            return FormOffset.None;
        }

        var shifts = new List<(double X, double Y)>();

        foreach (var anchor in template.Anchors)
        {
            var expected = Pdf.PdfTextNormalization.Normalize(anchor.Text);
            if (expected.Length == 0)
            {
                continue;
            }

            // 同じ文字が複数あるときは、指定した位置にいちばん近いものを使う。
            var best = lines
                .Where(line => Pdf.PdfTextNormalization.Normalize(read(line)).Contains(
                    expected, StringComparison.Ordinal))
                .OrderBy(line => Distance(line.Box, anchor.Area))
                .FirstOrDefault();

            if (best is null)
            {
                continue;
            }

            var limit = Math.Max(anchor.Area.Height, 1) * MaximumAnchorShift;
            var dx = best.Box.X - anchor.Area.X;
            var dy = best.Box.Y - anchor.Area.Y;

            if (Math.Abs(dx) <= limit && Math.Abs(dy) <= limit)
            {
                shifts.Add((dx, dy));
            }
        }

        if (shifts.Count == 0)
        {
            return FormOffset.None;
        }

        // 外れ値に引きずられないよう中央値を使う。
        return new FormOffset(Median(shifts.Select(s => s.X)), Median(shifts.Select(s => s.Y)));
    }

    /// <summary>
    /// 1 ページ分の項目を読み取る。
    /// **指定した項目の数と、返す件数は必ず同じ**(読めなかったものも返す)。
    /// </summary>
    public static IReadOnlyList<FormFieldReading> Read(
        FormTemplate template,
        IReadOnlyList<OcrRawLine> lines,
        FormOffset offset,
        Func<OcrRawLine, (string Text, double Confidence)> read)
    {
        var readings = new List<FormFieldReading>(template.Fields.Count);

        foreach (var field in template.Fields)
        {
            var area = offset.Apply(field.Area);

            if (field.Kind == FormFieldKind.Choice)
            {
                // 印は文字として読ませない。別の経路(画素の判定)で埋める。
                readings.Add(new FormFieldReading(
                    field.Name, field.Kind, string.Empty, 0, area, WasFound: false,
                    "印の判定はこのあとで行います"));
                continue;
            }

            var inside = lines
                .Where(line => Overlap(line.Box, area) >= MinimumOverlap)
                .OrderBy(line => line.Box.CenterY)
                .ThenBy(line => line.Box.X)
                .ToList();

            if (inside.Count == 0)
            {
                // ここが要。読み取り領域が 1 つも見つからなくても、項目は消さない。
                readings.Add(new FormFieldReading(
                    field.Name, field.Kind, string.Empty, 0, area, WasFound: false,
                    "この場所から文字を読み取れませんでした"));
                continue;
            }

            var parts = new List<string>();
            var confidence = 1.0;
            foreach (var line in inside)
            {
                var (text, score) = read(line);
                if (text.Length > 0)
                {
                    parts.Add(text);
                }

                confidence = Math.Min(confidence, score);
            }

            var joined = string.Join(" ", parts).Trim();

            readings.Add(joined.Length == 0
                ? new FormFieldReading(
                    field.Name, field.Kind, string.Empty, 0, area, WasFound: false,
                    "この場所から文字を読み取れませんでした")
                : new FormFieldReading(
                    field.Name, field.Kind, joined, confidence,
                    ScanTableBuilder.Union(inside.Select(line => line.Box)),
                    WasFound: true,
                    "指定した場所から読み取りました"));
        }

        return readings;
    }

    /// <summary>2 つの領域の重なりを、小さいほうに対する割合で返す。</summary>
    internal static double Overlap(OcrBox box, OcrBox area)
    {
        var width = Math.Min(box.Right, area.Right) - Math.Max(box.X, area.X);
        var height = Math.Min(box.Bottom, area.Bottom) - Math.Max(box.Y, area.Y);
        if (width <= 0 || height <= 0)
        {
            return 0;
        }

        var smaller = Math.Min(box.Width * box.Height, area.Width * area.Height);
        return smaller <= 0 ? 0 : width * height / smaller;
    }

    private static double Distance(OcrBox a, OcrBox b)
        => Math.Abs(a.X - b.X) + Math.Abs(a.Y - b.Y);

    private static double Median(IEnumerable<double> values)
    {
        var sorted = values.OrderBy(value => value).ToList();
        return sorted.Count == 0 ? 0 : sorted[sorted.Count / 2];
    }
}
