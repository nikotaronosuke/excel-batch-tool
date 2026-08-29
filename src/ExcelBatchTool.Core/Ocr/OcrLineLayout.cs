namespace ExcelBatchTool.Core.Ocr;

/// <summary>
/// bbox から「ページの中の行」を組み立てる。
///
/// OCR は行の位置を勝手に並べてくれないので、縦の重なりで同じ行にまとめ、
/// 行の中は左から右へ並べる。ページ番号と行番号を失わないことがここの目的
/// (Phase 2F-A の「ページ / 行 / 内容」と同じ形へ合流させる)。
/// </summary>
public static class OcrLineLayout
{
    /// <summary>同じ行とみなす縦のずれ(行の高さに対する割合)。</summary>
    private const double SameLineTolerance = 0.6;

    /// <summary>この幅より離れていたら、行の中でも区切って空白を入れる。</summary>
    private const double SpaceGapRatio = 0.5;

    public sealed record Line(int LineNumber, IReadOnlyList<int> RegionIndexes);

    /// <summary>領域を行へまとめる。戻り値は元の配列に対する添字。</summary>
    public static IReadOnlyList<Line> BuildLines(IReadOnlyList<OcrRawLine> regions)
    {
        if (regions.Count == 0)
        {
            return [];
        }

        var ordered = Enumerable.Range(0, regions.Count)
            .OrderBy(index => regions[index].Box.CenterY)
            .ThenBy(index => regions[index].Box.X)
            .ToList();

        var lines = new List<List<int>>();
        var current = new List<int> { ordered[0] };
        var currentCenter = regions[ordered[0]].Box.CenterY;
        var currentHeight = regions[ordered[0]].Box.Height;

        foreach (var index in ordered.Skip(1))
        {
            var box = regions[index].Box;
            var reference = Math.Max(Math.Min(currentHeight, box.Height), 1);

            if (Math.Abs(box.CenterY - currentCenter) <= reference * SameLineTolerance)
            {
                current.Add(index);

                // 行の代表値は、まとめた領域の平均で更新する(1 件目に引きずられない)。
                currentCenter = current.Average(i => regions[i].Box.CenterY);
                currentHeight = current.Average(i => regions[i].Box.Height);
                continue;
            }

            lines.Add(current);
            current = [index];
            currentCenter = box.CenterY;
            currentHeight = box.Height;
        }

        lines.Add(current);

        return lines
            .Select((line, order) => new Line(
                order + 1,
                line.OrderBy(index => regions[index].Box.X).ToList()))
            .ToList();
    }
}
