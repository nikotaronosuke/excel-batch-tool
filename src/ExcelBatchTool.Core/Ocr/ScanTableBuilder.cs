namespace ExcelBatchTool.Core.Ocr;

/// <summary>1 つのセルと、それを組み立てた読み取り位置。</summary>
public sealed record ScanTableCell(int Row, int Column, string Text, OcrBox Box, double Confidence)
{
    public bool IsEmpty => Text.Length == 0;

    /// <summary>
    /// 縦に離れた複数の読み取りを 1 つのセルへまとめた。
    ///
    /// 罫線が拾えず 2 行分が 1 つの区画に入ったときにこうなる。実測では
    /// 「A0017」と「A0018」が「A0017A0018」という 1 セルになり、自信 99.6% で
    /// 自動確定していた(行の区切りを間違えているので、自信も一致も当てにならない)。
    /// 区画を割り直しても解消しなかったものだけがここへ残る。
    /// </summary>
    public bool IsMerged { get; init; }
}

/// <summary>スキャンされた表を行・列へ戻した結果。</summary>
public sealed record ScanTable(
    int RowCount,
    int ColumnCount,
    IReadOnlyList<ScanTableCell> Cells,
    bool FromRulings)
{
    /// <summary>出力用の行の並び(空セルは空文字)。</summary>
    public IReadOnlyList<string[]> ToRows()
    {
        var rows = new string[RowCount][];
        for (var row = 0; row < RowCount; row++)
        {
            rows[row] = new string[ColumnCount];
            for (var column = 0; column < ColumnCount; column++)
            {
                rows[row][column] = string.Empty;
            }
        }

        foreach (var cell in Cells)
        {
            if (cell.Row < RowCount && cell.Column < ColumnCount)
            {
                rows[cell.Row][cell.Column] = cell.Text;
            }
        }

        return rows;
    }
}

/// <summary>
/// スキャンされた表を行・列へ戻す。
///
/// 罫線が印刷されているなら**罫線そのものから格子を作る**(2F-R で、表構造の
/// 学習モデルが 41 行の表で列を 1 つずらして全滅したのに対し、この方法は
/// 94.4% を出した)。罫線が無い表は、読み取り位置の x 座標のまとまりから列を決める。
///
/// どちらも「文字をつなげた 1 本の文章」にはしない。行と列を保ったまま出す。
/// </summary>
public static class ScanTableBuilder
{
    /// <summary>行として同じ高さとみなす、行の高さに対する割合。</summary>
    private const double SameRowTolerance = 0.6;

    /// <summary>列の左端が同じとみなす、文字の高さに対する割合。</summary>
    private const double SameColumnTolerance = 1.5;

    /// <summary>表とみなすのに必要な、最低限の行数と列数。</summary>
    public const int MinimumRows = 2;

    public const int MinimumColumns = 2;

    /// <summary>
    /// 罫線の格子へ読み取りを割り当てる。
    /// <paramref name="rowLines"/> / <paramref name="columnLines"/> は罫線の座標。
    /// </summary>
    public static ScanTable? FromRulings(
        IReadOnlyList<OcrRawLine> lines,
        IReadOnlyList<double> rowLines,
        IReadOnlyList<double> columnLines,
        Func<OcrRawLine, (string Text, double Confidence)> read)
    {
        if (rowLines.Count - 1 < MinimumRows || columnLines.Count - 1 < MinimumColumns)
        {
            return null;
        }

        // 外枠の罫線は拾えないことがある(細い線が画素の境目に来ると消える)。
        // 見つかった罫線の外側に文字があるなら、そこにも区切りがあったものとして足す。
        // こうしないと、いちばん左と右の列がまるごと落ちる(実測で 4 列 → 2 列になった)。
        var columns = Extend(
            columnLines, lines.Select(line => (line.Box.CenterX, line.Box.X, line.Box.Right)));
        var rows = SplitTallBands(
            Extend(rowLines, lines.Select(line => (line.Box.CenterY, line.Box.Y, line.Box.Bottom))),
            lines);

        var rowCount = rows.Count - 1;
        var columnCount = columns.Count - 1;

        var pieces = new List<(int Row, int Column, OcrRawLine Line)>();
        foreach (var line in lines)
        {
            var row = Band(rows, line.Box.CenterY);
            var column = Band(columns, line.Box.CenterX);
            if (row >= 0 && row < rowCount && column >= 0 && column < columnCount)
            {
                pieces.Add((row, column, line));
            }
        }

        return Assemble(pieces, rowCount, columnCount, read, fromRulings: true);
    }

    /// <summary>
    /// 罫線の外側に文字があるなら、その外側にも区切りを足す。
    /// 罫線から遠く離れた文字(表の外の注記など)までは拾わない。
    /// </summary>
    internal static List<double> Extend(
        IReadOnlyList<double> lines,
        IEnumerable<(double Center, double Start, double End)> extents)
    {
        var result = lines.OrderBy(value => value).ToList();
        if (result.Count < 2)
        {
            return result;
        }

        var first = result[0];
        var last = result[^1];

        // 足す区切りは、いちばん外側のセル 1 つ分までにとどめる。
        var margin = (last - first) / Math.Max(result.Count - 1, 1);

        // 罫線をまたいでいるだけの文字(枠のすぐ内側にある見出しなど)で
        // 区切りを増やさないよう、判断には**中心**を使う。
        var inside = extents
            .Where(extent => extent.Center > first - margin && extent.Center < last + margin)
            .ToList();

        if (inside.Count == 0)
        {
            return result;
        }

        if (inside.Any(extent => extent.Center < first))
        {
            result.Insert(0, inside.Min(extent => extent.Start) - 1);
        }

        if (inside.Any(extent => extent.Center > last))
        {
            result.Add(inside.Max(extent => extent.End) + 1);
        }

        return result;
    }

    /// <summary>
    /// 罫線が無い表を、読み取り位置から組み立てる。
    ///
    /// 行は縦の重なりでまとめ、列は**先頭行(見出し)の左端**を基準に決める。
    /// 見出しが取れないときは、全行の左端をまとめて列の位置を決める。
    /// </summary>
    public static ScanTable? FromAlignment(
        IReadOnlyList<OcrRawLine> lines,
        Func<OcrRawLine, (string Text, double Confidence)> read)
    {
        if (lines.Count == 0)
        {
            return null;
        }

        var rows = GroupRows(lines);
        if (rows.Count < MinimumRows)
        {
            return null;
        }

        var columnStarts = ColumnStarts(rows, lines);
        if (columnStarts.Count < MinimumColumns)
        {
            return null;
        }

        var pieces = new List<(int Row, int Column, OcrRawLine Line)>();
        for (var row = 0; row < rows.Count; row++)
        {
            foreach (var index in rows[row])
            {
                var column = NearestColumn(columnStarts, lines[index].Box.X);
                pieces.Add((row, column, lines[index]));
            }
        }

        return Assemble(pieces, rows.Count, columnStarts.Count, read, fromRulings: false);
    }

    private static ScanTable Assemble(
        List<(int Row, int Column, OcrRawLine Line)> pieces,
        int rows,
        int columns,
        Func<OcrRawLine, (string Text, double Confidence)> read,
        bool fromRulings)
    {
        var cells = new List<ScanTableCell>();

        foreach (var group in pieces.GroupBy(piece => (piece.Row, piece.Column)))
        {
            var ordered = group.OrderBy(piece => piece.Line.Box.Y)
                .ThenBy(piece => piece.Line.Box.X)
                .ToList();

            var parts = new List<string>();
            var confidence = 1.0;
            foreach (var piece in ordered)
            {
                var (text, score) = read(piece.Line);
                if (text.Length > 0)
                {
                    parts.Add(text);
                }

                confidence = Math.Min(confidence, score);
            }

            cells.Add(new ScanTableCell(
                group.Key.Row,
                group.Key.Column,
                // セル内で行が分かれていても 1 つの値にまとめる。
                string.Concat(parts),
                Union(ordered.Select(piece => piece.Line.Box)),
                parts.Count == 0 ? 0 : confidence)
            {
                IsMerged = HasVerticalGap(ordered.Select(piece => piece.Line.Box)),
            });
        }

        return new ScanTable(rows, columns, cells, fromRulings);
    }

    /// <summary>縦の重なりで行にまとめる。戻り値は元の配列に対する添字。</summary>
    internal static List<List<int>> GroupRows(IReadOnlyList<OcrRawLine> lines)
    {
        if (lines.Count == 0)
        {
            return [];
        }

        var ordered = Enumerable.Range(0, lines.Count)
            .OrderBy(index => lines[index].Box.CenterY)
            .ToList();

        var rows = new List<List<int>>();
        var current = new List<int> { ordered[0] };
        var center = lines[ordered[0]].Box.CenterY;
        var height = lines[ordered[0]].Box.Height;

        foreach (var index in ordered.Skip(1))
        {
            var box = lines[index].Box;
            var reference = Math.Max(Math.Min(height, box.Height), 1);

            if (Math.Abs(box.CenterY - center) <= reference * SameRowTolerance)
            {
                current.Add(index);
                center = current.Average(i => lines[i].Box.CenterY);
                height = current.Average(i => lines[i].Box.Height);
                continue;
            }

            rows.Add(current);
            current = [index];
            center = box.CenterY;
            height = box.Height;
        }

        rows.Add(current);
        return rows;
    }

    /// <summary>
    /// 列の左端を決める。見出しの行がいちばん素直なので、まずそれを使う。
    /// 見出しより多くの塊を持つ行があれば、そちらに合わせて足りない列を補う。
    /// </summary>
    internal static List<double> ColumnStarts(
        List<List<int>> rows, IReadOnlyList<OcrRawLine> lines)
    {
        if (rows.Count == 0)
        {
            return [];
        }

        var widest = rows.OrderByDescending(row => row.Count).First();
        var starts = widest
            .Select(index => lines[index].Box.X)
            .OrderBy(x => x)
            .ToList();

        var tolerance = Math.Max(
            widest.Average(index => lines[index].Box.Height) * SameColumnTolerance, 4);

        var merged = new List<double>();
        foreach (var start in starts)
        {
            if (merged.Count == 0 || start - merged[^1] > tolerance)
            {
                merged.Add(start);
            }
        }

        return merged;
    }

    private static int NearestColumn(List<double> starts, double x)
    {
        var best = 0;
        var bestDistance = double.MaxValue;
        for (var index = 0; index < starts.Count; index++)
        {
            var distance = Math.Abs(starts[index] - x);
            if (distance < bestDistance)
            {
                bestDistance = distance;
                best = index;
            }
        }

        return best;
    }

    /// <summary>座標がどの罫線の間にあるか。外なら -1。</summary>
    /// <summary>
    /// 縦に離れた読み取りが混じっているか。重なっていれば同じ行の続き
    /// (ふりがな等)、離れていれば別の行を巻き込んでいる。
    /// </summary>
    internal static bool HasVerticalGap(IEnumerable<OcrBox> boxes)
    {
        var ordered = boxes.OrderBy(box => box.Y).ToList();
        for (var index = 1; index < ordered.Count; index++)
        {
            if (ordered[index].Y > ordered[index - 1].Bottom)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// 罫線が拾えずに 2 行分が 1 区画へ入ったところを割り直す。
    ///
    /// 実測では、表のいちばん下のほうで細い罫線が落ち、GT の 4 行が 2 区画に
    /// 収まって「A0017A0018」のような値ができていた。区画の高さが他より明らかに
    /// 高く、その中の文字が縦に離れた塊へ分かれるときだけ、塊の間へ区切りを足す。
    ///
    /// 高さで判断するのは、セル内で折り返した長い文章を割ってしまわないため
    /// (折り返しは区画の高さを増やさない)。
    /// </summary>
    internal static List<double> SplitTallBands(
        List<double> boundaries, IReadOnlyList<OcrRawLine> lines)
    {
        if (boundaries.Count < 3)
        {
            return boundaries;
        }

        var heights = new List<double>();
        for (var index = 1; index < boundaries.Count; index++)
        {
            heights.Add(boundaries[index] - boundaries[index - 1]);
        }

        var median = heights.OrderBy(height => height).ElementAt(heights.Count / 2);
        if (median <= 0)
        {
            return boundaries;
        }

        var added = new List<double>();
        for (var index = 1; index < boundaries.Count; index++)
        {
            var top = boundaries[index - 1];
            var bottom = boundaries[index];
            if (bottom - top < median * 1.5)
            {
                continue;
            }

            // この区画に入る文字を、縦に離れた塊へ分ける。
            var inside = lines
                .Where(line => line.Box.CenterY > top && line.Box.CenterY < bottom)
                .OrderBy(line => line.Box.Y)
                .ToList();

            var clusters = new List<(double Top, double Bottom)>();
            foreach (var line in inside)
            {
                if (clusters.Count > 0 && line.Box.Y <= clusters[^1].Bottom)
                {
                    clusters[^1] = (clusters[^1].Top, Math.Max(clusters[^1].Bottom, line.Box.Bottom));
                }
                else
                {
                    clusters.Add((line.Box.Y, line.Box.Bottom));
                }
            }

            for (var cluster = 1; cluster < clusters.Count; cluster++)
            {
                added.Add((clusters[cluster - 1].Bottom + clusters[cluster].Top) / 2);
            }
        }

        if (added.Count == 0)
        {
            return boundaries;
        }

        return [.. boundaries.Concat(added).OrderBy(value => value)];
    }

    internal static int Band(IReadOnlyList<double> lines, double value)
    {
        for (var index = 0; index + 1 < lines.Count; index++)
        {
            if (value >= lines[index] && value < lines[index + 1])
            {
                return index;
            }
        }

        return -1;
    }

    internal static OcrBox Union(IEnumerable<OcrBox> boxes)
    {
        double minX = double.MaxValue, minY = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue;
        var any = false;

        foreach (var box in boxes)
        {
            any = true;
            minX = Math.Min(minX, box.X);
            minY = Math.Min(minY, box.Y);
            maxX = Math.Max(maxX, box.Right);
            maxY = Math.Max(maxY, box.Bottom);
        }

        return any ? new OcrBox(minX, minY, maxX - minX, maxY - minY) : default;
    }
}
