using System.Diagnostics;
using OpenCvSharp;

namespace PdfBench;

/// <summary>
/// H(代替)。スキャンされた罫線表を、学習モデルではなく罫線検出で復元する。
///
/// SLANet(表構造モデル)は 41 行の表で列格子が 1 列ずれて全滅したため、
/// 「罫線が印刷されているなら、罫線そのものから格子を作る」古典的な方法を測る。
/// 手順: 二値化 → 長い水平線・垂直線をモルフォロジーで抽出 → 線の位置から
/// 行 y・列 x を確定 → 全面 OCR の各領域を中心座標で格子へ割り当てる。
/// </summary>
public static class GridTableStage
{
    public static void Run(string fixtures, string outDir)
    {
        var gt = Json.Load<List<TablePageGt>>(Path.Combine(fixtures, "gt", "table.json"));
        var result = new StageResult { Stage = "grid-ocr-table-scan" };

        var bytes = File.ReadAllBytes(Path.Combine(fixtures, "table-scan-clean.pdf"));
        var pageCount = PDFtoImage.Conversion.GetPageCount(bytes);
        var timer = Stopwatch.StartNew();

        var cellTotal = 0;
        var cellExact = 0;
        var rowCountOk = 0;

        using var engine = PaddleStages.CreateEngine();
        for (var page = 0; page < pageCount; page++)
        {
            using var mat = PaddleStages.RenderToMat(bytes, page, 300);

            var (rowLines, columnLines) = FindGrid(mat);
            var ocr = engine.Run(mat);

            // 格子(行 × 列)へ OCR 領域を割り当てる。
            var rows = rowLines.Count - 1;
            var columns = columnLines.Count - 1;
            var cells = new string[Math.Max(rows, 0), Math.Max(columns, 0)];
            var pieces = new List<(int Row, int Column, double X, string Text)>();

            foreach (var region in ocr.Regions)
            {
                var cx = region.Rect.Center.X;
                var cy = region.Rect.Center.Y;
                var row = Between(rowLines, cy);
                var column = Between(columnLines, cx);
                if (row >= 0 && row < rows && column >= 0 && column < columns)
                {
                    pieces.Add((row, column, cx, region.Text));
                }
            }

            foreach (var group in pieces.GroupBy(piece => (piece.Row, piece.Column)))
            {
                cells[group.Key.Row, group.Key.Column] = TextMetrics.Strip(
                    string.Concat(group.OrderBy(piece => piece.X).Select(piece => piece.Text)));
            }

            var expected = gt[page].Rows;
            if (rows == expected.Count)
            {
                rowCountOk++;
            }
            else if (result.Failures.Count < 30)
            {
                result.Failures.Add($"p{page + 1} rows {rows}/{expected.Count} cols {columns}");
            }

            for (var r = 0; r < expected.Count; r++)
            {
                for (var c = 0; c < 4; c++)
                {
                    cellTotal++;
                    var want = TextMetrics.Strip(expected[r][c]);
                    var got = r < rows && c < columns ? cells[r, c] ?? string.Empty : "(なし)";
                    if (got == want)
                    {
                        cellExact++;
                    }
                    else if (result.Failures.Count < 60)
                    {
                        result.Failures.Add($"p{page + 1} r{r}c{c} '{want}' -> '{got}'");
                    }
                }
            }
        }

        timer.Stop();
        result.Seconds = timer.Elapsed.TotalSeconds;
        result.Pages = pageCount;
        result.Metrics["cellExact"] = (double)cellExact / cellTotal;
        result.Metrics["cellTotal"] = cellTotal;
        result.Metrics["rowCountOkPages"] = rowCountOk;
        result.Metrics["secPerPage"] = timer.Elapsed.TotalSeconds / pageCount;
        result.Save(outDir);
    }

    /// <summary>罫線(長い水平線・垂直線)の位置を求める。</summary>
    private static (List<double> RowLines, List<double> ColumnLines) FindGrid(Mat source)
    {
        using var gray = source.CvtColor(ColorConversionCodes.BGR2GRAY);
        using var binary = new Mat();
        Cv2.AdaptiveThreshold(
            gray, binary, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.BinaryInv, 25, 15);

        // 横に長い成分だけを残す → 水平罫線。
        using var horizontalKernel = Cv2.GetStructuringElement(
            MorphShapes.Rect, new Size(Math.Max(20, source.Width / 30), 1));
        using var horizontal = binary.MorphologyEx(MorphTypes.Open, horizontalKernel);

        using var verticalKernel = Cv2.GetStructuringElement(
            MorphShapes.Rect, new Size(1, Math.Max(20, source.Height / 40)));
        using var vertical = binary.MorphologyEx(MorphTypes.Open, verticalKernel);

        return (LinePositions(horizontal, horizontalIsRows: true),
            LinePositions(vertical, horizontalIsRows: false));
    }

    /// <summary>線画像を行(列)方向に足し込み、山の位置を線の座標として拾う。</summary>
    private static List<double> LinePositions(Mat lines, bool horizontalIsRows)
    {
        using var reduced = new Mat();
        Cv2.Reduce(
            lines, reduced,
            horizontalIsRows ? ReduceDimension.Column : ReduceDimension.Row,
            ReduceTypes.Avg, MatType.CV_32F);

        var length = horizontalIsRows ? lines.Rows : lines.Cols;
        var values = new float[length];
        for (var index = 0; index < length; index++)
        {
            values[index] = horizontalIsRows
                ? reduced.At<float>(index, 0)
                : reduced.At<float>(0, index);
        }

        // 山(平均輝度が高い帯)の中心を線の位置にする。
        var positions = new List<double>();
        var threshold = 40f;
        var start = -1;
        for (var index = 0; index < length; index++)
        {
            if (values[index] >= threshold)
            {
                if (start < 0)
                {
                    start = index;
                }
            }
            else if (start >= 0)
            {
                positions.Add((start + index - 1) / 2.0);
                start = -1;
            }
        }

        if (start >= 0)
        {
            positions.Add((start + length - 1) / 2.0);
        }

        return positions;
    }

    /// <summary>座標がどの線間(セル帯)にあるか。線の外なら -1。</summary>
    private static int Between(List<double> lines, double value)
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
}
