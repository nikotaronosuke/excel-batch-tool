using System.Diagnostics;
using SkiaSharp;

namespace PdfBench;

/// <summary>
/// G. チェックボックス。OCR ではなく、固定位置の画素解析で判定する。
/// スキャン(きれい)を対象にし、劣化版でどこまで落ちるかも測る。
/// </summary>
public static class CheckboxStage
{
    public static void Run(string fixtures, string outDir)
    {
        RunOne(fixtures, outDir, "form-scan-clean.pdf", 300, "checkbox-clean");
        RunOne(fixtures, outDir, "form-scan-degraded.pdf", 150, "checkbox-degraded");
    }

    private static void RunOne(string fixtures, string outDir, string pdfName, int dpi, string stage)
    {
        var gt = Json.Load<List<FormPageGt>>(Path.Combine(fixtures, "gt", "form.json"));
        var result = new StageResult { Stage = stage };
        var timer = Stopwatch.StartNew();

        var bytes = File.ReadAllBytes(Path.Combine(fixtures, pdfName));
        var pageCount = PDFtoImage.Conversion.GetPageCount(bytes);

        var total = 0;
        var exact = 0;

        for (var page = 0; page < pageCount; page++)
        {
            using var bitmap = PDFtoImage.Conversion.ToImage(
                bytes, page: page, options: new PDFtoImage.RenderOptions(Dpi: dpi));

            var expected = gt[page].Checkbox;
            var detected = Detect(bitmap, dpi);

            total++;
            if (detected == expected)
            {
                exact++;
            }
            else if (result.Failures.Count < 40)
            {
                result.Failures.Add($"p{page + 1} expected {expected} got {detected}");
            }
        }

        timer.Stop();
        result.Seconds = timer.Elapsed.TotalSeconds;
        result.Pages = pageCount;
        result.Metrics["exact"] = (double)exact / total;
        result.Metrics["total"] = total;
        result.Save(outDir);
    }

    /// <summary>
    /// 箱の中の黒画素率と、ラベル周辺(丸囲み用)の黒画素率から、どの選択肢かを決める。
    /// </summary>
    private static string Detect(SKBitmap bitmap, int dpi)
    {
        var scale = dpi / 72f;
        var bestOption = "none";
        var bestScore = 0.0;

        foreach (var (label, boxX) in Layouts.CheckOptions)
        {
            // 箱の内側(枠線を除く)。
            var inner = DarkRatio(bitmap,
                (int)((boxX + 3) * scale), (int)((Layouts.CheckY + 3) * scale),
                (int)((Layouts.CheckBoxSize - 6) * scale));

            // ラベルの左右上下の余白(丸囲みの線が通るところ)。文字自体は含めない。
            var ringTop = DarkRatio(bitmap,
                (int)((boxX + Layouts.CheckBoxSize + 2) * scale),
                (int)((Layouts.CheckY - 5) * scale),
                (int)(44 * scale), (int)(4 * scale));
            var ringBottom = DarkRatio(bitmap,
                (int)((boxX + Layouts.CheckBoxSize + 2) * scale),
                (int)((Layouts.CheckY + Layouts.CheckBoxSize + 1) * scale),
                (int)(44 * scale), (int)(4 * scale));

            var score = Math.Max(inner, Math.Max(ringTop, ringBottom));
            if (score > bestScore)
            {
                bestScore = score;
                bestOption = label;
            }
        }

        // どの選択肢にも十分な黒が無ければ無印。
        return bestScore >= 0.06 ? bestOption : "none";
    }

    private static double DarkRatio(SKBitmap bitmap, int x, int y, int size)
        => DarkRatio(bitmap, x, y, size, size);

    private static double DarkRatio(SKBitmap bitmap, int x, int y, int width, int height)
    {
        var dark = 0;
        var total = 0;
        for (var py = y; py < y + height && py < bitmap.Height; py++)
        {
            for (var px = x; px < x + width && px < bitmap.Width; px++)
            {
                total++;
                var color = bitmap.GetPixel(px, py);
                if (color.Red + color.Green + color.Blue < 3 * 128)
                {
                    dark++;
                }
            }
        }

        return total == 0 ? 0 : (double)dark / total;
    }
}
