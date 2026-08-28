using System.Diagnostics;
using Tesseract;

namespace PdfBench;

/// <summary>
/// Tesseract(baseline 比較)。jpn.traineddata(Apache-2.0)を作業フォルダーへ置いて使う。
/// PaddleOCR に対する優位が無ければ製品の本命にはしない。
/// </summary>
public static class TesseractStage
{
    public static void Run(string workDir, string fixtures, string outDir)
    {
        var tessdata = Path.Combine(workDir, "tessdata");
        if (!File.Exists(Path.Combine(tessdata, "jpn.traineddata")))
        {
            Console.WriteLine("tessdata/jpn.traineddata not found: " + tessdata);
            Console.WriteLine("download it first (benchmark only), e.g. tessdata_fast jpn.traineddata");
            return;
        }

        RunForms(tessdata, fixtures, outDir, degraded: false, maxPages: 30);
        RunForms(tessdata, fixtures, outDir, degraded: true, maxPages: 30);
    }

    private static void RunForms(
        string tessdata, string fixtures, string outDir, bool degraded, int maxPages)
    {
        var gt = Json.Load<List<FormPageGt>>(Path.Combine(fixtures, "gt", "form.json"));
        var pdfName = degraded ? "form-scan-degraded.pdf" : "form-scan-clean.pdf";
        var result = new StageResult { Stage = degraded ? "tess-form-degraded" : "tess-form-clean" };

        var bytes = File.ReadAllBytes(Path.Combine(fixtures, pdfName));
        var pageCount = Math.Min(PDFtoImage.Conversion.GetPageCount(bytes), maxPages);
        var timer = Stopwatch.StartNew();

        var fieldTotal = 0;
        var fieldExact = 0;
        var confidences = new List<float>();

        using var engine = new TesseractEngine(tessdata, "jpn", EngineMode.LstmOnly);
        for (var page = 0; page < pageCount; page++)
        {
            byte[] png;
            using (var bitmap = PDFtoImage.Conversion.ToImage(
                bytes, page: page, options: new PDFtoImage.RenderOptions(Dpi: degraded ? 150 : 300)))
            using (var image = SkiaSharp.SKImage.FromBitmap(bitmap))
            using (var encoded = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100))
            {
                png = encoded.ToArray();
            }

            using var pix = Pix.LoadFromMemory(png);
            using var recognized = engine.Process(pix);
            var text = recognized.GetText();
            confidences.Add(recognized.GetMeanConfidence());

            foreach (var (name, value) in gt[page].Fields)
            {
                if (value.Length == 0)
                {
                    continue;
                }

                fieldTotal++;
                if (TextMetrics.AppearsExactly(text, value))
                {
                    fieldExact++;
                }
                else if (result.Failures.Count < 30)
                {
                    result.Failures.Add($"p{page + 1} {name}='{value}'");
                }
            }
        }

        timer.Stop();
        result.Seconds = timer.Elapsed.TotalSeconds;
        result.Pages = pageCount;
        result.Metrics["fieldExact"] = (double)fieldExact / fieldTotal;
        result.Metrics["fieldTotal"] = fieldTotal;
        result.Metrics["meanConfidence"] = confidences.Average();
        result.Metrics["secPerPage"] = timer.Elapsed.TotalSeconds / pageCount;
        result.Save(outDir);
    }
}
