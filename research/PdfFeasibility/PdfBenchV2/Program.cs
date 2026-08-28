using System.Diagnostics;
using OpenCvSharp;
using PdfBench;
using Sdcb.PaddleInference;
using Sdcb.PaddleOCR;
using Sdcb.PaddleOCR.Models;
using Sdcb.PaddleOCR.Models.Online;

namespace PdfBenchV2;

/// <summary>
/// Sdcb 2.x + Paddle 2.x runtime での測定(日本語専用 rec + SLANet 表構造)。
/// fixture と GT は第 1 スタックが生成したものをそのまま使う。すべて架空データ。
/// モデルはベンチマーク中のみダウンロードする(製品では同梱し、実行時 DL はしない)。
/// </summary>
public static class Program
{
    private static FullOcrModel? _model;

    public static int Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("usage: PdfBenchV2 <work-dir> <v2-text|v2-form|v2-table|v2-handwriting|v2-reflect> [degraded]");
            return 1;
        }

        var workDir = Path.GetFullPath(args[0]);
        var fixtures = Path.Combine(workDir, "fixtures");
        var outDir = Path.Combine(workDir, "out");
        Directory.CreateDirectory(outDir);

        switch (args[1])
        {
            case "v2-reflect":
                foreach (var property in typeof(OnlineFullModels)
                    .GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static))
                {
                    Console.WriteLine("OnlineFullModels." + property.Name);
                }

                break;

            case "v2-text":
                RunTextScan(fixtures, outDir, degraded: args.Contains("degraded"));
                break;

            case "v2-form":
                RunForm(fixtures, outDir, degraded: args.Contains("degraded"));
                break;

            case "v2-table":
                RunTable(fixtures, outDir);
                break;

            case "v2-handwriting":
                RunHandwriting(fixtures, outDir);
                break;

            default:
                Console.WriteLine("unknown command: " + args[1]);
                return 1;
        }

        return 0;
    }

    private static FullOcrModel Model()
    {
        if (_model is null)
        {
            _model = OnlineFullModels.JapanV4.DownloadAsync().GetAwaiter().GetResult();
            Console.WriteLine("model: JapanV4 (v2 stack)");
        }

        return _model;
    }

    private static PaddleOcrAll CreateEngine()
        => new(Model(), PaddleDevice.Mkldnn())
        {
            AllowRotateDetection = false,
            Enable180Classification = false,
        };

    private static Mat RenderToMat(byte[] pdfBytes, int page, int dpi)
    {
        using var bitmap = PDFtoImage.Conversion.ToImage(
            pdfBytes, page: page, options: new PDFtoImage.RenderOptions(Dpi: dpi));
        using var image = SkiaSharp.SKImage.FromBitmap(bitmap);
        using var encoded = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
        return Cv2.ImDecode(encoded.ToArray(), ImreadModes.Color);
    }

    private static void RunTextScan(string fixtures, string outDir, bool degraded)
    {
        var gt = Json.Load<List<TextPageGt>>(Path.Combine(fixtures, "gt", "text.json"));
        var pdfName = degraded ? "text-scan-degraded.pdf" : "text-scan-clean.pdf";
        var result = new StageResult
        {
            Stage = (degraded ? "paddle-text-degraded" : "paddle-text-clean") + "-v2jp",
        };

        var bytes = File.ReadAllBytes(Path.Combine(fixtures, pdfName));
        var pageCount = PDFtoImage.Conversion.GetPageCount(bytes);
        var timer = Stopwatch.StartNew();

        var fieldTotal = 0;
        var fieldExact = 0;

        using var engine = CreateEngine();
        for (var page = 0; page < pageCount; page++)
        {
            using var mat = RenderToMat(bytes, page, degraded ? 150 : 300);
            var ocr = engine.Run(mat);
            var text = string.Concat(ocr.Regions.Select(region => region.Text));

            foreach (var (name, value) in gt[page].Fields)
            {
                fieldTotal++;
                if (TextMetrics.AppearsExactly(text, value))
                {
                    fieldExact++;
                }
                else if (result.Failures.Count < 40)
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
        result.Metrics["secPerPage"] = timer.Elapsed.TotalSeconds / pageCount;
        result.Save(outDir);
    }

    private static void RunForm(string fixtures, string outDir, bool degraded)
    {
        var gt = Json.Load<List<FormPageGt>>(Path.Combine(fixtures, "gt", "form.json"));
        var pdfName = degraded ? "form-scan-degraded.pdf" : "form-scan-clean.pdf";
        var result = new StageResult
        {
            Stage = (degraded ? "paddle-form-full-degraded" : "paddle-form-full") + "-v2jp",
        };

        var bytes = File.ReadAllBytes(Path.Combine(fixtures, pdfName));
        var pageCount = PDFtoImage.Conversion.GetPageCount(bytes);
        var timer = Stopwatch.StartNew();

        var perField = new Dictionary<string, (int Total, int Exact)>();
        var confidences = new List<(double Score, bool Correct)>();

        using var engine = CreateEngine();
        for (var page = 0; page < pageCount; page++)
        {
            using var mat = RenderToMat(bytes, page, degraded ? 150 : 300);
            var ocr = engine.Run(mat);
            var text = string.Concat(ocr.Regions.Select(region => region.Text));

            foreach (var (name, value) in gt[page].Fields)
            {
                if (value.Length == 0)
                {
                    continue;
                }

                var entry = perField.GetValueOrDefault(name);
                var hit = TextMetrics.AppearsExactly(text, value);
                perField[name] = (entry.Total + 1, entry.Exact + (hit ? 1 : 0));

                if (!hit && result.Failures.Count < 40)
                {
                    result.Failures.Add($"p{page + 1} {name}='{value}'");
                }

                var best = ocr.Regions
                    .OrderByDescending(region => TextMetrics.CharacterAccuracy(value, region.Text))
                    .FirstOrDefault();
                if (best.Text is not null)
                {
                    confidences.Add((best.Score, hit));
                }
            }
        }

        timer.Stop();
        result.Seconds = timer.Elapsed.TotalSeconds;
        result.Pages = pageCount;

        var total = perField.Values.Sum(v => v.Total);
        var exact = perField.Values.Sum(v => v.Exact);
        result.Metrics["fieldExact"] = (double)exact / total;
        result.Metrics["fieldTotal"] = total;
        foreach (var (name, value) in perField)
        {
            result.Metrics["field:" + name] = (double)value.Exact / value.Total;
        }

        result.Metrics["secPerPage"] = timer.Elapsed.TotalSeconds / pageCount;

        foreach (var threshold in new[] { 0.90, 0.95, 0.98 })
        {
            var auto = confidences.Count(c => c.Score >= threshold);
            var wrongAuto = confidences.Count(c => c.Score >= threshold && !c.Correct);
            result.Metrics[$"auto@{threshold:0.00}"] = (double)auto / confidences.Count;
            result.Metrics[$"wrongAuto@{threshold:0.00}"] = (double)wrongAuto / confidences.Count;
        }

        result.Save(outDir);
    }

    private static void RunTable(string fixtures, string outDir)
    {
        var gt = Json.Load<List<TablePageGt>>(Path.Combine(fixtures, "gt", "table.json"));
        var result = new StageResult { Stage = "paddle-table-scan-v2jp" };

        var bytes = File.ReadAllBytes(Path.Combine(fixtures, "table-scan-clean.pdf"));
        var pageCount = PDFtoImage.Conversion.GetPageCount(bytes);
        var timer = Stopwatch.StartNew();

        var cellTotal = 0;
        var cellExact = 0;
        var rowCountOk = 0;

        var tableModel = OnlineTableRecognitionModel.ChineseMobileV2_SLANET
            .DownloadAsync().GetAwaiter().GetResult();

        using var engine = CreateEngine();
        using var tableRecognizer = new PaddleOcrTableRecognizer(tableModel);

        for (var page = 0; page < pageCount; page++)
        {
            using var mat = RenderToMat(bytes, page, 300);
            var structure = tableRecognizer.Run(mat);
            var ocr = engine.Run(mat);
            var html = structure.RebuildTable(ocr);

            var rows = ParseHtmlTable(html);
            var expected = gt[page].Rows;

            if (rows.Count == expected.Count)
            {
                rowCountOk++;
            }
            else if (result.Failures.Count < 30)
            {
                result.Failures.Add($"p{page + 1} rows {rows.Count}/{expected.Count}");
            }

            for (var r = 0; r < expected.Count; r++)
            {
                for (var c = 0; c < 4; c++)
                {
                    cellTotal++;
                    var want = TextMetrics.Strip(expected[r][c]);
                    if (r < rows.Count && c < rows[r].Count && rows[r][c] == want)
                    {
                        cellExact++;
                    }
                    else if (result.Failures.Count < 60)
                    {
                        var got = r < rows.Count && c < rows[r].Count ? rows[r][c] : "(なし)";
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

    private static void RunHandwriting(string fixtures, string outDir)
    {
        var gt = Json.Load<List<TextPageGt>>(Path.Combine(fixtures, "gt", "handwriting.json"));
        var result = new StageResult { Stage = "paddle-handwriting-v2jp" };

        var bytes = File.ReadAllBytes(Path.Combine(fixtures, "handwriting-scan.pdf"));
        var pageCount = PDFtoImage.Conversion.GetPageCount(bytes);
        var timer = Stopwatch.StartNew();

        var fieldTotal = 0;
        var fieldExact = 0;

        using var engine = CreateEngine();
        for (var page = 0; page < pageCount; page++)
        {
            using var mat = RenderToMat(bytes, page, 300);
            var ocr = engine.Run(mat);
            var text = string.Concat(ocr.Regions.Select(region => region.Text));

            foreach (var (_, value) in gt[page].Fields)
            {
                fieldTotal++;
                if (TextMetrics.AppearsExactly(text, value))
                {
                    fieldExact++;
                }
            }
        }

        timer.Stop();
        result.Seconds = timer.Elapsed.TotalSeconds;
        result.Pages = pageCount;
        result.Metrics["fieldExact"] = (double)fieldExact / fieldTotal;
        result.Save(outDir);
    }

    private static List<List<string>> ParseHtmlTable(string html)
    {
        var rows = new List<List<string>>();
        foreach (var rowHtml in html.Split("<tr>", StringSplitOptions.RemoveEmptyEntries))
        {
            if (!rowHtml.Contains("<td"))
            {
                continue;
            }

            var cells = new List<string>();
            foreach (var cellHtml in rowHtml.Split("<td", StringSplitOptions.RemoveEmptyEntries).Skip(1))
            {
                var start = cellHtml.IndexOf('>') + 1;
                var end = cellHtml.IndexOf("</td>", StringComparison.Ordinal);
                if (start > 0 && end >= start)
                {
                    cells.Add(TextMetrics.Strip(cellHtml[start..end]));
                }
            }

            if (cells.Count > 0)
            {
                rows.Add(cells);
            }
        }

        return rows;
    }
}
