using System.Reflection;

namespace PdfBench;

/// <summary>
/// Phase 2F-R: PDF 対応の実現性ベンチマーク。
/// すべて架空データ。結果は JSON(out/)へ保存し、外部送信は一切しない。
/// </summary>
public static class Program
{
    public static int Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("usage: PdfBench <work-dir> <command>");
            Console.WriteLine("commands: gen | pdfpig | tabula | detect | reflect | " +
                "paddle-scan | paddle-degraded | paddle-form | paddle-area | paddle-table | " +
                "paddle-handwriting | checkbox | tess | sizes");
            return 1;
        }

        var workDir = Path.GetFullPath(args[0]);
        PaddleStages.UseJapanRecognition = args.Contains("jp");
        var fixtures = Path.Combine(workDir, "fixtures");
        var outDir = Path.Combine(workDir, "out");
        Directory.CreateDirectory(workDir);
        Directory.CreateDirectory(outDir);

        switch (args[1])
        {
            case "gen":
                FixtureGen.GenerateAll(fixtures);
                break;

            case "preview":
                foreach (var (name, page, dpi) in new[]
                {
                    ("form-scan-clean.pdf", 0, 120),
                    ("form-scan-degraded.pdf", 1, 120),
                    ("table-scan-clean.pdf", 0, 100),
                    ("handwriting-scan.pdf", 0, 120),
                    ("text.pdf", 0, 100),
                })
                {
                    var pdf = File.ReadAllBytes(Path.Combine(fixtures, name));
                    using var bitmap = PDFtoImage.Conversion.ToImage(
                        pdf, page: page, options: new PDFtoImage.RenderOptions(Dpi: dpi));
                    using var image = SkiaSharp.SKImage.FromBitmap(bitmap);
                    using var png = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 90);
                    File.WriteAllBytes(
                        Path.Combine(workDir, "preview-" + name.Replace(".pdf", "") + ".png"),
                        png.ToArray());
                }

                Console.WriteLine("previews written");
                break;

            case "pdfpig":
                PdfPigTextStage.Run(fixtures, outDir);
                break;

            case "tabula":
                TabulaStage.Run(fixtures, outDir, lined: true);
                TabulaStage.Run(fixtures, outDir, lined: false);
                HeaderGuidedStage.Run(fixtures, outDir, lined: true);
                HeaderGuidedStage.Run(fixtures, outDir, lined: false);
                break;

            case "detect":
                DetectStage.Run(fixtures, outDir);
                break;

            case "reflect":
                Reflect();
                break;

            case "dump":
            {
                using var document = UglyToad.PdfPig.PdfDocument.Open(
                    Path.Combine(fixtures, args.Length > 2 ? args[2] : "text.pdf"));
                var page = document.GetPage(1);
                var lines = new List<string>
                {
                    "letters=" + page.Letters.Count,
                    "--- ContentOrderTextExtractor ---",
                    UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor.ContentOrderTextExtractor
                        .GetText(page),
                    "--- words ---",
                };
                lines.AddRange(page.GetWords().Take(60).Select(word =>
                    $"({word.BoundingBox.Left:F0},{word.BoundingBox.Bottom:F0}) '{word.Text}'"));
                File.WriteAllText(Path.Combine(workDir, "dump.txt"), string.Join("\n", lines));
                Console.WriteLine("dumped");
                break;
            }

            case "dump-tabula":
            {
                using var document = UglyToad.PdfPig.PdfDocument.Open(
                    Path.Combine(fixtures, "table-lined.pdf"),
                    new UglyToad.PdfPig.ParsingOptions { ClipPaths = true });
                var pageArea = Tabula.ObjectExtractor.Extract(document, 1);
                var lines = new List<string>();
                foreach (var (algorithmName, algorithm) in new (string, Tabula.Extractors.IExtractionAlgorithm)[]
                {
                    ("spreadsheet", new Tabula.Extractors.SpreadsheetExtractionAlgorithm()),
                    ("basic", new Tabula.Extractors.BasicExtractionAlgorithm()),
                })
                {
                    var tables = algorithm.Extract(pageArea);
                    lines.Add($"== {algorithmName}: tables={tables.Count}");
                    foreach (var table in tables.Take(2))
                    {
                        lines.Add($"  rows={table.RowCount} cols={table.ColumnCount}");
                        foreach (var row in table.Rows.Take(6))
                        {
                            lines.Add("  | " + string.Join(" | ", row.Select(c => c.GetText())));
                        }
                    }
                }

                File.WriteAllText(Path.Combine(workDir, "dump-tabula.txt"), string.Join("\n", lines));
                Console.WriteLine("dumped");
                break;
            }

            case "paddle-scan":
                PaddleStages.RunTextScan(fixtures, outDir, degraded: false);
                break;

            case "paddle-degraded":
                PaddleStages.RunTextScan(fixtures, outDir, degraded: true);
                break;

            case "paddle-form":
                PaddleStages.RunFormFullPage(fixtures, outDir, degraded: args.Contains("degraded"));
                break;

            case "paddle-area":
                PaddleStages.RunFormFixedArea(fixtures, outDir, degraded: args.Contains("degraded"));
                break;

            case "paddle-table":
                PaddleStages.RunScannedTable(fixtures, outDir);
                break;

            case "grid-table":
                GridTableStage.Run(fixtures, outDir);
                break;

            case "paddle-handwriting":
                PaddleStages.RunHandwriting(fixtures, outDir);
                break;

            case "checkbox":
                CheckboxStage.Run(fixtures, outDir);
                break;

            case "paddle-dump":
            {
                var pdf = File.ReadAllBytes(Path.Combine(
                    fixtures,
                    args.Skip(2).FirstOrDefault(a => a.EndsWith(".pdf")) ?? "text-scan-clean.pdf"));
                using var engine = PaddleStages.CreateEngine();
                using var mat = PaddleStages.RenderToMat(pdf, 0, 300);
                var ocr = engine.Run(mat);
                File.WriteAllLines(
                    Path.Combine(workDir, "paddle-dump.txt"),
                    ocr.Regions.Select(region => $"{region.Score:F3} '{region.Text}'"));
                Console.WriteLine("dumped " + ocr.Regions.Length + " regions");
                break;
            }

            case "tess":
                TesseractStage.Run(workDir, fixtures, outDir);
                break;

            default:
                Console.WriteLine("unknown command: " + args[1]);
                return 1;
        }

        return 0;
    }

    /// <summary>Sdcb のモデル一覧を確認する(どの言語・版が同梱できるか)。</summary>
    private static void Reflect()
    {
        foreach (var assemblyName in new[]
        {
            "Sdcb.PaddleOCR.Models.Local",
            "Sdcb.PaddleOCR.Models.LocalV5",
            "Sdcb.PaddleOCR.Models.Online",
            "Sdcb.PaddleOCR",
        })
        {
            try
            {
                var assembly = Assembly.Load(assemblyName);
                Console.WriteLine("== " + assemblyName + " " + assembly.GetName().Version);
                foreach (var type in assembly.GetExportedTypes().OrderBy(t => t.FullName))
                {
                    var statics = type
                        .GetProperties(BindingFlags.Public | BindingFlags.Static)
                        .Select(p => p.Name)
                        .ToList();

                    if (statics.Count > 0)
                    {
                        Console.WriteLine("  " + type.FullName + ": " + string.Join(", ", statics));
                    }
                    else if (type.IsClass && type.Name.Contains("Recognizer"))
                    {
                        Console.WriteLine("  " + type.FullName);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("== " + assemblyName + " load failed: " + ex.Message);
            }
        }
    }
}
