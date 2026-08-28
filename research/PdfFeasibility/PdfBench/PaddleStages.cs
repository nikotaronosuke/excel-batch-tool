using System.Diagnostics;
using OpenCvSharp;
using Sdcb.PaddleInference;
using Sdcb.PaddleOCR;
using Sdcb.PaddleOCR.Models;
using Sdcb.PaddleOCR.Models.Local;

namespace PdfBench;

/// <summary>C/D/E/F/H/I. PaddleOCR(ローカル推論)によるスキャン PDF の測定。</summary>
public static class PaddleStages
{
    /// <summary>
    /// 日本語のモデル。同梱パッケージにある Japan 系のうち、いちばん新しい版を使う
    /// (パッケージ更新で名前が変わっても追従できるよう、名前は実行時に解決する)。
    /// </summary>
    /// <summary>
    /// true なら認識モデルを日本語専用(japan_PP-OCRv4)にする。
    /// ChineseV5(多言語)は日本語文で「支・一」等を高 confidence のまま脱落させる
    /// 誤確定を確認したため、両方を測って比べる。検出は言語非依存の ChineseV5 を共用。
    /// </summary>
    internal static bool UseJapanRecognition { get; set; }

    private static string ModelSuffix => UseJapanRecognition ? "-jp" : string.Empty;

    private static FullOcrModel? _model;

    internal static FullOcrModel Model()
    {
        if (_model is not null)
        {
            return _model;
        }

        if (UseJapanRecognition)
        {
            // ベンチマーク中のみモデルを取得する(製品では同梱し、実行時 DL はしない)。
            var recognition = Sdcb.PaddleOCR.Models.Online.LocalDictOnlineRecognizationModel.JapanV4
                .DownloadAsync().GetAwaiter().GetResult();
            Console.WriteLine("model: det=ChineseV5 rec=JapanV4");
            _model = new FullOcrModel(
                LocalDetectionModel.ChineseV5, LocalClassificationModel.ChineseMobileV2, recognition);
        }
        else
        {
            Console.WriteLine("model: ChineseV5 (multilingual incl. Japanese)");
            _model = LocalFullModels.ChineseV5;
        }

        return _model;
    }

    // JapanV4(v2 形式のモデル)は Paddle 3.x の oneDNN 経路で
    // 「OneDnnContext does not have the input Filter」で失敗するため、
    // 日本語 rec のときだけ oneDNN を使わない Openblas に切り替える。
    internal static Action<PaddleConfig> Device()
        => UseJapanRecognition ? PaddleDevice.Openblas() : PaddleDevice.Mkldnn();

    internal static PaddleOcrAll CreateEngine()
        => new(Model(), Device())
        {
            AllowRotateDetection = false,
            Enable180Classification = false,
        };

    internal static Mat RenderToMat(byte[] pdfBytes, int page, int dpi)
    {
        using var bitmap = PDFtoImage.Conversion.ToImage(
            pdfBytes, page: page, options: new PDFtoImage.RenderOptions(Dpi: dpi));
        using var image = SkiaSharp.SKImage.FromBitmap(bitmap);
        using var encoded = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
        return Cv2.ImDecode(encoded.ToArray(), ImreadModes.Color);
    }

    /// <summary>D/E. 文章スキャンのフィールド一致率(値がページのどこかに完全一致で現れるか)。</summary>
    public static void RunTextScan(string fixtures, string outDir, bool degraded)
    {
        var gt = Json.Load<List<TextPageGt>>(Path.Combine(fixtures, "gt", "text.json"));
        var pdfName = degraded ? "text-scan-degraded.pdf" : "text-scan-clean.pdf";
        var result = new StageResult { Stage = (degraded ? "paddle-text-degraded" : "paddle-text-clean") + ModelSuffix };

        var bytes = File.ReadAllBytes(Path.Combine(fixtures, pdfName));
        var pageCount = PDFtoImage.Conversion.GetPageCount(bytes);
        var timer = Stopwatch.StartNew();

        var fieldTotal = 0;
        var fieldExact = 0;
        var charAccuracy = new List<double>();

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

                charAccuracy.Add(BestWindowAccuracy(text, value));
            }
        }

        timer.Stop();
        result.Seconds = timer.Elapsed.TotalSeconds;
        result.Pages = pageCount;
        result.Metrics["fieldExact"] = (double)fieldExact / fieldTotal;
        result.Metrics["fieldTotal"] = fieldTotal;
        result.Metrics["charAccuracy"] = charAccuracy.Average();
        result.Metrics["secPerPage"] = timer.Elapsed.TotalSeconds / pageCount;
        result.Save(outDir);
    }

    /// <summary>F. 固定帳票を全ページ・全面 OCR し、フィールドの完全一致を測る。</summary>
    public static void RunFormFullPage(string fixtures, string outDir, bool degraded)
    {
        var gt = Json.Load<List<FormPageGt>>(Path.Combine(fixtures, "gt", "form.json"));
        var pdfName = degraded ? "form-scan-degraded.pdf" : "form-scan-clean.pdf";
        var result = new StageResult
        {
            Stage = (degraded ? "paddle-form-full-degraded" : "paddle-form-full") + ModelSuffix,
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

                // その値に最も近い OCR 行の confidence を対応付ける。
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
        AddConfidenceSweep(result, confidences);
        result.Save(outDir);
    }

    /// <summary>
    /// F(比較)。固定帳票の値の位置だけを切り出して、1 行認識(rec のみ)で読む。
    /// 全面 OCR との精度・速度・誤確定を比べる。
    /// </summary>
    public static void RunFormFixedArea(string fixtures, string outDir, bool degraded)
    {
        var gt = Json.Load<List<FormPageGt>>(Path.Combine(fixtures, "gt", "form.json"));
        var pdfName = degraded ? "form-scan-degraded.pdf" : "form-scan-clean.pdf";
        var dpi = degraded ? 150 : 300;
        var result = new StageResult
        {
            Stage = (degraded ? "paddle-form-area-degraded" : "paddle-form-area") + ModelSuffix,
        };

        var bytes = File.ReadAllBytes(Path.Combine(fixtures, pdfName));
        var pageCount = PDFtoImage.Conversion.GetPageCount(bytes);
        var timer = Stopwatch.StartNew();

        var perField = new Dictionary<string, (int Total, int Exact)>();
        var confidences = new List<(double Score, bool Correct)>();

        // 素朴な「切り出し + 1 行認識(rec のみ)」はクロップ余白に過敏だった
        // (余白 21pt で 87%、16pt で 60%)。切り出した領域へ det+rec を掛ける方式にする。
        using var engine = CreateEngine();
        var scale = dpi / 72f;

        for (var page = 0; page < pageCount; page++)
        {
            using var mat = RenderToMat(bytes, page, dpi);

            foreach (var (name, _, valueX, y, width) in Layouts.FormFields)
            {
                var expected = gt[page].Fields[name];
                if (expected.Length == 0)
                {
                    continue;
                }

                // 値の領域 + 余白(det が行を見つけられるだけの周囲)。
                var rect = new Rect(
                    (int)((valueX - 4) * scale),
                    (int)((y - 20) * scale),
                    (int)((width + 8) * scale),
                    (int)(30 * scale));
                rect = rect.Intersect(new Rect(0, 0, mat.Width, mat.Height));

                using var crop = new Mat(mat, rect);
                var ocr = engine.Run(crop);
                var text = string.Concat(ocr.Regions
                    .OrderBy(region => region.Rect.Center.X)
                    .Select(region => region.Text));
                var score = ocr.Regions.Length == 0
                    ? 0
                    : ocr.Regions.Min(region => region.Score);

                var entry = perField.GetValueOrDefault(name);
                var hit = TextMetrics.ExactIgnoringSpaces(expected, text);
                perField[name] = (entry.Total + 1, entry.Exact + (hit ? 1 : 0));
                confidences.Add((score, hit));

                if (!hit && result.Failures.Count < 40)
                {
                    result.Failures.Add(
                        $"p{page + 1} {name}='{expected}' -> '{TextMetrics.Strip(text)}'");
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
        AddConfidenceSweep(result, confidences);
        result.Save(outDir);
    }

    /// <summary>H. スキャンされた表: PaddleOCR + SLANet(表構造)で行列へ復元する。</summary>
    public static void RunScannedTable(string fixtures, string outDir)
    {
        var gt = Json.Load<List<TablePageGt>>(Path.Combine(fixtures, "gt", "table.json"));
        var result = new StageResult { Stage = "paddle-table-scan" + ModelSuffix };

        var bytes = File.ReadAllBytes(Path.Combine(fixtures, "table-scan-clean.pdf"));
        var pageCount = PDFtoImage.Conversion.GetPageCount(bytes);
        var timer = Stopwatch.StartNew();

        var cellTotal = 0;
        var cellExact = 0;
        var rowCountOk = 0;

        using var engine = CreateEngine();
        // SLANet は旧形式のモデルで、Paddle 3.x の oneDNN 経路では
        // 「OneDnnContext does not have the input Filter」で失敗する。Openblas で動かす。
        using var tableRecognizer = new PaddleOcrTableRecognizer(
            LocalTableRecognitionModel.ChineseMobileV2_SLANET, PaddleDevice.Openblas());

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
            else if (result.Failures.Count < 40)
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

    /// <summary>I. 手書き風の characterization(GO 条件には含めない)。</summary>
    public static void RunHandwriting(string fixtures, string outDir)
    {
        var gt = Json.Load<List<TextPageGt>>(Path.Combine(fixtures, "gt", "handwriting.json"));
        var result = new StageResult { Stage = "paddle-handwriting" + ModelSuffix };

        var bytes = File.ReadAllBytes(Path.Combine(fixtures, "handwriting-scan.pdf"));
        var pageCount = PDFtoImage.Conversion.GetPageCount(bytes);
        var timer = Stopwatch.StartNew();

        var fieldTotal = 0;
        var fieldExact = 0;
        var accuracy = new List<double>();

        using var engine = CreateEngine();
        for (var page = 0; page < pageCount; page++)
        {
            using var mat = RenderToMat(bytes, page, 300);
            var ocr = engine.Run(mat);
            var text = string.Concat(ocr.Regions.Select(region => region.Text));

            foreach (var (name, value) in gt[page].Fields)
            {
                fieldTotal++;
                if (TextMetrics.AppearsExactly(text, value))
                {
                    fieldExact++;
                }

                accuracy.Add(BestWindowAccuracy(text, value));
            }
        }

        timer.Stop();
        result.Seconds = timer.Elapsed.TotalSeconds;
        result.Pages = pageCount;
        result.Metrics["fieldExact"] = (double)fieldExact / fieldTotal;
        result.Metrics["charAccuracy"] = accuracy.Average();
        result.Save(outDir);
    }

    /// <summary>
    /// confidence の閾値を変えたときの、自動確定 / 要確認 / 誤確定の割合。
    /// 重要なのは「間違ったのに閾値を超えた(誤確定)」の少なさ。
    /// </summary>
    private static void AddConfidenceSweep(
        StageResult result, List<(double Score, bool Correct)> confidences)
    {
        foreach (var threshold in new[] { 0.90, 0.95, 0.98 })
        {
            var auto = confidences.Count(c => c.Score >= threshold);
            var wrongAuto = confidences.Count(c => c.Score >= threshold && !c.Correct);
            result.Metrics[$"auto@{threshold:0.00}"] = (double)auto / confidences.Count;
            result.Metrics[$"wrongAuto@{threshold:0.00}"] = (double)wrongAuto / confidences.Count;
        }

        // 間違いの confidence 分布(高 confidence 誤りの実態)。
        var wrong = confidences.Where(c => !c.Correct).Select(c => c.Score).OrderBy(s => s).ToList();
        result.Metrics["wrongCount"] = wrong.Count;
        if (wrong.Count > 0)
        {
            result.Metrics["wrongScoreMax"] = wrong[^1];
            result.Metrics["wrongScoreMedian"] = wrong[wrong.Count / 2];
        }
    }

    /// <summary>値と最も一致する近傍を全文から探したときの文字正解率(部分一致の把握用)。</summary>
    private static double BestWindowAccuracy(string text, string value)
    {
        var stripped = TextMetrics.Strip(text);
        var target = TextMetrics.Strip(value);
        if (stripped.Contains(target, StringComparison.Ordinal))
        {
            return 1;
        }

        var best = 0.0;
        for (var start = 0; start + target.Length <= stripped.Length; start += Math.Max(1, target.Length / 4))
        {
            var window = stripped.Substring(start, Math.Min(target.Length, stripped.Length - start));
            best = Math.Max(best, TextMetrics.CharacterAccuracy(target, window));
        }

        return best;
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
