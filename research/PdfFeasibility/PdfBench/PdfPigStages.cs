using System.Diagnostics;
using Tabula;
using Tabula.Extractors;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace PdfBench;

/// <summary>A. PdfPig: born-digital の文字と読み順。</summary>
public static class PdfPigTextStage
{
    public static void Run(string dir, string outDir)
    {
        var gt = Json.Load<List<TextPageGt>>(Path.Combine(dir, "gt", "text.json"));
        var result = new StageResult { Stage = "pdfpig-text" };
        var timer = Stopwatch.StartNew();

        var fieldTotal = 0;
        var fieldExact = 0;
        var charAccuracies = new List<double>();

        using (var document = PdfDocument.Open(Path.Combine(dir, "text.pdf")))
        {
            foreach (var pageGt in gt)
            {
                var page = document.GetPage(pageGt.Page);
                var text = ContentOrderTextExtractor.GetText(page);

                foreach (var (name, value) in pageGt.Fields)
                {
                    fieldTotal++;
                    var expected = $"{name}:{value}";
                    if (TextMetrics.AppearsExactly(text, expected))
                    {
                        fieldExact++;
                    }
                    else
                    {
                        result.Failures.Add($"p{pageGt.Page} {name}");
                    }

                    charAccuracies.Add(TextMetrics.AppearsExactly(text, value) ? 1 : 0);
                }

                // 読み順: タイトル → フィールド → 本文 の順で現れるか。
                var ordered = TextMetrics.Strip(text);
                var title = ordered.IndexOf(TextMetrics.Strip("月次報告書"), StringComparison.Ordinal);
                var field = ordered.IndexOf(TextMetrics.Strip("会社名"), StringComparison.Ordinal);
                var body = ordered.IndexOf(TextMetrics.Strip("本書は動作確認"), StringComparison.Ordinal);
                if (!(title >= 0 && title < field && field < body))
                {
                    result.Failures.Add($"p{pageGt.Page} reading-order");
                }
            }

            result.Pages = gt.Count;
        }

        timer.Stop();
        result.Seconds = timer.Elapsed.TotalSeconds;
        result.Metrics["fieldExact"] = (double)fieldExact / fieldTotal;
        result.Metrics["fieldTotal"] = fieldTotal;
        result.Metrics["readingOrderOk"] = result.Failures.Count(f => f.Contains("reading-order")) == 0;
        result.Save(outDir);
    }
}

/// <summary>B/C. PdfPig + Tabula: born-digital の表(罫線あり / なし)。</summary>
public static class TabulaStage
{
    public static void Run(string dir, string outDir, bool lined)
    {
        var gtName = lined ? "table.json" : "table-borderless.json";
        var pdfName = lined ? "table-lined.pdf" : "table-borderless.pdf";
        var gt = Json.Load<List<TablePageGt>>(Path.Combine(dir, "gt", gtName));

        var result = new StageResult { Stage = lined ? "tabula-lined" : "tabula-borderless" };
        var timer = Stopwatch.StartNew();

        var cellTotal = 0;
        var cellExact = 0;
        var rowCountOk = 0;
        var columnCountOk = 0;
        var headerOk = 0;

        using (var document = PdfDocument.Open(Path.Combine(dir, pdfName),
            new ParsingOptions { ClipPaths = true }))
        {
            IExtractionAlgorithm algorithm = lined
                ? new SpreadsheetExtractionAlgorithm()
                : new BasicExtractionAlgorithm();

            foreach (var pageGt in gt)
            {
                var pageArea = ObjectExtractor.Extract(document, pageGt.Page);
                var page = document.GetPage(pageGt.Page);
                var tables = algorithm.Extract(pageArea);

                // 外枠だけの「1 巨大セル」の格子を除き、実際の格子を選ぶ。
                var table = tables
                    .Where(t => t.Rows.Count > 0
                        && t.Rows[0].Count(c => TextMetrics.Strip(c.GetText()).Length is > 0 and < 40) >= 3)
                    .OrderByDescending(t => t.RowCount)
                    .FirstOrDefault()
                    ?? tables.OrderByDescending(t => t.RowCount).FirstOrDefault();

                // tabula-sharp は連続する同一文字を重複除去で潰すことがある(188→18)。
                // 構造(セル矩形)は Tabula、文字は PdfPig の letter から詰め直す hybrid で測る。
                var actual = table is null
                    ? []
                    : table.Rows
                        .Select(row => row
                            .Select(cell => TextMetrics.Strip(LettersInBox(page, cell.BoundingBox)))
                            .ToArray())
                        .Where(row => row.Any(cell => cell.Length > 0))
                        // 表全体を囲む外枠が「1 巨大セルの行」として混ざるので落とす。
                        .Where(row => !row.Any(cell => cell.Length > 60))
                        .ToList();

                // 罫線なし(basic)は空の列ができるので、全行で空の列を落とす。
                if (actual.Count > 0)
                {
                    var width = actual.Max(row => row.Length);
                    var keep = Enumerable.Range(0, width)
                        .Where(c => actual.Any(row => c < row.Length && row[c].Length > 0))
                        .ToArray();
                    actual = actual
                        .Select(row => keep.Select(c => c < row.Length ? row[c] : string.Empty).ToArray())
                        .ToList();
                }

                if (actual.Count == pageGt.Rows.Count)
                {
                    rowCountOk++;
                }
                else
                {
                    result.Failures.Add($"p{pageGt.Page} rows {actual.Count}/{pageGt.Rows.Count}");
                }

                if (actual.Count > 0 && actual.All(row => row.Length == 4))
                {
                    columnCountOk++;
                }

                if (actual.Count > 0
                    && actual[0].SequenceEqual(Layouts.TableHeaders.Select(TextMetrics.Strip)))
                {
                    headerOk++;
                }

                for (var r = 0; r < pageGt.Rows.Count; r++)
                {
                    for (var c = 0; c < 4; c++)
                    {
                        cellTotal++;
                        var expected = TextMetrics.Strip(pageGt.Rows[r][c]);
                        if (r < actual.Count && c < actual[r].Length && actual[r][c] == expected)
                        {
                            cellExact++;
                        }
                        else if (result.Failures.Count < 60)
                        {
                            var got = r < actual.Count && c < actual[r].Length ? actual[r][c] : "(なし)";
                            result.Failures.Add($"p{pageGt.Page} r{r}c{c} '{expected}' -> '{got}'");
                        }
                    }
                }
            }

            result.Pages = gt.Count;
        }

        timer.Stop();
        result.Seconds = timer.Elapsed.TotalSeconds;
        result.Metrics["cellExact"] = (double)cellExact / cellTotal;
        result.Metrics["cellTotal"] = cellTotal;
        result.Metrics["rowCountOkPages"] = rowCountOk;
        result.Metrics["columnCountOkPages"] = columnCountOk;
        result.Metrics["headerOkPages"] = headerOk;
        result.Save(outDir);
    }

    /// <summary>セル矩形の中にある PdfPig の letter を、左から順に並べて文字列にする。</summary>
    internal static string LettersInBox(Page page, UglyToad.PdfPig.Core.PdfRectangle box)
    {
        var inside = page.Letters
            .Where(letter =>
            {
                var x = (letter.GlyphRectangle.Left + letter.GlyphRectangle.Right) / 2;
                var y = (letter.GlyphRectangle.Top + letter.GlyphRectangle.Bottom) / 2;
                return x >= box.Left - 2 && x <= box.Right + 2
                    && y >= box.Bottom - 2 && y <= box.Top + 2;
            })
            .OrderBy(letter => letter.GlyphRectangle.Left);

        return string.Concat(inside.Select(letter => letter.Value));
    }
}

/// <summary>
/// 製品実装の候補: ヘッダー行の文字位置から列の左端を決め、以降の行を位置で割り当てる
/// 自前の再構成。罫線のあり・なしに依存しない。
/// </summary>
public static class HeaderGuidedStage
{
    public static void Run(string dir, string outDir, bool lined)
    {
        var gtName = lined ? "table.json" : "table-borderless.json";
        var pdfName = lined ? "table-lined.pdf" : "table-borderless.pdf";
        var gt = Json.Load<List<TablePageGt>>(Path.Combine(dir, "gt", gtName));

        var result = new StageResult
        {
            Stage = lined ? "header-guided-lined" : "header-guided-borderless",
        };
        var timer = Stopwatch.StartNew();

        var cellTotal = 0;
        var cellExact = 0;
        var rowCountOk = 0;

        using (var document = PdfDocument.Open(Path.Combine(dir, pdfName)))
        {
            foreach (var pageGt in gt)
            {
                var page = document.GetPage(pageGt.Page);

                // ベースライン(下端)のクラスタリングで行を作る(PDF は下原点なので上から)。
                var rows = page.Letters
                    .GroupBy(letter => Math.Round(letter.GlyphRectangle.Bottom / 4))
                    .OrderByDescending(group => group.Key)
                    .Select(group => group.OrderBy(letter => letter.GlyphRectangle.Left).ToList())
                    .ToList();

                // 1 行目 = ヘッダー。文字の途切れ(> 6pt)を列の境目とみなす。
                var columns = new List<double>();
                double? lastRight = null;
                foreach (var letter in rows[0])
                {
                    if (lastRight is null || letter.GlyphRectangle.Left - lastRight > 6)
                    {
                        columns.Add(letter.GlyphRectangle.Left);
                    }

                    lastRight = Math.Max(lastRight ?? 0, letter.GlyphRectangle.Right);
                }

                var actual = rows
                    .Select(letters =>
                    {
                        var cells = new string[columns.Count];
                        foreach (var group in letters.GroupBy(letter => ColumnOf(columns, letter)))
                        {
                            cells[group.Key] = TextMetrics.Strip(string.Concat(
                                group.OrderBy(l => l.GlyphRectangle.Left).Select(l => l.Value)));
                        }

                        return cells.Select(cell => cell ?? string.Empty).ToArray();
                    })
                    .Where(row => row.Any(cell => cell.Length > 0))
                    .ToList();

                if (actual.Count == pageGt.Rows.Count)
                {
                    rowCountOk++;
                }

                for (var r = 0; r < pageGt.Rows.Count; r++)
                {
                    for (var c = 0; c < 4; c++)
                    {
                        cellTotal++;
                        var expected = TextMetrics.Strip(pageGt.Rows[r][c]);
                        if (r < actual.Count && c < actual[r].Length && actual[r][c] == expected)
                        {
                            cellExact++;
                        }
                        else if (result.Failures.Count < 40)
                        {
                            var got = r < actual.Count && c < actual[r].Length
                                ? actual[r][c]
                                : "(なし)";
                            result.Failures.Add($"p{pageGt.Page} r{r}c{c} '{expected}' -> '{got}'");
                        }
                    }
                }
            }

            result.Pages = gt.Count;
        }

        timer.Stop();
        result.Seconds = timer.Elapsed.TotalSeconds;
        result.Metrics["cellExact"] = (double)cellExact / cellTotal;
        result.Metrics["cellTotal"] = cellTotal;
        result.Metrics["rowCountOkPages"] = rowCountOk;
        result.Save(outDir);
    }

    private static int ColumnOf(List<double> columns, Letter letter)
    {
        var x = letter.GlyphRectangle.Left + 0.5;
        for (var index = columns.Count - 1; index >= 0; index--)
        {
            if (x >= columns[index] - 1)
            {
                return index;
            }
        }

        return 0;
    }
}

/// <summary>PDF の種類の自動判定(埋め込みテキスト / 表候補 / スキャン / 混在)。</summary>
public static class DetectStage
{
    public static void Run(string dir, string outDir)
    {
        var result = new StageResult { Stage = "detect" };
        var timer = Stopwatch.StartNew();

        // ファイル → 期待する判定。
        (string File, string Expected)[] cases =
        [
            ("text.pdf", "text"),
            ("table-lined.pdf", "table"),
            ("table-borderless.pdf", "text-or-table"),
            ("form.pdf", "text"),
            ("form-scan-clean.pdf", "scan"),
            ("form-scan-degraded.pdf", "scan"),
            ("table-scan-clean.pdf", "scan"),
            ("text-scan-clean.pdf", "scan"),
        ];

        var ok = 0;
        foreach (var (file, expected) in cases)
        {
            var detected = Detect(Path.Combine(dir, file));
            var matched = expected switch
            {
                "text-or-table" => detected is "text" or "table",
                _ => detected == expected,
            };

            if (matched)
            {
                ok++;
            }
            else
            {
                result.Failures.Add($"{file}: expected {expected} got {detected}");
            }

            result.Metrics[file] = detected;
        }

        timer.Stop();
        result.Seconds = timer.Elapsed.TotalSeconds;
        result.Pages = cases.Length;
        result.Metrics["correct"] = ok + "/" + cases.Length;
        result.Save(outDir);
    }

    /// <summary>1 ページ目の文字数・画像の被覆・罫線の本数から種類を決める。</summary>
    public static string Detect(string path)
    {
        using var document = PdfDocument.Open(path);
        var page = document.GetPage(1);

        var letters = page.Letters.Count;
        var images = page.GetImages().ToList();
        var pageArea = page.Width * page.Height;
        var imageCoverage = images.Sum(image => image.Bounds.Width * image.Bounds.Height) / pageArea;

        if (letters < 10 && imageCoverage > 0.5)
        {
            return "scan";
        }

        if (letters >= 10 && imageCoverage > 0.5)
        {
            return "mixed";
        }

        // 細長い水平の描画(罫線)が多ければ表候補。
        var horizontalLines = page.Paths.Count(path =>
            path.GetBoundingRectangle() is { } box && box.Height < 2 && box.Width > 100);

        return horizontalLines >= 5 ? "table" : "text";
    }
}
