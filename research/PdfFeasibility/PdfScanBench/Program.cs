using System.Diagnostics;
using System.Text;
using ExcelBatchTool.Core.Ocr;

// Phase 2F-B2 の実測。製品の経路(OcrPack -> PdfScanReader -> 確認)そのままで測る。
var packDir = args[0];
var work = args.Length > 1 ? args[1] : Path.Combine(Path.GetTempPath(), "b2bench");
var only = args.Length > 2 ? args[2] : string.Empty;
Directory.CreateDirectory(work);

var status = OcrPack.Inspect(packDir);
if (!status.IsUsable)
{
    Console.Error.WriteLine(status.Message);
    return 1;
}

using var engine = OcrPack.Load(status);
Console.WriteLine($"engine: {engine.Info.MultiModel} + {engine.Info.JapanModel} / "
    + $"{engine.Info.Runtime} / {engine.Info.Backend} / {engine.Info.Dpi}dpi");
Console.WriteLine();

if (only is "" or "deskew")
{
    MeasureDeskew();
}

if (only is "" or "table")
{
    MeasureTable("罫線あり 10p", ruled: true, pages: 10, rows: 20, tilt: 0, degraded: false);
    MeasureTable("罫線あり 傾き 2°", ruled: true, pages: 3, rows: 20, tilt: 2, degraded: false);
    MeasureTable("罫線あり 劣化", ruled: true, pages: 3, rows: 20, tilt: 0, degraded: true);
    MeasureTable("罫線なし 10p", ruled: false, pages: 10, rows: 20, tilt: 0, degraded: false);
}

if (only is "" or "form")
{
    MeasureForm("帳票 120p", pages: 120, shift: 0, tilt: 0, scale: 1.0);
    MeasureForm("帳票 ずれ ±10px", pages: 30, shift: 10, tilt: 0, scale: 1.0);
    MeasureForm("帳票 傾き ±2°", pages: 30, shift: 0, tilt: 2, scale: 1.0);
    MeasureForm("帳票 拡大 2%", pages: 30, shift: 0, tilt: 0, scale: 1.02);
}

if (only is "" or "mark")
{
    MeasureMarks("印 clean", degraded: false, tilt: 0);
    MeasureMarks("印 傾き 2°", degraded: false, tilt: 2);
    MeasureMarks("印 劣化", degraded: true, tilt: 0);
}

return 0;

void MeasureDeskew()
{
    double[] angles = [0, 0.5, 1, 2, 3, 5, -1, -3, -5];
    var pdf = Path.Combine(work, "tilted.pdf");
    Fixtures.TiltedText(pdf, angles);

    using var source = engine.Open(pdf);
    Console.WriteLine("傾きの推定  実際     推定     誤差   直す?");
    Console.WriteLine("----------  -------  -------  -----  -----");

    var errors = new List<double>();
    for (var index = 0; index < angles.Length; index++)
    {
        var probe = source.Probe(index + 1, CancellationToken.None);
        var error = Math.Abs(probe.SkewDegrees - angles[index]);
        errors.Add(error);

        Console.WriteLine(
            $"            {angles[index],6:F2}°  {probe.SkewDegrees,6:F2}°  {error,5:F2}°  "
            + $"{(DeskewPolicy.ShouldDeskew(probe.SkewDegrees, probe.SkewReliable) ? "する" : "しない")}"
            + $"{(probe.SkewReliable ? string.Empty : "  (推定が不安定)")}");
    }

    Console.WriteLine($"  角度の誤差: 平均 {errors.Average():F2}° / 最大 {errors.Max():F2}°");
    Console.WriteLine();
}

void MeasureTable(string label, bool ruled, int pages, int rows, double tilt, bool degraded)
{
    var pdf = Path.Combine(work, $"table-{label.GetHashCode():X8}.pdf");
    if (ruled)
    {
        Fixtures.RuledTable(pdf, pages, rows, tilt, degraded);
    }
    else
    {
        Fixtures.BorderlessTable(pdf, pages, rows, tilt);
    }

    var pageList = Enumerable.Range(1, pages).ToList();
    var timer = Stopwatch.StartNew();
    var reading = new PdfScanReader().Read(
        engine, pdf, pageList, new OcrReadOptions { Mode = OcrReadMode.Table });
    timer.Stop();

    if (reading.Issues.Count > 0)
    {
        Console.WriteLine($"{label}: " + string.Join(" | ", reading.Issues.Select(i => i.Message)));
        return;
    }

    int total = 0, exact = 0, autoAccepted = 0, falseAuto = 0, rowsOk = 0;

    foreach (var page in pageList)
    {
        var truth = Fixtures.TableTruth(page, rows);
        var cells = reading.Items.Where(item => item.PageNumber == page).ToList();
        var gotRows = cells.Count == 0 ? 0 : cells.Max(item => item.Row!.Value) + 1;
        if (gotRows == truth.Count)
        {
            rowsOk++;
        }

        for (var r = 0; r < truth.Count; r++)
        {
            for (var c = 0; c < 4; c++)
            {
                total++;
                var want = Strip(truth[r][c]);
                var cell = cells.FirstOrDefault(item => item.Row == r && item.Column == c);
                var got = Strip(cell?.Text ?? string.Empty);
                var correct = got == want;
                if (correct)
                {
                    exact++;
                }

                if (cell?.InitialStatus == OcrItemStatus.AutoAccepted)
                {
                    autoAccepted++;
                    if (!correct)
                    {
                        falseAuto++;
                    }
                }
            }
        }
    }

    if (Environment.GetEnvironmentVariable("DUMP") == "1")
    {
        using (var probeSource = engine.Open(pdf))
        {
            var pr = probeSource.Probe(1, CancellationToken.None);
            Console.WriteLine($"  -- probe p1: skew={pr.SkewDegrees:F2} reliable={pr.SkewReliable} "
                + $"h={pr.HorizontalRulings} v={pr.VerticalRulings} under={pr.UnderlineCount} "
                + $"deskew={DeskewPolicy.ShouldDeskew(pr.SkewDegrees, pr.SkewReliable)}");
            var page1 = probeSource.Read(1, 0, CancellationToken.None);
            Console.WriteLine($"  -- rulings p1: rows={page1.RowRulings.Count} "
                + $"cols={page1.ColumnRulings.Count} "
                + $"colX=[{string.Join(", ", page1.ColumnRulings.Select(v => v.ToString("F0")))}]");
            Console.WriteLine($"     lines={page1.Lines.Count} "
                + $"firstBoxes=[{string.Join(", ", page1.Lines.Take(6).Select(l => $"{l.Box.X:F0}"))}]");
        }

        var truth1 = Fixtures.TableTruth(1, rows);
        Console.WriteLine($"  -- {label} p1: 期待 {truth1.Count} 行 --");
        foreach (var group in reading.Items.Where(i => i.PageNumber == 1)
            .GroupBy(i => i.Row!.Value).OrderBy(g => g.Key).Take(4))
        {
            var got = string.Join(" | ", group.OrderBy(i => i.Column).Select(i => $"[{i.Column}]{i.Text}"));
            var want = group.Key < truth1.Count ? string.Join(" | ", truth1[group.Key]) : "(なし)";
            Console.WriteLine($"     r{group.Key}: got {got}");
            Console.WriteLine($"          want {want}");
        }
    }

    Console.WriteLine(
        $"{label,-16} セル {exact,5}/{total,-5} = {(double)exact / total,6:P1}  "
        + $"行数一致 {rowsOk}/{pages}  自動確定 {(double)autoAccepted / total,5:P0}  "
        + $"誤確定 {falseAuto}  {timer.Elapsed.TotalSeconds / pages,5:F2} s/page");
}

void MeasureForm(string label, int pages, double shift, double tilt, double scale)
{
    var pdf = Path.Combine(work, $"form-{label.GetHashCode():X8}.pdf");
    Fixtures.Forms(pdf, pages, out var truth, shift, tilt, scale);

    var template = BuildTemplate();
    var pageList = Enumerable.Range(1, pages).ToList();

    GC.Collect();
    GC.WaitForPendingFinalizers();
    var before = GC.GetTotalMemory(true);
    var timer = Stopwatch.StartNew();
    var reading = new PdfScanReader().Read(
        engine, pdf, pageList,
        new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = template });
    timer.Stop();
    var mb = Math.Max((GC.GetTotalMemory(false) - before) / 1048576.0, 0);

    if (reading.Issues.Count > 0)
    {
        Console.WriteLine($"{label}: " + string.Join(" | ", reading.Issues.Select(i => i.Message)));
        return;
    }

    var expected = pages * template.Fields.Count;
    int exact = 0, autoAccepted = 0, falseAuto = 0;

    foreach (var page in pageList)
    {
        var fields = truth[page - 1];
        foreach (var item in reading.Items.Where(item => item.PageNumber == page))
        {
            var want = Strip(fields[item.FieldName!]);
            var correct = Strip(item.Text) == want;
            if (correct)
            {
                exact++;
            }

            if (item.InitialStatus == OcrItemStatus.AutoAccepted)
            {
                autoAccepted++;
                if (!correct)
                {
                    falseAuto++;
                }
            }
        }
    }

    var coverage = (double)reading.Items.Count / expected;

    Console.WriteLine(
        $"{label,-16} 項目 {exact,5}/{expected,-5} = {(double)exact / expected,6:P1}  "
        + $"網羅 {coverage,6:P1}  自動確定 {(double)autoAccepted / expected,5:P0}  "
        + $"誤確定 {falseAuto}  見つからない {reading.InitiallyMissingCount,4}  "
        + $"{timer.Elapsed.TotalSeconds / pages,5:F2} s/page  {mb,4:F0} MB");
}

void MeasureMarks(string label, bool degraded, double tilt)
{
    const int pages = 30;
    var pdf = Path.Combine(work, $"mark-{label.GetHashCode():X8}.pdf");
    Fixtures.Marks(pdf, pages, out var truth, degraded, tilt);

    var template = new FormTemplate
    {
        Name = "架空の回答票",
        Fields =
        [
            new FormField
            {
                Name = "回答",
                Area = new OcrBox(300, 800, 1400, 80),
                Kind = FormFieldKind.Choice,
                Choices = [.. Fixtures.MarkBoxes().Select(box =>
                    new FormChoice(box.Label, new OcrBox(box.X, box.Y, box.Size, box.Size)))],
            },
        ],
    };

    var pageList = Enumerable.Range(1, pages).ToList();
    var reading = new PdfScanReader().Read(
        engine, pdf, pageList,
        new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = template });

    if (reading.Issues.Count > 0)
    {
        Console.WriteLine($"{label}: " + string.Join(" | ", reading.Issues.Select(i => i.Message)));
        return;
    }

    int exact = 0, autoAccepted = 0, falseAuto = 0;
    string[] styleNames = ["check", "fill", "circle", "cross", "none"];
    var perStyle = new Dictionary<string, (int Total, int Exact)>();

    foreach (var page in pageList)
    {
        var want = truth[page - 1];
        var item = reading.Items.First(item => item.PageNumber == page);
        var correct = item.Text == want;

        var style = styleNames[page % styleNames.Length];
        var stat = perStyle.GetValueOrDefault(style);
        perStyle[style] = (stat.Total + 1, stat.Exact + (correct ? 1 : 0));
        if (!correct && Environment.GetEnvironmentVariable("DUMP") == "1")
        {
            Console.WriteLine($"     p{page} {style}: want \"{want}\" got \"{item.Text}\" "
                + $"({item.Reason})");
        }
        if (correct)
        {
            exact++;
        }

        if (item.InitialStatus == OcrItemStatus.AutoAccepted)
        {
            autoAccepted++;
            if (!correct)
            {
                falseAuto++;
            }
        }
    }

    Console.WriteLine(
        $"{label,-16} 印 {exact,3}/{pages,-3} = {(double)exact / pages,6:P1}  "
        + $"自動確定 {(double)autoAccepted / pages,5:P0}  誤確定 {falseAuto}  "
        + string.Join(" ", perStyle.OrderBy(e => e.Key)
            .Select(e => $"{e.Key} {e.Value.Exact}/{e.Value.Total}")));
}

static FormTemplate BuildTemplate()
{
    var anchor = Fixtures.AnchorArea();
    return new FormTemplate
    {
        Name = "架空の売上報告",
        Fields = [.. Fixtures.FormAreas().Select(area => new FormField
        {
            Name = area.Name,
            Area = new OcrBox(area.X, area.Y, area.Width, area.Height),
            Kind = area.Name is "売上" ? FormFieldKind.NumberLike
                : area.Name is "店舗コード" ? FormFieldKind.Code
                : FormFieldKind.Text,
        })],
        Anchors =
        [
            new FormAnchor(
                anchor.Text,
                new OcrBox(anchor.X, anchor.Y, anchor.Width, anchor.Height)),
        ],
    };
}

static string Strip(string text)
{
    var normalized = text.Normalize(NormalizationForm.FormKC);
    var builder = new StringBuilder(normalized.Length);
    foreach (var character in normalized)
    {
        if (!char.IsWhiteSpace(character))
        {
            builder.Append(character);
        }
    }

    return builder.ToString();
}
