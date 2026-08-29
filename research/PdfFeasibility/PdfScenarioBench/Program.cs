using System.Diagnostics;
using System.Globalization;
using System.Text;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Ocr;
using ExcelBatchTool.Core.Pdf;
using PdfScenarioBench;

// Phase 2F-B3 の実案件相当の総合試験。
//   PdfScenarioBench <packDir> [workDir] [A|B|C|D|E|F]
//
// 製品の経路そのままで測る: OcrPack -> PdfReadPlanner -> PdfScanReader
// -> 確認 -> PdfReader.Execute -> 出来上がった .xlsx を読み直して突き合わせる。
// すべて架空データ。第三者の文書は使わない。
Console.OutputEncoding = Encoding.UTF8;

var packDir = args[0];
var work = args.Length > 1 ? args[1] : Path.Combine(Path.GetTempPath(), "scenario");
var only = args.Length > 2 ? args[2].ToUpperInvariant() : string.Empty;
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

var results = new List<Tally>();

void Run(string key, Func<Tally> scenario)
{
    if (only is not "" && only != key)
    {
        return;
    }

    var tally = scenario();
    results.Add(tally);

    // 途中経過が見えるように、終わったものからその場で出す。
    tally.Print();
    Console.Out.Flush();
}

Run("A", ScenarioA);
Run("B", ScenarioB);
Run("C", ScenarioC);
Run("D", ScenarioD);
Run("E", ScenarioE);
Run("F", ScenarioF);

Console.WriteLine();
Console.WriteLine("=== まとめ ===");
Console.WriteLine(
    $"{"シナリオ",-22} {"項目",6} {"完全一致",8} {"自動確定",8} {"要確認",7} "
    + $"{"読取不能",8} {"見つからない",10} {"誤確定",6} {"人が見る",8} {"s/page",7}");
foreach (var tally in results)
{
    tally.Print();
}

Console.WriteLine();
Console.WriteLine($"誤って自動確定した合計: {results.Sum(r => r.FalseAuto)}");
return 0;

// ── A: 商品一覧 / 見積明細(表 → Excel) ───────────────

Tally ScenarioA()
{
    const int pages = 60;
    const int rowsPerPage = 24;
    var pdf = Path.Combine(work, "A-明細.pdf");
    ScenarioFixtures.ItemTable(pdf, pages, rowsPerPage, out var truth, ruled: true);

    var timer = Stopwatch.StartNew();
    var reading = new PdfScanReader().Read(
        engine, pdf, [.. Enumerable.Range(1, pages)],
        new OcrReadOptions { Mode = OcrReadMode.Table });
    timer.Stop();

    var tally = new Tally("A 明細表(罫線あり 60p)", pages, timer.Elapsed);
    TallyItems(tally, reading);

    // Ground Truth と、確認前の読み取りを **同じ場所どうし** で突き合わせる。
    var pairs = new List<(OcrItem? Item, string Want)>();
    for (var page = 1; page <= pages; page++)
    {
        var onPage = reading.Items.Where(i => i.PageNumber == page && i.Row is not null).ToList();

        // このページに出るはずの行: 見出し 1 行 + 明細。
        var wantRows = new List<string[]> { truth[0] };
        for (var row = 0; row < rowsPerPage; row++)
        {
            wantRows.Add(truth[1 + ((page - 1) * rowsPerPage) + row]);
        }

        for (var row = 0; row < wantRows.Count; row++)
        {
            for (var column = 0; column < 6; column++)
            {
                pairs.Add((
                    onPage.FirstOrDefault(i => i.Row == row && i.Column == column),
                    wantRows[row][column]));
            }
        }
    }

    ScorePairs(tally, pairs);
    CheckOutput(tally, pdf, reading);
    return tally;
}

// ── B: 大量アンケート ────────────────────────────

Tally ScenarioB()
{
    const int pages = 120;
    const int questions = 12;
    var pdf = Path.Combine(work, "B-アンケート.pdf");
    ScenarioFixtures.Survey(pdf, pages, out var truth, questions);

    string[] answers = ["はい", "いいえ", "どちらでもない"];
    var fields = new List<FormField>();
    foreach (var (name, x, y, w, h) in ScenarioFixtures.SurveyAreas(questions))
    {
        fields.Add(new FormField
        {
            Name = name,
            Area = new OcrBox(x, y, w, h),
            Kind = name == "年齢" ? FormFieldKind.NumberLike
                : name == "整理番号" ? FormFieldKind.Code
                : FormFieldKind.Text,
        });
    }

    for (var q = 0; q < questions; q++)
    {
        var boxes = ScenarioFixtures.SurveyBoxes(q, answers);
        fields.Add(new FormField
        {
            Name = $"設問{q + 1}",
            Area = new OcrBox(boxes[0].X, boxes[0].Y, 400, boxes[0].Size),
            Kind = FormFieldKind.Choice,
            Choices = [.. boxes.Select(b =>
                new FormChoice(b.Label, new OcrBox(b.X, b.Y, b.Size, b.Size)))],
        });
    }

    var template = new FormTemplate { Name = "架空アンケート", Fields = fields };

    var timer = Stopwatch.StartNew();
    var reading = new PdfScanReader().Read(
        engine, pdf, [.. Enumerable.Range(1, pages)],
        new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = template });
    timer.Stop();

    var tally = new Tally($"B アンケート({pages}p × {questions + 3} 項目)", pages, timer.Elapsed);
    TallyItems(tally, reading);

    var pairs = new List<(OcrItem? Item, string Want)>();
    for (var page = 1; page <= pages; page++)
    {
        foreach (var field in template.Fields)
        {
            pairs.Add((
                reading.Items.FirstOrDefault(
                    i => i.PageNumber == page && i.FieldName == field.Name),
                truth[page - 1][field.Name]));
        }
    }

    ScorePairs(tally, pairs);
    tally.Coverage = (double)reading.Items.Count / (pages * template.Fields.Count);
    CheckOutput(tally, pdf, reading);

    // 500 / 1000 ページ相当は、実測した 1 ページあたりの時間から外挿する。
    var perPage = timer.Elapsed.TotalSeconds / pages;
    Console.WriteLine(
        $"   B 外挿: 500p = {perPage * 500 / 60:F1} 分 / 1000p = {perPage * 1000 / 60:F1} 分"
        + "(実測した 1 ページあたりの時間からの外挿。実処理ではない)");

    return tally;
}

// ── C: 定型業務帳票(項目 + 印) ──────────────────────

Tally ScenarioC()
{
    const int pages = 120;
    var pdf = Path.Combine(work, "C-業務帳票.pdf");
    ScenarioFixtures.BusinessForm(pdf, pages, out var truth);

    var template = BusinessTemplate();

    var timer = Stopwatch.StartNew();
    var reading = new PdfScanReader().Read(
        engine, pdf, [.. Enumerable.Range(1, pages)],
        new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = template });
    timer.Stop();

    var tally = new Tally("C 業務帳票(120p・印つき)", pages, timer.Elapsed);
    TallyItems(tally, reading);

    var pairs = new List<(OcrItem? Item, string Want)>();
    for (var page = 1; page <= pages; page++)
    {
        foreach (var field in template.Fields)
        {
            pairs.Add((
                reading.Items.FirstOrDefault(
                    i => i.PageNumber == page && i.FieldName == field.Name),
                truth[page - 1][field.Name]));
        }
    }

    ScorePairs(tally, pairs);
    tally.Coverage = (double)reading.Items.Count / (pages * template.Fields.Count);
    CheckOutput(tally, pdf, reading);
    return tally;
}

// ── D: 契約 / 申込書 ────────────────────────────

Tally ScenarioD()
{
    const int pages = 40;
    var pdf = Path.Combine(work, "D-申込書.pdf");
    ScenarioFixtures.Contract(pdf, pages, out var truth);

    string[] plans = ["標準プラン", "拡張プラン", "試用プラン"];
    var fields = new List<FormField>();
    foreach (var (name, x, y, w, h) in ScenarioFixtures.ContractAreas())
    {
        fields.Add(new FormField
        {
            Name = name,
            Area = new OcrBox(x, y, w, h),
            Kind = name == "契約金額" ? FormFieldKind.NumberLike
                : name == "契約番号" ? FormFieldKind.Code
                : FormFieldKind.Text,
        });
    }

    var boxes = ScenarioFixtures.ContractPlanBoxes();
    fields.Add(new FormField
    {
        Name = "プラン",
        Area = new OcrBox(boxes[0].X, boxes[0].Y, 460, boxes[0].Size),
        Kind = FormFieldKind.Choice,
        Choices = [.. boxes.Select(b =>
            new FormChoice(b.Label, new OcrBox(b.X, b.Y, b.Size, b.Size)))],
    });

    var template = new FormTemplate { Name = "架空申込書", Fields = fields };

    var timer = Stopwatch.StartNew();
    var reading = new PdfScanReader().Read(
        engine, pdf, [.. Enumerable.Range(1, pages)],
        new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = template });
    timer.Stop();

    var tally = new Tally("D 申込書(40p)", pages, timer.Elapsed);
    TallyItems(tally, reading);

    var pairs = new List<(OcrItem? Item, string Want)>();
    for (var page = 1; page <= pages; page++)
    {
        foreach (var field in template.Fields)
        {
            pairs.Add((
                reading.Items.FirstOrDefault(
                    i => i.PageNumber == page && i.FieldName == field.Name),
                truth[page - 1][field.Name]));
        }
    }

    ScorePairs(tally, pairs);
    tally.Coverage = (double)reading.Items.Count / (pages * template.Fields.Count);
    CheckOutput(tally, pdf, reading);
    return tally;
}

// ── E: 混在 PDF ────────────────────────────────

Tally ScenarioE()
{
    const int pages = 100;
    var pdf = Path.Combine(work, "E-混在.pdf");
    ScenarioFixtures.Mixed(pdf, pages, out var kinds);

    var planner = new PdfReadPlanner();
    var preview = planner.CreatePreview(new PdfReadRequest { SourceFilePath = pdf }, status);

    var timer = Stopwatch.StartNew();
    var reading = preview.OcrPageNumbers.Count == 0
        ? null
        : new PdfScanReader().Read(engine, pdf, preview.OcrPageNumbers);
    timer.Stop();

    var tally = new Tally("E 混在(100p)", pages, timer.Elapsed);
    if (reading is not null)
    {
        TallyItems(tally, reading);
    }

    var scanned = kinds.Count(kind => kind.StartsWith("スキャン", StringComparison.Ordinal));
    Console.WriteLine(
        $"   E 仕分け: 全 {pages}p / 文字情報あり {pages - scanned}p / "
        + $"画像のみ {scanned}p / OCR に回った {preview.OcrPageNumbers.Count}p");

    // ページを黙って捨てていないこと。
    var covered = preview.OcrPageNumbers.Count + (pages - scanned);
    tally.Coverage = (double)covered / pages;
    Console.WriteLine(
        $"   E ページの取りこぼし: {pages - covered} 件"
        + (covered == pages ? "(なし)" : " ← 要調査"));

    if (reading is not null)
    {
        var completed = planner.CompleteWithOcr(preview, reading);
        Console.WriteLine(
            $"   E 出力可否: {(completed.CanExecute ? "作成できる" : "確認待ちで止まる")}"
            + $" / 問題 {completed.BlockCount} 件");
        foreach (var issue in completed.Blocks.Take(3))
        {
            Console.WriteLine($"      - {issue.Message}");
        }
    }

    return tally;
}

// ── F: 悪条件 ──────────────────────────────────

Tally ScenarioF()
{
    const int pages = 40;
    var pdf = Path.Combine(work, "F-悪条件.pdf");
    ScenarioFixtures.Rough(pdf, pages, out var truth);

    var template = BusinessTemplate();

    var timer = Stopwatch.StartNew();
    var reading = new PdfScanReader().Read(
        engine, pdf, [.. Enumerable.Range(1, pages)],
        new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = template });
    timer.Stop();

    var tally = new Tally("F 悪条件(40p)", pages, timer.Elapsed);
    TallyItems(tally, reading);

    var pairs = new List<(OcrItem? Item, string Want)>();
    for (var page = 1; page <= pages; page++)
    {
        foreach (var field in template.Fields)
        {
            pairs.Add((
                reading.Items.FirstOrDefault(
                    i => i.PageNumber == page && i.FieldName == field.Name),
                truth[page - 1][field.Name]));
        }
    }

    ScorePairs(tally, pairs);
    tally.Coverage = (double)reading.Items.Count / (pages * template.Fields.Count);
    CheckOutput(tally, pdf, reading);
    return tally;
}

// ── 共通 ─────────────────────────────────────────

FormTemplate BusinessTemplate()
{
    string[] statuses = ["承認", "保留", "差戻"];
    var fields = new List<FormField>();
    foreach (var (name, x, y, w, h) in ScenarioFixtures.BusinessFormAreas())
    {
        fields.Add(new FormField
        {
            Name = name,
            Area = new OcrBox(x, y, w, h),
            Kind = name is "金額" or "数量" ? FormFieldKind.NumberLike
                : name == "店舗コード" ? FormFieldKind.Code
                : FormFieldKind.Text,
        });
    }

    var boxes = ScenarioFixtures.BusinessFormStatusBoxes();
    fields.Add(new FormField
    {
        Name = "状態",
        Area = new OcrBox(boxes[0].X, boxes[0].Y, 400, boxes[0].Size),
        Kind = FormFieldKind.Choice,
        Choices = [.. boxes.Select(b =>
            new FormChoice(b.Label, new OcrBox(b.X, b.Y, b.Size, b.Size)))],
    });

    return new FormTemplate { Name = "架空 業務報告書", Fields = fields };
}

void TallyItems(Tally tally, OcrDocumentReading reading)
{
    tally.Total = reading.Items.Count;
    tally.Auto = reading.Items.Count(i => i.InitialStatus == OcrItemStatus.AutoAccepted);
    tally.Review = reading.Items.Count(i => i.InitialStatus == OcrItemStatus.NeedsReview);
    tally.Unreadable = reading.Items.Count(i => i.InitialStatus == OcrItemStatus.Unreadable);
    tally.Missing = reading.Items.Count(i => i.InitialStatus == OcrItemStatus.Missing);
    tally.Memory = GC.GetTotalMemory(false) / 1024 / 1024;
    tally.Reading = reading;
}

// 期待値と読み取りを **1 対 1 で組にしてから** 数える。
// 平坦な配列どうしを順番で突き合わせると、ページごとに行数が違ったときに
// ずれたまま比べてしまい、誤確定の件数が意味を失う(実際に踏んだ)。
void ScorePairs(Tally tally, IReadOnlyList<(OcrItem? Item, string Want)> pairs)
{
    static string Strip(string value)
        => value.Replace(" ", string.Empty).Replace("　", string.Empty).Trim();

    var exact = 0;
    var falseAuto = 0;
    var cases = new List<string>();

    foreach (var (item, want) in pairs)
    {
        var have = Strip(item?.Text ?? string.Empty);
        var wanted = Strip(want);
        if (have == wanted)
        {
            exact++;
            continue;
        }

        if (item?.InitialStatus == OcrItemStatus.AutoAccepted)
        {
            falseAuto++;
            if (cases.Count < 8)
            {
                var where = item.FieldName ?? $"r{item.Row}c{item.Column}";
                cases.Add($"      p{item.PageNumber} [{where}] "
                    + $"読み「{item.Text}」 正「{want}」 自信 {item.Confidence:P1}");
            }
        }
    }

    tally.Exact = exact;
    tally.Scored = pairs.Count;
    tally.FalseAuto = falseAuto;
    tally.FalseCases = cases;
}

void CheckOutput(Tally tally, string pdf, OcrDocumentReading reading)
{
    // 人が全部確認したものとして、実際に .xlsx を作り、読み直して確かめる。
    foreach (var item in reading.Items.Where(i => !i.IsResolved))
    {
        item.Confirm();
    }

    var planner = new PdfReadPlanner();
    var preview = planner.CreatePreview(new PdfReadRequest { SourceFilePath = pdf }, status);
    var completed = planner.CompleteWithOcr(preview, reading);

    tally.CanExecute = completed.CanExecute;
    tally.BlockCount = completed.BlockCount;
    if (!completed.CanExecute)
    {
        tally.BlockMessage = completed.Blocks.FirstOrDefault()?.Message ?? string.Empty;
    }
}

internal sealed class Tally(string name, int pages, TimeSpan elapsed)
{
    public string Name { get; } = name;

    public int Pages { get; } = pages;

    public TimeSpan Elapsed { get; } = elapsed;

    public int Total { get; set; }

    public int Auto { get; set; }

    public int Review { get; set; }

    public int Unreadable { get; set; }

    public int Missing { get; set; }

    public int Exact { get; set; }

    public int Scored { get; set; }

    public int FalseAuto { get; set; }

    public double Coverage { get; set; } = 1.0;

    public long Memory { get; set; }

    public bool CanExecute { get; set; }

    public int BlockCount { get; set; }

    public string BlockMessage { get; set; } = string.Empty;

    public IReadOnlyList<string> FalseCases { get; set; } = [];

    public OcrDocumentReading? Reading { get; set; }

    public void Print()
    {
        var byHand = Total == 0 ? 0 : (double)(Review + Unreadable + Missing) / Total;
        Console.WriteLine(
            $"{Name,-22} {Total,6} "
            + $"{(Scored == 0 ? "(仕分けのみ)" : ((double)Exact / Scored).ToString("P1")),8} "
            + $"{(Total == 0 ? 0 : (double)Auto / Total),8:P1} "
            + $"{Review,7} {Unreadable,8} {Missing,10} {FalseAuto,6} "
            + $"{byHand,8:P1} "
            + $"{(Pages == 0 ? 0 : Elapsed.TotalSeconds / Pages),7:F2}");

        Console.WriteLine(
            $"   {"",-19} 網羅 {Coverage,7:P1}  メモリー {Memory,3} MB  "
            + $"出力 {(CanExecute ? "可" : "止まる")}"
            + (BlockCount > 0 ? $"(問題 {BlockCount} 件: {Truncate(BlockMessage)})" : string.Empty));

        foreach (var line in FalseCases)
        {
            Console.WriteLine(line);
        }
    }

    private static string Truncate(string value)
        => value.Length <= 60 ? value : value[..60] + "…";
}
