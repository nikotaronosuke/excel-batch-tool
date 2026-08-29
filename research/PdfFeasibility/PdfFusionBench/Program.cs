using PdfBench;

// Phase 2F-B1: 二重読み統合方式を GT の完全一致率で選ぶ。すべて架空データ。
if (args.Length < 2)
{
    Console.WriteLine("usage: PdfFusionBench <work-dir> <capture|fuse>");
    return 1;
}

var workDir = Path.GetFullPath(args[0]);
var fixtures = Path.Combine(workDir, "fixtures");
var outDir = Path.Combine(workDir, "fusion");
Directory.CreateDirectory(outDir);

// 対象: 帳票 clean 120p / 帳票 degraded 30p / 文章 clean 5p / 文章 degraded 5p。
var targets = new (string Pdf, string Gt, int Dpi, int MaxPages, string Key)[]
{
    ("form-scan-clean.pdf", "form.json", 300, 120, "form-clean"),
    ("form-scan-degraded.pdf", "form.json", 150, 30, "form-degraded"),
    ("text-scan-clean.pdf", "text.json", 300, 5, "text-clean"),
    ("text-scan-degraded.pdf", "text.json", 150, 5, "text-degraded"),
};

if (args[1] == "capture")
{
    foreach (var target in targets)
    {
        Console.WriteLine($"capture {target.Key} ...");
        var capture = OcrCapture.Run(
            Path.Combine(fixtures, target.Pdf), target.Dpi, target.MaxPages);
        Json.Save(Path.Combine(outDir, target.Key + ".json"), capture);
        Console.WriteLine(
            $"  {capture.Regions.Count} regions, {capture.Seconds:F1}s " +
            $"({capture.Seconds / capture.Pages:F2} s/page)");
    }

    return 0;
}

if (args[1] != "fuse")
{
    Console.WriteLine("unknown command");
    return 1;
}

// 実測で選んだ統合方式。per-field の内訳はこの方式についてだけ出す。
const string Chosen = "agree-then-charclass-strict";

var report = new List<object>();

foreach (var target in targets)
{
    var capturePath = Path.Combine(outDir, target.Key + ".json");
    if (!File.Exists(capturePath))
    {
        Console.WriteLine($"skip {target.Key}: no capture");
        continue;
    }

    var capture = Json.Load<DualCapture>(capturePath);
    var fields = LoadFields(Path.Combine(fixtures, "gt", target.Gt), capture.Pages);

    Console.WriteLine();
    Console.WriteLine($"### {target.Key}  ({capture.Pages}p, {fields.Count} fields, " +
        $"{capture.Seconds / capture.Pages:F2} s/page)");
    Console.WriteLine(
        "strategy                   thr   exact   auto   falseAuto  reviewCatch  missed");
    Console.WriteLine(
        "-------------------------- ----- ------- ------ ---------- ------------ ------");

    foreach (var strategy in Fusion.All)
    {
        foreach (var threshold in new[] { 0.90, 0.95, 0.98 })
        {
            var fused = capture.Regions
                .Select(region => strategy.Fuse(region, threshold))
                .ToList();

            var byPage = fused
                .GroupBy(region => region.Page)
                .ToDictionary(group => group.Key, group => group.ToList());

            var total = 0;
            var exact = 0;
            var autoAccepted = 0;
            var falseAutoAccepted = 0;
            var reviewCaught = 0;
            var missed = 0;
            var perField = new Dictionary<string, (int Total, int Exact)>(StringComparer.Ordinal);

            foreach (var (page, name, value) in fields)
            {
                if (!byPage.TryGetValue(page, out var pageRegions))
                {
                    continue;
                }

                total++;
                var pageText = string.Concat(pageRegions.Select(region => region.Text));
                var correct = TextMetrics.AppearsExactly(pageText, value);
                if (correct)
                {
                    exact++;
                }

                var stat = perField.GetValueOrDefault(name);
                perField[name] = (stat.Total + 1, stat.Exact + (correct ? 1 : 0));

                // その値を読んだはずの領域 = 値にいちばん近い領域。
                var owner = pageRegions
                    .OrderByDescending(region => TextMetrics.CharacterAccuracy(value, region.Text))
                    .First();

                if (owner.Status == FusedStatus.AutoAccepted)
                {
                    autoAccepted++;
                    if (!correct)
                    {
                        falseAutoAccepted++;
                    }
                }
                else if (!correct)
                {
                    reviewCaught++;
                }

                if (!correct && owner.Status == FusedStatus.AutoAccepted)
                {
                    missed++;
                }

            }

            var wrong = total - exact;
            Console.WriteLine(
                $"{strategy.Name,-26} {threshold:0.00}  " +
                $"{(double)exact / total,6:P1}  {(double)autoAccepted / total,5:P0}  " +
                $"{(double)falseAutoAccepted / total,9:P2}  " +
                $"{(wrong == 0 ? 1 : (double)reviewCaught / wrong),11:P1}  {missed,5}");

            if (strategy.Name == Chosen && threshold == 0.98)
            {
                foreach (var (field, stat) in perField.OrderBy(entry => entry.Key, StringComparer.Ordinal))
                {
                    Console.WriteLine(
                        $"      field {field,-12} {(double)stat.Exact / stat.Total,7:P1}  ({stat.Exact}/{stat.Total})");
                }
            }

            report.Add(new
            {
                target = target.Key,
                strategy = strategy.Name,
                threshold,
                fields = total,
                exact = (double)exact / total,
                autoAccepted = (double)autoAccepted / total,
                falseAutoAccepted = (double)falseAutoAccepted / total,
                falseAutoAcceptedCount = falseAutoAccepted,
                reviewCatchRate = wrong == 0 ? 1 : (double)reviewCaught / wrong,
            });
        }
    }
}

Json.Save(Path.Combine(outDir, "fusion-report.json"), report);
Console.WriteLine();
Console.WriteLine("saved fusion-report.json");
return 0;

static List<(int Page, string Name, string Value)> LoadFields(string gtPath, int pages)
{
    var result = new List<(int, string, string)>();
    if (gtPath.EndsWith("form.json", StringComparison.Ordinal))
    {
        foreach (var page in Json.Load<List<FormPageGt>>(gtPath).Take(pages))
        {
            foreach (var (name, value) in page.Fields.Where(f => f.Value.Length > 0))
            {
                result.Add((page.Page, name, value));
            }
        }
    }
    else
    {
        foreach (var page in Json.Load<List<TextPageGt>>(gtPath).Take(pages))
        {
            foreach (var (name, value) in page.Fields.Where(f => f.Value.Length > 0))
            {
                result.Add((page.Page, name, value));
            }
        }
    }

    return result;
}
