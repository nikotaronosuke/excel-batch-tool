using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;
using ExcelBatchTool.Core.Ocr;

namespace ExcelBatchTool.Core.Pdf;

/// <summary>
/// PDF を読んで「Excel で扱える構造」にする計画を立てる。
/// データ元は読み取りのみ。全ページを安全に抽出できないときは、部分的な出力を作らない。
/// </summary>
public sealed class PdfReadPlanner
{
    /// <param name="pack">
    /// Offline OCR Pack の状態。null なら OCR は使えないものとして扱う
    /// (Pack が無くても、文字情報のある PDF はこれまでどおり読める)。
    /// </param>
    public PdfReadPreview CreatePreview(PdfReadRequest request, OcrPackStatus? pack = null)
    {
        var issues = new List<MergeIssue>();
        var sourceFileName = Path.GetFileName(request.SourceFilePath);

        if (!string.Equals(
            Path.GetExtension(request.SourceFilePath), ".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return Blocked(issues, sourceFileName, "PDF ファイル(.pdf)を選んでください。");
        }

        var (scan, failure) = PdfDocumentReader.Inspect(request.SourceFilePath);
        if (scan is null)
        {
            return Blocked(issues, sourceFileName, failure!.Message);
        }

        if (scan.Kind == PdfDocumentKind.Unknown)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, "この PDF の内容を判定できませんでした。"));
            return WithKind(scan, issues, sourceFileName, request, []);
        }

        var plans = scan.Pages
            .Select(page => new PdfPagePlan(page.Page, Route(page.Kind)))
            .ToList();

        // スキャンされたページがあるなら、まず OCR で読み取る段階へ進む。
        if (plans.Any(plan => plan.Route == PdfPageRoute.Scan))
        {
            return PlanOcr(request, scan, plans, issues, sourceFileName, pack);
        }

        var (outputPath, auditPath, outputFileName) = ResolveOutput(request, issues);

        SourceSnapshot snapshot;
        try
        {
            snapshot = MutationSnapshot.Take(request.SourceFilePath);
        }
        catch (Exception ex)
        {
            return Blocked(issues, sourceFileName, $"PDF ファイルを読み取れません: {ex.Message}");
        }

        if (scan.Kind == PdfDocumentKind.Table)
        {
            return BuildTable(request, scan, plans, issues, sourceFileName, outputFileName,
                outputPath, auditPath, snapshot);
        }

        return BuildText(request, scan, plans, issues, sourceFileName, outputFileName, outputPath,
            auditPath, snapshot);
    }

    private static PdfPageRoute Route(PdfPageKind kind) => kind switch
    {
        PdfPageKind.Text => PdfPageRoute.BornDigitalText,
        PdfPageKind.Table => PdfPageRoute.BornDigitalTable,
        PdfPageKind.ImageOnly => PdfPageRoute.Scan,
        _ => PdfPageRoute.Unknown,
    };

    /// <summary>
    /// スキャンされたページがある PDF の計画。
    ///
    /// OCR は数分かかるので、プレビューの段階では実行しない。ここでは
    /// 「OCR で読み取れる状態か」だけを決め、実際の読み取りは利用者が始める。
    /// </summary>
    private static PdfReadPreview PlanOcr(
        PdfReadRequest request,
        PdfDocumentReader.PdfScan scan,
        IReadOnlyList<PdfPagePlan> plans,
        List<MergeIssue> issues,
        string sourceFileName,
        OcrPackStatus? pack)
    {
        var scanPages = plans.Where(plan => plan.Route == PdfPageRoute.Scan).ToList();

        if (pack is null || !pack.IsUsable)
        {
            var what = scanPages.Count == plans.Count
                ? "この PDF はスキャン画像です。読み取りには OCR が必要です。"
                : $"{plans.Count:N0} ページのうち {scanPages.Count:N0} ページ"
                    + $"({Describe(scanPages.Select(plan => plan.Page).ToList())})がスキャン画像です。"
                    + "読み取りには OCR が必要です。";

            var why = pack?.Message ?? OcrPackStatus.Missing(string.Empty).Message;
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, what + why));
            return WithKind(scan, issues, sourceFileName, request, plans);
        }

        // 表のページとスキャンのページが混ざると、行と列の意味が揃わない。
        // 無理に 1 つの表へまとめず、この段階では扱わない。
        if (plans.Any(plan => plan.Route == PdfPageRoute.BornDigitalTable))
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                "表のページとスキャンのページが混在しています。"
                    + "行と列の意味を安全に揃えられないため、1 つの表にはまとめません。"));
            return WithKind(scan, issues, sourceFileName, request, plans);
        }

        return new PdfReadPreview
        {
            Kind = scan.Kind,
            PageCount = scan.Pages.Count,
            Issues = issues,
            SourceFileName = sourceFileName,
            Request = request,
            PagePlans = plans,
            Stage = issues.Any(issue => issue.Severity == MergeIssueSeverity.Block)
                ? PdfReadStage.Blocked
                : PdfReadStage.NeedsOcr,
        };
    }

    private static PdfReadPreview BuildText(
        PdfReadRequest request,
        PdfDocumentReader.PdfScan scan,
        IReadOnlyList<PdfPagePlan> plans,
        List<MergeIssue> issues,
        string sourceFileName,
        string outputFileName,
        string outputPath,
        string auditPath,
        SourceSnapshot snapshot)
    {
        var (lines, failure) = PdfDocumentReader.ReadLines(request.SourceFilePath);
        if (lines is null)
        {
            return Blocked(issues, sourceFileName, failure!.Message);
        }

        if (lines.Count == 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, "この PDF から文字を取り出せませんでした。"));
        }

        return new PdfReadPreview
        {
            Kind = PdfDocumentKind.Text,
            PageCount = scan.Pages.Count,
            Lines = lines,
            Issues = issues,
            SourceFileName = sourceFileName,
            OutputFileName = outputFileName,
            OutputPath = outputPath,
            AuditPath = auditPath,
            Snapshot = snapshot,
            Request = request,
            PagePlans = plans,
            Stage = PdfReadStage.Ready,
        };
    }

    private static PdfReadPreview BuildTable(
        PdfReadRequest request,
        PdfDocumentReader.PdfScan scan,
        IReadOnlyList<PdfPagePlan> plans,
        List<MergeIssue> issues,
        string sourceFileName,
        string outputFileName,
        string outputPath,
        string auditPath,
        SourceSnapshot snapshot)
    {
        var (table, failure) = PdfDocumentReader.ReadTable(request.SourceFilePath);
        if (table is null)
        {
            return Blocked(issues, sourceFileName, failure!.Message);
        }

        var rows = JoinPages(table, issues);

        if (rows.Count == 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, "この PDF から表を取り出せませんでした。"));
        }

        return new PdfReadPreview
        {
            Kind = PdfDocumentKind.Table,
            PageCount = scan.Pages.Count,
            TableRows = rows,
            TableFromRulings = table.FromRulings,
            Issues = issues,
            SourceFileName = sourceFileName,
            OutputFileName = outputFileName,
            OutputPath = outputPath,
            AuditPath = auditPath,
            Snapshot = snapshot,
            Request = request,
            PagePlans = plans,
            Stage = PdfReadStage.Ready,
        };
    }

    /// <summary>
    /// ページごとの表を 1 つにつなぐ。
    ///
    /// 2 ページ目以降の先頭行が 1 ページ目のヘッダーと**完全に一致**する場合だけ、
    /// 繰り返しのヘッダーとみなして落とす。少しでも違えばデータ行として残す
    /// (勝手な推測で行を消さない)。
    /// </summary>
    private static List<string[]> JoinPages(PdfTableResult table, List<MergeIssue> issues)
    {
        var rows = new List<string[]>();
        string[]? header = null;
        var removedHeaders = 0;

        foreach (var page in table.Pages)
        {
            if (page.Rows.Count == 0)
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Warning,
                    $"{page.Page} ページ目からは表を取り出せませんでした。"));
                continue;
            }

            var pageRows = page.Rows;

            if (header is null)
            {
                header = pageRows[0];
            }
            else if (pageRows[0].SequenceEqual(header, StringComparer.Ordinal))
            {
                pageRows = [.. pageRows.Skip(1)];
                removedHeaders++;
            }

            rows.AddRange(pageRows);
        }

        if (removedHeaders > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Warning,
                $"2 ページ目以降で同じ項目名の行が {removedHeaders:N0} 回繰り返されていたため、"
                    + "見出しの繰り返しとして 1 つにまとめました。"));
        }

        // 列数が揃わないページがあれば、いちばん広い列数に合わせて空欄で埋める。
        if (rows.Count > 0)
        {
            var width = rows.Max(row => row.Length);
            if (rows.Any(row => row.Length != width))
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Warning,
                    $"列の数がページによって違います。いちばん多い {width:N0} 列に合わせ、"
                        + "足りないところは空欄にしました。内容を確認してください。"));

                rows = rows
                    .Select(row => row.Length == width
                        ? row
                        : [.. row.Concat(Enumerable.Repeat(string.Empty, width - row.Length))])
                    .ToList();
            }
        }

        return rows;
    }

    /// <summary>出力先を決める。上書きはしない。</summary>
    private static (string OutputPath, string AuditPath, string OutputFileName) ResolveOutput(
        PdfReadRequest request, List<MergeIssue> issues)
    {
        var directory = Path.GetDirectoryName(Path.GetFullPath(request.SourceFilePath)) ?? string.Empty;
        var baseName = Path.GetFileNameWithoutExtension(request.SourceFilePath);
        var extension = request.OutputFormat == PdfOutputFormat.Xlsx ? ".xlsx" : ".csv";
        var outputFileName = baseName + request.OutputSuffix + extension;
        var outputPath = Path.Combine(directory, outputFileName);
        var auditPath = outputPath + ".audit.json";

        if (request.OutputSuffix.Length == 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                "出力名が空です。元のファイルと同じ名前にならないよう、付ける文字を入れてください。"));
        }
        else if (outputFileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"出力名「{request.OutputSuffix}」にファイル名として使えない文字が含まれています。"));
        }
        else if (File.Exists(outputPath))
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"「{outputFileName}」はすでにあります。既存のファイルは上書きしません。"));
        }
        else if (File.Exists(auditPath))
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"「{Path.GetFileName(auditPath)}」はすでにあります。既存のファイルは上書きしません。"));
        }

        return (outputPath, auditPath, outputFileName);
    }

    private static string Describe(IReadOnlyList<int> pages)
        => pages.Count <= 5
            ? string.Join(" / ", pages.Select(page => $"{page} ページ目"))
            : string.Join(" / ", pages.Take(5).Select(page => $"{page} ページ目")) + " ほか";

    private static PdfReadPreview WithKind(
        PdfDocumentReader.PdfScan scan,
        List<MergeIssue> issues,
        string sourceFileName,
        PdfReadRequest request,
        IReadOnlyList<PdfPagePlan> plans)
        => new()
        {
            Kind = scan.Kind,
            PageCount = scan.Pages.Count,
            Issues = issues,
            SourceFileName = sourceFileName,
            Request = request,
            PagePlans = plans,
            Stage = PdfReadStage.Blocked,
        };

    private static PdfReadPreview Blocked(
        List<MergeIssue> issues, string sourceFileName, string message)
    {
        issues.Add(new MergeIssue(MergeIssueSeverity.Block, message));
        return new PdfReadPreview
        {
            Kind = PdfDocumentKind.Unknown,
            Issues = issues,
            SourceFileName = sourceFileName,
            Stage = PdfReadStage.Blocked,
        };
    }

    /// <summary>
    /// OCR の結果と、文字情報のあるページを 1 つの「ページ / 行 / 内容」へまとめる。
    ///
    /// 出力できるのは、確認が必要な項目がひとつも残っていないときだけ。
    /// 未確認が残っている間は Block を付けたまま返す(黙って飛ばさない)。
    /// </summary>
    public PdfReadPreview CompleteWithOcr(PdfReadPreview preview, OcrDocumentReading reading)
    {
        if (preview.Request is not { } request)
        {
            return preview;
        }

        var issues = new List<MergeIssue>(preview.Issues.Where(
            issue => issue.Severity != MergeIssueSeverity.Block));
        issues.AddRange(reading.Issues);

        var lines = new List<PdfTextLine>();

        // 文字情報のあるページは OCR を通さない(元から正しい文字を画像化して落とさない)。
        var bornDigitalPages = preview.PagePlans
            .Where(plan => plan.Route == PdfPageRoute.BornDigitalText)
            .Select(plan => plan.Page)
            .ToHashSet();

        if (bornDigitalPages.Count > 0)
        {
            var (embedded, failure) = PdfDocumentReader.ReadLines(request.SourceFilePath);
            if (embedded is null)
            {
                return Blocked(issues, preview.SourceFileName, failure!.Message);
            }

            lines.AddRange(embedded.Where(line => bornDigitalPages.Contains(line.Page)));
        }

        lines.AddRange(ToLines(reading));

        var ordered = lines
            .OrderBy(line => line.Page)
            .ThenBy(line => line.Line)
            .ToList();

        if (reading.UnresolvedCount > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"確認が済んでいない項目が {reading.UnresolvedCount:N0} 件あります"
                    + $"(要確認 {reading.NeedsReviewCount:N0} 件 / "
                    + $"読取不能 {reading.UnreadableCount:N0} 件)。"
                    + "すべて確認してから出力してください。"));
        }

        if (ordered.Count == 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, "この PDF から文字を取り出せませんでした。"));
        }

        var (outputPath, auditPath, outputFileName) = ResolveOutput(request, issues);

        SourceSnapshot snapshot;
        try
        {
            snapshot = MutationSnapshot.Take(request.SourceFilePath);
        }
        catch (Exception ex)
        {
            return Blocked(issues, preview.SourceFileName, $"PDF ファイルを読み取れません: {ex.Message}");
        }

        return new PdfReadPreview
        {
            Kind = preview.Kind,
            PageCount = preview.PageCount,
            Lines = ordered,
            Issues = issues,
            SourceFileName = preview.SourceFileName,
            OutputFileName = outputFileName,
            OutputPath = outputPath,
            AuditPath = auditPath,
            Snapshot = snapshot,
            Request = request,
            PagePlans = preview.PagePlans,
            OcrReading = reading,
            Stage = issues.Any(issue => issue.Severity == MergeIssueSeverity.Block)
                ? PdfReadStage.Blocked
                : PdfReadStage.Ready,
        };
    }

    /// <summary>確認済みの読み取りを「ページ / 行 / 内容」にする。修正した文字を使う。</summary>
    internal static IReadOnlyList<PdfTextLine> ToLines(OcrDocumentReading reading)
        => reading.Items
            .GroupBy(item => (item.PageNumber, item.LineNumber))
            .OrderBy(group => group.Key.PageNumber)
            .ThenBy(group => group.Key.LineNumber)
            .Select(group => new PdfTextLine(
                group.Key.PageNumber,
                group.Key.LineNumber,
                JoinItems(group.OrderBy(item => item.IndexInLine).ToList())))
            .Where(line => line.Text.Length > 0)
            .ToList();

    /// <summary>
    /// 同じ行の項目をつなぐ。離れている項目のあいだには空白を 1 つ入れて、
    /// ラベルと値がくっつかないようにする。
    /// </summary>
    private static string JoinItems(IReadOnlyList<OcrItem> items)
    {
        var parts = new List<string>();
        double? previousRight = null;
        double previousHeight = 0;

        foreach (var item in items)
        {
            var text = item.FinalText;
            if (text.Length == 0)
            {
                continue;
            }

            if (previousRight is { } right)
            {
                var reference = Math.Max(Math.Max(previousHeight, item.BoundingBox.Height), 1);
                if (item.BoundingBox.X - right > reference * 0.5)
                {
                    parts.Add(" ");
                }
            }

            parts.Add(text);
            previousRight = item.BoundingBox.Right;
            previousHeight = item.BoundingBox.Height;
        }

        return string.Concat(parts).Trim();
    }
}
