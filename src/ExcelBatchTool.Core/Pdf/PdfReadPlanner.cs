using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Pdf;

/// <summary>
/// PDF を読んで「Excel で扱える構造」にする計画を立てる。
/// データ元は読み取りのみ。全ページを安全に抽出できないときは、部分的な出力を作らない。
/// </summary>
public sealed class PdfReadPlanner
{
    public PdfReadPreview CreatePreview(PdfReadRequest request)
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

        // 判定の結果、この段階で扱えないものはここで止める(空のファイルを作らない)。
        switch (scan.Kind)
        {
            case PdfDocumentKind.Scan:
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    "この PDF はスキャン画像です。文字情報が含まれていないため、"
                        + "読み取りには OCR(次の段階で対応)が必要です。"));
                return WithKind(scan, issues, sourceFileName, request);

            case PdfDocumentKind.Mixed:
            {
                var imagePages = scan.Pages
                    .Where(page => page.Kind == PdfPageKind.ImageOnly)
                    .Select(page => page.Page)
                    .ToList();

                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"{scan.Pages.Count:N0} ページのうち {imagePages.Count:N0} ページ"
                        + $"({Describe(imagePages)})がスキャン画像です。OCR(次の段階で対応)が必要です。"
                        + "一部のページだけを取り出したファイルは作りません。"));
                return WithKind(scan, issues, sourceFileName, request);
            }

            case PdfDocumentKind.Unknown:
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block, "この PDF の内容を判定できませんでした。"));
                return WithKind(scan, issues, sourceFileName, request);
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
            return BuildTable(request, scan, issues, sourceFileName, outputFileName, outputPath,
                auditPath, snapshot);
        }

        return BuildText(request, scan, issues, sourceFileName, outputFileName, outputPath,
            auditPath, snapshot);
    }

    private static PdfReadPreview BuildText(
        PdfReadRequest request,
        PdfDocumentReader.PdfScan scan,
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
        };
    }

    private static PdfReadPreview BuildTable(
        PdfReadRequest request,
        PdfDocumentReader.PdfScan scan,
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
        PdfReadRequest request)
        => new()
        {
            Kind = scan.Kind,
            PageCount = scan.Pages.Count,
            Issues = issues,
            SourceFileName = sourceFileName,
            Request = request,
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
        };
    }
}
