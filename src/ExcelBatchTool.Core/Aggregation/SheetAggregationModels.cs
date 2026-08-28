using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>集約対象として選択された 1 つの Worksheet と、その出力シート名。</summary>
public sealed record SheetSelection(string FilePath, string SheetName, string? OutputSheetName = null);

/// <summary>集約対象 1 件分の計画。</summary>
public sealed class SheetAggregationPlan
{
    public required string FilePath { get; init; }

    public required string FileName { get; init; }

    public required string SheetName { get; init; }

    /// <summary>出力 Workbook 内でのシート名。</summary>
    public required string OutputSheetName { get; init; }

    /// <summary>元シートが非表示だった。出力でも非表示にする。</summary>
    public bool IsHidden { get; init; }

    /// <summary>このシート単体で集約できない。</summary>
    public bool IsBlocked { get; init; }

    /// <summary>出力 Workbook 内での並び順(1 始まり)。</summary>
    public int Order { get; init; }

    public string SourceDisplay => $"{FileName} / {SheetName}";
}

/// <summary>集約前のプレビュー(検証結果)。</summary>
public sealed class SheetAggregationPreview
{
    public required IReadOnlyList<SheetAggregationPlan> Sheets { get; init; }

    public IReadOnlyList<MergeIssue> Issues { get; init; } = Array.Empty<MergeIssue>();

    public int WorkbookCount => Sheets
        .Select(sheet => sheet.FilePath)
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .Count();

    public int SheetCount => Sheets.Count;

    public IEnumerable<MergeIssue> Blocks => Issues.Where(issue => issue.Severity == MergeIssueSeverity.Block);

    public IEnumerable<MergeIssue> Warnings => Issues.Where(issue => issue.Severity == MergeIssueSeverity.Warning);

    public int BlockCount => Blocks.Count();

    public int WarningCount => Warnings.Count();

    public bool HasBlocks => BlockCount > 0;

    /// <summary>集約を実行できるか。Block が 1 件でもあれば実行できない。</summary>
    public bool CanExecute => !HasBlocks && Sheets.Count > 0;
}

/// <summary>集約の実行結果。</summary>
public sealed class SheetAggregationResult
{
    public required bool Success { get; init; }

    public required string Message { get; init; }

    public string? OutputPath { get; init; }

    public int WorkbookCount { get; init; }

    public int SheetCount { get; init; }

    internal static SheetAggregationResult Failed(string message) => new()
    {
        Success = false,
        Message = message,
    };
}

/// <summary>集約の進捗通知。</summary>
public sealed record SheetAggregationProgress(int CompletedSheets, int TotalSheets);
