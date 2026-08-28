using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>集約対象として選択された 1 つの Worksheet と、その出力シート名。</summary>
public sealed record SheetSelection(string FilePath, string SheetName, string? OutputSheetName = null);

/// <summary>
/// 出力へ引き継ぐ印刷・ページレイアウト情報の概要。出力後の検証でも使う。
/// 範囲文字列はシート名を含まない(出力シート名で組み立て直す)。
/// </summary>
public sealed record PrintLayoutSummary
{
    public bool HasPageSetupProperties { get; init; }

    public bool HasPrintOptions { get; init; }

    public bool HasPageMargins { get; init; }

    public bool HasPageSetup { get; init; }

    public bool HasHeaderFooter { get; init; }

    public int RowBreakCount { get; init; }

    public int ColumnBreakCount { get; init; }

    public IReadOnlyList<string> PrintAreaRanges { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> PrintTitleRanges { get; init; } = Array.Empty<string>();

    public bool HasPrintArea => PrintAreaRanges.Count > 0;

    public bool HasPrintTitles => PrintTitleRanges.Count > 0;

    public bool IsEmpty => !HasPageSetupProperties && !HasPrintOptions && !HasPageMargins
        && !HasPageSetup && !HasHeaderFooter && RowBreakCount == 0 && ColumnBreakCount == 0
        && !HasPrintArea && !HasPrintTitles;
}

/// <summary>
/// 出力へ書き込む準備が整ったハイパーリンク 1 件。
/// 別シート宛のリンク先は、出力シート名で組み立て直したあとの文字列。
/// </summary>
public sealed record ResolvedHyperlink
{
    public required string Reference { get; init; }

    /// <summary>Web / メールへの外部リンクの絶対 URI。内部リンクでは null。</summary>
    public string? ExternalTarget { get; init; }

    /// <summary>内部リンクの参照先、または外部リンクの文書内アンカー。</summary>
    public string? Location { get; init; }

    public string? Tooltip { get; init; }

    public string? Display { get; init; }

    public bool IsExternal => ExternalTarget is not null;
}

/// <summary>集約対象 1 件分の計画。</summary>
public sealed class SheetAggregationPlan
{
    public required string FilePath { get; init; }

    public required string FileName { get; init; }

    public required string SheetName { get; init; }

    /// <summary>出力 Workbook 内でのシート名。</summary>
    public required string OutputSheetName { get; init; }

    /// <summary>元シートの表示状態。出力でもそのまま再現する。</summary>
    public SheetVisibility Visibility { get; init; } = SheetVisibility.Visible;

    /// <summary>表示されていない(hidden または veryHidden)。</summary>
    public bool IsHidden => Visibility != SheetVisibility.Visible;

    /// <summary>表示状態の日本語表記。</summary>
    public string VisibilityDisplay => Visibility switch
    {
        SheetVisibility.Hidden => "非表示",
        SheetVisibility.VeryHidden => "非常に非表示",
        _ => "表示",
    };

    /// <summary>このシート単体で集約できない。</summary>
    public bool IsBlocked { get; init; }

    /// <summary>出力 Workbook 内での並び順(1 始まり)。</summary>
    public int Order { get; init; }

    /// <summary>出力へ引き継ぐ印刷・ページレイアウト情報。</summary>
    public PrintLayoutSummary PrintLayout { get; init; } = new();

    /// <summary>出力へ引き継ぐハイパーリンク(リンク先は解決済み)。</summary>
    public IReadOnlyList<ResolvedHyperlink> Hyperlinks { get; init; } = Array.Empty<ResolvedHyperlink>();

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
