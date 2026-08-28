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

/// <summary>出力へ引き継ぐ入力規則 1 件の概要(出力後の検証に使う)。</summary>
public sealed record DataValidationSummary(
    string Sqref,
    string Type,
    string? Formula1,
    string? Formula2);

/// <summary>出力へ作る、ブック全体を対象とする名前定義(候補一覧の参照先)。</summary>
public sealed record OutputDefinedName(string Name, string RefersTo);

/// <summary>参照元を解決済みの x14 リスト入力規則。</summary>
public sealed record ResolvedX14ListValidation(string Sqref, string ListSource);

/// <summary>
/// 出力へ引き継ぐ条件付き書式のルール 1 件の概要(出力後の検証に使う)。
/// 値が null の項目は「元ファイルに属性が無い」ことを表す。出力にも書かない。
/// </summary>
public sealed record ConditionalFormattingRuleSummary
{
    /// <summary>ルールの種類(duplicateValues / uniqueValues / top10 / aboveAverage)。</summary>
    public required string Type { get; init; }

    /// <summary>優先順位。元の値をそのまま使う(振り直さない)。</summary>
    public required int Priority { get; init; }

    /// <summary>このルールが成立したら以降のルールを評価しない。</summary>
    public bool? StopIfTrue { get; init; }

    /// <summary>top10: 上位/下位の件数(または割合)。</summary>
    public uint? Rank { get; init; }

    /// <summary>top10: 件数ではなく割合で指定する。</summary>
    public bool? Percent { get; init; }

    /// <summary>top10: 上位ではなく下位を対象にする。</summary>
    public bool? Bottom { get; init; }

    /// <summary>aboveAverage: 平均より上(false なら平均より下)。</summary>
    public bool? AboveAverage { get; init; }

    /// <summary>aboveAverage: 平均と等しい値も含める。</summary>
    public bool? EqualAverage { get; init; }

    /// <summary>aboveAverage: 標準偏差いくつ分か。</summary>
    public int? StandardDeviation { get; init; }

    /// <summary>書式に含まれる項目("font,fill" など)。出力後の照合に使う。</summary>
    public string FormatChildren { get; init; } = string.Empty;

    /// <summary>書式の表示形式(formatCode)。指定が無ければ null。</summary>
    public string? FormatNumberCode { get; init; }
}

/// <summary>出力へ引き継ぐ条件付き書式(1 つの適用範囲とそのルール群)の概要。</summary>
public sealed record ConditionalFormattingSummary
{
    /// <summary>適用範囲(空白区切りの A1 形式)。元の範囲をそのまま使う。</summary>
    public required string Sqref { get; init; }

    public IReadOnlyList<ConditionalFormattingRuleSummary> Rules { get; init; }
        = Array.Empty<ConditionalFormattingRuleSummary>();
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

    /// <summary>出力へ引き継ぐ入力規則の概要。</summary>
    public IReadOnlyList<DataValidationSummary> DataValidations { get; init; }
        = Array.Empty<DataValidationSummary>();

    /// <summary>出力へ引き継ぐ x14 リスト入力規則(参照元は解決済み)。</summary>
    public IReadOnlyList<ResolvedX14ListValidation> X14ListValidations { get; init; }
        = Array.Empty<ResolvedX14ListValidation>();

    /// <summary>出力へ引き継ぐ条件付き書式の概要。</summary>
    public IReadOnlyList<ConditionalFormattingSummary> ConditionalFormattings { get; init; }
        = Array.Empty<ConditionalFormattingSummary>();

    public string SourceDisplay => $"{FileName} / {SheetName}";
}

/// <summary>集約前のプレビュー(検証結果)。</summary>
public sealed class SheetAggregationPreview
{
    public required IReadOnlyList<SheetAggregationPlan> Sheets { get; init; }

    public IReadOnlyList<MergeIssue> Issues { get; init; } = Array.Empty<MergeIssue>();

    /// <summary>出力ブックへ作る、ブック全体を対象とする名前定義。</summary>
    public IReadOnlyList<OutputDefinedName> DefinedNames { get; init; } = Array.Empty<OutputDefinedName>();

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
