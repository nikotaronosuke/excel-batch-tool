namespace ExcelBatchTool.Core.Merge;

/// <summary>統合プレビューで報告される問題の深刻度。</summary>
public enum MergeIssueSeverity
{
    /// <summary>統合は実行できるが、利用者が知っておくべきこと。</summary>
    Warning = 0,

    /// <summary>統合を実行できない。1 件でもあれば実行不可。</summary>
    Block = 1,
}

/// <summary>統合プレビューで報告される 1 件の問題。</summary>
public sealed record MergeIssue(
    MergeIssueSeverity Severity,
    string Message,
    string? FileName = null,
    string? SheetName = null)
{
    public string Glyph => Severity == MergeIssueSeverity.Block ? "✖" : "⚠";

    /// <summary>「ファイル名 / シート名」形式の場所表示。全体に関わる問題では空。</summary>
    public string Location => (FileName, SheetName) switch
    {
        (null, _) => string.Empty,
        (var file, null) => file!,
        var (file, sheet) => $"{file} / {sheet}",
    };
}

/// <summary>統合対象として選択された 1 つの Workbook / Worksheet。</summary>
public sealed record MergeSourceSelection(string FilePath, string SheetName);

/// <summary>統合の設定。</summary>
public sealed class MergeOptions
{
    /// <summary>出力先頭に「元ファイル」列を追加する。</summary>
    public bool IncludeSourceFileColumn { get; init; } = true;

    /// <summary>出力先頭に「元シート」列を追加する。</summary>
    public bool IncludeSourceSheetColumn { get; init; } = true;

    public string SourceFileColumnName { get; init; } = "元ファイル";

    public string SourceSheetColumnName { get; init; } = "元シート";

    public string OutputSheetName { get; init; } = "統合結果";
}

/// <summary>統合対象 1 件分の計画。</summary>
public sealed class MergeSourcePlan
{
    public required string FilePath { get; init; }

    public required string FileName { get; init; }

    public required string SheetName { get; init; }

    /// <summary>この Sheet の Header(trim 後、列順)。</summary>
    public IReadOnlyList<string> Headers { get; init; } = Array.Empty<string>();

    /// <summary>Header 以降の、完全空行を除いたデータ行数。</summary>
    public int DataRowCount { get; init; }

    /// <summary>この Sheet 単体で Block された(統合できない)。</summary>
    public bool IsBlocked { get; init; }

    /// <summary>
    /// この Sheet の Header 位置(0 始まり)から、出力データ列位置(0 始まり)への対応。
    /// metadata 列は含まない。Block された Sheet では空。
    /// </summary>
    public IReadOnlyList<int> ColumnMap { get; init; } = Array.Empty<int>();
}

/// <summary>統合前のプレビュー(検証結果)。</summary>
public sealed class MergePreview
{
    public required IReadOnlyList<MergeSourcePlan> Sources { get; init; }

    /// <summary>metadata 列を除いた出力列(基準 Sheet の Header 順)。</summary>
    public IReadOnlyList<string> DataHeaders { get; init; } = Array.Empty<string>();

    /// <summary>metadata 列を含む最終的な出力列。</summary>
    public IReadOnlyList<string> OutputHeaders { get; init; } = Array.Empty<string>();

    /// <summary>先頭に追加される metadata 列の数(0〜2)。</summary>
    public int MetadataColumnCount { get; init; }

    public int WorkbookCount => Sources.Count;

    public int SheetCount => Sources.Count;

    /// <summary>入力側のデータ行数の合計(完全空行を除く)。</summary>
    public int InputDataRowCount { get; init; }

    /// <summary>出力予定の行数(Header 1 行 + データ行)。</summary>
    public int OutputRowCount => InputDataRowCount + 1;

    public IReadOnlyList<MergeIssue> Issues { get; init; } = Array.Empty<MergeIssue>();

    public IEnumerable<MergeIssue> Blocks => Issues.Where(i => i.Severity == MergeIssueSeverity.Block);

    public IEnumerable<MergeIssue> Warnings => Issues.Where(i => i.Severity == MergeIssueSeverity.Warning);

    public int BlockCount => Blocks.Count();

    public int WarningCount => Warnings.Count();

    public bool HasBlocks => BlockCount > 0;

    /// <summary>統合を実行できるか。Block が 1 件でもあれば実行できない。</summary>
    public bool CanExecute => !HasBlocks && Sources.Count > 0 && OutputHeaders.Count > 0;
}

/// <summary>統合の実行結果。</summary>
public sealed class MergeExecutionResult
{
    public required bool Success { get; init; }

    public required string Message { get; init; }

    /// <summary>成功時の出力パス。</summary>
    public string? OutputPath { get; init; }

    public int WorkbookCount { get; init; }

    public int SheetCount { get; init; }

    /// <summary>出力したデータ行数(Header 行を含まない)。</summary>
    public int DataRowCount { get; init; }

    internal static MergeExecutionResult Failed(string message) => new()
    {
        Success = false,
        Message = message,
    };
}

/// <summary>統合の進捗通知。</summary>
public sealed record MergeProgress(int CompletedSources, int TotalSources, int WrittenDataRows);
