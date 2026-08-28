using System.Globalization;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Mutation;

/// <summary>一括変更で書き込める値の種類。</summary>
public enum CellWriteKind
{
    /// <summary>文字列(InlineString として書く)。</summary>
    Text = 0,

    /// <summary>数値。</summary>
    Number,

    /// <summary>値を消して空欄にする(セルと書式は残す)。</summary>
    Blank,
}

/// <summary>変更対象として選ばれた 1 つの Worksheet。</summary>
public sealed record CellMutationTarget(string FilePath, string SheetName);

/// <summary>入力セットの 1 項目。「どのセルへ、どんな種類の、どの値を入れるか」。</summary>
public sealed record CellMutationOperationRequest
{
    /// <summary>書き込む位置。A1 形式の単一セルのみ。</summary>
    public required string CellReference { get; init; }

    public CellWriteKind WriteKind { get; init; } = CellWriteKind.Text;

    /// <summary>文字列を書く場合の値。</summary>
    public string? TextValue { get; init; }

    /// <summary>数値を書く場合の値(利用者の入力文字列。検証は Planner が行う)。</summary>
    public string? NumberText { get; init; }
}

/// <summary>
/// 一括変更の指定内容。選択したすべてのシートへ、同じ入力セット
/// (複数セル × それぞれの値)を書き込む。
/// </summary>
public sealed record CellMutationRequest
{
    /// <summary>変更対象(ファイルとシートの組)。</summary>
    public required IReadOnlyList<CellMutationTarget> Targets { get; init; }

    /// <summary>入力セット。同じセルの重複指定は許可しない。</summary>
    public required IReadOnlyList<CellMutationOperationRequest> Operations { get; init; }

    /// <summary>出力ファイル名に付ける接尾辞。</summary>
    public string OutputSuffix { get; init; } = CellMutationDefaults.OutputSuffix;
}

/// <summary>Phase 2A の既定値。</summary>
public static class CellMutationDefaults
{
    public const string OutputSuffix = "_変更済み";

    /// <summary>変更内容を書き出す控えファイルの拡張子。</summary>
    public const string AuditExtension = ".audit.json";
}

/// <summary>実行直前に元ファイルが変わっていないか確かめるための控え。</summary>
public sealed record SourceSnapshot(string Sha256, long Length, DateTime LastWriteUtc);

/// <summary>変更対象 1 件(ファイル × シート × セル)の計画。</summary>
public sealed record CellMutationTargetPlan
{
    public required string FilePath { get; init; }

    public required string FileName { get; init; }

    public required string SheetName { get; init; }

    public required string CellReference { get; init; }

    /// <summary>この変更で書き込む値(解釈済み)。</summary>
    internal NewCellValue NewValue { get; init; }

    /// <summary>現在の値の表示用文字列。数式は評価しない。</summary>
    public string CurrentValueDisplay { get; init; } = string.Empty;

    /// <summary>現在の値の種類(控えファイルに書く名前)。</summary>
    public string CurrentTypeName { get; init; } = "blank";

    /// <summary>変更後の値の種類(控えファイルに書く名前)。</summary>
    public string NewTypeName { get; init; } = "blank";

    /// <summary>変更後の値の表示用文字列。</summary>
    public string NewValueDisplay { get; init; } = string.Empty;

    /// <summary>この値をどこから持ってきたか(データ元から転記した場合のみ)。</summary>
    public MutationProvenance? Provenance { get; init; }

    /// <summary>出力ファイル名(元と同じフォルダー)。</summary>
    public required string OutputFileName { get; init; }

    public string? BlockReason { get; init; }

    public bool IsBlocked => BlockReason is not null;

    /// <summary>現在の値と新しい値が型を含めて同じ(書き換え不要)。</summary>
    public bool IsNoOp { get; init; }

    public string StatusText => BlockReason ?? (IsNoOp ? "変更なし(現在の値と同じです)" : "変更できます");

    public string Glyph => IsBlocked ? "✖" : IsNoOp ? "－" : "✅";
}

/// <summary>出力 1 ファイル分の計画(同じ Workbook の複数シートをまとめる)。</summary>
public sealed record CellMutationFilePlan
{
    public required string FilePath { get; init; }

    public required string FileName { get; init; }

    public required string OutputFileName { get; init; }

    /// <summary>出力ファイルのフルパス。UI へは出さない。</summary>
    public required string OutputPath { get; init; }

    /// <summary>変更内容の控えファイルのフルパス。</summary>
    public required string AuditPath { get; init; }

    public required SourceSnapshot Snapshot { get; init; }

    /// <summary>転記元のデータファイル(データ元から転記した場合のみ)。</summary>
    public MutationDataSourceInfo? DataSource { get; init; }

    /// <summary>転記先の表の読み方(表同士の突合更新の場合のみ)。</summary>
    public MutationTargetTableInfo? TargetTable { get; init; }

    /// <summary>このファイルで実際に書き換えるシート(No-op と Block を除く)。</summary>
    public IReadOnlyList<CellMutationTargetPlan> Changes { get; init; }
        = Array.Empty<CellMutationTargetPlan>();
}

/// <summary>
/// この値をデータ元のどこから取ったか。表同士の突合更新では、
/// 転記先のどの列・どの行を更新したかも持つ。
/// </summary>
public sealed record MutationProvenance(
    string SourceColumn,
    string Key,
    int SourceRowNumber,
    string? TargetColumn = null,
    int? TargetRowNumber = null);

/// <summary>転記先の表の読み方(控えファイルに残す)。</summary>
public sealed record MutationTargetTableInfo(int HeaderRow, string KeyColumn);

/// <summary>転記に使ったデータ元の情報(控えファイルに残す)。パスは持たない。</summary>
public sealed record MutationDataSourceInfo
{
    public required string FileName { get; init; }

    public required string Sha256 { get; init; }

    /// <summary>"xlsx" または "csv"。</summary>
    public required string Type { get; init; }

    /// <summary>.xlsx の場合の読み取り元シート。</summary>
    public string? SheetName { get; init; }

    /// <summary>項目名の行(1 始まり)。</summary>
    public int HeaderRow { get; init; }

    /// <summary>照合に使った項目名。</summary>
    public required string KeyColumn { get; init; }
}

/// <summary>実行直前にデータ元が変わっていないか確かめるための控え。</summary>
internal sealed record MutationDataSourceCheck(string FilePath, string FileName, SourceSnapshot Snapshot);

/// <summary>実行前のプレビュー。</summary>
public sealed class CellMutationPreview
{
    public required IReadOnlyList<CellMutationTargetPlan> Targets { get; init; }

    /// <summary>実際に出力を作るファイル(変更が 1 件もないファイルは含まない)。</summary>
    public required IReadOnlyList<CellMutationFilePlan> Files { get; init; }

    public IReadOnlyList<MergeIssue> Issues { get; init; } = Array.Empty<MergeIssue>();

    /// <summary>実行直前に読み直して変化を確かめるデータ元(転記の場合のみ)。</summary>
    internal MutationDataSourceCheck? DataSourceCheck { get; init; }

    public IEnumerable<MergeIssue> Blocks => Issues.Where(issue => issue.Severity == MergeIssueSeverity.Block);

    public IEnumerable<MergeIssue> Warnings => Issues.Where(issue => issue.Severity == MergeIssueSeverity.Warning);

    public int BlockCount => Blocks.Count();

    public int WarningCount => Warnings.Count();

    public int ChangeCount => Files.Sum(file => file.Changes.Count);

    public int NoOpCount => Targets.Count(target => !target.IsBlocked && target.IsNoOp);

    public int OutputFileCount => Files.Count;

    public bool HasBlocks => BlockCount > 0;

    /// <summary>実行できるか。Block が 1 件でもあれば実行できない。変更が 0 件でも実行しない。</summary>
    public bool CanExecute => !HasBlocks && ChangeCount > 0;
}

/// <summary>一括変更の実行結果。</summary>
public sealed class CellMutationResult
{
    public required bool Success { get; init; }

    public required string Message { get; init; }

    /// <summary>作成した出力ファイル名(フルパスではない)。</summary>
    public IReadOnlyList<string> OutputFileNames { get; init; } = Array.Empty<string>();

    public int ChangedCellCount { get; init; }

    internal static CellMutationResult Failed(string message) => new()
    {
        Success = false,
        Message = message,
    };
}

/// <summary>一括変更の進捗通知。</summary>
public sealed record CellMutationProgress(int CompletedFiles, int TotalFiles);

/// <summary>セル値の表示用文字列。数式は評価しない。</summary>
internal static class CellValueDisplay
{
    public const string Blank = "(空欄)";

    public static string Of(MergeCellValue value) => value.Kind switch
    {
        MergeValueKind.Blank => Blank,
        MergeValueKind.Text => value.Text ?? Blank,
        MergeValueKind.Boolean => value.Boolean ? "TRUE" : "FALSE",
        MergeValueKind.Number => value.Number.ToString(CultureInfo.InvariantCulture),
        MergeValueKind.Time => MergeCellValue.SerialToDateTime(value.Number)?
            .ToString("HH:mm:ss", CultureInfo.InvariantCulture) ?? Number(value),
        MergeValueKind.Date => MergeCellValue.SerialToDateTime(value.Number)?
            .ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? Number(value),
        _ => MergeCellValue.SerialToDateTime(value.Number)?
            .ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) ?? Number(value),
    };

    private static string Number(MergeCellValue value)
        => value.Number.ToString(CultureInfo.InvariantCulture);

    /// <summary>控えファイルに書く型名。</summary>
    public static string TypeNameOf(MergeCellValue value) => value.Kind switch
    {
        MergeValueKind.Text => "text",
        MergeValueKind.Number => "number",
        MergeValueKind.Boolean => "boolean",
        MergeValueKind.Date => "date",
        MergeValueKind.DateTime => "datetime",
        MergeValueKind.Time => "time",
        _ => "blank",
    };
}
