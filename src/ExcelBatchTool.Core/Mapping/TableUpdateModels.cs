using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Mapping;

/// <summary>「データ元のこの項目を、転記先のこの項目へ」1 件。</summary>
public sealed record TableColumnMappingRequest
{
    /// <summary>データ元の項目名(ヘッダーの文字列)。</summary>
    public required string SourceColumn { get; init; }

    /// <summary>転記先の項目名(ヘッダーの文字列)。名前は違ってよい。</summary>
    public required string TargetColumn { get; init; }

    /// <summary>文字 / 数値のみ。</summary>
    public CellWriteKind WriteKind { get; init; } = CellWriteKind.Text;
}

/// <summary>表同士の突合更新(Phase 2C2)の指定内容。</summary>
public sealed record TableUpdateBatchRequest
{
    /// <summary>データ元のファイル(.xlsx / .csv)。1 回の実行につき 1 つ。</summary>
    public required string SourceFilePath { get; init; }

    /// <summary>.xlsx の場合の読み取り元シート。</summary>
    public string? SourceSheetName { get; init; }

    /// <summary>データ元の項目名の行(1 始まり)。.csv では 1 固定。</summary>
    public int SourceHeaderRow { get; init; } = 1;

    /// <summary>照合に使うデータ元の項目名。</summary>
    public required string SourceKeyColumn { get; init; }

    /// <summary>転記先の項目名の行(1 始まり)。すべての転記先シートで同じ。</summary>
    public int TargetHeaderRow { get; init; } = 1;

    /// <summary>照合に使う転記先の項目名。データ元と名前が違ってよい。</summary>
    public required string TargetKeyColumn { get; init; }

    /// <summary>転記先(ファイルとシートの組)。</summary>
    public required IReadOnlyList<CellMutationTarget> Targets { get; init; }

    public required IReadOnlyList<TableColumnMappingRequest> Mappings { get; init; }

    public string OutputSuffix { get; init; } = TableUpdateDefaults.OutputSuffix;
}

/// <summary>Phase 2C2 の既定値。</summary>
public static class TableUpdateDefaults
{
    public const string OutputSuffix = "_更新済み";
}

/// <summary>キーの突合の集計(利用者が「何件更新されないか」を把握するための数字)。</summary>
public sealed record TableMatchSummary
{
    /// <summary>データ元でキーが入っていた行の数。</summary>
    public int SourceKeyedRowCount { get; init; }

    /// <summary>転記先でキーが入っていた行の数(全シート合計)。</summary>
    public int TargetKeyedRowCount { get; init; }

    /// <summary>両側に存在し、更新に使えるキーの数。</summary>
    public int MatchedKeyCount { get; init; }

    /// <summary>データ元にだけ存在するキーの数(行の追加はしない)。</summary>
    public int SourceOnlyKeyCount { get; init; }

    /// <summary>転記先にだけ存在するキーの数(行の削除・空欄化はしない)。</summary>
    public int TargetOnlyKeyCount { get; init; }

    /// <summary>どちらかの側で 2 件以上あったキーの数。</summary>
    public int DuplicateKeyCount { get; init; }

    /// <summary>キーが空欄で読み飛ばした行の数(両側合計)。</summary>
    public int BlankKeyRowCount { get; init; }
}

/// <summary>表同士の突合更新のプレビュー(共通の変更計画 + 突合の集計)。</summary>
public sealed class TableUpdatePreview
{
    /// <summary>実行に使う共通の変更計画(CellMutator へそのまま渡す)。</summary>
    public required CellMutationPreview Mutation { get; init; }

    public required TableMatchSummary Summary { get; init; }
}
