using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Mapping;

/// <summary>データ元の種類。</summary>
public enum SourceFileKind
{
    Xlsx = 0,
    Csv,
}

/// <summary>「データ元のこの項目を、転記先のこのセルへ」1 件。</summary>
public sealed record SourceMappingRequest
{
    /// <summary>データ元の項目名(ヘッダーの文字列)。</summary>
    public required string SourceColumn { get; init; }

    /// <summary>転記先のセル。A1 形式の単一セルのみ。</summary>
    public required string TargetCell { get; init; }

    /// <summary>文字 / 数値のみ。空欄はデータ元から供給する値ではないので扱わない。</summary>
    public CellWriteKind WriteKind { get; init; } = CellWriteKind.Text;
}

/// <summary>表からの転記(Phase 2C1)の指定内容。</summary>
public sealed record SourceMappingBatchRequest
{
    /// <summary>データ元のファイル(.xlsx / .csv)。1 回の実行につき 1 つ。</summary>
    public required string SourceFilePath { get; init; }

    /// <summary>.xlsx の場合の読み取り元シート。</summary>
    public string? SourceSheetName { get; init; }

    /// <summary>項目名の行(1 始まり)。.csv では 1 固定。</summary>
    public int HeaderRow { get; init; } = 1;

    /// <summary>照合に使うデータ元の項目名。</summary>
    public required string KeyColumn { get; init; }

    /// <summary>転記先シートで、キーが入っているセル。</summary>
    public required string TargetKeyCell { get; init; }

    /// <summary>転記先(ファイルとシートの組)。</summary>
    public required IReadOnlyList<CellMutationTarget> Targets { get; init; }

    public required IReadOnlyList<SourceMappingRequest> Mappings { get; init; }

    public string OutputSuffix { get; init; } = SourceMappingDefaults.OutputSuffix;
}

/// <summary>Phase 2C1 の既定値。</summary>
public static class SourceMappingDefaults
{
    public const string OutputSuffix = "_転記済み";
}

/// <summary>データ元の項目名の読み取り結果(画面向け)。</summary>
public sealed record SourceColumnsResult
{
    public IReadOnlyList<string> Columns { get; init; } = Array.Empty<string>();

    /// <summary>CSV の場合に判定した文字コード(表示用)。</summary>
    public string? EncodingName { get; init; }

    public string? Error { get; init; }

    public bool IsSuccess => Error is null;
}
