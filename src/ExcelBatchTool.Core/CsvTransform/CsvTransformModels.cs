using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.CsvTransform;

/// <summary>出力する 1 列に、何を入れるか。</summary>
public enum CsvValueSourceKind
{
    /// <summary>データ元の項目をそのまま入れる。</summary>
    SourceColumn = 0,

    /// <summary>すべての行に同じ文字を入れる。</summary>
    FixedText,

    /// <summary>すべての行を空欄にする。</summary>
    Blank,
}

/// <summary>作る CSV の文字コード。</summary>
public enum CsvOutputEncoding
{
    /// <summary>UTF-8(BOM あり)。Windows の Excel で開いても文字化けしにくい。</summary>
    Utf8Bom = 0,

    /// <summary>UTF-8(BOM なし)。</summary>
    Utf8,

    /// <summary>Shift_JIS(CP932)。</summary>
    ShiftJis,
}

/// <summary>引用符の付け方。</summary>
public enum CsvQuoteMode
{
    /// <summary>区切り記号・引用符・改行を含む項目だけを囲む。</summary>
    Minimal = 0,

    /// <summary>すべての項目を囲む。</summary>
    All,
}

/// <summary>出力する 1 列の指定。</summary>
public sealed record CsvOutputColumnRequest
{
    /// <summary>出力する CSV の項目名。</summary>
    public required string OutputName { get; init; }

    public CsvValueSourceKind ValueSourceKind { get; init; }

    /// <summary>データ元から取るときの項目名。</summary>
    public string? SourceColumn { get; init; }

    /// <summary>固定値のときに全行へ入れる文字。</summary>
    public string? FixedValue { get; init; }
}

/// <summary>CSV 変換 1 回分の指定。</summary>
public sealed record CsvTransformRequest
{
    public required string SourceFilePath { get; init; }

    /// <summary>.xlsx のデータ元で読むシート。CSV では使わない。</summary>
    public string? SourceSheetName { get; init; }

    /// <summary>項目名の行(1 始まり)。CSV では 1 行目固定。</summary>
    public int HeaderRow { get; init; } = 1;

    public IReadOnlyList<CsvOutputColumnRequest> Columns { get; init; } = [];

    public CsvOutputEncoding Encoding { get; init; } = CsvOutputEncoding.Utf8Bom;

    public CsvQuoteMode QuoteMode { get; init; } = CsvQuoteMode.Minimal;

    public string OutputSuffix { get; init; } = CsvTransformDefaults.OutputSuffix;
}

/// <summary>CSV 変換の既定値。</summary>
public static class CsvTransformDefaults
{
    public const string OutputSuffix = "_変換済み";

    /// <summary>プレビューに載せる行数。全行を画面へ展開しない。</summary>
    public const int SampleRowCount = 20;
}

/// <summary>プレビューに出す「出力する 1 列」の説明。</summary>
public sealed record CsvOutputColumnPlan
{
    public required string OutputName { get; init; }

    public required CsvValueSourceKind ValueSourceKind { get; init; }

    public string? SourceColumn { get; init; }

    public string? FixedValue { get; init; }

    /// <summary>「元の項目: 価格」「固定値: 1」のような表示。</summary>
    public string SourceDisplay => ValueSourceKind switch
    {
        CsvValueSourceKind.SourceColumn => $"元の項目: {SourceColumn}",
        CsvValueSourceKind.FixedText => $"固定値: {FixedValue}",
        _ => "空欄",
    };
}

/// <summary>CSV 変換のプレビュー。</summary>
public sealed class CsvTransformPreview
{
    public required IReadOnlyList<CsvOutputColumnPlan> Columns { get; init; }

    /// <summary>データ元の項目名。</summary>
    public IReadOnlyList<string> SourceColumns { get; init; } = [];

    /// <summary>先頭だけを見せるための試し読み(全行は保持しない)。</summary>
    public IReadOnlyList<CsvSampleRow> SampleRows { get; init; } = [];

    public IReadOnlyList<MergeIssue> Issues { get; init; } = [];

    /// <summary>データ元で読んだ行(項目名の行を除く。空行は数えない)。</summary>
    public int SourceRowCount { get; init; }

    /// <summary>作る CSV の行(項目名の行を除く)。</summary>
    public int OutputRowCount { get; init; }

    /// <summary>すべての項目が空欄で読み飛ばした行。</summary>
    public int BlankRowCount { get; init; }

    public string SourceFileName { get; init; } = string.Empty;

    public string? SourceEncodingName { get; init; }

    public string OutputFileName { get; init; } = string.Empty;

    internal string OutputPath { get; init; } = string.Empty;

    internal string AuditPath { get; init; } = string.Empty;

    internal SourceSnapshot? Snapshot { get; init; }

    internal CsvTransformRequest? Request { get; init; }

    public IEnumerable<MergeIssue> Blocks => Issues.Where(issue => issue.Severity == MergeIssueSeverity.Block);

    public IEnumerable<MergeIssue> Warnings
        => Issues.Where(issue => issue.Severity == MergeIssueSeverity.Warning);

    public int BlockCount => Blocks.Count();

    public int WarningCount => Warnings.Count();

    public bool HasBlocks => BlockCount > 0;

    public bool CanExecute => !HasBlocks && Columns.Count > 0 && OutputRowCount > 0;
}

/// <summary>プレビューに出す 1 行(出力する列だけ)。</summary>
public sealed record CsvSampleRow(int RowNumber, IReadOnlyList<string> Values)
{
    /// <summary>画面で 1 列に収めるための表示(項目は「 | 」で区切る)。</summary>
    public string Display => string.Join(" | ", Values);
}

/// <summary>CSV 変換の実行結果。</summary>
public sealed record CsvTransformResult
{
    public required bool Success { get; init; }

    public string Message { get; init; } = string.Empty;

    /// <summary>作ったファイル名(パスは持たない)。</summary>
    public IReadOnlyList<string> OutputFileNames { get; init; } = [];

    public int RowCount { get; init; }

    public static CsvTransformResult Failed(string message) => new() { Success = false, Message = message };
}
