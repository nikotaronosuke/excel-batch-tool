using ExcelBatchTool.Core.CsvTransform;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Pdf;

/// <summary>PDF の自動判定の結果。</summary>
public enum PdfDocumentKind
{
    /// <summary>判定できない(開けない・空など)。</summary>
    Unknown = 0,

    /// <summary>文字情報を持つ、通常の文字 PDF。</summary>
    Text,

    /// <summary>文字情報を持つ、表の PDF。</summary>
    Table,

    /// <summary>画像だけ(スキャン)の PDF。OCR が必要なため現在の段階では対象外。</summary>
    Scan,

    /// <summary>文字のページと画像だけのページが混在。全ページを安全に抽出できないため対象外。</summary>
    Mixed,
}

/// <summary>出力の形式。</summary>
public enum PdfOutputFormat
{
    Xlsx = 0,
    Csv,
}

/// <summary>PDF 読み取り 1 回分の指定。ViewModel に直書きせず、将来レシピ化できる形に分けておく。</summary>
public sealed record PdfReadRequest
{
    public required string SourceFilePath { get; init; }

    public PdfOutputFormat OutputFormat { get; init; } = PdfOutputFormat.Xlsx;

    /// <summary>CSV のときの文字コード。</summary>
    public CsvOutputEncoding CsvEncoding { get; init; } = CsvOutputEncoding.Utf8Bom;

    /// <summary>CSV のときの引用符の付け方。</summary>
    public CsvQuoteMode CsvQuoteMode { get; init; } = CsvQuoteMode.Minimal;

    public string OutputSuffix { get; init; } = PdfReadDefaults.OutputSuffix;
}

public static class PdfReadDefaults
{
    public const string OutputSuffix = "_PDF抽出";

    /// <summary>このページ数を超える PDF は動作未確認として扱う。</summary>
    public const int MaxPages = 2000;

    /// <summary>1 ページの文字数がこれ未満なら「文字情報のあるページ」とみなさない。</summary>
    public const int MinLettersPerTextPage = 10;

    /// <summary>プレビューに出す行数の上限(全行は出力時に使う)。</summary>
    public const int PreviewRowLimit = 500;
}

/// <summary>通常の文字 PDF の 1 行(ページ番号・ページ内の行番号を失わない)。</summary>
public sealed record PdfTextLine(int Page, int Line, string Text);

/// <summary>PDF 読み取りのプレビュー。</summary>
public sealed class PdfReadPreview
{
    public PdfDocumentKind Kind { get; init; }

    /// <summary>画面に出す判定名。</summary>
    public string KindDisplay => Kind switch
    {
        PdfDocumentKind.Text => "通常の文字 PDF",
        PdfDocumentKind.Table => "表 PDF",
        PdfDocumentKind.Scan => "スキャン / 画像 PDF(OCR が必要)",
        PdfDocumentKind.Mixed => "文字と画像が混在(OCR が必要)",
        _ => "判定できない PDF",
    };

    public int PageCount { get; init; }

    /// <summary>通常の文字 PDF のとき: 全行。</summary>
    public IReadOnlyList<PdfTextLine> Lines { get; init; } = [];

    /// <summary>表 PDF のとき: ヘッダー + 全データ行。</summary>
    public IReadOnlyList<string[]> TableRows { get; init; } = [];

    /// <summary>表 PDF のとき、罫線から取ったか(false = 罫線なしを位置から再構成)。</summary>
    public bool TableFromRulings { get; init; }

    public IReadOnlyList<MergeIssue> Issues { get; init; } = [];

    public string SourceFileName { get; init; } = string.Empty;

    public string OutputFileName { get; init; } = string.Empty;

    internal string OutputPath { get; init; } = string.Empty;

    internal string AuditPath { get; init; } = string.Empty;

    internal SourceSnapshot? Snapshot { get; init; }

    internal PdfReadRequest? Request { get; init; }

    public IEnumerable<MergeIssue> Blocks => Issues.Where(issue => issue.Severity == MergeIssueSeverity.Block);

    public IEnumerable<MergeIssue> Warnings
        => Issues.Where(issue => issue.Severity == MergeIssueSeverity.Warning);

    public int BlockCount => Blocks.Count();

    public int WarningCount => Warnings.Count();

    public bool HasBlocks => BlockCount > 0;

    /// <summary>出力する行数(文字 PDF は行、表 PDF はヘッダー込みの行)。</summary>
    public int OutputRowCount => Kind switch
    {
        PdfDocumentKind.Text => Lines.Count,
        PdfDocumentKind.Table => TableRows.Count,
        _ => 0,
    };

    public bool CanExecute => !HasBlocks
        && Kind is PdfDocumentKind.Text or PdfDocumentKind.Table
        && OutputRowCount > 0;
}

/// <summary>PDF 読み取りの実行結果。</summary>
public sealed record PdfReadResult
{
    public required bool Success { get; init; }

    public string Message { get; init; } = string.Empty;

    public IReadOnlyList<string> OutputFileNames { get; init; } = [];

    public static PdfReadResult Failed(string message) => new() { Success = false, Message = message };
}
