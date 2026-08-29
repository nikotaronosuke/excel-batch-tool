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

/// <summary>いま何ができる状態か。</summary>
public enum PdfReadStage
{
    /// <summary>そのまま出力できる(文字情報だけ、または OCR の確認が終わった)。</summary>
    Ready = 0,

    /// <summary>スキャンされたページがあるので、先に OCR で読み取る必要がある。</summary>
    NeedsOcr,

    /// <summary>この段階では扱えない。</summary>
    Blocked,
}

/// <summary>ページ 1 枚をどう処理するか。document 全体で 1 種類と決めつけない。</summary>
public enum PdfPageRoute
{
    /// <summary>文字情報のある文章ページ。PdfPig で読む。</summary>
    BornDigitalText = 0,

    /// <summary>文字情報のある表ページ。PdfPig + 罫線で読む。</summary>
    BornDigitalTable,

    /// <summary>画像だけのページ。OCR が必要。</summary>
    Scan,

    /// <summary>判定できないページ。</summary>
    Unknown,
}

public sealed record PdfPagePlan(int Page, PdfPageRoute Route);

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

    /// <summary>
    /// OCR に回せるページ数の上限。
    ///
    /// 文字情報のある PDF と違い、OCR は 1 ページあたり実測 1.6〜3.8 秒かかる。
    /// 2,000 ページを OCR に回すと 1〜2 時間になり、途中経過も分からないまま
    /// 待たせることになる。無制限に走らせず、ここで区切って
    /// 「何ページまでなら扱えるか」を先に伝える。
    /// </summary>
    public const int MaxOcrPages = 1000;

    /// <summary>
    /// これを超えるページ数の OCR は時間がかかることを先に知らせる(止めはしない)。
    /// </summary>
    public const int SlowOcrPageWarning = 200;

    /// <summary>
    /// 1 ページを 300dpi で描いたときの画素数の上限。
    ///
    /// A4 は約 870 万画素。A0 相当までは扱えるようにしつつ、
    /// それ以上の巨大なページでメモリーを使い切らないように区切る。
    /// </summary>
    public const long MaxRenderedPixelsPerPage = 80_000_000;

    /// <summary>OCR の 1 ページあたりの実測時間(見込みを伝えるためだけに使う)。</summary>
    public const double OcrSecondsPerPage = 2.0;

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

    /// <summary>いま何ができる状態か。</summary>
    public PdfReadStage Stage { get; init; } = PdfReadStage.Blocked;

    /// <summary>ページごとの処理の振り分け。document 全体で 1 種類と決めつけない。</summary>
    public IReadOnlyList<PdfPagePlan> PagePlans { get; init; } = [];

    /// <summary>OCR が必要なページ番号。</summary>
    public IReadOnlyList<int> OcrPageNumbers
        => PagePlans.Where(plan => plan.Route == PdfPageRoute.Scan).Select(plan => plan.Page).ToList();

    /// <summary>OCR の確認結果(OCR を通した場合だけ入る)。</summary>
    public Ocr.OcrDocumentReading? OcrReading { get; init; }

    /// <summary>出力する行数(文字 PDF は行、表 PDF はヘッダー込みの行)。</summary>
    public int OutputRowCount => Kind == PdfDocumentKind.Table ? TableRows.Count : Lines.Count;

    public bool CanExecute => Stage == PdfReadStage.Ready
        && !HasBlocks
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
