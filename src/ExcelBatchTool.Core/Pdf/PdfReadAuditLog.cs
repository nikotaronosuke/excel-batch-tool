using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using ExcelBatchTool.Core.CsvTransform;
using ExcelBatchTool.Core.Mutation;
using ExcelBatchTool.Core.Ocr;

namespace ExcelBatchTool.Core.Pdf;

/// <summary>
/// 作ったファイルの隣に置く「どの PDF から何を取り出したか」の控え。
/// 完全にローカルのファイルで、外部へは送らない。パスは書かずファイル名だけを残す。
/// PDF の本文そのものは控えへ複製しない(件数と設定だけを残す)。
/// </summary>
internal static class PdfReadAuditLog
{
    /// <summary>文字情報だけから作った場合(Phase 2F-A と同じ内容)。</summary>
    public const int SchemaVersion = 1;

    /// <summary>
    /// OCR を通した場合。ocr の節が増えるぶん版を上げる
    /// (「版 1 が意味すること」を後から変えない)。
    /// </summary>
    public const int OcrSchemaVersion = 2;

    public const string OperationName = "pdf-extract";

    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
    };

    public static void Write(
        Stream stream,
        PdfReadPreview preview,
        PdfReadRequest request,
        SourceSnapshot snapshot,
        int rowCount,
        int columnCount)
    {
        var reading = preview.OcrReading;

        var document = new AuditDocument
        {
            SchemaVersion = reading is null ? SchemaVersion : OcrSchemaVersion,
            CreatedAtUtc = DateTimeOffset.UtcNow.ToString("o"),
            Operation = OperationName,
            Source = new AuditSource
            {
                FileName = preview.SourceFileName,
                Sha256 = snapshot.Sha256,
                Length = snapshot.Length,
                PageCount = preview.PageCount,
                DetectedKind = KindName(preview.Kind),
                ExtractionMethod = ExtractionMethod(preview),
                Normalization = "NFKC",
            },
            Output = new AuditOutput
            {
                FileName = preview.OutputFileName,
                Type = request.OutputFormat == PdfOutputFormat.Xlsx ? "xlsx" : "csv",
                Encoding = request.OutputFormat == PdfOutputFormat.Csv
                    ? CsvEncodingName(request.CsvEncoding)
                    : null,
                QuoteMode = request.OutputFormat == PdfOutputFormat.Csv
                    ? request.CsvQuoteMode == CsvQuoteMode.All ? "all" : "minimal"
                    : null,
                RowCount = rowCount,
                ColumnCount = columnCount,
            },
            // 読み取った文字そのものは残さない。件数と、どのモデルで読んだかだけ。
            Ocr = reading is null ? null : new AuditOcr
            {
                MultiModel = reading.EngineInfo.MultiModel,
                JapanModel = reading.EngineInfo.JapanModel,
                Runtime = reading.EngineInfo.Runtime,
                Backend = reading.EngineInfo.Backend,
                Dpi = reading.EngineInfo.Dpi,
                AutoAcceptThreshold = OcrFusion.AutoAcceptThreshold,
                OcrPageCount = reading.OcrPages.Count,
                Mode = ModeName(reading.ResolvedMode),
                DeskewedPageCount = reading.NeedsDeskewPages.Count,
                ScanTablePageCount = reading.TableLikePages.Count,
                FormPageCount = reading.FormPages.Count,
                ExpectedFieldCount = reading.FieldNames.Count == 0
                    ? 0
                    : reading.FieldNames.Count * reading.FormPages.Count,
                ItemCount = reading.Items.Count,
                AutoAcceptedCount = reading.InitiallyAutoAcceptedCount,
                NeedsReviewCount = reading.InitiallyNeedsReviewCount,
                UnreadableCount = reading.InitiallyUnreadableCount,
                MissingCount = reading.InitiallyMissingCount,
                UserConfirmedCount = reading.UserConfirmedCount,
                UserEditedCount = reading.UserEditedCount,
            },
            Warnings = [.. preview.Warnings.Select(issue => issue.Message)],
        };

        JsonSerializer.Serialize(stream, document, Options);
    }

    private static string ModeName(Ocr.OcrReadMode mode) => mode switch
    {
        Ocr.OcrReadMode.Table => "table",
        Ocr.OcrReadMode.FixedForm => "fixed-form",
        _ => "lines",
    };

    internal static string KindName(PdfDocumentKind kind) => kind switch
    {
        PdfDocumentKind.Text => "text",
        PdfDocumentKind.Table => "table",
        PdfDocumentKind.Scan => "scan",
        PdfDocumentKind.Mixed => "mixed",
        _ => "unknown",
    };

    private static string ExtractionMethod(PdfReadPreview preview)
    {
        if (preview.OcrReading is not null)
        {
            return preview.PagePlans.Any(plan => plan.Route == PdfPageRoute.BornDigitalText)
                ? "ocr-dual-read+pdfpig-lines"
                : "ocr-dual-read";
        }

        return preview.Kind switch
        {
            PdfDocumentKind.Text => "pdfpig-lines",
            PdfDocumentKind.Table => preview.TableFromRulings
                ? "ruling-grid+pdfpig-letters"
                : "header-guided+pdfpig-letters",
            _ => "none",
        };
    }

    private static string CsvEncodingName(CsvOutputEncoding encoding) => encoding switch
    {
        CsvOutputEncoding.Utf8Bom => "utf-8-bom",
        CsvOutputEncoding.Utf8 => "utf-8",
        _ => "shift_jis",
    };

    internal sealed class AuditDocument
    {
        public int SchemaVersion { get; init; }

        public string CreatedAtUtc { get; init; } = string.Empty;

        public string Operation { get; init; } = string.Empty;

        public AuditSource Source { get; init; } = new();

        public AuditOutput Output { get; init; } = new();

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public AuditOcr? Ocr { get; init; }

        public IReadOnlyList<string> Warnings { get; init; } = [];
    }

    /// <summary>OCR を通した場合だけ書く。読み取った文字そのものは入れない。</summary>
    internal sealed class AuditOcr
    {
        public string MultiModel { get; init; } = string.Empty;

        public string JapanModel { get; init; } = string.Empty;

        public string Runtime { get; init; } = string.Empty;

        public string Backend { get; init; } = string.Empty;

        public int Dpi { get; init; }

        public double AutoAcceptThreshold { get; init; }

        public int OcrPageCount { get; init; }

        /// <summary>どう組み立てたか(文章 / 表 / 帳票)。</summary>
        public string Mode { get; init; } = string.Empty;

        /// <summary>傾きを直したページ数。</summary>
        public int DeskewedPageCount { get; init; }

        /// <summary>表として読んだページ数。</summary>
        public int ScanTablePageCount { get; init; }

        /// <summary>帳票として読んだページ数。</summary>
        public int FormPageCount { get; init; }

        /// <summary>帳票として読むときに指定した項目の総数(ページ数 × 項目数)。</summary>
        public int ExpectedFieldCount { get; init; }

        public int ItemCount { get; init; }

        public int AutoAcceptedCount { get; init; }

        public int NeedsReviewCount { get; init; }

        public int UnreadableCount { get; init; }

        /// <summary>読む場所は分かっていたのに、何も読み取れなかった件数。</summary>
        public int MissingCount { get; init; }

        public int UserConfirmedCount { get; init; }

        public int UserEditedCount { get; init; }
    }

    internal sealed class AuditSource
    {
        public string FileName { get; init; } = string.Empty;

        public string Sha256 { get; init; } = string.Empty;

        public long Length { get; init; }

        public int PageCount { get; init; }

        public string DetectedKind { get; init; } = string.Empty;

        public string ExtractionMethod { get; init; } = string.Empty;

        public string Normalization { get; init; } = string.Empty;
    }

    internal sealed class AuditOutput
    {
        public string FileName { get; init; } = string.Empty;

        public string Type { get; init; } = string.Empty;

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Encoding { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? QuoteMode { get; init; }

        public int RowCount { get; init; }

        public int ColumnCount { get; init; }
    }
}
