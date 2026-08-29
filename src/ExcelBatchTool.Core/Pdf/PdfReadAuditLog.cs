using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using ExcelBatchTool.Core.CsvTransform;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Pdf;

/// <summary>
/// 作ったファイルの隣に置く「どの PDF から何を取り出したか」の控え。
/// 完全にローカルのファイルで、外部へは送らない。パスは書かずファイル名だけを残す。
/// PDF の本文そのものは控えへ複製しない(件数と設定だけを残す)。
/// </summary>
internal static class PdfReadAuditLog
{
    public const int SchemaVersion = 1;

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
        var document = new AuditDocument
        {
            SchemaVersion = SchemaVersion,
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
            Warnings = [.. preview.Warnings.Select(issue => issue.Message)],
        };

        JsonSerializer.Serialize(stream, document, Options);
    }

    internal static string KindName(PdfDocumentKind kind) => kind switch
    {
        PdfDocumentKind.Text => "text",
        PdfDocumentKind.Table => "table",
        PdfDocumentKind.Scan => "scan",
        PdfDocumentKind.Mixed => "mixed",
        _ => "unknown",
    };

    private static string ExtractionMethod(PdfReadPreview preview) => preview.Kind switch
    {
        PdfDocumentKind.Text => "pdfpig-lines",
        PdfDocumentKind.Table => preview.TableFromRulings
            ? "ruling-grid+pdfpig-letters"
            : "header-guided+pdfpig-letters",
        _ => "none",
    };

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

        public IReadOnlyList<string> Warnings { get; init; } = [];
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
