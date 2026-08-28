using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.CsvTransform;

/// <summary>
/// 作った CSV の隣に置く「何をどう変換したか」の控え。
/// 完全にローカルのファイルで、外部へは送らない。
/// パスは書かずファイル名だけを残す。
///
/// Excel の書き換えの控え(schemaVersion 1 / 2 / 3)とは操作の意味が違うので、
/// 同じ形式へ押し込まず、csv-transform 用の別の形式にしている。
/// </summary>
internal static class CsvTransformAuditLog
{
    public const int SchemaVersion = 1;

    public const string OperationName = "csv-transform";

    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
    };

    public static void Write(
        string path,
        CsvTransformPreview preview,
        CsvTransformRequest request,
        SourceSnapshot snapshot,
        int rowCount)
    {
        var document = new AuditDocument
        {
            SchemaVersion = SchemaVersion,
            CreatedAt = DateTimeOffset.Now.ToString("o"),
            Operation = OperationName,
            Source = new AuditSource
            {
                FileName = preview.SourceFileName,
                Sha256 = snapshot.Sha256,
                Type = CsvTransformPlanner.KindOf(request.SourceFilePath) == Mapping.SourceFileKind.Csv
                    ? "csv"
                    : "xlsx",
                SheetName = request.SourceSheetName,
                HeaderRow = request.HeaderRow,
                EncodingName = preview.SourceEncodingName,
            },
            Output = new AuditOutput
            {
                FileName = preview.OutputFileName,
                Encoding = EncodingName(request.Encoding),
                QuoteMode = request.QuoteMode == CsvQuoteMode.All ? "all" : "minimal",
                LineEnding = "crlf",
                RowCount = rowCount,
                ColumnCount = preview.Columns.Count,
            },
            Columns = [.. preview.Columns.Select(column => new AuditColumn
            {
                OutputName = column.OutputName,
                SourceKind = SourceKindName(column.ValueSourceKind),
                SourceColumn = column.ValueSourceKind == CsvValueSourceKind.SourceColumn
                    ? column.SourceColumn
                    : null,
                FixedValue = column.ValueSourceKind == CsvValueSourceKind.FixedText
                    ? column.FixedValue
                    : null,
            })],
        };

        File.WriteAllText(path, JsonSerializer.Serialize(document, Options));
    }

    internal static string EncodingName(CsvOutputEncoding encoding) => encoding switch
    {
        CsvOutputEncoding.Utf8Bom => "utf-8-bom",
        CsvOutputEncoding.Utf8 => "utf-8",
        _ => "shift_jis",
    };

    internal static string SourceKindName(CsvValueSourceKind kind) => kind switch
    {
        CsvValueSourceKind.SourceColumn => "source-column",
        CsvValueSourceKind.FixedText => "fixed-text",
        _ => "blank",
    };

    internal sealed class AuditDocument
    {
        public int SchemaVersion { get; init; }

        public string CreatedAt { get; init; } = string.Empty;

        public string Operation { get; init; } = string.Empty;

        public AuditSource Source { get; init; } = new();

        public AuditOutput Output { get; init; } = new();

        public IReadOnlyList<AuditColumn> Columns { get; init; } = [];
    }

    internal sealed class AuditSource
    {
        public string FileName { get; init; } = string.Empty;

        public string Sha256 { get; init; } = string.Empty;

        public string Type { get; init; } = string.Empty;

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? SheetName { get; init; }

        public int HeaderRow { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? EncodingName { get; init; }
    }

    internal sealed class AuditOutput
    {
        public string FileName { get; init; } = string.Empty;

        public string Encoding { get; init; } = string.Empty;

        public string QuoteMode { get; init; } = string.Empty;

        public string LineEnding { get; init; } = string.Empty;

        public int RowCount { get; init; }

        public int ColumnCount { get; init; }
    }

    internal sealed class AuditColumn
    {
        public string OutputName { get; init; } = string.Empty;

        public string SourceKind { get; init; } = string.Empty;

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? SourceColumn { get; init; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? FixedValue { get; init; }
    }
}
