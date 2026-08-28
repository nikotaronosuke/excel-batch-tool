using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

namespace ExcelBatchTool.Core.Mutation;

/// <summary>
/// 出力 Workbook の隣に置く「何をどう変えたか」の控え。
/// 完全にローカルのファイルで、外部へは送らない。
/// パスは書かずファイル名だけを残す(置き場所を移しても意味が変わらないようにするため)。
/// </summary>
internal static class MutationAuditLog
{
    /// <summary>控えファイルの書式が変わったら上げる。</summary>
    public const int SchemaVersion = 1;

    public const string OperationName = "set-cell-value";

    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        // 日本語をそのまま読めるようにする(HTML 記号だけは従来どおり退避)。
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
    };

    public static void Write(
        string path,
        CellMutationFilePlan file,
        IReadOnlyList<AppliedChange> applied,
        string outputSha256)
    {
        var document = new AuditDocument
        {
            SchemaVersion = SchemaVersion,
            CreatedAt = DateTimeOffset.Now.ToString("o"),
            SourceFileName = file.FileName,
            OutputFileName = file.OutputFileName,
            SourceSha256 = file.Snapshot.Sha256,
            OutputSha256 = outputSha256,
            Operation = OperationName,
            Changes = [.. applied.Select(item => new AuditChange
            {
                SheetName = item.Change.SheetName,
                Cell = item.Change.CellReference,
                OldValue = item.Change.CurrentValueDisplay,
                OldType = item.Change.CurrentTypeName,
                NewValue = item.Change.NewValueDisplay,
                NewType = item.Change.NewTypeName,
            })],
        };

        File.WriteAllText(path, JsonSerializer.Serialize(document, Options));
    }

    internal sealed class AuditDocument
    {
        public int SchemaVersion { get; init; }

        public string CreatedAt { get; init; } = string.Empty;

        public string SourceFileName { get; init; } = string.Empty;

        public string OutputFileName { get; init; } = string.Empty;

        public string SourceSha256 { get; init; } = string.Empty;

        public string OutputSha256 { get; init; } = string.Empty;

        public string Operation { get; init; } = string.Empty;

        public IReadOnlyList<AuditChange> Changes { get; init; } = Array.Empty<AuditChange>();
    }

    internal sealed class AuditChange
    {
        public string SheetName { get; init; } = string.Empty;

        public string Cell { get; init; } = string.Empty;

        public string OldValue { get; init; } = string.Empty;

        public string OldType { get; init; } = string.Empty;

        public string NewValue { get; init; } = string.Empty;

        public string NewType { get; init; } = string.Empty;
    }
}
