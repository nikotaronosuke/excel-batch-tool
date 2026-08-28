using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;

namespace ExcelBatchTool.Core.Mutation;

/// <summary>
/// 出力 Workbook の隣に置く「何をどう変えたか」の控え。
/// 完全にローカルのファイルで、外部へは送らない。
/// パスは書かずファイル名だけを残す(置き場所を移しても意味が変わらないようにするため)。
/// </summary>
internal static class MutationAuditLog
{
    /// <summary>手入力した値を書いたときの控えの版。</summary>
    public const int SchemaVersion = 1;

    /// <summary>データ元から転記したときの控えの版(データ元と行の情報が増える)。</summary>
    public const int MappingSchemaVersion = 2;

    /// <summary>表同士の突合更新の控えの版(転記先の表と行の情報が増える)。</summary>
    public const int TableUpdateSchemaVersion = 3;

    public const string OperationName = "set-cell-value";

    public const string MappingOperationName = "map-source-to-cells";

    public const string TableUpdateOperationName = "map-source-table-to-target-table";

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
        var isTableUpdate = file.TargetTable is not null;
        var isMapping = file.DataSource is not null;

        var document = new AuditDocument
        {
            SchemaVersion = isTableUpdate ? TableUpdateSchemaVersion
                : isMapping ? MappingSchemaVersion
                : SchemaVersion,
            CreatedAt = DateTimeOffset.Now.ToString("o"),
            SourceFileName = file.FileName,
            OutputFileName = file.OutputFileName,
            SourceSha256 = file.Snapshot.Sha256,
            OutputSha256 = outputSha256,
            Operation = isTableUpdate ? TableUpdateOperationName
                : isMapping ? MappingOperationName
                : OperationName,
            TargetTable = file.TargetTable is { } table
                ? new AuditTargetTable { HeaderRow = table.HeaderRow, KeyColumn = table.KeyColumn }
                : null,
            DataSource = file.DataSource is { } source
                ? new AuditDataSource
                {
                    FileName = source.FileName,
                    Sha256 = source.Sha256,
                    Type = source.Type,
                    SheetName = source.SheetName,
                    HeaderRow = source.HeaderRow,
                    KeyColumn = source.KeyColumn,
                }
                : null,
            Changes = [.. applied.Select(item => new AuditChange
            {
                SheetName = item.Change.SheetName,
                Cell = item.Change.CellReference,
                OldValue = item.Change.CurrentValueDisplay,
                OldType = item.Change.CurrentTypeName,
                NewValue = item.Change.NewValueDisplay,
                NewType = item.Change.NewTypeName,
                SourceColumn = item.Change.Provenance?.SourceColumn,
                TargetColumn = item.Change.Provenance?.TargetColumn,
                Key = item.Change.Provenance?.Key,
                SourceRowNumber = item.Change.Provenance?.SourceRowNumber,
                TargetRowNumber = item.Change.Provenance?.TargetRowNumber,
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

        /// <summary>転記の場合のみ。手入力のときは書かない。</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public AuditDataSource? DataSource { get; init; }

        /// <summary>表同士の突合更新の場合のみ。</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public AuditTargetTable? TargetTable { get; init; }

        public IReadOnlyList<AuditChange> Changes { get; init; } = Array.Empty<AuditChange>();
    }

    /// <summary>転記に使ったデータ元。パスは書かずファイル名だけ。</summary>
    internal sealed class AuditDataSource
    {
        public string FileName { get; init; } = string.Empty;

        public string Sha256 { get; init; } = string.Empty;

        public string Type { get; init; } = string.Empty;

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? SheetName { get; init; }

        public int HeaderRow { get; init; }

        public string KeyColumn { get; init; } = string.Empty;
    }

    /// <summary>転記先の表の読み方。</summary>
    internal sealed class AuditTargetTable
    {
        public int HeaderRow { get; init; }

        public string KeyColumn { get; init; } = string.Empty;
    }

    internal sealed class AuditChange
    {
        public string SheetName { get; init; } = string.Empty;

        public string Cell { get; init; } = string.Empty;

        public string OldValue { get; init; } = string.Empty;

        public string OldType { get; init; } = string.Empty;

        public string NewValue { get; init; } = string.Empty;

        public string NewType { get; init; } = string.Empty;

        /// <summary>転記の場合のみ: データ元のどの項目から取ったか。</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? SourceColumn { get; init; }

        /// <summary>表同士の突合更新の場合のみ: 転記先のどの項目を更新したか。</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? TargetColumn { get; init; }

        /// <summary>転記の場合のみ: 照合に使ったキー。</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Key { get; init; }

        /// <summary>転記の場合のみ: データ元の行番号(CSV はレコード番号)。</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? SourceRowNumber { get; init; }

        /// <summary>表同士の突合更新の場合のみ: 転記先の行番号。</summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? TargetRowNumber { get; init; }
    }
}
