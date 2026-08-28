using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelBatchTool.Core.Mapping;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Recipes;

/// <summary>
/// 保存できる処理設定の種類。
/// JSON に書く値は C# 名と切り離して固定し、後から名前を変えても読めなくならないようにする。
/// </summary>
public enum RecipeType
{
    /// <summary>「4. セルをまとめて変更」の設定。</summary>
    CellInputSet,

    /// <summary>「5. 表から転記」の設定。</summary>
    SourceToFixedCells,

    /// <summary>「6. 表を突合して更新」の設定。</summary>
    SourceTableToTargetTable,
}

/// <summary>JSON 上の固定文字列と C# 側の型との対応。</summary>
public static class RecipeJsonNames
{
    public const string CellInputSet = "cell-input-set";
    public const string SourceToFixedCells = "source-to-fixed-cells";
    public const string SourceTableToTargetTable = "source-table-to-target-table";

    public const string Text = "text";
    public const string Number = "number";
    public const string Blank = "blank";

    public const string Xlsx = "xlsx";
    public const string Csv = "csv";

    public static string Of(RecipeType type) => type switch
    {
        RecipeType.CellInputSet => CellInputSet,
        RecipeType.SourceToFixedCells => SourceToFixedCells,
        _ => SourceTableToTargetTable,
    };

    public static string Of(CellWriteKind kind) => kind switch
    {
        CellWriteKind.Number => Number,
        CellWriteKind.Blank => Blank,
        _ => Text,
    };

    public static string Of(SourceFileKind kind)
        => kind == SourceFileKind.Csv ? Csv : Xlsx;
}

/// <summary>固定文字列として保存した値を読むための共通処理。</summary>
internal static class RecipeJsonReader
{
    /// <summary>文字列以外が入っていたら読み取りエラーにする(黙って既定値にしない)。</summary>
    public static string? ReadText(ref Utf8JsonReader reader)
        => reader.TokenType == JsonTokenType.String
            ? reader.GetString()
            : throw new JsonException("文字列で保存されているはずの値が読み取れません。");
}

/// <summary>レシピの種類を固定文字列として読み書きする。未知の値は読み取りエラーにする。</summary>
internal sealed class RecipeTypeConverter : JsonConverter<RecipeType>
{
    public override RecipeType Read(ref Utf8JsonReader reader, Type _, JsonSerializerOptions options)
        => RecipeJsonReader.ReadText(ref reader) switch
        {
            RecipeJsonNames.CellInputSet => RecipeType.CellInputSet,
            RecipeJsonNames.SourceToFixedCells => RecipeType.SourceToFixedCells,
            RecipeJsonNames.SourceTableToTargetTable => RecipeType.SourceTableToTargetTable,
            var other => throw new JsonException($"未知の処理の種類です: {other}"),
        };

    public override void Write(Utf8JsonWriter writer, RecipeType value, JsonSerializerOptions options)
        => writer.WriteStringValue(RecipeJsonNames.Of(value));
}

/// <summary>値の種類を固定文字列として読み書きする。未知の値は読み取りエラーにする。</summary>
internal sealed class CellWriteKindConverter : JsonConverter<CellWriteKind>
{
    public override CellWriteKind Read(ref Utf8JsonReader reader, Type _, JsonSerializerOptions options)
        => RecipeJsonReader.ReadText(ref reader) switch
        {
            RecipeJsonNames.Text => CellWriteKind.Text,
            RecipeJsonNames.Number => CellWriteKind.Number,
            RecipeJsonNames.Blank => CellWriteKind.Blank,
            var other => throw new JsonException($"未知の値の種類です: {other}"),
        };

    public override void Write(Utf8JsonWriter writer, CellWriteKind value, JsonSerializerOptions options)
        => writer.WriteStringValue(RecipeJsonNames.Of(value));
}

/// <summary>データ元の種類を固定文字列として読み書きする。未知の値は読み取りエラーにする。</summary>
internal sealed class SourceFileKindConverter : JsonConverter<SourceFileKind>
{
    public override SourceFileKind Read(ref Utf8JsonReader reader, Type _, JsonSerializerOptions options)
        => RecipeJsonReader.ReadText(ref reader) switch
        {
            RecipeJsonNames.Xlsx => SourceFileKind.Xlsx,
            RecipeJsonNames.Csv => SourceFileKind.Csv,
            var other => throw new JsonException($"未知のデータ元の種類です: {other}"),
        };

    public override void Write(Utf8JsonWriter writer, SourceFileKind value, JsonSerializerOptions options)
        => writer.WriteStringValue(RecipeJsonNames.Of(value));
}

/// <summary>「4. セルをまとめて変更」で保存する 1 行。</summary>
public sealed class RecipeOperation
{
    /// <summary>変更する位置(A1 形式の単一セル)。</summary>
    public string Cell { get; init; } = string.Empty;

    public CellWriteKind Kind { get; init; }

    /// <summary>入力する値。「空欄」の行では書かない。</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; init; }
}

/// <summary>「5. 表から転記」で保存する対応付けの 1 行。</summary>
public sealed class RecipeCellMapping
{
    public string SourceColumn { get; init; } = string.Empty;

    /// <summary>転記先の固定セル(A1 形式)。</summary>
    public string TargetCell { get; init; } = string.Empty;

    public CellWriteKind Kind { get; init; }
}

/// <summary>「6. 表を突合して更新」で保存する対応付けの 1 行。</summary>
public sealed class RecipeColumnMapping
{
    public string SourceColumn { get; init; } = string.Empty;

    public string TargetColumn { get; init; } = string.Empty;

    public CellWriteKind Kind { get; init; }
}

/// <summary>「4. セルをまとめて変更」の設定。対象のファイル・シートは含まない。</summary>
public sealed class CellInputSetRecipe
{
    public IReadOnlyList<RecipeOperation> Operations { get; init; } = [];

    public string OutputSuffix { get; init; } = string.Empty;
}

/// <summary>「5. 表から転記」の設定。データ元・転記先のファイルは含まない。</summary>
public sealed class SourceToFixedCellsRecipe
{
    public SourceFileKind SourceFileKind { get; init; }

    /// <summary>.xlsx のデータ元で使うシート名。CSV では書かない。</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SourceSheetName { get; init; }

    public int HeaderRow { get; init; } = 1;

    public string SourceKeyColumn { get; init; } = string.Empty;

    /// <summary>転記先シートで、キーが入っているセル。</summary>
    public string TargetKeyCell { get; init; } = string.Empty;

    public IReadOnlyList<RecipeCellMapping> Mappings { get; init; } = [];

    public string OutputSuffix { get; init; } = string.Empty;
}

/// <summary>「6. 表を突合して更新」の設定。データ元・転記先のファイルは含まない。</summary>
public sealed class SourceTableToTargetTableRecipe
{
    public SourceFileKind SourceFileKind { get; init; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SourceSheetName { get; init; }

    public int SourceHeaderRow { get; init; } = 1;

    public string SourceKeyColumn { get; init; } = string.Empty;

    public int TargetHeaderRow { get; init; } = 1;

    public string TargetKeyColumn { get; init; } = string.Empty;

    public IReadOnlyList<RecipeColumnMapping> Mappings { get; init; } = [];

    public string OutputSuffix { get; init; } = string.Empty;
}

/// <summary>
/// 名前を付けて保存した 1 つの処理設定。
/// 「今回使うファイル」は含めない(毎月ファイル名が変わっても同じ設定を使えるようにするため)。
/// </summary>
public sealed record SavedRecipe
{
    /// <summary>内部の識別子。画面には出さず、更新・削除の対象を特定するためだけに使う。</summary>
    public string Id { get; init; } = string.Empty;

    public string Name { get; init; } = string.Empty;

    public RecipeType Type { get; init; }

    /// <summary>保存した日時(ISO 8601)。表示用で、処理の判断には使わない。</summary>
    public string CreatedAt { get; init; } = string.Empty;

    /// <summary>最後に更新した日時(ISO 8601)。表示用で、処理の判断には使わない。</summary>
    public string UpdatedAt { get; init; } = string.Empty;

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public CellInputSetRecipe? CellInputSet { get; init; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public SourceToFixedCellsRecipe? SourceToFixedCells { get; init; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public SourceTableToTargetTableRecipe? SourceTableToTargetTable { get; init; }
}

/// <summary>レシピファイル全体。</summary>
public sealed class RecipeDocument
{
    public int SchemaVersion { get; init; }

    public IReadOnlyList<SavedRecipe> Recipes { get; init; } = [];
}
