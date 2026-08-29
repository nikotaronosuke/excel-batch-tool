using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelBatchTool.Core.CsvTransform;
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

    /// <summary>「7. CSV を変換」の設定。</summary>
    CsvTransform,

    /// <summary>「8. PDF を読み取る」の設定。</summary>
    PdfRead,
}

/// <summary>JSON 上の固定文字列と C# 側の型との対応。</summary>
public static class RecipeJsonNames
{
    public const string CellInputSet = "cell-input-set";
    public const string SourceToFixedCells = "source-to-fixed-cells";
    public const string SourceTableToTargetTable = "source-table-to-target-table";
    public const string CsvTransform = "csv-transform";
    public const string PdfRead = "pdf-read";

    public const string Text = "text";
    public const string Number = "number";
    public const string Blank = "blank";

    public const string Xlsx = "xlsx";
    public const string Csv = "csv";

    public const string SourceColumnValue = "source-column";
    public const string FixedTextValue = "fixed-text";
    public const string BlankValue = "blank";

    public const string Utf8BomEncoding = "utf-8-bom";
    public const string Utf8Encoding = "utf-8";
    public const string ShiftJisEncoding = "shift_jis";

    public const string MinimalQuotes = "minimal";
    public const string AllQuotes = "all";

    public static string Of(RecipeType type) => type switch
    {
        RecipeType.CellInputSet => CellInputSet,
        RecipeType.SourceToFixedCells => SourceToFixedCells,
        RecipeType.SourceTableToTargetTable => SourceTableToTargetTable,
        RecipeType.PdfRead => PdfRead,
        _ => CsvTransform,
    };

    public static string Of(CsvValueSourceKind kind) => kind switch
    {
        CsvValueSourceKind.FixedText => FixedTextValue,
        CsvValueSourceKind.Blank => BlankValue,
        _ => SourceColumnValue,
    };

    public static string Of(CsvOutputEncoding encoding) => encoding switch
    {
        CsvOutputEncoding.Utf8 => Utf8Encoding,
        CsvOutputEncoding.ShiftJis => ShiftJisEncoding,
        _ => Utf8BomEncoding,
    };

    public static string Of(CsvQuoteMode mode)
        => mode == CsvQuoteMode.All ? AllQuotes : MinimalQuotes;

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
            RecipeJsonNames.CsvTransform => RecipeType.CsvTransform,
            RecipeJsonNames.PdfRead => RecipeType.PdfRead,
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

/// <summary>出力する列の入れ方を固定文字列として読み書きする。</summary>
internal sealed class CsvValueSourceKindConverter : JsonConverter<CsvValueSourceKind>
{
    public override CsvValueSourceKind Read(
        ref Utf8JsonReader reader, Type _, JsonSerializerOptions options)
        => RecipeJsonReader.ReadText(ref reader) switch
        {
            RecipeJsonNames.SourceColumnValue => CsvValueSourceKind.SourceColumn,
            RecipeJsonNames.FixedTextValue => CsvValueSourceKind.FixedText,
            RecipeJsonNames.BlankValue => CsvValueSourceKind.Blank,
            var other => throw new JsonException($"未知の項目の入れ方です: {other}"),
        };

    public override void Write(
        Utf8JsonWriter writer, CsvValueSourceKind value, JsonSerializerOptions options)
        => writer.WriteStringValue(RecipeJsonNames.Of(value));
}

/// <summary>出力の文字コードを固定文字列として読み書きする。</summary>
internal sealed class CsvOutputEncodingConverter : JsonConverter<CsvOutputEncoding>
{
    public override CsvOutputEncoding Read(
        ref Utf8JsonReader reader, Type _, JsonSerializerOptions options)
        => RecipeJsonReader.ReadText(ref reader) switch
        {
            RecipeJsonNames.Utf8BomEncoding => CsvOutputEncoding.Utf8Bom,
            RecipeJsonNames.Utf8Encoding => CsvOutputEncoding.Utf8,
            RecipeJsonNames.ShiftJisEncoding => CsvOutputEncoding.ShiftJis,
            var other => throw new JsonException($"未知の文字コードです: {other}"),
        };

    public override void Write(
        Utf8JsonWriter writer, CsvOutputEncoding value, JsonSerializerOptions options)
        => writer.WriteStringValue(RecipeJsonNames.Of(value));
}

/// <summary>引用符の付け方を固定文字列として読み書きする。</summary>
internal sealed class CsvQuoteModeConverter : JsonConverter<CsvQuoteMode>
{
    public override CsvQuoteMode Read(ref Utf8JsonReader reader, Type _, JsonSerializerOptions options)
        => RecipeJsonReader.ReadText(ref reader) switch
        {
            RecipeJsonNames.MinimalQuotes => CsvQuoteMode.Minimal,
            RecipeJsonNames.AllQuotes => CsvQuoteMode.All,
            var other => throw new JsonException($"未知の引用符の付け方です: {other}"),
        };

    public override void Write(Utf8JsonWriter writer, CsvQuoteMode value, JsonSerializerOptions options)
        => writer.WriteStringValue(RecipeJsonNames.Of(value));
}

/// <summary>「7. CSV を変換」で保存する出力列 1 件。</summary>
public sealed record RecipeCsvColumn
{
    public string OutputName { get; init; } = string.Empty;

    public CsvValueSourceKind ValueSourceKind { get; init; }

    /// <summary>データ元から取るときの項目名。</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SourceColumn { get; init; }

    /// <summary>固定値のときに全行へ入れる文字。</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FixedValue { get; init; }
}

/// <summary>「7. CSV を変換」の設定。データ元のファイルは含まない。</summary>
public sealed record CsvTransformRecipe
{
    public SourceFileKind SourceFileKind { get; init; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SourceSheetName { get; init; }

    public int HeaderRow { get; init; } = 1;

    public IReadOnlyList<RecipeCsvColumn> OutputColumns { get; init; } = [];

    public CsvOutputEncoding Encoding { get; init; }

    public CsvQuoteMode QuoteMode { get; init; }

    public string OutputSuffix { get; init; } = string.Empty;
}

/// <summary>「4. セルをまとめて変更」で保存する 1 行。</summary>
public sealed record RecipeOperation
{
    /// <summary>変更する位置(A1 形式の単一セル)。</summary>
    public string Cell { get; init; } = string.Empty;

    public CellWriteKind Kind { get; init; }

    /// <summary>入力する値。「空欄」の行では書かない。</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; init; }
}

/// <summary>「5. 表から転記」で保存する対応付けの 1 行。</summary>
public sealed record RecipeCellMapping
{
    public string SourceColumn { get; init; } = string.Empty;

    /// <summary>転記先の固定セル(A1 形式)。</summary>
    public string TargetCell { get; init; } = string.Empty;

    public CellWriteKind Kind { get; init; }
}

/// <summary>「6. 表を突合して更新」で保存する対応付けの 1 行。</summary>
public sealed record RecipeColumnMapping
{
    public string SourceColumn { get; init; } = string.Empty;

    public string TargetColumn { get; init; } = string.Empty;

    public CellWriteKind Kind { get; init; }
}

/// <summary>「4. セルをまとめて変更」の設定。対象のファイル・シートは含まない。</summary>
public sealed record CellInputSetRecipe
{
    public IReadOnlyList<RecipeOperation> Operations { get; init; } = [];

    public string OutputSuffix { get; init; } = string.Empty;
}

/// <summary>「5. 表から転記」の設定。データ元・転記先のファイルは含まない。</summary>
public sealed record SourceToFixedCellsRecipe
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
public sealed record SourceTableToTargetTableRecipe
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

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public CsvTransformRecipe? CsvTransform { get; init; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public PdfReadRecipe? PdfRead { get; init; }
}

/// <summary>
/// 「8. PDF を読み取る」の設定。
///
/// **PDF そのものに関わるものは何も保存しない。**
/// 元のファイル名・保存場所・読み取った文字・ページの中身・個人情報の実値は入れない。
/// 入れるのは「どう読むか」の設定だけで、同じ様式の帳票を毎月受け取るときに
/// 項目を作り直さずに済むようにするためのもの。
/// </summary>
public sealed record PdfReadRecipe
{
    /// <summary>読み取り方(文章 / 表 / 同じ様式の帳票)。</summary>
    public string ReadMode { get; init; } = string.Empty;

    /// <summary>出力形式(Excel / CSV)。</summary>
    public string OutputFormat { get; init; } = string.Empty;

    public string OutputSuffix { get; init; } = string.Empty;

    /// <summary>CSV のときの文字コードと引用の扱い。</summary>
    public CsvOutputEncoding Encoding { get; init; }

    public CsvQuoteMode QuoteMode { get; init; }

    /// <summary>同じ様式の帳票として読むときの項目。</summary>
    public IReadOnlyList<PdfReadRecipeField> Fields { get; init; } = [];
}

/// <summary>帳票の 1 項目の設定。読み取った値は保存しない。</summary>
public sealed record PdfReadRecipeField
{
    public string Name { get; init; } = string.Empty;

    /// <summary>項目の種類(そのままの文字 / 数量・金額 / コード / 選択)。</summary>
    public string Kind { get; init; } = string.Empty;

    public bool IsRequired { get; init; }

    /// <summary>読み取る場所(300dpi のページ座標)。</summary>
    public double X { get; init; }

    public double Y { get; init; }

    public double Width { get; init; }

    public double Height { get; init; }

    /// <summary>選択項目の選択肢(位置つき)。</summary>
    public IReadOnlyList<PdfReadRecipeChoice> Choices { get; init; } = [];
}

public sealed record PdfReadRecipeChoice
{
    public string Label { get; init; } = string.Empty;

    public double X { get; init; }

    public double Y { get; init; }

    public double Width { get; init; }

    public double Height { get; init; }
}

/// <summary>レシピファイル全体。</summary>
public sealed record RecipeDocument
{
    public int SchemaVersion { get; init; }

    public IReadOnlyList<SavedRecipe> Recipes { get; init; } = [];
}
