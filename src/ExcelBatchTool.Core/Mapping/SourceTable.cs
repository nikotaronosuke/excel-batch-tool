using System.Globalization;

namespace ExcelBatchTool.Core.Mapping;

/// <summary>データ元のセル 1 つの種類。</summary>
internal enum SourceValueKind
{
    /// <summary>値が無い。</summary>
    Blank = 0,

    /// <summary>素の文字列。</summary>
    Text,

    /// <summary>そのまま読み取ってよい数値。</summary>
    Number,

    /// <summary>読み取れない(数式・日付書式・リッチテキストなど)。</summary>
    Unsupported,
}

/// <summary>データ元のセル 1 つ。読み取れないものは理由を持つ。</summary>
internal readonly record struct SourceValue(
    SourceValueKind Kind, string? Text, double Number, string? Reason)
{
    public static SourceValue Blank() => new(SourceValueKind.Blank, null, 0, null);

    public static SourceValue OfText(string text) => new(SourceValueKind.Text, text, 0, null);

    public static SourceValue OfNumber(double number) => new(SourceValueKind.Number, null, number, null);

    public static SourceValue Unsupported(string reason) => new(SourceValueKind.Unsupported, null, 0, reason);

    public bool IsBlank => Kind == SourceValueKind.Blank;

    /// <summary>プレビュー用の表示文字列。</summary>
    public string Display => Kind switch
    {
        SourceValueKind.Text => Text ?? string.Empty,
        SourceValueKind.Number => Number.ToString(CultureInfo.InvariantCulture),
        SourceValueKind.Unsupported => "(読み取れません)",
        _ => "(空欄)",
    };
}

/// <summary>データ元の 1 行(必要な列だけ)。</summary>
internal sealed record SourceRow(int RowNumber, IReadOnlyList<SourceValue> Values);

/// <summary>項目名(ヘッダー)の読み取り結果。</summary>
internal sealed record SourceHeaderResult
{
    public IReadOnlyList<string> Columns { get; init; } = Array.Empty<string>();

    /// <summary>CSV の場合に判定した文字コード(表示用)。</summary>
    public string? EncodingName { get; init; }

    public string? Error { get; init; }

    public bool IsSuccess => Error is null;

    public static SourceHeaderResult Failed(string error) => new() { Error = error };
}

/// <summary>
/// データ元の値を、転記先へ書く値へ変換する。推測での型変換はしない。
/// 固定セルへの転記(2C1)と表同士の突合更新(2C2)で同じ規則を使う。
/// </summary>
internal static class SourceValueConversion
{
    /// <summary>変換できない場合は「キー「k」の「列名」…」に続く形の理由を返す。</summary>
    public static bool TryConvert(
        SourceValue value,
        Mutation.CellWriteKind writeKind,
        SourceFileKind kind,
        out Mutation.NewCellValue newValue,
        out string? reason)
    {
        newValue = default;
        reason = null;

        if (value.Kind == SourceValueKind.Unsupported)
        {
            reason = $"は{value.Reason}。";
            return false;
        }

        if (value.IsBlank)
        {
            reason = "が空欄です。現在のバージョンでは、空欄を転記してセルを消すことはしません。";
            return false;
        }

        if (writeKind == Mutation.CellWriteKind.Text)
        {
            if (value.Kind != SourceValueKind.Text)
            {
                reason = "は数値です。文字として転記するには、データ元を文字列にしてください。";
                return false;
            }

            newValue = Mutation.NewCellValue.OfText(value.Text!);
            return true;
        }

        if (value.Kind == SourceValueKind.Number)
        {
            newValue = Mutation.NewCellValue.OfNumber(value.Number);
            return true;
        }

        // CSV は値がすべて文字列なので、数値として読めるかここで確かめる。
        if (kind == SourceFileKind.Csv
            && double.TryParse(
                value.Text,
                System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture, out var number)
            && double.IsFinite(number))
        {
            newValue = Mutation.NewCellValue.OfNumber(number);
            return true;
        }

        reason = kind == SourceFileKind.Csv
            ? $"「{value.Text}」を数値として読み取れません。"
            : "は文字列です。数値として転記するには、データ元を数値にしてください。";
        return false;
    }
}

/// <summary>キー列だけを読んだ結果(表同士の突合更新の 1 パス目)。</summary>
internal sealed record SourceKeyScan
{
    /// <summary>非空欄のキー(重複していたものも含む)。</summary>
    public IReadOnlySet<string> Keys { get; init; } = new HashSet<string>(StringComparer.Ordinal);

    /// <summary>データ元で 2 件以上あったキー。</summary>
    public IReadOnlySet<string> DuplicateKeys { get; init; } = new HashSet<string>(StringComparer.Ordinal);

    /// <summary>キーが入っていた行の数(重複行も数える)。</summary>
    public int KeyedRowCount { get; init; }

    /// <summary>キーも使用対象の項目もすべて空欄だった行の数。</summary>
    public int BlankRowCount { get; init; }

    /// <summary>キーが空欄なのに使用対象の項目には値があった行の数。</summary>
    public int BlankKeyWithValueCount { get; init; }

    public string? Error { get; init; }

    public bool IsSuccess => Error is null;

    public static SourceKeyScan Failed(string error) => new() { Error = error };
}

/// <summary>必要なキーに一致する行を集めた結果。</summary>
internal sealed record SourceMatchResult
{
    /// <summary>キー → 一致した 1 行(重複していたキーは含めない)。</summary>
    public IReadOnlyDictionary<string, SourceRow> RowsByKey { get; init; }
        = new Dictionary<string, SourceRow>(StringComparer.Ordinal);

    /// <summary>データ元で 2 件以上あったキー。</summary>
    public IReadOnlyCollection<string> DuplicateKeys { get; init; } = Array.Empty<string>();

    /// <summary>キーが空欄で、他の項目にも値が無かった行の数(読み飛ばし)。</summary>
    public int BlankRowCount { get; init; }

    /// <summary>キーが空欄なのに他の項目には値があった行の数(読み飛ばすが知らせる)。</summary>
    public int BlankKeyWithValueCount { get; init; }

    /// <summary>今回どの転記先にも使わなかった行の数。</summary>
    public int UnusedRowCount { get; init; }

    public string? Error { get; init; }

    public bool IsSuccess => Error is null;

    public static SourceMatchResult Failed(string error) => new() { Error = error };
}
