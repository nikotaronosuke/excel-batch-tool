using System.Globalization;

namespace ExcelBatchTool.Core.Merge;

/// <summary>統合時に扱うセル値の種類。</summary>
public enum MergeValueKind
{
    Blank = 0,
    Text,
    Number,
    Boolean,
    Date,
    DateTime,
    Time,
}

/// <summary>
/// 統合時に受け渡すセル値。日付は「1900 date system の serial 値」に正規化して保持する
/// (1904 date system の入力は読み取り時に +1462 して変換済み)。
/// </summary>
public readonly record struct MergeCellValue
{
    public MergeValueKind Kind { get; private init; }

    public string? Text { get; private init; }

    /// <summary>数値、または日付・時刻の serial 値。</summary>
    public double Number { get; private init; }

    public bool Boolean { get; private init; }

    public static MergeCellValue Blank => new() { Kind = MergeValueKind.Blank };

    public bool IsBlank => Kind == MergeValueKind.Blank;

    public static MergeCellValue FromText(string? text)
        => string.IsNullOrEmpty(text)
            ? Blank
            : new MergeCellValue { Kind = MergeValueKind.Text, Text = text };

    public static MergeCellValue FromNumber(double value)
        => new() { Kind = MergeValueKind.Number, Number = value };

    public static MergeCellValue FromBoolean(bool value)
        => new() { Kind = MergeValueKind.Boolean, Boolean = value };

    /// <summary>日付・時刻。<paramref name="serial"/> は 1900 date system の serial 値。</summary>
    public static MergeCellValue FromDateSerial(double serial, MergeValueKind kind)
        => new() { Kind = kind, Number = serial };

    /// <summary>Header 比較などに使う文字列表現(trim 前の生の値)。</summary>
    public string ToDisplayString() => Kind switch
    {
        MergeValueKind.Blank => string.Empty,
        MergeValueKind.Text => Text ?? string.Empty,
        MergeValueKind.Boolean => Boolean ? "TRUE" : "FALSE",
        MergeValueKind.Number => Number.ToString(CultureInfo.InvariantCulture),
        _ => SerialToDateTime(Number)?.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
             ?? Number.ToString(CultureInfo.InvariantCulture),
    };

    /// <summary>1900 date system の serial 値を DateTime へ変換する(範囲外は null)。</summary>
    public static DateTime? SerialToDateTime(double serial)
    {
        if (serial < 0 || serial > 2958465)
        {
            return null;
        }

        // serial 60 は Excel 上の存在しない 1900-02-29。60 未満は 1899-12-31 起点で数える。
        return serial < 60
            ? new DateTime(1899, 12, 31).AddDays(serial)
            : DateTime.FromOADate(serial);
    }

    /// <summary>DateTime を 1900 date system の serial 値へ変換する。</summary>
    public static double DateTimeToSerial(DateTime value) => value.ToOADate();
}
