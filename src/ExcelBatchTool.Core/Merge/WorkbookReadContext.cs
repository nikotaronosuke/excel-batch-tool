using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBatchTool.Core.Merge;

/// <summary>数値セルの表示形式の分類。</summary>
internal enum NumberFormatKind
{
    /// <summary>日付・時刻ではない。</summary>
    None = 0,
    Date,
    DateTime,
    Time,

    /// <summary>日付か数値か確信が持てない(月/分が区別できない書式など)。</summary>
    Ambiguous,
}

/// <summary>
/// 1 つの Workbook を読む際に必要な共有情報(共有文字列・表示形式・date system)。
/// セル値の解釈をここに集約する。
/// </summary>
internal sealed class WorkbookReadContext
{
    private readonly string[] _sharedStrings;
    private readonly NumberFormatKind[] _styleKinds;
    private readonly bool _date1904;

    private WorkbookReadContext(string[] sharedStrings, NumberFormatKind[] styleKinds, bool date1904)
    {
        _sharedStrings = sharedStrings;
        _styleKinds = styleKinds;
        _date1904 = date1904;
    }

    public static WorkbookReadContext Create(WorkbookPart workbookPart)
    {
        var sharedStrings = LoadSharedStrings(workbookPart);
        var styleKinds = LoadStyleKinds(workbookPart);
        var date1904 = workbookPart.Workbook?.WorkbookProperties?.Date1904?.Value ?? false;
        return new WorkbookReadContext(sharedStrings, styleKinds, date1904);
    }

    /// <summary>セルの値を読む。日付は 1900 date system の serial 値へ正規化する。</summary>
    /// <param name="ambiguousNumberFormat">
    /// 日付か数値か確信が持てない表示形式だったときに true。値は数値として返す。
    /// </param>
    public MergeCellValue ReadCell(Cell cell, out bool ambiguousNumberFormat)
    {
        ambiguousNumberFormat = false;
        var dataType = cell.DataType?.Value;

        if (dataType == CellValues.SharedString)
        {
            var raw = cell.CellValue?.InnerText;
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index)
                && index >= 0 && index < _sharedStrings.Length
                    ? MergeCellValue.FromText(_sharedStrings[index])
                    : MergeCellValue.Blank;
        }

        if (dataType == CellValues.InlineString)
        {
            return MergeCellValue.FromText(cell.InlineString?.InnerText);
        }

        if (dataType == CellValues.String)
        {
            return MergeCellValue.FromText(cell.CellValue?.InnerText);
        }

        if (dataType == CellValues.Boolean)
        {
            var raw = cell.CellValue?.InnerText;
            if (string.IsNullOrEmpty(raw))
            {
                return MergeCellValue.Blank;
            }

            return MergeCellValue.FromBoolean(raw == "1" || string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase));
        }

        if (dataType == CellValues.Error)
        {
            // エラー値(#REF! 等)は文字列としてそのまま持ち越す。
            return MergeCellValue.FromText(cell.CellValue?.InnerText);
        }

        if (dataType == CellValues.Date)
        {
            var raw = cell.CellValue?.InnerText;
            return DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
                ? MergeCellValue.FromDateSerial(
                    MergeCellValue.DateTimeToSerial(parsed),
                    parsed.TimeOfDay == TimeSpan.Zero ? MergeValueKind.Date : MergeValueKind.DateTime)
                : MergeCellValue.FromText(raw);
        }

        // 型指定なし = 数値。
        var text = cell.CellValue?.InnerText;
        if (string.IsNullOrEmpty(text))
        {
            return MergeCellValue.Blank;
        }

        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
        {
            return MergeCellValue.FromText(text);
        }

        var kind = KindForStyle(cell.StyleIndex?.Value);
        if (kind == NumberFormatKind.Ambiguous)
        {
            ambiguousNumberFormat = true;
            return MergeCellValue.FromNumber(number);
        }

        if (kind is NumberFormatKind.None)
        {
            return MergeCellValue.FromNumber(number);
        }

        var serial = _date1904 ? number + 1462 : number;
        if (serial < 0 || serial > 2958465)
        {
            // date として解釈できない範囲。誤変換せず数値のまま扱う。
            return MergeCellValue.FromNumber(number);
        }

        return MergeCellValue.FromDateSerial(serial, kind switch
        {
            NumberFormatKind.Date => MergeValueKind.Date,
            NumberFormatKind.DateTime => MergeValueKind.DateTime,
            _ => MergeValueKind.Time,
        });
    }

    private NumberFormatKind KindForStyle(uint? styleIndex)
    {
        if (styleIndex is not { } index || index >= _styleKinds.Length)
        {
            return NumberFormatKind.None;
        }

        return _styleKinds[index];
    }

    private static string[] LoadSharedStrings(WorkbookPart workbookPart)
    {
        var part = workbookPart.SharedStringTablePart;
        if (part is null)
        {
            return [];
        }

        var values = new List<string>();
        using var reader = OpenXmlReader.Create(part);
        while (reader.Read())
        {
            if (reader.IsStartElement && reader.ElementType == typeof(SharedStringItem))
            {
                values.Add(((SharedStringItem)reader.LoadCurrentElement()!).InnerText);
            }
        }

        return [.. values];
    }

    private static NumberFormatKind[] LoadStyleKinds(WorkbookPart workbookPart)
    {
        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        var cellFormats = stylesheet?.CellFormats;
        if (cellFormats is null)
        {
            return [];
        }

        var customFormats = new Dictionary<uint, string>();
        foreach (var format in stylesheet!.NumberingFormats?.Elements<NumberingFormat>() ?? [])
        {
            if (format.NumberFormatId?.Value is { } id && format.FormatCode?.Value is { } code)
            {
                customFormats[id] = code;
            }
        }

        var kinds = new List<NumberFormatKind>();
        foreach (var format in cellFormats.Elements<CellFormat>())
        {
            var id = format.NumberFormatId?.Value ?? 0;
            kinds.Add(customFormats.TryGetValue(id, out var code)
                ? ClassifyFormatCode(code)
                : ClassifyBuiltInFormat(id));
        }

        return [.. kinds];
    }

    /// <summary>組み込み表示形式 ID の分類(ECMA-376 の既定 numFmtId)。</summary>
    private static NumberFormatKind ClassifyBuiltInFormat(uint id) => id switch
    {
        14 or 15 or 16 or 17 => NumberFormatKind.Date,
        22 => NumberFormatKind.DateTime,
        18 or 19 or 20 or 21 or 45 or 46 or 47 => NumberFormatKind.Time,
        // 東アジア向けの日付・時刻書式。
        27 or 28 or 29 or 30 or 31 or 36 or 50 or 51 or 52 or 53 or 54 or 57 or 58 => NumberFormatKind.Date,
        32 or 33 or 34 or 35 or 55 or 56 => NumberFormatKind.Time,
        _ => NumberFormatKind.None,
    };

    /// <summary>
    /// ユーザー定義の表示形式コードを分類する。リテラル("..." / \x / [..])を取り除いてから
    /// 日付・時刻トークンを探す。月(m)しか無い場合は分(minute)と区別できないため Ambiguous。
    /// </summary>
    internal static NumberFormatKind ClassifyFormatCode(string formatCode)
    {
        if (string.IsNullOrWhiteSpace(formatCode))
        {
            return NumberFormatKind.None;
        }

        // 正の値のセクションだけを見る。
        var section = SplitFirstSection(formatCode);
        var tokens = StripLiterals(section, out var hasElapsedTimeToken);

        var hasDatePart = false;
        var hasTimePart = hasElapsedTimeToken;
        var hasMonthOrMinute = false;

        for (var i = 0; i < tokens.Length; i++)
        {
            switch (char.ToLowerInvariant(tokens[i]))
            {
                case 'y':
                case 'd':
                case 'e':
                case 'g':
                    hasDatePart = true;
                    break;
                case 'h':
                case 's':
                    hasTimePart = true;
                    break;
                case 'm':
                    hasMonthOrMinute = true;
                    break;
            }
        }

        if (hasDatePart && hasTimePart)
        {
            return NumberFormatKind.DateTime;
        }

        if (hasDatePart)
        {
            return NumberFormatKind.Date;
        }

        if (hasTimePart)
        {
            return NumberFormatKind.Time;
        }

        return hasMonthOrMinute ? NumberFormatKind.Ambiguous : NumberFormatKind.None;
    }

    private static string SplitFirstSection(string formatCode)
    {
        var inQuote = false;
        for (var i = 0; i < formatCode.Length; i++)
        {
            var c = formatCode[i];
            if (c == '"')
            {
                inQuote = !inQuote;
            }
            else if (c == '\\')
            {
                i++;
            }
            else if (c == ';' && !inQuote)
            {
                return formatCode[..i];
            }
        }

        return formatCode;
    }

    private static string StripLiterals(string formatCode, out bool hasElapsedTimeToken)
    {
        hasElapsedTimeToken = false;
        var builder = new StringBuilder(formatCode.Length);

        for (var i = 0; i < formatCode.Length; i++)
        {
            var c = formatCode[i];
            switch (c)
            {
                case '"':
                    i++;
                    while (i < formatCode.Length && formatCode[i] != '"')
                    {
                        i++;
                    }

                    break;

                case '\\':
                    i++; // 次の 1 文字はリテラル。
                    break;

                case '[':
                {
                    var end = formatCode.IndexOf(']', i + 1);
                    if (end < 0)
                    {
                        i = formatCode.Length;
                        break;
                    }

                    // [h] / [mm] / [ss] は経過時間トークンで、時刻として扱う。
                    var inner = formatCode[(i + 1)..end];
                    if (inner.Length > 0 && inner.All(ch => ch is 'h' or 'H' or 'm' or 'M' or 's' or 'S'))
                    {
                        hasElapsedTimeToken = true;
                    }

                    i = end;
                    break;
                }

                default:
                    // AM/PM・A/P は時刻マーカー。'M' を月と誤認しないよう読み飛ばす。
                    if (TryConsumeAmPm(formatCode, ref i))
                    {
                        hasElapsedTimeToken = true;
                        break;
                    }

                    builder.Append(c);
                    break;
            }
        }

        return builder.ToString();
    }

    private static readonly string[] AmPmMarkers = ["AM/PM", "A/P"];

    private static bool TryConsumeAmPm(string formatCode, ref int index)
    {
        foreach (var marker in AmPmMarkers)
        {
            if (index + marker.Length <= formatCode.Length
                && string.Compare(formatCode, index, marker, 0, marker.Length, StringComparison.OrdinalIgnoreCase) == 0)
            {
                index += marker.Length - 1;
                return true;
            }
        }

        return false;
    }
}
