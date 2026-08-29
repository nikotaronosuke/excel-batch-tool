using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace ExcelBatchTool.Core.Pdf;

/// <summary>
/// 取り出した内容を新しい .xlsx として書く。Microsoft Excel は不要。
///
/// 見た目を再現するのではなく、Excel で扱える構造にするのが目的なので、
/// 書式・数式は作らない。値は原則そのままの文字として書き、
/// **意味が変わらないと確実に言えるときだけ**数値にする
/// (先頭 0 の商品コード・電話番号・郵便番号風の文字列を数値にしない)。
/// </summary>
internal static class PdfWorkbookWriter
{
    public const string SheetName = "PDF抽出";

    public static void Write(Stream stream, IReadOnlyList<string[]> rows)
    {
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        using (var writer = OpenXmlWriter.Create(worksheetPart))
        {
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            for (var index = 0; index < rows.Count; index++)
            {
                var rowIndex = (uint)(index + 1);
                writer.WriteStartElement(new Row { RowIndex = rowIndex });

                for (var column = 0; column < rows[index].Length; column++)
                {
                    WriteCell(writer, rows[index][column], column + 1, rowIndex);
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        workbookPart.Workbook.AppendChild(new Sheets(new Sheet
        {
            Name = SheetName,
            SheetId = 1,
            Id = workbookPart.GetIdOfPart(worksheetPart),
        }));

        workbookPart.Workbook.Save();
    }

    private static void WriteCell(OpenXmlWriter writer, string value, int column, uint rowIndex)
    {
        var reference = CellRangeParser.ColumnIndexToLetters(column) + rowIndex.ToString(
            CultureInfo.InvariantCulture);

        if (value.Length == 0)
        {
            writer.WriteElement(new Cell { CellReference = reference });
            return;
        }

        if (IsSafeNumber(value, out var number))
        {
            writer.WriteElement(new Cell
            {
                CellReference = reference,
                CellValue = new CellValue(number.ToString(CultureInfo.InvariantCulture)),
            });
            return;
        }

        writer.WriteStartElement(new Cell
        {
            CellReference = reference,
            DataType = CellValues.InlineString,
        });
        writer.WriteStartElement(new InlineString());
        writer.WriteElement(new Text(value));
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    /// <summary>
    /// 数値として書いても意味が変わらないか。迷うものはすべて文字のままにする。
    /// 先頭 0(0123)・記号入り(000-1234-5678、1,200、A001)・長い桁は数値にしない。
    /// </summary>
    internal static bool IsSafeNumber(string value, out double number)
    {
        number = 0;

        if (value.Length == 0 || value.Length > 15)
        {
            return false;
        }

        foreach (var character in value)
        {
            if (!char.IsAsciiDigit(character) && character != '-' && character != '.')
            {
                return false;
            }
        }

        // 先頭 0 は桁として意味を持つことがある(0123 / 007)。
        var digits = value.StartsWith('-') ? value[1..] : value;
        if (digits.Length > 1 && digits[0] == '0' && !digits.StartsWith("0.", StringComparison.Ordinal))
        {
            return false;
        }

        if (!double.TryParse(
            value, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
            CultureInfo.InvariantCulture, out number) || !double.IsFinite(number))
        {
            return false;
        }

        // 書き戻したときに同じ文字になるものだけを数値として扱う。
        return number.ToString(CultureInfo.InvariantCulture) == value;
    }

    /// <summary>書いた .xlsx を開き直し、形式と中身が指定どおりか確かめる。</summary>
    public static string? Verify(string path, IReadOnlyList<string[]> expected)
    {
        try
        {
            using var document = SpreadsheetDocument.Open(path, isEditable: false);

            if (new OpenXmlValidator().Validate(document).Take(1).Any())
            {
                return "作成した Excel ファイルの形式に問題があります。作成を取り消しました。";
            }

            var workbookPart = document.WorkbookPart;
            var sheet = workbookPart?.Workbook?.Sheets?.Elements<Sheet>().FirstOrDefault();
            if (workbookPart is null || sheet?.Id?.Value is not { } relationshipId
                || workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                return "作成した Excel ファイルを読み直せませんでした。作成を取り消しました。";
            }

            var actual = ReadRows(worksheetPart, workbookPart);

            if (actual.Count != expected.Count)
            {
                return $"作成した Excel ファイルの行数({actual.Count:N0})が"
                    + $"予定({expected.Count:N0})と違います。作成を取り消しました。";
            }

            for (var row = 0; row < expected.Count; row++)
            {
                if (actual[row].Count != expected[row].Length)
                {
                    return $"作成した Excel ファイルの {row + 1} 行目の列数が違います。作成を取り消しました。";
                }

                for (var column = 0; column < expected[row].Length; column++)
                {
                    if (actual[row][column] != expected[row][column])
                    {
                        return $"作成した Excel ファイルの {row + 1} 行目の内容が指定と違います。"
                            + "作成を取り消しました。";
                    }
                }
            }

            return null;
        }
        catch (Exception ex)
        {
            return $"作成した Excel ファイルを確認できませんでした: {ex.Message}";
        }
    }

    /// <summary>読み直した内容を、書いたときと同じ文字列の形で取り出す。</summary>
    private static List<List<string>> ReadRows(WorksheetPart worksheetPart, WorkbookPart workbookPart)
    {
        var rows = new List<List<string>>();

        using var reader = OpenXmlReader.Create(worksheetPart);
        while (reader.Read())
        {
            if (!reader.IsStartElement || reader.ElementType != typeof(Row))
            {
                continue;
            }

            var row = (Row)reader.LoadCurrentElement()!;
            var cells = new List<(int Column, string Value)>();

            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value is not { } reference
                    || !CellRangeParser.TryParseCell(reference, out var column, out _))
                {
                    continue;
                }

                cells.Add((column, ReadCell(cell, workbookPart)));
            }

            var width = cells.Count == 0 ? 0 : cells.Max(cell => cell.Column);
            var values = Enumerable.Repeat(string.Empty, width).ToList();
            foreach (var (column, value) in cells)
            {
                values[column - 1] = value;
            }

            rows.Add(values);
        }

        return rows;
    }

    private static string ReadCell(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString?.Text?.Text ?? string.Empty;
        }

        if (cell.DataType?.Value == CellValues.SharedString
            && int.TryParse(cell.CellValue?.Text, out var index))
        {
            return workbookPart.SharedStringTablePart?.SharedStringTable
                ?.Elements<SharedStringItem>().ElementAtOrDefault(index)?.Text?.Text ?? string.Empty;
        }

        return cell.CellValue?.Text ?? string.Empty;
    }
}
