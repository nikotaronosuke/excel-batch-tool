using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBatchTool.Core.Tests;

/// <summary>転記元テスト用のセル。値は string / int / double / null / Formula / Styled。</summary>
internal sealed record SourceTestCell(object? Value, int? StyleId = null)
{
    // 引数名を明示する。省略すると record のコピーコンストラクターが選ばれ、
    // その引数変換でこの演算子自身が再び呼ばれて無限再帰になる。
    public static implicit operator SourceTestCell(string? text) => new(Value: text);

    public static implicit operator SourceTestCell(int number) => new(Value: number);

    public static implicit operator SourceTestCell(double number) => new(Value: number);
}

/// <summary>転記元テスト用の数式セル。</summary>
internal sealed record SourceTestFormula(string Formula, string CachedValue);

/// <summary>転記元テスト用の .xlsx / .csv を架空データのみで生成する。</summary>
internal static class TestSourceTableFactory
{
    /// <summary>
    /// 転記元の .xlsx を作る。<paramref name="rows"/> は項目名の行を含む(先頭が
    /// <paramref name="headerRow"/> 行目に置かれる)。
    /// </summary>
    public static void CreateXlsx(
        string path,
        string sheetName,
        IReadOnlyList<SourceTestCell[]> rows,
        int headerRow = 1,
        IReadOnlyList<MutationTestStyle>? styles = null,
        bool addSecondSheet = false,
        bool addStyleSchemaError = false)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = new Sheets();
        workbookPart.Workbook.AppendChild(sheets);

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = BuildStylesheet(styles ?? []);

        if (addStyleSchemaError)
        {
            // numFmt には formatCode が必須。開けるが Excel の形式としては不正な状態にする。
            stylesPart.Stylesheet.InsertAt(
                new NumberingFormats(new NumberingFormat { NumberFormatId = 200U }) { Count = 1U }, 0);
        }

        stylesPart.Stylesheet.Save();

        var sharedStrings = new List<SharedStringItem>();

        AddSheet(workbookPart, sheets, sheetName, rows, headerRow, sharedStrings, 1);

        if (addSecondSheet)
        {
            AddSheet(workbookPart, sheets, "参考", [[(SourceTestCell)"メモ"]], 1, sharedStrings, 2);
        }

        var sharedPart = workbookPart.AddNewPart<SharedStringTablePart>();
        sharedPart.SharedStringTable = new SharedStringTable(
            sharedStrings.Select(item => item.CloneNode(true)))
        {
            Count = (uint)sharedStrings.Count,
            UniqueCount = (uint)sharedStrings.Count,
        };
        sharedPart.SharedStringTable.Save();

        workbookPart.Workbook.Save();
    }

    /// <summary>転記元の .csv を作る。文字コードを明示できる。</summary>
    public static void CreateCsv(
        string path,
        IReadOnlyList<string> lines,
        string encodingName = "utf-8",
        bool withBom = false,
        string newLine = "\r\n")
    {
        var text = string.Join(newLine, lines) + newLine;
        var encoding = GetEncoding(encodingName, withBom);
        File.WriteAllBytes(path, encoding.GetPreamble().Concat(encoding.GetBytes(text)).ToArray());
    }

    private static Encoding GetEncoding(string name, bool withBom)
    {
        if (string.Equals(name, "utf-8", StringComparison.OrdinalIgnoreCase))
        {
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: withBom);
        }

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(932);
    }

    private static void AddSheet(
        WorkbookPart workbookPart,
        Sheets sheets,
        string sheetName,
        IReadOnlyList<SourceTestCell[]> rows,
        int headerRow,
        List<SharedStringItem> sharedStrings,
        uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();

        for (var offset = 0; offset < rows.Count; offset++)
        {
            var rowIndex = (uint)(headerRow + offset);
            var row = new Row { RowIndex = rowIndex };
            var values = rows[offset];

            for (var column = 0; column < values.Length; column++)
            {
                if (BuildCell(values[column], column + 1, rowIndex, sharedStrings) is { } cell)
                {
                    row.Append(cell);
                }
            }

            sheetData.Append(row);
        }

        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = sheetName,
        });
    }

    private static Cell? BuildCell(
        SourceTestCell spec, int column, uint rowIndex, List<SharedStringItem> sharedStrings)
    {
        if (spec.Value is null && spec.StyleId is null)
        {
            return null; // セルそのものを作らない。
        }

        var cell = new Cell
        {
            CellReference = $"{CellRangeParser.ColumnIndexToLetters(column)}{rowIndex}",
        };

        if (spec.StyleId is { } styleId)
        {
            cell.StyleIndex = (uint)styleId;
        }

        switch (spec.Value)
        {
            case null:
                break;

            case SourceTestFormula formula:
                cell.CellFormula = new CellFormula(formula.Formula);
                cell.CellValue = new CellValue(formula.CachedValue);
                break;

            case InlineSourceText inline:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(inline.Text));
                break;

            case RichSourceText rich:
                var richIndex = sharedStrings.Count;
                sharedStrings.Add(new SharedStringItem(
                    new Run(new RunProperties(new Bold()), new Text(rich.Bold)),
                    new Run(new Text(rich.Plain))));
                cell.DataType = CellValues.SharedString;
                cell.CellValue = new CellValue(richIndex.ToString(CultureInfo.InvariantCulture));
                break;

            case bool flag:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(flag ? "1" : "0");
                break;

            case string text:
                var index = sharedStrings.Count;
                sharedStrings.Add(new SharedStringItem(new Text(text)));
                cell.DataType = CellValues.SharedString;
                cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
                break;

            default:
                cell.CellValue = new CellValue(
                    Convert.ToDouble(spec.Value, CultureInfo.InvariantCulture)
                        .ToString(CultureInfo.InvariantCulture));
                break;
        }

        return cell;
    }

    private static Stylesheet BuildStylesheet(IReadOnlyList<MutationTestStyle> styles)
    {
        var numberingFormats = new NumberingFormats();
        var cellFormats = new CellFormats(new CellFormat
        {
            NumberFormatId = 0U,
            FontId = 0U,
            FillId = 0U,
            BorderId = 0U,
            FormatId = 0U,
        });

        foreach (var style in styles)
        {
            if (style.FormatCode is { } code)
            {
                numberingFormats.Append(new NumberingFormat
                {
                    NumberFormatId = style.NumberFormatId,
                    FormatCode = code,
                });
            }

            cellFormats.Append(new CellFormat
            {
                NumberFormatId = style.NumberFormatId,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
                ApplyNumberFormat = true,
            });
        }

        var stylesheet = new Stylesheet();
        if (numberingFormats.Any())
        {
            numberingFormats.Count = (uint)numberingFormats.Count();
            stylesheet.Append(numberingFormats);
        }

        stylesheet.Append(new Fonts(new Font(new FontSize { Val = 11D })) { Count = 1U });
        stylesheet.Append(new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2U });
        stylesheet.Append(new Borders(new Border()) { Count = 1U });
        stylesheet.Append(new CellStyleFormats(
            new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U }) { Count = 1U });

        cellFormats.Count = (uint)cellFormats.Count();
        stylesheet.Append(cellFormats);
        stylesheet.Append(new CellStyles(
            new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U }) { Count = 1U });

        return stylesheet;
    }
}

/// <summary>共有文字列ではなくセル内に直接書く文字列。</summary>
internal sealed record InlineSourceText(string Text);

/// <summary>文字ごとに書式を持つ文字列(リッチテキスト)。</summary>
internal sealed record RichSourceText(string Bold, string Plain);
