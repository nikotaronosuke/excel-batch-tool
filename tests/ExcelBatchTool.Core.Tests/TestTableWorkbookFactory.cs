using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Charts = DocumentFormat.OpenXml.Drawing.Charts;

namespace ExcelBatchTool.Core.Tests;

/// <summary>日付か数値か区別できない表示形式("mm")を持たせたい値。</summary>
internal sealed record AmbiguousFormatted(double Value);

/// <summary>時刻(小数部のみの serial)として書き込みたい値。</summary>
internal sealed record TimeOfDayValue(TimeSpan Value);

/// <summary>テスト用ワークシートの内容。すべて架空データ。</summary>
internal sealed class TestSheetSpec
{
    public required string Name { get; init; }

    public string[] Headers { get; init; } = [];

    /// <summary>
    /// データ行。長さ 0 の配列は「完全空行」を表す。
    /// 値は string / int / double / bool / DateTime / TimeOfDayValue / AmbiguousFormatted / null(空セル)。
    /// </summary>
    public object?[][] Rows { get; init; } = [];

    /// <summary>結合セルの範囲(例 "A1:B2")。</summary>
    public string? MergeReference { get; init; }

    /// <summary>データ末尾に数式セルを 1 つ追加する。</summary>
    public bool AddFormulaCell { get; init; }

    public bool AddChart { get; init; }

    public bool AddImage { get; init; }
}

/// <summary>
/// 統合テスト用の .xlsx を架空データのみで生成する。実業務データは使用しない。
/// </summary>
internal static class TestTableWorkbookFactory
{
    private const uint StyleDefault = 0;
    private const uint StyleDate = 1;
    private const uint StyleDateTime = 2;
    private const uint StyleTime = 3;
    private const uint StyleAmbiguous = 4;

    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    /// <summary>ヘッダー + データ行を持つ 1 シートの Workbook を作る。</summary>
    public static void CreateTable(
        string path,
        string sheetName,
        string[] headers,
        object?[][] rows,
        bool useSharedStrings = false,
        bool date1904 = false)
        => Create(path, [new TestSheetSpec { Name = sheetName, Headers = headers, Rows = rows }],
            useSharedStrings, date1904);

    public static void Create(
        string path,
        IReadOnlyList<TestSheetSpec> sheets,
        bool useSharedStrings = false,
        bool date1904 = false)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        if (date1904)
        {
            workbookPart.Workbook.AppendChild(new WorkbookProperties { Date1904 = true });
        }

        workbookPart.Workbook.AppendChild(new Sheets());

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = BuildStylesheet();
        stylesPart.Stylesheet.Save();

        var sharedStrings = useSharedStrings ? new List<string>() : null;

        uint sheetId = 1;
        foreach (var spec in sheets)
        {
            AddSheet(workbookPart, spec, sheetId++, sharedStrings, date1904);
        }

        if (sharedStrings is not null)
        {
            var part = workbookPart.AddNewPart<SharedStringTablePart>();
            part.SharedStringTable = new SharedStringTable(
                sharedStrings.Select(value => new SharedStringItem(new Text(value))))
            {
                Count = (uint)sharedStrings.Count,
                UniqueCount = (uint)sharedStrings.Count,
            };
            part.SharedStringTable.Save();
        }
    }

    private static void AddSheet(
        WorkbookPart workbookPart,
        TestSheetSpec spec,
        uint sheetId,
        List<string>? sharedStrings,
        bool date1904)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();

        uint rowIndex = 1;
        if (spec.Headers.Length > 0)
        {
            var headerRow = new Row { RowIndex = rowIndex };
            for (var i = 0; i < spec.Headers.Length; i++)
            {
                if (spec.Headers[i].Length == 0)
                {
                    continue; // 空ヘッダーはセル自体を書かない。
                }

                headerRow.Append(BuildCell(Reference(i + 1, rowIndex), spec.Headers[i], sharedStrings, date1904));
            }

            sheetData.Append(headerRow);
        }

        foreach (var values in spec.Rows)
        {
            rowIndex++;
            var row = new Row { RowIndex = rowIndex };
            for (var i = 0; i < values.Length; i++)
            {
                if (values[i] is null)
                {
                    continue;
                }

                row.Append(BuildCell(Reference(i + 1, rowIndex), values[i]!, sharedStrings, date1904));
            }

            sheetData.Append(row); // 空行も row 要素として残す。
        }

        if (spec.AddFormulaCell)
        {
            rowIndex++;
            var row = new Row { RowIndex = rowIndex };
            row.Append(new Cell
            {
                CellReference = Reference(1, rowIndex),
                CellFormula = new CellFormula("1+1"),
                CellValue = new CellValue("2"),
            });
            sheetData.Append(row);
        }

        var dimension = new SheetDimension
        {
            Reference = $"A1:{CellRangeParser.ColumnIndexToLetters(Math.Max(1, spec.Headers.Length))}{Math.Max(1, rowIndex)}",
        };

        var worksheet = new Worksheet(dimension, sheetData);

        if (spec.MergeReference is { } mergeReference)
        {
            worksheet.InsertAfter(new MergeCells(new MergeCell { Reference = mergeReference }), sheetData);
        }

        worksheetPart.Worksheet = worksheet;

        if (spec.AddChart || spec.AddImage)
        {
            var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();

            if (spec.AddChart)
            {
                var chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = new Charts.ChartSpace(new Charts.Chart(new Charts.PlotArea()));
            }

            if (spec.AddImage)
            {
                var imagePart = drawingsPart.AddImagePart("image/png");
                using var stream = new MemoryStream(TinyPng);
                imagePart.FeedData(stream);
            }

            worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
        }

        workbookPart.Workbook!.Sheets!.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = spec.Name,
        });
    }

    private static Cell BuildCell(string reference, object value, List<string>? sharedStrings, bool date1904)
        => value switch
        {
            string text when sharedStrings is not null => new Cell
            {
                CellReference = reference,
                DataType = CellValues.SharedString,
                CellValue = new CellValue(AddSharedString(sharedStrings, text).ToString()),
            },

            string text => new Cell
            {
                CellReference = reference,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(text)),
            },

            bool flag => new Cell
            {
                CellReference = reference,
                DataType = CellValues.Boolean,
                CellValue = new CellValue(flag ? "1" : "0"),
            },

            DateTime date => new Cell
            {
                CellReference = reference,
                StyleIndex = date.TimeOfDay == TimeSpan.Zero ? StyleDate : StyleDateTime,
                CellValue = new CellValue(ToSerial(date, date1904).ToString(System.Globalization.CultureInfo.InvariantCulture)),
            },

            TimeOfDayValue time => new Cell
            {
                CellReference = reference,
                StyleIndex = StyleTime,
                CellValue = new CellValue(time.Value.TotalDays.ToString(System.Globalization.CultureInfo.InvariantCulture)),
            },

            AmbiguousFormatted ambiguous => new Cell
            {
                CellReference = reference,
                StyleIndex = StyleAmbiguous,
                CellValue = new CellValue(ambiguous.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)),
            },

            _ => new Cell
            {
                CellReference = reference,
                StyleIndex = StyleDefault,
                CellValue = new CellValue(Convert.ToDouble(value).ToString(System.Globalization.CultureInfo.InvariantCulture)),
            },
        };

    private static double ToSerial(DateTime value, bool date1904)
        => date1904 ? value.ToOADate() - 1462 : value.ToOADate();

    private static int AddSharedString(List<string> sharedStrings, string text)
    {
        var index = sharedStrings.IndexOf(text);
        if (index >= 0)
        {
            return index;
        }

        sharedStrings.Add(text);
        return sharedStrings.Count - 1;
    }

    private static string Reference(int column, uint row)
        => $"{CellRangeParser.ColumnIndexToLetters(column)}{row}";

    private static Stylesheet BuildStylesheet() => new(
        new NumberingFormats(
            new NumberingFormat { NumberFormatId = 165U, FormatCode = "mm" })
        { Count = 1U },
        new Fonts(new Font(new FontSize { Val = 11D })) { Count = 1U },
        new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }))
        { Count = 2U },
        new Borders(new Border()) { Count = 1U },
        new CellStyleFormats(new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U }) { Count = 1U },
        new CellFormats(
            new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U },
            new CellFormat { NumberFormatId = 14U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyNumberFormat = true },
            new CellFormat { NumberFormatId = 22U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyNumberFormat = true },
            new CellFormat { NumberFormatId = 21U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyNumberFormat = true },
            new CellFormat { NumberFormatId = 165U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyNumberFormat = true })
        { Count = 5U });
}
