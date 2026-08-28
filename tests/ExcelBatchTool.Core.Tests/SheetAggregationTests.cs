using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Aggregation;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 1B.1: 複数 Workbook の Worksheet を 1 つの新規 Workbook へ集約する。
/// すべて架空データで生成した Workbook を使う。
/// </summary>
public sealed class SheetAggregationTests
{
    [Fact]
    public void Aggregate_SheetsFromTwoWorkbooks_ProducesOneWorkbookWithBothSheets()
    {
        using var dir = new TempDir();
        var osaka = dir.File("大阪.xlsx");
        var kyoto = dir.File("京都.xlsx");
        TestSheetWorkbookFactory.Create(osaka, [Sheet("売上", [["商品", "金額"], ["架空A", 100]])]);
        TestSheetWorkbookFactory.Create(kyoto, [Sheet("在庫", [["商品", "数量"], ["架空B", 5]])]);

        var output = dir.File("月次資料.xlsx");
        var result = Aggregate(output, (osaka, "売上"), (kyoto, "在庫"));

        Assert.True(result.Success, result.Message);
        Assert.Equal(2, result.WorkbookCount);
        Assert.Equal(2, result.SheetCount);
        Assert.Equal(new[] { "売上", "在庫" }, ReadSheetNames(output));
        Assert.Equal(new[] { "商品", "金額" }, ReadRowTexts(output, "売上")[0]);
        Assert.Equal(new[] { "架空B", "5" }, ReadRowTexts(output, "在庫")[1]);
    }

    [Fact]
    public void Aggregate_MultipleSheetsFromOneWorkbook_KeepsAllOfThem()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            Sheet("売上", [["A", 1]]),
            Sheet("在庫", [["B", 2]]),
            Sheet("メモ", [["C", 3]]),
        ]);

        var output = dir.File("out.xlsx");
        var result = Aggregate(output, (path, "売上"), (path, "在庫"), (path, "メモ"));

        Assert.True(result.Success, result.Message);
        Assert.Equal(1, result.WorkbookCount);
        Assert.Equal(3, result.SheetCount);
        Assert.Equal(new[] { "売上", "在庫", "メモ" }, ReadSheetNames(output));
    }

    [Fact]
    public void Aggregate_OutputSheetOrder_FollowsTheSelectionOrder()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        TestSheetWorkbookFactory.Create(a, [Sheet("一", [["1"]]), Sheet("二", [["2"]])]);
        TestSheetWorkbookFactory.Create(b, [Sheet("三", [["3"]])]);

        var preview = CreatePreview((a, "一"), (a, "二"), (b, "三"));
        Assert.Equal(new[] { 1, 2, 3 }, preview.Sheets.Select(sheet => sheet.Order));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);
        Assert.Equal(new[] { "一", "二", "三" }, ReadSheetNames(output));
    }

    [Fact]
    public void Aggregate_CellTypes_ArePreserved()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        var date = new DateTime(2026, 3, 15);
        TestSheetWorkbookFactory.Create(path,
            [Sheet("表", [["架空テキスト", 123.5, true, new Styled(date, 0)]])],
            styles: [new TestStyle { BuiltinNumberFormatId = 14U }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var row = ReadRow(output, "表", 1);
        Assert.Equal(MergeValueKind.Text, row[0].Kind);
        Assert.Equal("架空テキスト", row[0].Text);
        Assert.Equal(MergeValueKind.Number, row[1].Kind);
        Assert.Equal(123.5, row[1].Number, 6);
        Assert.Equal(MergeValueKind.Boolean, row[2].Kind);
        Assert.True(row[2].Boolean);
        Assert.Equal(MergeValueKind.Date, row[3].Kind);
        Assert.Equal(date, MergeCellValue.SerialToDateTime(row[3].Number));
    }

    [Fact]
    public void Aggregate_MixedDateSystems_KeepTheSameCalendarDate()
    {
        using var dir = new TempDir();
        var system1900 = dir.File("1900.xlsx");
        var system1904 = dir.File("1904.xlsx");
        var date = new DateTime(2026, 3, 15);
        TestStyle[] styles = [new TestStyle { BuiltinNumberFormatId = 14U }];

        TestSheetWorkbookFactory.Create(system1900, [Sheet("表", [[new Styled(date, 0)]])], styles);
        TestSheetWorkbookFactory.Create(system1904, [Sheet("表", [[new Styled(date, 0)]])], styles, date1904: true);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (system1900, "表"), (system1904, "表")).Success);

        foreach (var sheetName in ReadSheetNames(output))
        {
            var row = ReadRow(output, sheetName, 1);
            Assert.Equal(MergeValueKind.Date, row[0].Kind);
            Assert.Equal(date, MergeCellValue.SerialToDateTime(row[0].Number));
        }
    }

    [Fact]
    public void Aggregate_CellStyles_ArePreserved()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestStyle[] styles =
        [
            new TestStyle { Bold = true },
            new TestStyle { FillArgb = "FFFF0000" },
            new TestStyle { ThinBorder = true },
            new TestStyle { HorizontalAlignment = "center" },
            new TestStyle { NumberFormatCode = "0.00\"円\"" },
        ];
        TestSheetWorkbookFactory.Create(path,
            [Sheet("表", [[new Styled("太字", 0), new Styled("赤", 1), new Styled("枠", 2), new Styled("中央", 3), new Styled(12.5, 4)]])],
            styles);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet!;
        var cells = ReadCells(workbookPart, "表").ToDictionary(cell => cell.CellReference!.Value!);

        Assert.True(FormatOf(stylesheet, cells["A1"]).ApplyFont?.Value);
        Assert.True(FontOf(stylesheet, cells["A1"]).Bold is not null);

        Assert.True(FormatOf(stylesheet, cells["B1"]).ApplyFill?.Value);
        var fill = FillOf(stylesheet, cells["B1"]).PatternFill!;
        Assert.Equal(PatternValues.Solid, fill.PatternType?.Value);
        Assert.Equal("FFFF0000", fill.ForegroundColor?.Rgb?.Value);

        Assert.True(FormatOf(stylesheet, cells["C1"]).ApplyBorder?.Value);
        Assert.Equal(BorderStyleValues.Thin, BorderOf(stylesheet, cells["C1"]).LeftBorder?.Style?.Value);

        Assert.Equal(HorizontalAlignmentValues.Center,
            FormatOf(stylesheet, cells["D1"]).Alignment?.Horizontal?.Value);

        var numberFormatId = FormatOf(stylesheet, cells["E1"]).NumberFormatId!.Value;
        var numberFormat = stylesheet.NumberingFormats!.Elements<NumberingFormat>()
            .Single(format => format.NumberFormatId!.Value == numberFormatId);
        Assert.Equal("0.00\"円\"", numberFormat.FormatCode?.Value);
    }

    [Fact]
    public void Aggregate_SameStyleIndexInDifferentWorkbooks_DoesNotCollide()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");

        // どちらの Workbook でも StyleIndex 1 だが、意味は違う。
        TestSheetWorkbookFactory.Create(a, [Sheet("A表", [[new Styled("太字", 0)]])],
            styles: [new TestStyle { Bold = true }]);
        TestSheetWorkbookFactory.Create(b, [Sheet("B表", [[new Styled("赤", 0)]])],
            styles: [new TestStyle { FillArgb = "FF00FF00" }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (a, "A表"), (b, "B表")).Success);

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet!;

        var aCell = ReadCells(workbookPart, "A表").Single(cell => cell.CellReference?.Value == "A1");
        var bCell = ReadCells(workbookPart, "B表").Single(cell => cell.CellReference?.Value == "A1");

        Assert.NotEqual(aCell.StyleIndex!.Value, bCell.StyleIndex!.Value);
        Assert.NotNull(FontOf(stylesheet, aCell).Bold);
        Assert.Null(FontOf(stylesheet, bCell).Bold);
        Assert.Equal("FF00FF00", FillOf(stylesheet, bCell).PatternFill?.ForegroundColor?.Rgb?.Value);
        Assert.NotEqual(PatternValues.Solid, FillOf(stylesheet, aCell).PatternFill?.PatternType?.Value);
    }

    [Fact]
    public void Aggregate_RowAndColumnSettings_ArePreserved()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["A1"], ["A2"], ["A3"]],
                RowProperties = [new TestRowProperty(1, Height: 42.5), new TestRowProperty(2, Hidden: true)],
                ColumnProperties = [new TestColumnProperty(1, 1, Width: 24.5), new TestColumnProperty(2, 3, Hidden: true)],
            },
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var worksheet = GetWorksheetPart(document.WorkbookPart!, "表").Worksheet!;

        var rows = worksheet.Descendants<Row>().ToDictionary(row => row.RowIndex!.Value);
        Assert.Equal(42.5, rows[1].Height!.Value, 3);
        Assert.True(rows[1].CustomHeight!.Value);
        Assert.True(rows[2].Hidden!.Value);
        Assert.Null(rows[3].Hidden);

        var columns = worksheet.Descendants<Column>().ToList();
        Assert.Equal(2, columns.Count);
        Assert.Equal(24.5, columns[0].Width!.Value, 3);
        Assert.True(columns[0].CustomWidth!.Value);
        Assert.True(columns[1].Hidden!.Value);
        Assert.Equal(2U, columns[1].Min!.Value);
        Assert.Equal(3U, columns[1].Max!.Value);
    }

    [Fact]
    public void Aggregate_MergedCellsAndFreezePane_ArePreserved()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["見出し", null, null], ["A", "B", "C"]],
                Merges = ["A1:C1"],
                FreezeTopRow = true,
            },
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var worksheet = GetWorksheetPart(document.WorkbookPart!, "表").Worksheet!;

        var merge = Assert.Single(worksheet.Descendants<MergeCell>());
        Assert.Equal("A1:C1", merge.Reference?.Value);

        var pane = Assert.Single(worksheet.Descendants<Pane>());
        Assert.Equal(1D, pane.VerticalSplit!.Value);
        Assert.Equal("A2", pane.TopLeftCell?.Value);
        Assert.Equal(PaneStateValues.Frozen, pane.State?.Value);
    }

    [Fact]
    public void Aggregate_HiddenSheet_KeepsItsVisibility()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            Sheet("表示", [["A"]]),
            new TestAggregationSheetSpec { Name = "非表示", IsHidden = true, Rows = [["B"]] },
        ]);

        var preview = CreatePreview((path, "表示"), (path, "非表示"));
        Assert.False(preview.Sheets[0].IsHidden);
        Assert.True(preview.Sheets[1].IsHidden);

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var sheets = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToList();

        Assert.Null(sheets[0].State);
        Assert.Equal(SheetStateValues.Hidden, sheets[1].State?.Value);
    }

    [Fact]
    public void Aggregate_SheetProtection_IsPreserved()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [new TestAggregationSheetSpec { Name = "保護", Rows = [["A"]], AddProtection = true }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "保護")).Success);

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var protection = Assert.Single(
            GetWorksheetPart(document.WorkbookPart!, "保護").Worksheet!.Descendants<SheetProtection>());
        Assert.True(protection.Sheet!.Value);
    }

    [Fact]
    public void Preview_SheetsWithTheSameName_GetDeterministicUniqueOutputNames()
    {
        using var dir = new TempDir();
        var a = dir.File("大阪.xlsx");
        var b = dir.File("京都.xlsx");
        var c = dir.File("神戸.xlsx");
        foreach (var path in new[] { a, b, c })
        {
            TestSheetWorkbookFactory.Create(path, [Sheet("売上", [["A", 1]])]);
        }

        var preview = CreatePreview((a, "売上"), (b, "売上"), (c, "売上"));

        Assert.True(preview.CanExecute);
        Assert.Equal(new[] { "売上", "売上 (2)", "売上 (3)" },
            preview.Sheets.Select(sheet => sheet.OutputSheetName));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);
        Assert.Equal(new[] { "売上", "売上 (2)", "売上 (3)" }, ReadSheetNames(output));
    }

    [Fact]
    public void Propose_LongSheetName_LeavesRoomForTheSuffix()
    {
        var longName = new string('あ', 31);

        var first = OutputSheetNameResolver.Propose(longName, []);
        var second = OutputSheetNameResolver.Propose(longName, [first]);

        Assert.Equal(31, first.Length);
        Assert.Equal(31, second.Length);
        Assert.EndsWith(" (2)", second);
        Assert.Null(OutputSheetNameResolver.Validate(second));
        Assert.NotEqual(first, second);
    }

    [Theory]
    [InlineData("", "空です")]
    [InlineData("   ", "空です")]
    [InlineData("これはとても長いシート名でありExcelの31文字制限を超えています", "31 文字まで")]
    [InlineData("売上/2026", "使えない文字")]
    [InlineData("売上[1]", "使えない文字")]
    [InlineData("'売上", "アポストロフィ")]
    [InlineData("History", "予約している")]
    public void Preview_InvalidOutputSheetName_IsBlocked(string name, string expectedFragment)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path, [Sheet("表", [["A"]])]);

        var preview = new SheetAggregationPlanner().CreatePreview([new(path, "表", name)]);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(expectedFragment));
    }

    [Fact]
    public void Preview_DuplicateOutputSheetName_IsBlocked()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        TestSheetWorkbookFactory.Create(a, [Sheet("売上", [["A"]])]);
        TestSheetWorkbookFactory.Create(b, [Sheet("在庫", [["B"]])]);

        var preview = new SheetAggregationPlanner().CreatePreview(
            [new(a, "売上", "まとめ"), new(b, "在庫", "まとめ")]);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("重複しています"));
    }

    [Fact]
    public void Preview_AllSelectedSheetsHidden_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            Sheet("表示", [["A"]]),
            new TestAggregationSheetSpec { Name = "非表示", IsHidden = true, Rows = [["B"]] },
        ]);

        var preview = CreatePreview((path, "非表示"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("すべて非表示"));
    }

    [Fact]
    public void Preview_NoSelection_IsBlocked()
    {
        var preview = new SheetAggregationPlanner().CreatePreview([]);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("選択されていません"));
    }

    [Fact]
    public void Output_PassesOpenXmlValidation()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        TestSheetWorkbookFactory.Create(a,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [[new Styled("見出し", 0), null], ["値", 1]],
                Merges = ["A1:B1"],
                FreezeTopRow = true,
                RowProperties = [new TestRowProperty(1, Height: 30)],
                ColumnProperties = [new TestColumnProperty(1, 2, Width: 18)],
            },
        ],
            styles: [new TestStyle { Bold = true, FillArgb = "FFDDEEFF", ThinBorder = true, HorizontalAlignment = "center" }]);
        TestSheetWorkbookFactory.Create(b, [Sheet("別表", [["X", 1]])],
            styles: [new TestStyle { NumberFormatCode = "#,##0" }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (a, "表"), (b, "別表")).Success);

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);

        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0, string.Join(" / ", errors.Take(5).Select(e => $"{e.Path?.XPath}: {e.Description}")));

        // Phase 0 の解析からも問題なく読める。
        var analysis = WorkbookAnalyzer.Analyze(output);
        Assert.Equal(AnalysisStatus.Succeeded, analysis.Status);
        Assert.Equal(2, analysis.Sheets.Count);
    }

    // --- helpers -------------------------------------------------------

    private static TestAggregationSheetSpec Sheet(string name, object?[][] rows)
        => new() { Name = name, Rows = rows };

    private static SheetAggregationPreview CreatePreview(params (string Path, string Sheet)[] selections)
        => new SheetAggregationPlanner().CreatePreview(
            [.. selections.Select(s => new SheetSelection(s.Path, s.Sheet))]);

    private static SheetAggregationResult Aggregate(string output, params (string Path, string Sheet)[] selections)
    {
        var preview = CreatePreview(selections);
        Assert.True(preview.CanExecute,
            string.Join(" / ", preview.Blocks.Select(issue => $"{issue.Location}: {issue.Message}")));
        return new SheetAggregator().Execute(preview, output);
    }

    internal static string[] ReadSheetNames(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        return [.. document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Select(sheet => sheet.Name!.Value!)];
    }

    private static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
    {
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>()
            .Single(s => s.Name?.Value == sheetName);
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    }

    private static IEnumerable<Cell> ReadCells(WorkbookPart workbookPart, string sheetName)
        => GetWorksheetPart(workbookPart, sheetName).Worksheet!.Descendants<Cell>();

    private static CellFormat FormatOf(Stylesheet stylesheet, Cell cell)
        => stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)(cell.StyleIndex?.Value ?? 0));

    private static Font FontOf(Stylesheet stylesheet, Cell cell)
        => stylesheet.Fonts!.Elements<Font>().ElementAt((int)(FormatOf(stylesheet, cell).FontId?.Value ?? 0));

    private static Fill FillOf(Stylesheet stylesheet, Cell cell)
        => stylesheet.Fills!.Elements<Fill>().ElementAt((int)(FormatOf(stylesheet, cell).FillId?.Value ?? 0));

    private static Border BorderOf(Stylesheet stylesheet, Cell cell)
        => stylesheet.Borders!.Elements<Border>().ElementAt((int)(FormatOf(stylesheet, cell).BorderId?.Value ?? 0));

    /// <summary>出力シートの指定行(1 始まり)を、Phase 1A の読み取り基盤で値として読む。</summary>
    internal static MergeCellValue[] ReadRow(string path, string sheetName, int rowIndex)
        => ReadRows(path, sheetName)[rowIndex - 1];

    internal static List<MergeCellValue[]> ReadRows(string path, string sheetName)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var context = WorkbookReadContext.Create(workbookPart);

        var rows = new List<MergeCellValue[]>();
        foreach (var row in GetWorksheetPart(workbookPart, sheetName).Worksheet!.Descendants<Row>())
        {
            var cells = row.Elements<Cell>().ToList();
            var width = cells.Count == 0
                ? 0
                : cells.Max(cell => CellRangeParser.TryParseCell(cell.CellReference?.Value ?? "A1", out var column, out _)
                    ? column
                    : 1);

            var values = new MergeCellValue[width];
            foreach (var cell in cells)
            {
                if (CellRangeParser.TryParseCell(cell.CellReference?.Value ?? string.Empty, out var column, out _))
                {
                    values[column - 1] = context.ReadCell(cell, out _);
                }
            }

            rows.Add(values);
        }

        return rows;
    }

    internal static List<string[]> ReadRowTexts(string path, string sheetName)
        => [.. ReadRows(path, sheetName).Select(row => row.Select(value => value.ToDisplayString()).ToArray())];
}
