using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Aggregation;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 1B.2A: 印刷設定・ページレイアウトの安全な移植。
/// 保持できるものは保持し、プリンター固有設定など安全に移せないものは Block する。
/// </summary>
public sealed class SheetAggregationPrintLayoutTests
{
    [Fact]
    public void Aggregate_PageMargins_KeepsEveryValue()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [new TestAggregationSheetSpec { Name = "表", Rows = [["A"]], AddPageMargins = true }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var margins = Assert.Single(Worksheet(output, "表").Descendants<PageMargins>());
        Assert.Equal(0.7D, margins.Left!.Value, 3);
        Assert.Equal(0.7D, margins.Right!.Value, 3);
        Assert.Equal(0.75D, margins.Top!.Value, 3);
        Assert.Equal(0.75D, margins.Bottom!.Value, 3);
        Assert.Equal(0.3D, margins.Header!.Value, 3);
        Assert.Equal(0.3D, margins.Footer!.Value, 3);
    }

    [Fact]
    public void Aggregate_PrintOptions_KeepsItsAttributes()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [new TestAggregationSheetSpec { Name = "表", Rows = [["A"]], AddPrintOptions = true }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var options = Assert.Single(Worksheet(output, "表").Descendants<PrintOptions>());
        Assert.True(options.HorizontalCentered!.Value);
    }

    [Fact]
    public void Aggregate_PageSetupWithoutPrinterSettings_KeepsOrientationScaleAndFit()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [new TestAggregationSheetSpec { Name = "表", Rows = [["A"]], AddPageSetup = true }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var setup = Assert.Single(Worksheet(output, "表").Descendants<PageSetup>());
        Assert.Equal(9U, setup.PaperSize!.Value);
        Assert.Equal(OrientationValues.Landscape, setup.Orientation!.Value);
        Assert.Equal(85U, setup.Scale!.Value);
        Assert.Equal(1U, setup.FitToWidth!.Value);
        Assert.Equal(0U, setup.FitToHeight!.Value);

        // プリンター設定パートへの参照は持ち込まない。
        Assert.Null(setup.Id?.Value);
    }

    [Fact]
    public void Preview_PageSetupWithPrinterSettings_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "月報",
                Rows = [["A"]],
                AddPageSetup = true,
                AddPrinterSettings = true,
            },
        ]);

        var preview = CreatePreview((path, "月報"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("プリンター固有の設定"));
    }

    [Fact]
    public void Aggregate_PageSetupProperties_KeepsFitToPage()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [new TestAggregationSheetSpec { Name = "表", Rows = [["A"]], AddPageSetupProperties = true }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var properties = Assert.Single(Worksheet(output, "表").Descendants<PageSetupProperties>());
        Assert.True(properties.FitToPage!.Value);
    }

    [Fact]
    public void Aggregate_TextHeaderFooter_IsPreserved()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [new TestAggregationSheetSpec { Name = "表", Rows = [["A"]], AddHeaderFooter = true }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var headerFooter = Assert.Single(Worksheet(output, "表").Descendants<HeaderFooter>());
        Assert.Equal("&L架空ヘッダー", headerFooter.OddHeader?.Text);
        Assert.Equal("&C架空フッター", headerFooter.OddFooter?.Text);
    }

    [Fact]
    public void Aggregate_DifferentOddEvenAndFirstHeaderFooter_ArePreserved()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["A"]],
                AddHeaderFooter = true,
                AddDistinctHeaderFooter = true,
            },
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var headerFooter = Assert.Single(Worksheet(output, "表").Descendants<HeaderFooter>());
        Assert.True(headerFooter.DifferentOddEven!.Value);
        Assert.True(headerFooter.DifferentFirst!.Value);
        Assert.False(headerFooter.ScaleWithDoc!.Value);
        Assert.False(headerFooter.AlignWithMargins!.Value);
        Assert.Equal("&R架空偶数ヘッダー", headerFooter.EvenHeader?.Text);
        Assert.Equal("&L架空偶数フッター", headerFooter.EvenFooter?.Text);
        Assert.Equal("&C架空先頭ヘッダー", headerFooter.FirstHeader?.Text);
        Assert.Equal("&R架空先頭フッター", headerFooter.FirstFooter?.Text);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void Preview_HeaderFooterWithImage_IsBlocked(bool useDrawing, bool useImageCode)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["A"]],
                AddHeaderFooter = useImageCode,
                AddHeaderFooterImageCode = useImageCode,
                AddHeaderFooterDrawing = useDrawing,
            },
        ]);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("ヘッダー・フッターに画像"));
    }

    [Fact]
    public void Aggregate_RowAndColumnBreaks_ArePreserved()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["A"]],
                AddRowBreaks = true,
                AddColumnBreaks = true,
            },
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var worksheet = Worksheet(output, "表");

        var rowBreaks = Assert.Single(worksheet.Descendants<RowBreaks>());
        Assert.Equal(new uint[] { 2U, 5U }, rowBreaks.Elements<Break>().Select(item => item.Id!.Value));
        Assert.Equal(2U, rowBreaks.Count!.Value);
        Assert.Equal(2U, rowBreaks.ManualBreakCount!.Value);

        var columnBreaks = Assert.Single(worksheet.Descendants<ColumnBreaks>());
        var columnBreak = Assert.Single(columnBreaks.Elements<Break>());
        Assert.Equal(2U, columnBreak.Id!.Value);
        Assert.True(columnBreak.ManualPageBreak!.Value);
    }

    [Fact]
    public void Preview_BrokenBreakDefinition_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["A"]],
                AddRowBreaks = true,
                AddBrokenBreak = true,
            },
        ]);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("改ページ位置の定義が壊れています"));
    }

    [Theory]
    [InlineData("'表'!$A$1:$F$100", "$A$1:$F$100")]
    [InlineData("表!$A$1:$F$100", "$A$1:$F$100")]
    [InlineData("'表'!$A$1:$C$10,'表'!$E$1:$G$10", "$A$1:$C$10,$E$1:$G$10")]
    public void Aggregate_PrintArea_IsKeptAndRewrittenForTheOutputSheet(string source, string expectedRanges)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [new TestAggregationSheetSpec { Name = "表", Rows = [["A"]], PrintArea = source }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var expected = string.Join(",", expectedRanges.Split(',').Select(range => $"'表'!{range}"));
        Assert.Equal(expected, DefinedNameText(output, "_xlnm.Print_Area", 0));
    }

    [Theory]
    [InlineData("'表'!$1:$2")]
    [InlineData("'表'!$A:$B")]
    [InlineData("'表'!$A:$B,'表'!$1:$2")]
    public void Aggregate_PrintTitles_AreKept(string source)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [new TestAggregationSheetSpec { Name = "表", Rows = [["A"]], PrintTitles = source }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        Assert.Equal(source, DefinedNameText(output, "_xlnm.Print_Titles", 0));
    }

    [Fact]
    public void Aggregate_RenamedOutputSheet_RewritesTheSheetNameInsideDefinedNames()
    {
        using var dir = new TempDir();
        var osaka = dir.File("大阪.xlsx");
        var kyoto = dir.File("京都.xlsx");
        TestSheetWorkbookFactory.Create(osaka,
            [new TestAggregationSheetSpec { Name = "売上", Rows = [["A"]], PrintArea = "'売上'!$A$1:$B$5" }]);
        TestSheetWorkbookFactory.Create(kyoto,
            [new TestAggregationSheetSpec { Name = "売上", Rows = [["B"]], PrintArea = "'売上'!$C$1:$D$9" }]);

        var preview = new SheetAggregationPlanner().CreatePreview(
        [
            new SheetSelection(osaka, "売上", "大阪_売上"),
            new SheetSelection(kyoto, "売上", "京都_売上"),
        ]);
        Assert.True(preview.CanExecute, string.Join(" / ", preview.Blocks.Select(issue => issue.Message)));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        Assert.Equal("'大阪_売上'!$A$1:$B$5", DefinedNameText(output, "_xlnm.Print_Area", 0));
        Assert.Equal("'京都_売上'!$C$1:$D$9", DefinedNameText(output, "_xlnm.Print_Area", 1));
    }

    [Theory]
    [InlineData("大阪 支店", "'大阪 支店'!$A$1:$B$2", "'大阪 支店'!$A$1:$B$2")]
    [InlineData("大阪'支店", "'大阪''支店'!$A$1:$B$2", "'大阪''支店'!$A$1:$B$2")]
    public void Aggregate_SheetNamesNeedingQuoting_AreEscapedCorrectly(
        string sheetName, string source, string expected)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [new TestAggregationSheetSpec { Name = sheetName, Rows = [["A"]], PrintArea = source }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, sheetName)).Success);

        Assert.Equal(expected, DefinedNameText(output, "_xlnm.Print_Area", 0));
    }

    [Fact]
    public void Aggregate_LocalSheetId_IsRemappedToTheOutputSheetIndex()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");

        // 元 Workbook では 6 枚目(index 5)のシートだけを選ぶ。
        var specs = new List<TestAggregationSheetSpec>();
        for (var i = 1; i <= 5; i++)
        {
            specs.Add(new TestAggregationSheetSpec { Name = $"その他{i}", Rows = [["x"]] });
        }

        specs.Add(new TestAggregationSheetSpec
        {
            Name = "対象",
            Rows = [["A"]],
            PrintArea = "'対象'!$A$1:$B$3",
        });

        TestSheetWorkbookFactory.Create(path, specs);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "対象")).Success);

        var definedName = Assert.Single(DefinedNames(output));
        Assert.Equal("_xlnm.Print_Area", definedName.Name?.Value);
        Assert.Equal(0U, definedName.LocalSheetId!.Value);
        Assert.Equal("'対象'!$A$1:$B$3", definedName.Text);
    }

    [Theory]
    [InlineData("'別シート'!$A$1:$B$2", "他のシート")]
    [InlineData("'[1]表'!$A$1:$B$2", "他のブック")]
    [InlineData("表:別シート!$A$1:$B$2", "複数シートにまたがって")]
    [InlineData("#REF!", "壊れています")]
    [InlineData("OFFSET(表!$A$1,0,0,5,5)", "解釈できません")]
    [InlineData("'表'!A1:B2", "対応していない参照形式")]
    public void Preview_UnsupportedPrintArea_IsBlocked(string reference, string expectedFragment)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec { Name = "表", Rows = [["A"]], PrintArea = reference },
            new TestAggregationSheetSpec { Name = "別シート", Rows = [["B"]] },
        ]);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(expectedFragment));
    }

    [Fact]
    public void Preview_OtherLocalDefinedName_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["A"]],
                LocalDefinedName = ("集計範囲", "'表'!$A$1:$B$2"),
            },
        ]);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("名前定義"));
    }

    [Fact]
    public void Preview_PrintSettingsOnlyOnAnUnselectedSheet_DoNotAffectTheSelectedSheet()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec { Name = "きれいな表", Rows = [["A"]] },
            new TestAggregationSheetSpec
            {
                Name = "印刷設定あり",
                Rows = [["B"]],
                AddPageMargins = true,
                AddPageSetup = true,
                AddPrinterSettings = true,
                PrintArea = "'印刷設定あり'!$A$1:$B$2",
            },
        ]);

        var preview = CreatePreview((path, "きれいな表"));

        Assert.True(preview.CanExecute, string.Join(" / ", preview.Blocks.Select(issue => issue.Message)));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);
        Assert.Empty(DefinedNames(output));
    }

    [Fact]
    public void Aggregate_TypicalWorkbookWithPrintSettings_IsSupportedAndValid()
    {
        using var dir = new TempDir();
        var a = dir.File("大阪.xlsx");
        var b = dir.File("京都.xlsx");

        // Excel が保存したときに自動で付く「pageMargins だけがある通常シート」。
        TestSheetWorkbookFactory.Create(a,
        [
            new TestAggregationSheetSpec { Name = "売上", Rows = [["商品", "金額"], ["架空A", 100]], AddPageMargins = true },
            new TestAggregationSheetSpec { Name = "在庫", Rows = [["商品", "数量"], ["架空B", 5]], AddPageMargins = true },
        ]);

        TestSheetWorkbookFactory.Create(b,
        [
            new TestAggregationSheetSpec
            {
                Name = "売上",
                Rows = [["商品", "金額"], ["架空C", 300]],
                AddPageMargins = true,
                AddPrintOptions = true,
                AddPageSetup = true,
                AddPageSetupProperties = true,
                AddHeaderFooter = true,
                AddRowBreaks = true,
                AddColumnBreaks = true,
                PrintArea = "'売上'!$A$1:$B$3",
                PrintTitles = "'売上'!$1:$1",
            },
        ]);

        var output = dir.File("out.xlsx");
        var result = Aggregate(output, (a, "売上"), (a, "在庫"), (b, "売上"));
        Assert.True(result.Success, result.Message);
        Assert.Equal(3, result.SheetCount);

        Assert.Equal("'売上 (2)'!$A$1:$B$3", DefinedNameText(output, "_xlnm.Print_Area", 2));
        Assert.Equal("'売上 (2)'!$1:$1", DefinedNameText(output, "_xlnm.Print_Titles", 2));

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join(" / ", errors.Take(5).Select(error => $"{error.Path?.XPath}: {error.Description}")));
    }

    [Fact]
    public void Aggregate_WithPrintSettings_DoesNotModifyInputWorkbooksAndKeepsVisibility()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表示",
                Rows = [["A"]],
                AddPageMargins = true,
                PrintArea = "'表示'!$A$1:$A$1",
            },
            new TestAggregationSheetSpec { Name = "非表示", Rows = [["B"]], IsHidden = true, AddPageMargins = true },
            new TestAggregationSheetSpec { Name = "非常に非表示", Rows = [["C"]], IsVeryHidden = true },
        ]);

        var before = Snapshot(path);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表示"), (path, "非表示"), (path, "非常に非表示")).Success);

        Assert.Equal(before, Snapshot(path));

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var sheets = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToList();
        Assert.Null(sheets[0].State);
        Assert.Equal(SheetStateValues.Hidden, sheets[1].State?.Value);
        Assert.Equal(SheetStateValues.VeryHidden, sheets[2].State?.Value);
    }

    // --- helpers -------------------------------------------------------

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

    private static Worksheet Worksheet(string path, string sheetName)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name?.Value == sheetName);
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet!;
    }

    private static List<DefinedName> DefinedNames(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        return document.WorkbookPart!.Workbook!.DefinedNames?.Elements<DefinedName>().ToList() ?? [];
    }

    private static string? DefinedNameText(string path, string name, uint localSheetId)
        => DefinedNames(path)
            .SingleOrDefault(definedName => definedName.Name?.Value == name
                && definedName.LocalSheetId?.Value == localSheetId)
            ?.Text;

    private static (string Hash, long Length, DateTime LastWriteUtc) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }
}
