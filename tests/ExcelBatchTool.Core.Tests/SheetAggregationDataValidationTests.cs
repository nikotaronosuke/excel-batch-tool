using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Aggregation;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 1B.2B2A: 標準の入力規則(x:dataValidation)のうち、
/// 意味を決定的に維持できるものだけを移植する。
/// </summary>
public sealed class SheetAggregationDataValidationTests
{
    [Fact]
    public void Aggregate_ListLiteral_KeepsTheFormulaTextExactly()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [SheetWith(new TestDataValidation("A1:A10", "list", Formula1: "\"赤,青,緑\"", ShowDropDown: true))]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var validation = SingleValidation(output, "表");
        Assert.Equal("list", validation.Type?.InnerText);
        Assert.Equal("\"赤,青,緑\"", validation.Formula1?.Text);
        Assert.Null(validation.Formula2);
        Assert.True(validation.ShowDropDown!.Value);
    }

    [Theory]
    [InlineData("$B$1:$B$10")]
    [InlineData("$A$1")]
    public void Aggregate_ListRangeOnTheSameSheet_IsKept(string formula)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [SheetWith(new TestDataValidation("A1:A10", "list", Formula1: formula))]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        Assert.Equal(formula, SingleValidation(output, "表").Formula1?.Text);
    }

    [Fact]
    public void Aggregate_MultipleTargetRanges_KeepsTheSqref()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [SheetWith(new TestDataValidation("A1:A10 C1:C10 E5", "list", Formula1: "\"はい,いいえ\""))]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        Assert.Equal("A1:A10 C1:C10 E5", SingleValidation(output, "表").SequenceOfReferences?.InnerText);
    }

    [Theory]
    [InlineData("whole", "equal", "10", null)]
    [InlineData("whole", "between", "1", "10")]
    [InlineData("decimal", "greaterThan", "1.5", null)]
    [InlineData("textLength", "lessThan", "20", null)]
    [InlineData("whole", "greaterThanOrEqual", "-5", null)]
    public void Aggregate_NumericRestrictions_AreKept(string type, string op, string formula1, string? formula2)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [SheetWith(new TestDataValidation("A1:A10", type, Operator: op, Formula1: formula1, Formula2: formula2))]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var validation = SingleValidation(output, "表");
        Assert.Equal(type, validation.Type?.InnerText);
        Assert.Equal(op, validation.Operator?.InnerText);
        Assert.Equal(formula1, validation.Formula1?.Text);
        Assert.Equal(formula2, validation.Formula2?.Text);
    }

    [Fact]
    public void Aggregate_DateRestrictionFrom1900Workbook_KeepsTheSerialAsIs()
    {
        using var dir = new TempDir();
        var path = dir.File("1900.xlsx");
        var serial = MergeCellValue.DateTimeToSerial(new DateTime(2026, 3, 15));
        TestSheetWorkbookFactory.Create(path,
        [
            SheetWith(new TestDataValidation(
                "A1:A10", "date", Operator: "greaterThan",
                Formula1: serial.ToString(CultureInfo.InvariantCulture))),
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var formula = SingleValidation(output, "表").Formula1!.Text;
        Assert.Equal(new DateTime(2026, 3, 15), MergeCellValue.SerialToDateTime(ParseSerial(formula)));
    }

    [Fact]
    public void Aggregate_DateRestrictionFrom1904Workbook_MeansTheSameDateInTheOutput()
    {
        using var dir = new TempDir();
        var date = new DateTime(2026, 3, 15);
        var serial1900 = MergeCellValue.DateTimeToSerial(date);
        var serial1904 = serial1900 - 1462;

        var from1900 = dir.File("1900.xlsx");
        var from1904 = dir.File("1904.xlsx");
        TestSheetWorkbookFactory.Create(from1900,
        [
            SheetWith(new TestDataValidation(
                "A1:A10", "date", Operator: "greaterThan",
                Formula1: serial1900.ToString(CultureInfo.InvariantCulture)), name: "表1900"),
        ]);
        TestSheetWorkbookFactory.Create(from1904,
        [
            SheetWith(new TestDataValidation(
                "A1:A10", "date", Operator: "greaterThan",
                Formula1: serial1904.ToString(CultureInfo.InvariantCulture)), name: "表1904"),
        ],
            date1904: true);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (from1900, "表1900"), (from1904, "表1904")).Success);

        // どちらのシートでも同じ日付条件を意味する。
        foreach (var sheetName in new[] { "表1900", "表1904" })
        {
            var formula = SingleValidation(output, sheetName).Formula1!.Text;
            Assert.Equal(date, MergeCellValue.SerialToDateTime(ParseSerial(formula)));
        }
    }

    [Fact]
    public void Aggregate_TimeRestriction_IsNotShiftedByTheDateSystem()
    {
        using var dir = new TempDir();
        var path = dir.File("1904.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [SheetWith(new TestDataValidation("A1:A10", "time", Operator: "greaterThan", Formula1: "0.5"))],
            date1904: true);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        Assert.Equal("0.5", SingleValidation(output, "表").Formula1?.Text);
    }

    [Fact]
    public void Aggregate_UserMessagesAndOptions_ArePreserved()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            SheetWith(new TestDataValidation(
                "A1:A10", "list",
                Formula1: "\"はい,いいえ\"",
                AllowBlank: true,
                ShowDropDown: true,
                ShowInputMessage: true,
                ShowErrorMessage: true,
                ErrorStyle: "warning",
                ImeMode: "hiragana",
                PromptTitle: "入力のヒント",
                Prompt: "一覧から選んでください",
                ErrorTitle: "入力エラー",
                Error: "一覧にない値です")),
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var validation = SingleValidation(output, "表");
        Assert.True(validation.AllowBlank!.Value);
        Assert.True(validation.ShowDropDown!.Value);
        Assert.True(validation.ShowInputMessage!.Value);
        Assert.True(validation.ShowErrorMessage!.Value);
        Assert.Equal("warning", validation.ErrorStyle?.InnerText);
        Assert.Equal("hiragana", validation.ImeMode?.InnerText);
        Assert.Equal("入力のヒント", validation.PromptTitle?.Value);
        Assert.Equal("一覧から選んでください", validation.Prompt?.Value);
        Assert.Equal("入力エラー", validation.ErrorTitle?.Value);
        Assert.Equal("一覧にない値です", validation.Error?.Value);
    }

    [Fact]
    public void Aggregate_ContainerAttributes_ArePreservedAndCountIsRecalculated()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["A"]],
                AddDataValidationContainerAttributes = true,
                DataValidations =
                [
                    new TestDataValidation("A1:A10", "list", Formula1: "\"はい,いいえ\""),
                    new TestDataValidation("B1:B10", "whole", Operator: "between", Formula1: "1", Formula2: "10"),
                ],
            },
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var container = Assert.Single(Worksheet(output, "表").Descendants<DataValidations>());
        Assert.Equal(2, container.Elements<DataValidation>().Count());

        // 元ファイルの count は 99 にしてある。出力では実件数へ振り直す。
        Assert.Equal(2U, container.Count!.Value);
        Assert.True(container.DisablePrompts!.Value);
        Assert.Equal(100U, container.XWindow!.Value);
        Assert.Equal(200U, container.YWindow!.Value);
    }

    public static TheoryData<string, string?, string?, string?, string> UnsupportedValidations => new()
    {
        { "list", null, "商品一覧", null, "名前定義" },
        { "list", null, "INDIRECT($E$4)", null, "関数や数式" },
        { "list", null, "Sheet2!$A$1:$A$5", null, "他のシート" },
        { "list", null, "'[1]別ブック'!$A$1:$A$5", null, "他のブック" },
        { "list", null, "#REF!", null, "壊れています" },
        { "custom", null, "=A1>0", null, "ユーザー設定の数式" },
        { "whole", "between", "1", null, "2 つ目の条件値がありません" },
        { "list", null, "\"はい,いいえ\"", "1", "構造が想定と異なります" },
        { "list", "between", "\"はい,いいえ\"", null, "意味を持たない条件設定" },
        { "whole", "equal", "$A$1", null, "名前定義" },
        { "decimal", "equal", "SUM(A1:A2)", null, "関数や数式" },
    };

    [Theory]
    [MemberData(nameof(UnsupportedValidations))]
    public void Preview_UnsupportedValidation_IsBlocked(
        string type, string? op, string? formula1, string? formula2, string expectedFragment)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [SheetWith(new TestDataValidation("A1:A10", type, Operator: op, Formula1: formula1, Formula2: formula2))]);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(expectedFragment));
    }

    [Fact]
    public void Preview_UnknownValidationType_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [SheetWith(new TestDataValidation("A1:A10", "fictionalType", Formula1: "1"))]);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("種類"));
    }

    [Theory]
    [InlineData("ZZZZ1:ZZZZ9", "適用範囲")]
    [InlineData("#REF!", "適用範囲")]
    [InlineData("A1:A10 #REF!", "適用範囲")]
    public void Preview_InvalidSqref_IsBlocked(string sqref, string expectedFragment)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [SheetWith(new TestDataValidation(sqref, "list", Formula1: "\"はい,いいえ\""))]);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(expectedFragment));
    }

    [Fact]
    public void Aggregate_SeveralSafeValidations_AreAllKept()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["A"]],
                DataValidations =
                [
                    new TestDataValidation("A1:A10", "list", Formula1: "\"はい,いいえ\""),
                    new TestDataValidation("B1:B10", "whole", Operator: "between", Formula1: "1", Formula2: "10"),
                    new TestDataValidation("C1:C10", "textLength", Operator: "lessThan", Formula1: "20"),
                ],
            },
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var validations = Worksheet(output, "表").Descendants<DataValidation>().ToList();
        Assert.Equal(3, validations.Count);
        Assert.Equal(
            new[] { "A1:A10", "B1:B10", "C1:C10" },
            validations.Select(item => item.SequenceOfReferences?.InnerText));
    }

    [Fact]
    public void Preview_OneUnsupportedValidation_BlocksTheWholeSheet()
    {
        using var dir = new TempDir();
        var path = dir.File("入力.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "商品",
                Rows = [["A"]],
                DataValidations =
                [
                    new TestDataValidation("A1:A10", "list", Formula1: "\"はい,いいえ\""),
                    new TestDataValidation("B1:B10", "whole", Operator: "between", Formula1: "1", Formula2: "10"),
                    new TestDataValidation("D2:D100", "list", Formula1: "商品一覧"),
                ],
            },
        ]);

        var preview = CreatePreview((path, "商品"));

        Assert.False(preview.CanExecute);
        var block = Assert.Single(preview.Blocks, issue => issue.Message.Contains("D2:D100"));
        Assert.Contains("商品一覧", block.Message);
    }

    [Fact]
    public void Preview_UnknownExtendedAttribute_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["A"]],
                AddUnknownDataValidationAttribute = true,
                DataValidations = [new TestDataValidation("A1:A10", "list", Formula1: "\"はい,いいえ\"")],
            },
        ]);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("対応していない設定"));
    }

    /// <summary>
    /// Phase 1B.2B2B で x14 リスト入力規則に対応したが、参照先シート(Sheet2)が
    /// 集約対象に無いため引き続き Block される(理由は「参照先が対象外」に変わる)。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Preview_X14DataValidationPointingOutsideTheSelection_IsBlocked(bool alsoStandard)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "入力欄",
                Rows = [["A"]],
                AddX14DataValidation = true,
                DataValidations = alsoStandard
                    ? [new TestDataValidation("A1:A10", "list", Formula1: "\"はい,いいえ\"")]
                    : [],
            },
        ]);

        var preview = CreatePreview((path, "入力欄"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("集約対象に含まれていない"));
    }

    [Fact]
    public void Analyze_X14DataValidation_IsReportedByThePhase0Analyzer()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [new TestAggregationSheetSpec { Name = "入力欄", Rows = [["A"]], AddX14DataValidation = true }]);

        var analysis = WorkbookAnalyzer.Analyze(path);

        Assert.Equal(AnalysisStatus.Succeeded, analysis.Status);
        Assert.Contains(analysis.Findings, finding => finding.Type == FindingType.DataValidation);
    }

    [Fact]
    public void Aggregate_ValidationWithHyperlinkAndPrintSettings_KeepsThemAllAndStaysValid()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "売上",
                Rows = [["商品", "金額"], ["架空A", 100]],
                AddPageMargins = true,
                PrintArea = "'売上'!$A$1:$B$2",
                Hyperlinks = [new TestHyperlink("A1", ExternalTarget: "https://example.invalid/")],
                DataValidations = [new TestDataValidation("B2:B100", "whole", Operator: "between", Formula1: "0", Formula2: "1000")],
            },
            new TestAggregationSheetSpec { Name = "在庫", Rows = [["商品"]], IsHidden = true },
            new TestAggregationSheetSpec { Name = "メモ", Rows = [["補足"]], IsVeryHidden = true },
        ]);

        var before = Snapshot(path);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "売上"), (path, "在庫"), (path, "メモ")).Success);

        Assert.Equal(before, Snapshot(path));

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;

        var sheets = workbookPart.Workbook!.Sheets!.Elements<Sheet>().ToList();
        Assert.Null(sheets[0].State);
        Assert.Equal(SheetStateValues.Hidden, sheets[1].State?.Value);
        Assert.Equal(SheetStateValues.VeryHidden, sheets[2].State?.Value);

        Assert.Contains(
            workbookPart.Workbook.DefinedNames!.Elements<DefinedName>(),
            name => name.Text == "'売上'!$A$1:$B$2");

        var worksheet = Worksheet(output, "売上");
        Assert.Single(worksheet.Descendants<PageMargins>());
        Assert.Single(worksheet.Descendants<Hyperlink>());
        Assert.Single(worksheet.Descendants<DataValidation>());

        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join(" / ", errors.Take(5).Select(error => $"{error.Path?.XPath}: {error.Description}")));
    }

    [Fact]
    public void Preview_UnsupportedValidationOnlyOnAnUnselectedSheet_DoesNotBlockTheSelectedSheet()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec { Name = "きれいな表", Rows = [["A"]] },
            SheetWith(new TestDataValidation("A1:A10", "list", Formula1: "商品一覧"), name: "問題あり"),
        ]);

        var preview = CreatePreview((path, "きれいな表"));

        Assert.True(preview.CanExecute, string.Join(" / ", preview.Blocks.Select(issue => issue.Message)));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);
    }

    // --- helpers -------------------------------------------------------

    private static TestAggregationSheetSpec SheetWith(TestDataValidation validation, string name = "表") => new()
    {
        Name = name,
        Rows = [["A"]],
        DataValidations = [validation],
    };

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

    private static DataValidation SingleValidation(string path, string sheetName)
        => Assert.Single(Worksheet(path, sheetName).Descendants<DataValidation>());

    private static double ParseSerial(string? text)
        => double.Parse(text!, NumberStyles.Float, CultureInfo.InvariantCulture);

    private static (string Hash, long Length, DateTime LastWriteUtc) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }
}
