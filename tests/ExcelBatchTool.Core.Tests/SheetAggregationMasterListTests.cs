using System.IO.Compression;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Aggregation;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 1B.2B2B: 別シートの候補一覧を参照するプルダウン(x14 リスト入力規則と
/// ブック全体の名前定義)の安全な移植。
/// </summary>
public sealed class SheetAggregationMasterListTests
{
    private const string X14Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private const string XmNamespace = "http://schemas.microsoft.com/office/excel/2006/main";
    private const string RevisionNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";

    // --- characterization ---------------------------------------------

    /// <summary>
    /// 出力される x14 入力規則の XML 構造(URI・名前空間・要素名)を固定する。
    /// 推測ではなく実際の出力を確認したうえで実装を進めるための土台。
    /// </summary>
    [Fact]
    public void Characterize_OutputX14Structure()
    {
        using var dir = new TempDir();
        var path = MasterWorkbook(dir, "注文.xlsx", "商品マスタ!$A$2:$A$50");

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "注文"), (path, "商品マスタ")).Success);

        using var zip = ZipFile.OpenRead(output);
        using var reader = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open());
        var xml = reader.ReadToEnd();

        Assert.Contains($"uri=\"{X14DataValidationScanner.ExtensionUri}\"", xml, StringComparison.Ordinal);
        Assert.Contains(X14Namespace, xml, StringComparison.Ordinal);
        Assert.Contains(XmNamespace, xml, StringComparison.Ordinal);
        Assert.Contains(":dataValidations", xml, StringComparison.Ordinal);
        Assert.Contains(":formula1", xml, StringComparison.Ordinal);
        Assert.Contains(":sqref", xml, StringComparison.Ordinal);

        // リビジョン識別子は持ち込まない(SDK の検証器が未宣言属性として扱うため)。
        Assert.DoesNotContain(RevisionNamespace, xml, StringComparison.Ordinal);
    }

    /// <summary>
    /// 名前定義の参照文字列に先頭の "=" が付かず、シート名が引用されることを固定する。
    /// </summary>
    [Fact]
    public void Characterize_OutputDefinedNameText()
    {
        using var dir = new TempDir();
        var path = NamedRangeWorkbook(dir, "注文.xlsx", "商品一覧", "'商品マスタ'!$A$2:$A$50");

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "注文"), (path, "商品マスタ")).Success);

        var definedName = Assert.Single(DefinedNames(output).Where(name => name.LocalSheetId is null));
        Assert.Equal("商品一覧", definedName.Name?.Value);
        Assert.Equal("'商品マスタ'!$A$2:$A$50", definedName.Text);
        Assert.DoesNotContain('=', definedName.Text);
    }

    // --- x14 direct cross-sheet reference ------------------------------

    [Fact]
    public void Aggregate_X14ListPointingAtASelectedMasterSheet_IsKept()
    {
        using var dir = new TempDir();
        var path = MasterWorkbook(dir, "注文.xlsx", "商品マスタ!$A$2:$A$50");

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "注文"), (path, "商品マスタ")).Success);

        Assert.Equal("'商品マスタ'!$A$2:$A$50", SingleX14ListSource(output, "注文"));
    }

    [Fact]
    public void Aggregate_MasterSheetRenamedInOutput_RewritesTheListSource()
    {
        using var dir = new TempDir();
        var path = MasterWorkbook(dir, "注文.xlsx", "商品マスタ!$A$2:$A$50");

        var preview = new SheetAggregationPlanner().CreatePreview(
        [
            new SheetSelection(path, "注文", "大阪_注文"),
            new SheetSelection(path, "商品マスタ", "大阪_商品マスタ"),
        ]);
        Assert.True(preview.CanExecute, Reasons(preview));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        Assert.Equal("'大阪_商品マスタ'!$A$2:$A$50", SingleX14ListSource(output, "大阪_注文"));
    }

    [Theory]
    [InlineData("商品 マスタ", "'商品 マスタ'!$A$2:$A$50", "'商品 マスタ'!$A$2:$A$50")]
    [InlineData("大阪'商品", "'大阪''商品'!$A$2:$A$50", "'大阪''商品'!$A$2:$A$50")]
    public void Aggregate_MasterSheetNameNeedingQuoting_IsEscaped(
        string masterName, string source, string expected)
    {
        using var dir = new TempDir();
        var path = dir.File("注文.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            OrderSheet(x14: new TestX14Validation("B2:B100", source)),
            new TestAggregationSheetSpec { Name = masterName, Rows = [["架空商品"]] },
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "注文"), (path, masterName)).Success);

        Assert.Equal(expected, SingleX14ListSource(output, "注文"));
    }

    [Fact]
    public void Aggregate_X14ListPointingAtItsOwnSheet_FollowsTheOutputName()
    {
        using var dir = new TempDir();
        var path = dir.File("注文.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [OrderSheet(x14: new TestX14Validation("B2:B100", "注文!$D$1:$D$5"))]);

        var preview = new SheetAggregationPlanner().CreatePreview(
            [new SheetSelection(path, "注文", "大阪_注文")]);
        Assert.True(preview.CanExecute, Reasons(preview));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        Assert.Equal("'大阪_注文'!$D$1:$D$5", SingleX14ListSource(output, "大阪_注文"));
    }

    [Theory]
    [InlineData("商品マスタ!$A$2:$A$50")]
    [InlineData("商品マスタ!$A$2:$Z$2")]
    public void Aggregate_SingleColumnOrRowMasterRange_IsKept(string source)
    {
        using var dir = new TempDir();
        var path = MasterWorkbook(dir, "注文.xlsx", source);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "注文"), (path, "商品マスタ")).Success);

        Assert.Equal($"'商品マスタ'!{source.Split('!')[1]}", SingleX14ListSource(output, "注文"));
    }

    [Fact]
    public void Preview_MasterSheetNotSelected_IsBlocked()
    {
        using var dir = new TempDir();
        var path = MasterWorkbook(dir, "注文.xlsx", "商品マスタ!$A$2:$A$50");

        var preview = CreatePreview((path, "注文"));

        Assert.False(preview.CanExecute);
        var block = Assert.Single(preview.Blocks, issue => issue.Message.Contains("商品マスタ"));
        Assert.Contains("集約対象に含まれていない", block.Message);

        // マスタも選べば実行できるようになる。
        var retry = CreatePreview((path, "注文"), (path, "商品マスタ"));
        Assert.True(retry.CanExecute, Reasons(retry));
    }

    public static TheoryData<string, string> UnsupportedX14Sources => new()
    {
        { "商品マスタ!$A$1:$C$10", "縦横に広がる範囲" },
        { "商品マスタ!A2:A50", "安全に集約できません" },
        { "'[1]別ブック'!$A$1:$A$5", "他のブック" },
        { "商品マスタ!#REF!", "壊れています" },
        { "商品マスタ:別表!$A$1:$A$5", "複数シート" },
        { "INDIRECT($E$4)", "安全に集約できません" },
        { "OFFSET($A$1,0,0,5,1)", "安全に集約できません" },
    };

    [Theory]
    [MemberData(nameof(UnsupportedX14Sources))]
    public void Preview_UnsupportedX14ListSource_IsBlocked(string source, string expectedFragment)
    {
        using var dir = new TempDir();
        var path = MasterWorkbook(dir, "注文.xlsx", source);

        var preview = CreatePreview((path, "注文"), (path, "商品マスタ"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(expectedFragment));
    }

    [Theory]
    [InlineData("custom")]
    [InlineData("whole")]
    [InlineData("date")]
    public void Preview_X14ValidationOfAnotherType_IsBlocked(string type)
    {
        using var dir = new TempDir();
        var path = dir.File("注文.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [OrderSheet(x14: new TestX14Validation("B2:B100", "1", Type: type))]);

        var preview = CreatePreview((path, "注文"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("新しい形式の入力規則"));
    }

    [Fact]
    public void Preview_SafeAndUnsupportedX14Mixed_BlocksTheWholeSheet()
    {
        using var dir = new TempDir();
        var path = dir.File("注文.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "注文",
                Rows = [["注文"]],
                X14Validations =
                [
                    new TestX14Validation("B2:B10", "商品マスタ!$A$2:$A$50"),
                    new TestX14Validation("C2:C10", "商品マスタ!$B$2:$B$50"),
                    new TestX14Validation("D2:D10", "1", Type: "custom"),
                ],
            },
            new TestAggregationSheetSpec { Name = "商品マスタ", Rows = [["架空商品"]] },
        ]);

        var preview = CreatePreview((path, "注文"), (path, "商品マスタ"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("新しい形式の入力規則"));
    }

    [Theory]
    [InlineData(true, false, false, "構造が想定と異なります")]
    [InlineData(false, true, false, "対応していない内容")]
    [InlineData(false, false, true, "対応していない設定")]
    public void Preview_MalformedX14Validation_IsBlocked(
        bool withFormula2, bool unknownChild, bool unknownAttribute, string expectedFragment)
    {
        using var dir = new TempDir();
        var path = MasterWorkbook(dir, "注文.xlsx", "商品マスタ!$A$2:$A$50",
            configure: spec => spec with
            {
                Formula2 = withFormula2 ? "商品マスタ!$B$2:$B$5" : null,
                AddUnknownChild = unknownChild,
                AddUnknownAttribute = unknownAttribute,
            });

        var preview = CreatePreview((path, "注文"), (path, "商品マスタ"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(expectedFragment));
    }

    [Theory]
    [InlineData("#REF!")]
    [InlineData("ZZZZ1:ZZZZ9")]
    public void Preview_InvalidX14Sqref_IsBlocked(string sqref)
    {
        using var dir = new TempDir();
        var path = dir.File("注文.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            OrderSheet(x14: new TestX14Validation(sqref, "商品マスタ!$A$2:$A$50")),
            new TestAggregationSheetSpec { Name = "商品マスタ", Rows = [["架空商品"]] },
        ]);

        var preview = CreatePreview((path, "注文"), (path, "商品マスタ"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("適用範囲"));
    }

    [Fact]
    public void Preview_UnsupportedExtensionAlongsideDataValidation_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("注文.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "注文",
                Rows = [["注文"]],
                AddUnsupportedExtension = true,
                X14Validations = [new TestX14Validation("B2:B10", "商品マスタ!$A$2:$A$50")],
            },
            new TestAggregationSheetSpec { Name = "商品マスタ", Rows = [["架空商品"]] },
        ]);

        var preview = CreatePreview((path, "注文"), (path, "商品マスタ"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("対応していない拡張情報"));
    }

    [Fact]
    public void Aggregate_SourceRevisionIdsAreNotCarriedOverAndDoNotCollide()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        const string sameUid = "{11111111-2222-3333-4444-555555555555}";

        foreach (var path in new[] { a, b })
        {
            TestSheetWorkbookFactory.Create(path,
            [
                OrderSheet(x14: new TestX14Validation("B2:B10", "商品マスタ!$A$2:$A$50", RevisionUid: sameUid)),
                new TestAggregationSheetSpec { Name = "商品マスタ", Rows = [["架空商品"]] },
            ]);
        }

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output,
            (a, "注文"), (a, "商品マスタ"), (b, "注文"), (b, "商品マスタ")).Success);

        using var zip = ZipFile.OpenRead(output);
        foreach (var entry in zip.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal)))
        {
            using var reader = new StreamReader(entry.Open());
            Assert.DoesNotContain(RevisionNamespace, reader.ReadToEnd(), StringComparison.Ordinal);
        }
    }

    // --- workbook-level defined names ----------------------------------

    [Fact]
    public void Aggregate_StandardListUsingAWorkbookName_KeepsBothTheRuleAndTheName()
    {
        using var dir = new TempDir();
        var path = NamedRangeWorkbook(dir, "注文.xlsx", "商品一覧", "'商品マスタ'!$A$2:$A$50");

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "注文"), (path, "商品マスタ")).Success);

        var validation = Assert.Single(Worksheet(output, "注文").Descendants<DataValidation>());
        Assert.Equal("商品一覧", validation.Formula1?.Text);

        var definedName = Assert.Single(DefinedNames(output).Where(name => name.LocalSheetId is null));
        Assert.Equal("'商品マスタ'!$A$2:$A$50", definedName.Text);
    }

    [Fact]
    public void Aggregate_X14ListUsingAWorkbookName_KeepsBothTheRuleAndTheName()
    {
        using var dir = new TempDir();
        var path = dir.File("注文.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            OrderSheet(x14: new TestX14Validation("B2:B100", "商品一覧")),
            new TestAggregationSheetSpec { Name = "商品マスタ", Rows = [["架空商品"]] },
        ],
            definedNames: [new TestDefinedName("商品一覧", "'商品マスタ'!$A$2:$A$50")]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "注文"), (path, "商品マスタ")).Success);

        Assert.Equal("商品一覧", SingleX14ListSource(output, "注文"));
        Assert.Equal("'商品マスタ'!$A$2:$A$50",
            Assert.Single(DefinedNames(output).Where(name => name.LocalSheetId is null)).Text);
    }

    [Fact]
    public void Aggregate_NameTargetSheetRenamed_RewritesTheRefersTo()
    {
        using var dir = new TempDir();
        var path = NamedRangeWorkbook(dir, "注文.xlsx", "商品一覧", "'商品マスタ'!$A$2:$A$50");

        var preview = new SheetAggregationPlanner().CreatePreview(
        [
            new SheetSelection(path, "注文", "大阪_注文"),
            new SheetSelection(path, "商品マスタ", "大阪_商品マスタ"),
        ]);
        Assert.True(preview.CanExecute, Reasons(preview));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        Assert.Equal("'大阪_商品マスタ'!$A$2:$A$50",
            Assert.Single(DefinedNames(output).Where(name => name.LocalSheetId is null)).Text);
    }

    [Fact]
    public void Preview_NameTargetSheetNotSelected_IsBlocked()
    {
        using var dir = new TempDir();
        var path = NamedRangeWorkbook(dir, "注文.xlsx", "商品一覧", "'商品マスタ'!$A$2:$A$50");

        var preview = CreatePreview((path, "注文"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("集約対象に含まれていない"));
    }

    public static TheoryData<string, string> UnsupportedNameTargets => new()
    {
        { "INDIRECT($E$4)", "安全に集約できません" },
        { "OFFSET('商品マスタ'!$A$1,0,0,5,1)", "安全に集約できません" },
        { "'商品マスタ'!A2:A50", "安全に集約できません" },
        { "'[1]別ブック'!$A$1:$A$5", "他のブック" },
        { "#REF!", "壊れています" },
        { "'商品マスタ'!$A$1:$C$10", "縦横に広がる範囲" },
        { "$A$1:$A$5", "シート名がありません" },
    };

    [Theory]
    [MemberData(nameof(UnsupportedNameTargets))]
    public void Preview_UnsupportedNameTarget_IsBlocked(string refersTo, string expectedFragment)
    {
        using var dir = new TempDir();
        var path = NamedRangeWorkbook(dir, "注文.xlsx", "商品一覧", refersTo);

        var preview = CreatePreview((path, "注文"), (path, "商品マスタ"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(expectedFragment));
    }

    [Fact]
    public void Preview_SheetScopedName_IsBlocked()
    {
        using var dir = new TempDir();
        var path = NamedRangeWorkbook(dir, "注文.xlsx", "商品一覧", "'商品マスタ'!$A$2:$A$50",
            localSheetId: 0U);

        var preview = CreatePreview((path, "注文"), (path, "商品マスタ"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("シート固有の名前"));
    }

    [Theory]
    [InlineData(true, null)]
    [InlineData(false, "架空の説明")]
    public void Preview_NameWithUnsupportedMetadata_IsBlocked(bool hidden, string? comment)
    {
        using var dir = new TempDir();
        var path = NamedRangeWorkbook(dir, "注文.xlsx", "商品一覧", "'商品マスタ'!$A$2:$A$50",
            hidden: hidden, comment: comment);

        var preview = CreatePreview((path, "注文"), (path, "商品マスタ"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("扱えない設定"));
    }

    [Fact]
    public void Aggregate_SameNameWithTheSameTargetInTwoWorkbooks_ProducesOneDefinition()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        foreach (var path in new[] { a, b })
        {
            NamedRangeWorkbook(dir, Path.GetFileName(path), "商品一覧", "'商品マスタ'!$A$2:$A$50");
        }

        // 出力名が同じになるよう、2 つ目のマスタは選ばない構成にはできないので、
        // 同じ出力名になる 1 ブック分だけを使って重複定義が起きないことを見る。
        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (a, "注文"), (a, "商品マスタ")).Success);

        Assert.Single(DefinedNames(output).Where(name => name.LocalSheetId is null));
    }

    [Fact]
    public void Preview_SameNamePointingAtDifferentTargets_IsBlocked()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        NamedRangeWorkbook(dir, "A.xlsx", "商品一覧", "'商品マスタ'!$A$2:$A$50");
        NamedRangeWorkbook(dir, "B.xlsx", "商品一覧", "'商品マスタ'!$B$2:$B$50");

        var preview = CreatePreview((a, "注文"), (a, "商品マスタ"), (b, "注文"), (b, "商品マスタ"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("別々の内容"));
    }

    // --- regression ----------------------------------------------------

    [Fact]
    public void Aggregate_MasterListWithEverythingElse_StaysValidAndLeavesInputsUntouched()
    {
        using var dir = new TempDir();
        var path = dir.File("注文.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "注文",
                Rows = [["商品", "数量"], ["架空A", 1]],
                AddPageMargins = true,
                PrintArea = "'注文'!$A$1:$B$2",
                Hyperlinks = [new TestHyperlink("A1", ExternalTarget: "https://example.invalid/")],
                DataValidations =
                [
                    new TestDataValidation("C2:C100", "list", Formula1: "商品一覧"),
                    new TestDataValidation("D2:D100", "whole", Operator: "between", Formula1: "1", Formula2: "10"),
                ],
                X14Validations = [new TestX14Validation("B2:B100", "商品マスタ!$A$2:$A$50")],
            },
            new TestAggregationSheetSpec { Name = "商品マスタ", Rows = [["架空商品"]], AddPageMargins = true },
            new TestAggregationSheetSpec { Name = "メモ", Rows = [["補足"]], IsVeryHidden = true },
        ],
            definedNames: [new TestDefinedName("商品一覧", "'商品マスタ'!$A$2:$A$50")]);

        var before = Snapshot(path);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "注文"), (path, "商品マスタ"), (path, "メモ")).Success);

        Assert.Equal(before, Snapshot(path));

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;

        var sheets = workbookPart.Workbook!.Sheets!.Elements<Sheet>().ToList();
        Assert.Equal(SheetStateValues.VeryHidden, sheets[2].State?.Value);

        var names = workbookPart.Workbook.DefinedNames!.Elements<DefinedName>().ToList();
        Assert.Contains(names, name => name.Text == "'注文'!$A$1:$B$2" && name.LocalSheetId is not null);
        Assert.Contains(names, name => name.Name?.Value == "商品一覧" && name.LocalSheetId is null);

        var worksheet = Worksheet(output, "注文");
        Assert.Single(worksheet.Descendants<PageMargins>());
        Assert.Single(worksheet.Descendants<Hyperlink>());
        Assert.Equal(2, worksheet.Descendants<DataValidation>().Count());
        Assert.Equal("'商品マスタ'!$A$2:$A$50", SingleX14ListSource(output, "注文"));

        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join(" / ", errors.Take(5).Select(error => $"{error.Path?.XPath}: {error.Description}")));
    }

    [Fact]
    public void Preview_UnsupportedMasterListOnlyOnAnUnselectedSheet_DoesNotBlockTheSelectedSheet()
    {
        using var dir = new TempDir();
        var path = dir.File("注文.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec { Name = "きれいな表", Rows = [["A"]] },
            OrderSheet(x14: new TestX14Validation("B2:B10", "INDIRECT($E$4)")),
        ]);

        var preview = CreatePreview((path, "きれいな表"));

        Assert.True(preview.CanExecute, Reasons(preview));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);
    }

    // --- helpers -------------------------------------------------------

    private static TestAggregationSheetSpec OrderSheet(TestX14Validation? x14 = null) => new()
    {
        Name = "注文",
        Rows = [["注文"]],
        X14Validations = x14 is null ? [] : [x14],
    };

    private static string MasterWorkbook(
        TempDir dir,
        string fileName,
        string listSource,
        Func<TestX14Validation, TestX14Validation>? configure = null)
    {
        var path = dir.File(fileName);
        var validation = new TestX14Validation("B2:B100", listSource);
        if (configure is not null)
        {
            validation = configure(validation);
        }

        TestSheetWorkbookFactory.Create(path,
        [
            OrderSheet(validation),
            new TestAggregationSheetSpec { Name = "商品マスタ", Rows = [["架空商品"]] },
        ]);

        return path;
    }

    private static string NamedRangeWorkbook(
        TempDir dir,
        string fileName,
        string name,
        string refersTo,
        uint? localSheetId = null,
        bool hidden = false,
        string? comment = null)
    {
        var path = dir.File(fileName);
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "注文",
                Rows = [["注文"]],
                DataValidations = [new TestDataValidation("C2:C100", "list", Formula1: name)],
            },
            new TestAggregationSheetSpec { Name = "商品マスタ", Rows = [["架空商品"]] },
        ],
            definedNames: [new TestDefinedName(name, refersTo, localSheetId, hidden, comment)]);

        return path;
    }

    private static SheetAggregationPreview CreatePreview(params (string Path, string Sheet)[] selections)
        => new SheetAggregationPlanner().CreatePreview(
            [.. selections.Select(s => new SheetSelection(s.Path, s.Sheet))]);

    private static SheetAggregationResult Aggregate(string output, params (string Path, string Sheet)[] selections)
    {
        var preview = CreatePreview(selections);
        Assert.True(preview.CanExecute, Reasons(preview));
        return new SheetAggregator().Execute(preview, output);
    }

    private static string Reasons(SheetAggregationPreview preview)
        => string.Join(" / ", preview.Blocks.Select(issue => $"{issue.Location}: {issue.Message}"));

    private static Worksheet Worksheet(string path, string sheetName)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name?.Value == sheetName);
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet!;
    }

    private static string? SingleX14ListSource(string path, string sheetName)
    {
        var validation = Assert.Single(Worksheet(path, sheetName).Descendants<X14.DataValidation>());
        return validation.DataValidationForumla1?.GetFirstChild<Xm.Formula>()?.Text;
    }

    private static List<DefinedName> DefinedNames(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        return document.WorkbookPart!.Workbook!.DefinedNames?.Elements<DefinedName>().ToList() ?? [];
    }

    private static (string Hash, long Length, DateTime LastWriteUtc) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }
}
