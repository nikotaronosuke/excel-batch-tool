using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Aggregation;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 1B.2B1: ハイパーリンクの安全な移植。
/// 外部リンクは出力側で relationship を張り直し、ブック内リンクは出力シート名へ書き直す。
/// 安全に維持できないリンクは黙って落とさず Block する。
/// </summary>
public sealed class SheetAggregationHyperlinkTests
{
    [Theory]
    [InlineData("https://example.invalid/help?id=10#section")]
    [InlineData("http://example.invalid/")]
    [InlineData("mailto:someone@example.invalid")]
    public void Aggregate_ExternalHyperlink_KeepsItsTarget(string target)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["リンク"]],
                Hyperlinks = [new TestHyperlink("A1", ExternalTarget: target)],
            },
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var (hyperlink, relationship) = SingleHyperlink(output, "表");
        Assert.Equal("A1", hyperlink.Reference?.Value);
        Assert.NotNull(relationship);
        Assert.True(relationship!.IsExternal);
        Assert.Equal(target, relationship.Uri?.OriginalString);
    }

    [Fact]
    public void Aggregate_ExternalHyperlink_GetsAFreshRelationshipIdButTheSameTarget()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");

        // 元ファイル側で先に別の relationship を作らせ、r:id がずれる状況にする。
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["リンク"]],
                AddChart = false,
                Hyperlinks =
                [
                    new TestHyperlink("A1", ExternalTarget: "https://example.invalid/one"),
                    new TestHyperlink("A2", ExternalTarget: "https://example.invalid/two"),
                    new TestHyperlink("A3", ExternalTarget: "https://example.invalid/one"),
                ],
            },
        ]);

        var sourceIds = SourceHyperlinkIds(path, "表");
        Assert.Equal(3, sourceIds.Count);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var links = HyperlinksOf(output, "表");
        Assert.Equal(3, links.Count);

        var targets = links.Select(link => TargetOf(output, "表", link)).ToList();
        Assert.Equal(
            new[] { "https://example.invalid/one", "https://example.invalid/two", "https://example.invalid/one" },
            targets);

        // 出力の r:id は互いに衝突せず、すべて解決できる。
        var outputIds = links.Select(link => link.Id!.Value!).ToList();
        Assert.Equal(outputIds.Count, outputIds.Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void Aggregate_TooltipAndDisplay_ArePreserved()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["リンク"]],
                Hyperlinks =
                [
                    new TestHyperlink(
                        "A1",
                        ExternalTarget: "https://example.invalid/",
                        Tooltip: "架空の説明",
                        Display: "架空の表示名"),
                ],
            },
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var (hyperlink, _) = SingleHyperlink(output, "表");
        Assert.Equal("架空の説明", hyperlink.Tooltip?.Value);
        Assert.Equal("架空の表示名", hyperlink.Display?.Value);
    }

    [Fact]
    public void Aggregate_SameRelationshipIdInTwoWorkbooks_DoesNotCollide()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        TestSheetWorkbookFactory.Create(a,
        [
            new TestAggregationSheetSpec
            {
                Name = "A表",
                Rows = [["リンク"]],
                Hyperlinks = [new TestHyperlink("A1", ExternalTarget: "https://example.invalid/a")],
            },
        ]);
        TestSheetWorkbookFactory.Create(b,
        [
            new TestAggregationSheetSpec
            {
                Name = "B表",
                Rows = [["リンク"]],
                Hyperlinks = [new TestHyperlink("A1", ExternalTarget: "https://example.invalid/b")],
            },
        ]);

        // relationship の ID は WorksheetPart ごとに独立して振られる。
        // 元ファイル側の ID をそのまま持ち込まず、出力側で解決できることを確かめる。
        Assert.Single(SourceHyperlinkIds(a, "A表"));
        Assert.Single(SourceHyperlinkIds(b, "B表"));

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (a, "A表"), (b, "B表")).Success);

        Assert.Equal("https://example.invalid/a", TargetOf(output, "A表", HyperlinksOf(output, "A表")[0]));
        Assert.Equal("https://example.invalid/b", TargetOf(output, "B表", HyperlinksOf(output, "B表")[0]));
    }

    [Fact]
    public void Aggregate_InternalLinkWithoutSheetName_IsKeptAsIs()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["リンク"]],
                Hyperlinks = [new TestHyperlink("A1", Location: "A10")],
            },
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var (hyperlink, relationship) = SingleHyperlink(output, "表");
        Assert.Equal("A10", hyperlink.Location?.Value);
        Assert.Null(relationship);
        Assert.Null(hyperlink.Id?.Value);
    }

    [Fact]
    public void Aggregate_InternalLinkToOwnSheet_FollowsTheOutputSheetName()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "売上",
                Rows = [["リンク"]],
                Hyperlinks = [new TestHyperlink("A1", Location: "'売上'!A10")],
            },
        ]);

        var preview = new SheetAggregationPlanner().CreatePreview(
            [new SheetSelection(path, "売上", "大阪_売上")]);
        Assert.True(preview.CanExecute, string.Join(" / ", preview.Blocks.Select(issue => issue.Message)));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        var (hyperlink, _) = SingleHyperlink(output, "大阪_売上");
        Assert.Equal("'大阪_売上'!A10", hyperlink.Location?.Value);
    }

    [Fact]
    public void Aggregate_OutputSheetNameWithApostrophe_IsEscapedInTheLocation()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "売上",
                Rows = [["リンク"]],
                Hyperlinks = [new TestHyperlink("A1", Location: "売上!A10")],
            },
        ]);

        var preview = new SheetAggregationPlanner().CreatePreview(
            [new SheetSelection(path, "売上", "大阪'第一")]);
        Assert.True(preview.CanExecute, string.Join(" / ", preview.Blocks.Select(issue => issue.Message)));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        var (hyperlink, _) = SingleHyperlink(output, "大阪'第一");
        Assert.Equal("'大阪''第一'!A10", hyperlink.Location?.Value);
    }

    [Fact]
    public void Aggregate_CrossSheetInternalLink_IsRewrittenToTheTargetOutputName()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "売上",
                Rows = [["リンク"]],
                Hyperlinks = [new TestHyperlink("A1", Location: "在庫!B5")],
            },
            new TestAggregationSheetSpec { Name = "在庫", Rows = [["在庫データ"]] },
        ]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "売上"), (path, "在庫")).Success);

        var (hyperlink, _) = SingleHyperlink(output, "売上");
        Assert.Equal("'在庫'!B5", hyperlink.Location?.Value);
    }

    [Fact]
    public void Aggregate_CrossSheetInternalLink_FollowsAUserEditedTargetName()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "売上",
                Rows = [["リンク"]],
                Hyperlinks = [new TestHyperlink("A1", Location: "'在庫'!B5")],
            },
            new TestAggregationSheetSpec { Name = "在庫", Rows = [["在庫データ"]] },
        ]);

        var preview = new SheetAggregationPlanner().CreatePreview(
        [
            new SheetSelection(path, "売上", "大阪_売上"),
            new SheetSelection(path, "在庫", "大阪_在庫"),
        ]);
        Assert.True(preview.CanExecute, string.Join(" / ", preview.Blocks.Select(issue => issue.Message)));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        var (hyperlink, _) = SingleHyperlink(output, "大阪_売上");
        Assert.Equal("'大阪_在庫'!B5", hyperlink.Location?.Value);
    }

    [Fact]
    public void Preview_CrossSheetLinkToAnUnselectedSheet_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "売上",
                Rows = [["リンク"]],
                Hyperlinks = [new TestHyperlink("A10", Location: "在庫!B5")],
            },
            new TestAggregationSheetSpec { Name = "在庫", Rows = [["在庫データ"]] },
        ]);

        var preview = CreatePreview((path, "売上"));

        Assert.False(preview.CanExecute);
        var block = Assert.Single(preview.Blocks, issue => issue.Message.Contains("在庫"));
        Assert.Contains("A10", block.Message);
        Assert.Contains("集約対象に含まれていない", block.Message);

        // 参照先も選べば実行できるようになる。
        var retry = CreatePreview((path, "売上"), (path, "在庫"));
        Assert.True(retry.CanExecute, string.Join(" / ", retry.Blocks.Select(issue => issue.Message)));
    }

    public static TheoryData<string?, string?, bool, bool, bool, string> UnsupportedLinks => new()
    {
        { null, "集計範囲", false, false, false, "名前定義" },
        { null, "'[1]別ブック'!A1", false, false, false, "他のブック" },
        { null, "売上:在庫!A1", false, false, false, "複数シートにまたがる" },
        { null, "#REF!", false, false, false, "壊れています" },
        { null, "'表'!ZZZZ99999999", false, false, false, "セル位置を解釈できません" },
        { "..\\別フォルダー\\資料.xlsx", null, true, false, false, "ローカルファイル" },
        { "file:///fictional/document.xlsx", null, false, false, false, "ローカルファイル" },
        { "file://server/share/資料.xlsx", null, false, false, false, "ローカルファイル" },
        { "ftp://example.invalid/file", null, false, false, false, "対応していません" },
        { null, null, false, true, false, "参照が壊れています" },
        { null, null, false, false, true, "対応していない形式" },
    };

    [Theory]
    [MemberData(nameof(UnsupportedLinks))]
    public void Preview_UnsupportedHyperlink_IsBlocked(
        string? externalTarget,
        string? location,
        bool relative,
        bool dangling,
        bool internalRelationship,
        string expectedFragment)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["リンク"]],
                Hyperlinks =
                [
                    new TestHyperlink(
                        "A1",
                        ExternalTarget: externalTarget,
                        Location: location,
                        ExternalTargetIsRelative: relative,
                        UseDanglingRelationshipId: dangling,
                        UseInternalRelationship: internalRelationship),
                ],
            },
        ]);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(expectedFragment));
    }

    [Fact]
    public void Preview_InvalidHyperlinkReference_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["リンク"]],
                Hyperlinks = [new TestHyperlink("#REF!", Location: "A10")],
            },
        ]);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("リンクの位置"));
    }

    [Fact]
    public void Preview_UnsupportedHyperlinkOnlyOnAnUnselectedSheet_DoesNotBlockTheSelectedSheet()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec { Name = "きれいな表", Rows = [["A"]] },
            new TestAggregationSheetSpec
            {
                Name = "問題あり",
                Rows = [["リンク"]],
                Hyperlinks = [new TestHyperlink("A1", ExternalTarget: "file:///fictional/document.xlsx")],
            },
        ]);

        var preview = CreatePreview((path, "きれいな表"));

        Assert.True(preview.CanExecute, string.Join(" / ", preview.Blocks.Select(issue => issue.Message)));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);
    }

    [Fact]
    public void Aggregate_HyperlinksWithPrintSettingsAndVisibility_StayValidAndLeaveInputsUntouched()
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
                PrintTitles = "'売上'!$1:$1",
                Hyperlinks =
                [
                    new TestHyperlink("A1", ExternalTarget: "https://example.invalid/"),
                    new TestHyperlink("A2", Location: "在庫!B5", Tooltip: "在庫へ"),
                ],
            },
            new TestAggregationSheetSpec
            {
                Name = "在庫",
                Rows = [["商品", "数量"]],
                AddPageMargins = true,
                IsHidden = true,
            },
            new TestAggregationSheetSpec { Name = "メモ", Rows = [["補足"]], IsVeryHidden = true },
        ]);

        var before = Snapshot(path);

        var output = dir.File("out.xlsx");
        var result = Aggregate(output, (path, "売上"), (path, "在庫"), (path, "メモ"));
        Assert.True(result.Success, result.Message);

        Assert.Equal(before, Snapshot(path));

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);

        // Phase 1B.2A の印刷設定と Phase 1B.1.1 の表示状態が維持されている。
        var sheets = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToList();
        Assert.Null(sheets[0].State);
        Assert.Equal(SheetStateValues.Hidden, sheets[1].State?.Value);
        Assert.Equal(SheetStateValues.VeryHidden, sheets[2].State?.Value);

        var definedNames = document.WorkbookPart.Workbook.DefinedNames!.Elements<DefinedName>().ToList();
        Assert.Contains(definedNames, name => name.Text == "'売上'!$A$1:$B$2");
        Assert.Contains(definedNames, name => name.Text == "'売上'!$1:$1");

        var links = HyperlinksOf(output, "売上");
        Assert.Equal(2, links.Count);
        Assert.Equal("'在庫'!B5", links[1].Location?.Value);
        Assert.Equal("在庫へ", links[1].Tooltip?.Value);

        var errors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join(" / ", errors.Take(5).Select(error => $"{error.Path?.XPath}: {error.Description}")));
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

    private static WorksheetPart WorksheetPartOf(WorkbookPart workbookPart, string sheetName)
    {
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name?.Value == sheetName);
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    }

    private static List<Hyperlink> HyperlinksOf(string path, string sheetName)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        return WorksheetPartOf(document.WorkbookPart!, sheetName).Worksheet!
            .Descendants<Hyperlink>().ToList();
    }

    private static (Hyperlink Hyperlink, HyperlinkRelationship? Relationship) SingleHyperlink(
        string path, string sheetName)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var worksheetPart = WorksheetPartOf(document.WorkbookPart!, sheetName);
        var hyperlink = Assert.Single(worksheetPart.Worksheet!.Descendants<Hyperlink>());
        var relationship = hyperlink.Id?.Value is { } id
            ? worksheetPart.HyperlinkRelationships.FirstOrDefault(item => item.Id == id)
            : null;
        return (hyperlink, relationship);
    }

    private static string? TargetOf(string path, string sheetName, Hyperlink hyperlink)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var worksheetPart = WorksheetPartOf(document.WorkbookPart!, sheetName);
        return worksheetPart.HyperlinkRelationships
            .FirstOrDefault(item => item.Id == hyperlink.Id?.Value)?.Uri?.OriginalString;
    }

    private static List<string> SourceHyperlinkIds(string path, string sheetName)
        => [.. HyperlinksOf(path, sheetName).Select(link => link.Id?.Value ?? string.Empty)];

    private static (string Hash, long Length, DateTime LastWriteUtc) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }
}
