using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using ExcelBatchTool.Core.Aggregation;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 1B.1 で保持できない要素は、黙って落とさず事前に Block することを確認する。
/// あわせて出力の安全性(入力非変更・上書き禁止・失敗時の後始末)も確認する。
/// </summary>
public sealed class SheetAggregationBlockTests
{
    public static TheoryData<string, string> UnsupportedSheetElements => new()
    {
        { "formula", "数式" },
        { "chart", "グラフ" },
        { "image", "画像" },
        { "table", "テーブル" },
        { "conditionalFormatting", "条件付き書式" },
        { "dataValidation", "入力規則" },
        { "autoFilter", "オートフィルター" },
        { "comment", "コメント" },
        { "richText", "リッチテキスト" },
    };

    [Theory]
    [MemberData(nameof(UnsupportedSheetElements))]
    public void Preview_SelectedSheetWithUnsupportedElement_IsBlocked(string element, string expectedFragment)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path, [BuildSheet("表", element)]);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(expectedFragment));
        Assert.All(preview.Blocks, issue => Assert.DoesNotContain("集約できます", issue.Message));
    }

    [Theory]
    [MemberData(nameof(UnsupportedSheetElements))]
    public void Preview_UnsupportedElementOnlyOnAnUnselectedSheet_DoesNotBlockTheSelectedSheet(
        string element, string expectedFragment)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec { Name = "きれいな表", Rows = [["A", 1]] },
            BuildSheet("問題あり", element),
        ]);

        var preview = CreatePreview((path, "きれいな表"));

        Assert.True(preview.CanExecute,
            string.Join(" / ", preview.Blocks.Select(issue => issue.Message)));
        Assert.DoesNotContain(preview.Issues, issue => issue.Message.Contains(expectedFragment));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);
        Assert.Equal(new[] { "きれいな表" }, SheetAggregationTests.ReadSheetNames(output));
    }

    [Fact]
    public void Preview_MacroWorkbook_IsBlockedAtWorkbookLevel()
    {
        using var dir = new TempDir();
        var path = dir.File("マクロあり.xlsx");
        TestSheetWorkbookFactory.Create(path, [new TestAggregationSheetSpec { Name = "表", Rows = [["A"]] }],
            addMacro: true);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("マクロ"));
    }

    [Fact]
    public void Preview_WorkbookWithExternalLink_IsBlockedAtWorkbookLevel()
    {
        using var dir = new TempDir();
        var path = dir.File("外部参照あり.xlsx");
        TestSheetWorkbookFactory.Create(path, [new TestAggregationSheetSpec { Name = "表", Rows = [["A"]] }],
            addExternalLink: true);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("外部参照"));
    }

    [Fact]
    public void Preview_UnreadableWorkbook_IsBlockedWithoutThrowing()
    {
        using var dir = new TempDir();
        var path = dir.File("corrupt.xlsx");
        TestWorkbookFactory.CreateCorrupt(path);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("読み取れません"));
    }

    [Fact]
    public void Preview_NonXlsxFile_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("legacy.xls");
        File.WriteAllText(path, "架空の旧形式ファイル");

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(".xlsx のみ"));
    }

    [Fact]
    public void Preview_ChartsheetSelected_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("グラフシート.xlsx");
        TestWorkbookFactory.CreateWithChart(path);

        // CreateWithChart は「グラフ元」という通常シートを作る。存在しないシート名の選択も拒否する。
        var preview = CreatePreview((path, "存在しないシート"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("見つかりません"));
    }

    [Fact]
    public void Preview_SameSheetSelectedTwice_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path, [new TestAggregationSheetSpec { Name = "表", Rows = [["A"]] }]);

        var preview = new SheetAggregationPlanner().CreatePreview(
            [new(path, "表"), new(path, "表")]);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("複数回選択"));
    }

    [Fact]
    public void Execute_WhenOutputFileAlreadyExists_DoesNotOverwriteIt()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path, [new TestAggregationSheetSpec { Name = "表", Rows = [["A"]] }]);

        var output = dir.File("out.xlsx");
        File.WriteAllText(output, "既存の内容(架空)");
        var before = Snapshot(output);

        var result = new SheetAggregator().Execute(CreatePreview((path, "表")), output);

        Assert.False(result.Success);
        Assert.Contains("上書きしません", result.Message);
        Assert.Equal(before, Snapshot(output));
    }

    [Fact]
    public void Execute_WhenOutputPathEqualsInputPath_IsRefused()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path, [new TestAggregationSheetSpec { Name = "表", Rows = [["A"]] }]);
        var before = Snapshot(path);

        var result = new SheetAggregator().Execute(CreatePreview((path, "表")), path);

        Assert.False(result.Success);
        Assert.Contains("入力ファイル", result.Message);
        Assert.Equal(before, Snapshot(path));
    }

    [Fact]
    public void Execute_WhenASourceDisappearsMidway_LeavesNoPartialOutput()
    {
        using var dir = new TempDir();
        var first = dir.File("1.xlsx");
        var second = dir.File("2.xlsx");
        var third = dir.File("3.xlsx");
        foreach (var (path, value) in new[] { (first, "A"), (second, "B"), (third, "C") })
        {
            TestSheetWorkbookFactory.Create(path, [new TestAggregationSheetSpec { Name = "表", Rows = [[value]] }]);
        }

        var preview = CreatePreview((first, "表"), (second, "表"), (third, "表"));
        Assert.True(preview.CanExecute);

        File.Delete(second);

        var output = dir.File("out.xlsx");
        var result = new SheetAggregator().Execute(preview, output);

        Assert.False(result.Success);
        Assert.Contains("出力ファイルは作成していません", result.Message);
        Assert.False(File.Exists(output));
        Assert.Empty(Directory.GetFiles(dir.Root, "~ebt-sheets-*"));
    }

    [Fact]
    public void AggregateAndPreview_DoNotModifyInputWorkbooks()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        TestSheetWorkbookFactory.Create(a,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [[new Styled("見出し", 0), 1]],
                Merges = ["A1:B1"],
                FreezeTopRow = true,
                RowProperties = [new TestRowProperty(1, Height: 30)],
                ColumnProperties = [new TestColumnProperty(1, 2, Width: 18)],
            },
        ],
            styles: [new TestStyle { Bold = true, FillArgb = "FFEEEEEE" }]);
        TestSheetWorkbookFactory.Create(b, [new TestAggregationSheetSpec { Name = "別表", Rows = [["X", 2]] }]);

        var inputs = new[] { a, b };
        var before = inputs.ToDictionary(path => path, Snapshot);

        foreach (var path in inputs)
        {
            WorkbookAnalyzer.Analyze(path);
        }

        var preview = CreatePreview((a, "表"), (b, "別表"));
        Assert.True(preview.CanExecute);

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        foreach (var path in inputs)
        {
            Assert.Equal(before[path], Snapshot(path));
        }
    }

    [Fact]
    public void Output_CanBeReopenedAsSpreadsheetDocument()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec { Name = "表1", Rows = [["A", 1]] },
            new TestAggregationSheetSpec { Name = "表2", Rows = [["B", 2]] },
        ]);

        var output = dir.File("out.xlsx");
        var preview = CreatePreview((path, "表1"), (path, "表2"));
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);

        Assert.NotNull(document.WorkbookPart);
        Assert.Equal(new[] { "表1", "表2" }, SheetAggregationTests.ReadSheetNames(output));
    }

    // --- helpers -------------------------------------------------------

    private static TestAggregationSheetSpec BuildSheet(string name, string element) => new()
    {
        Name = name,
        Rows = [["A", 1], ["B", 2]],
        AddFormula = element == "formula",
        AddChart = element == "chart",
        AddImage = element == "image",
        AddTable = element == "table",
        AddConditionalFormatting = element == "conditionalFormatting",
        AddDataValidation = element == "dataValidation",
        AddHyperlink = element == "hyperlink",
        AddAutoFilter = element == "autoFilter",
        AddComment = element == "comment",
        AddRichTextCell = element == "richText",
    };

    private static SheetAggregationPreview CreatePreview(params (string Path, string Sheet)[] selections)
        => new SheetAggregationPlanner().CreatePreview(
            [.. selections.Select(s => new SheetSelection(s.Path, s.Sheet))]);

    private static (string Hash, long Length, DateTime LastWriteUtc) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }
}
