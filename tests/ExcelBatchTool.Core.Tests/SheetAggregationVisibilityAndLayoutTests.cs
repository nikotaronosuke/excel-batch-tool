using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelBatchTool.Core.Aggregation;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 1B.1.1: シートの表示状態(Visible / Hidden / VeryHidden)を正確に保持し、
/// 保持しない印刷・ページレイアウト情報は黙って落とさず Block することを確認する。
/// </summary>
public sealed class SheetAggregationVisibilityAndLayoutTests
{
    [Theory]
    [InlineData(SheetVisibility.Visible)]
    [InlineData(SheetVisibility.Hidden)]
    [InlineData(SheetVisibility.VeryHidden)]
    public void Aggregate_SheetVisibility_IsPreservedExactly(SheetVisibility visibility)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            // 出力に表示シートが 1 枚も無いと Excel が開けないので、必ず表示シートを 1 枚含める。
            new TestAggregationSheetSpec { Name = "表示", Rows = [["A"]] },
            new TestAggregationSheetSpec
            {
                Name = "対象",
                Rows = [["B"]],
                IsHidden = visibility == SheetVisibility.Hidden,
                IsVeryHidden = visibility == SheetVisibility.VeryHidden,
            },
        ]);

        var preview = CreatePreview((path, "表示"), (path, "対象"));
        Assert.True(preview.CanExecute, string.Join(" / ", preview.Blocks.Select(issue => issue.Message)));
        Assert.Equal(SheetVisibility.Visible, preview.Sheets[0].Visibility);
        Assert.Equal(visibility, preview.Sheets[1].Visibility);

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        var states = ReadSheetStates(output);
        Assert.Null(states["表示"]);
        Assert.Equal(ExpectedState(visibility), states["対象"]);

        // Phase 0 の解析からも同じ表示状態として読める。
        var analysis = WorkbookAnalyzer.Analyze(output);
        Assert.Equal(visibility, analysis.Sheets.Single(sheet => sheet.Name == "対象").Visibility);
    }

    [Fact]
    public void Aggregate_HiddenAndVeryHiddenTogether_KeepsThemDistinct()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec { Name = "表示", Rows = [["A"]] },
            new TestAggregationSheetSpec { Name = "非表示", Rows = [["B"]], IsHidden = true },
            new TestAggregationSheetSpec { Name = "非常に非表示", Rows = [["C"]], IsVeryHidden = true },
        ]);

        var output = dir.File("out.xlsx");
        var preview = CreatePreview((path, "表示"), (path, "非表示"), (path, "非常に非表示"));
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        var states = ReadSheetStates(output);
        Assert.Null(states["表示"]);
        Assert.Equal(SheetStateValues.Hidden, states["非表示"]);
        Assert.Equal(SheetStateValues.VeryHidden, states["非常に非表示"]);

        Assert.Equal(new[] { "表示", "非表示", "非常に非表示" },
            preview.Sheets.Select(sheet => sheet.VisibilityDisplay));
    }

    [Fact]
    public void Preview_OnlyVeryHiddenSheetsSelected_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec { Name = "表示", Rows = [["A"]] },
            new TestAggregationSheetSpec { Name = "非常に非表示", Rows = [["B"]], IsVeryHidden = true },
        ]);

        var preview = CreatePreview((path, "非常に非表示"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("すべて非表示"));
    }

    public static TheoryData<string> PrintLayoutElements => new()
    {
        "pageMargins",
        "pageSetup",
        "printOptions",
        "headerFooter",
        "rowBreaks",
        "columnBreaks",
        "pageSetupProperties",
    };

    [Theory]
    [MemberData(nameof(PrintLayoutElements))]
    public void Preview_SelectedSheetWithPrintLayoutInformation_IsBlocked(string element)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path, [BuildSheet("表", element)]);

        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("印刷設定・ページレイアウト"));
    }

    [Theory]
    [MemberData(nameof(PrintLayoutElements))]
    public void Preview_PrintLayoutOnlyOnAnUnselectedSheet_DoesNotBlockTheSelectedSheet(string element)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec { Name = "きれいな表", Rows = [["A", 1]] },
            BuildSheet("印刷設定あり", element),
        ]);

        var preview = CreatePreview((path, "きれいな表"));

        Assert.True(preview.CanExecute, string.Join(" / ", preview.Blocks.Select(issue => issue.Message)));
        Assert.DoesNotContain(preview.Issues, issue => issue.Message.Contains("印刷設定・ページレイアウト"));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);
    }

    [Fact]
    public void Preview_TabColorAndPhoneticSettings_WarnWithoutBlocking()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
        [
            new TestAggregationSheetSpec
            {
                Name = "表",
                Rows = [["A", 1]],
                AddTabColor = true,
                AddPhoneticProperties = true,
            },
        ]);

        var preview = CreatePreview((path, "表"));

        Assert.True(preview.CanExecute, string.Join(" / ", preview.Blocks.Select(issue => issue.Message)));
        Assert.Contains(preview.Warnings, issue => issue.Message.Contains("シート見出しの色"));
        Assert.Contains(preview.Warnings, issue => issue.Message.Contains("ふりがな"));

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);
    }

    [Fact]
    public void AggregateWithVisibilityAndLayoutChecks_DoesNotModifyInputWorkbooks()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        TestSheetWorkbookFactory.Create(a,
        [
            new TestAggregationSheetSpec { Name = "表示", Rows = [["A"]] },
            new TestAggregationSheetSpec { Name = "非表示", Rows = [["B"]], IsHidden = true },
            new TestAggregationSheetSpec { Name = "非常に非表示", Rows = [["C"]], IsVeryHidden = true },
        ]);
        TestSheetWorkbookFactory.Create(b,
            [new TestAggregationSheetSpec { Name = "印刷設定あり", Rows = [["D"]], AddPageSetup = true }]);

        var inputs = new[] { a, b };
        var before = inputs.ToDictionary(path => path, Snapshot);

        // Block されるファイルもプレビューに通す。
        var blocked = CreatePreview((a, "表示"), (b, "印刷設定あり"));
        Assert.False(blocked.CanExecute);

        var preview = CreatePreview((a, "表示"), (a, "非表示"), (a, "非常に非表示"));
        Assert.True(preview.CanExecute);

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        foreach (var path in inputs)
        {
            Assert.Equal(before[path], Snapshot(path));
        }
    }

    // --- helpers -------------------------------------------------------

    private static TestAggregationSheetSpec BuildSheet(string name, string element) => new()
    {
        Name = name,
        Rows = [["A", 1]],
        AddPageMargins = element == "pageMargins",
        AddPageSetup = element == "pageSetup",
        AddPrintOptions = element == "printOptions",
        AddHeaderFooter = element == "headerFooter",
        AddRowBreaks = element == "rowBreaks",
        AddColumnBreaks = element == "columnBreaks",
        AddPageSetupProperties = element == "pageSetupProperties",
    };

    private static SheetAggregationPreview CreatePreview(params (string Path, string Sheet)[] selections)
        => new SheetAggregationPlanner().CreatePreview(
            [.. selections.Select(s => new SheetSelection(s.Path, s.Sheet))]);

    private static SheetStateValues? ExpectedState(SheetVisibility visibility) => visibility switch
    {
        SheetVisibility.Hidden => SheetStateValues.Hidden,
        SheetVisibility.VeryHidden => SheetStateValues.VeryHidden,
        _ => null,
    };

    private static Dictionary<string, SheetStateValues?> ReadSheetStates(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        return document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>()
            .ToDictionary(sheet => sheet.Name!.Value!, sheet => sheet.State?.Value);
    }

    private static (string Hash, long Length, DateTime LastWriteUtc) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }
}
