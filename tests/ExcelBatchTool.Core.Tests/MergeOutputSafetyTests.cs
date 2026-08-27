using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Tests;

/// <summary>Phase 1A の出力安全性(上書き禁止・入力非変更・失敗時に半端な出力を残さない)。</summary>
public sealed class MergeOutputSafetyTests
{
    private sealed record FileSnapshot(string Sha256, long Length, DateTime LastWriteTimeUtc);

    private static readonly MergeOptions Options = new();

    [Fact]
    public void Output_CanBeOpenedAsSpreadsheetDocument()
    {
        using var dir = new TempDir();
        var source = dir.File("A.xlsx");
        TestTableWorkbookFactory.CreateTable(source, "表", ["商品", "売上"], [["A", 100]]);

        var output = dir.File("out.xlsx");
        Assert.True(Merge(source, output).Success);

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);

        var workbookPart = document.WorkbookPart;
        Assert.NotNull(workbookPart);
        var sheet = Assert.Single(workbookPart!.Workbook!.Sheets!.Elements<Sheet>());
        Assert.Equal("統合結果", sheet.Name?.Value);

        var analysis = WorkbookAnalyzer.Analyze(output);
        Assert.Equal(AnalysisStatus.Succeeded, analysis.Status);
        Assert.Single(analysis.Sheets);
    }

    [Fact]
    public void Execute_WhenOutputFileAlreadyExists_DoesNotOverwriteIt()
    {
        using var dir = new TempDir();
        var source = dir.File("A.xlsx");
        TestTableWorkbookFactory.CreateTable(source, "表", ["商品"], [["A"]]);

        var output = dir.File("out.xlsx");
        File.WriteAllText(output, "既存の内容(架空)");
        var before = TakeSnapshot(output);

        var result = Merge(source, output);

        Assert.False(result.Success);
        Assert.Contains("上書きしません", result.Message);
        Assert.Equal(before, TakeSnapshot(output));
        Assert.Equal("既存の内容(架空)", File.ReadAllText(output));
    }

    [Fact]
    public void Execute_WhenOutputPathEqualsInputPath_IsRefused()
    {
        using var dir = new TempDir();
        var source = dir.File("A.xlsx");
        TestTableWorkbookFactory.CreateTable(source, "表", ["商品"], [["A"]]);
        var before = TakeSnapshot(source);

        var result = Merge(source, source);

        Assert.False(result.Success);
        Assert.Contains("入力ファイル", result.Message);
        Assert.Equal(before, TakeSnapshot(source));
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
            TestTableWorkbookFactory.CreateTable(path, "表", ["商品"], [[value]]);
        }

        var preview = new MergePlanner().CreatePreview(
            [new(first, "表"), new(second, "表"), new(third, "表")], new(first, "表"), Options);
        Assert.True(preview.CanExecute);

        // プレビュー後・実行中に 2 番目のファイルが失われるケース。
        File.Delete(second);

        var output = dir.File("out.xlsx");
        var result = new TableMerger().Execute(preview, Options, output);

        Assert.False(result.Success);
        Assert.Contains("出力ファイルは作成していません", result.Message);
        Assert.False(File.Exists(output));
        Assert.Empty(Directory.GetFiles(dir.Root, "~ebt-merge-*"));
    }

    [Fact]
    public void AnalyzePreviewAndMerge_DoNotModifyInputWorkbooks()
    {
        using var dir = new TempDir();
        var inputs = new List<string>();

        var plain = dir.File("通常.xlsx");
        TestTableWorkbookFactory.CreateTable(plain, "表", ["商品", "売上"], [["A", 100], ["B", 200]]);
        inputs.Add(plain);

        var reordered = dir.File("列順違い.xlsx");
        TestTableWorkbookFactory.CreateTable(reordered, "表", ["売上", "商品"], [[300, "C"]]);
        inputs.Add(reordered);

        var decorated = dir.File("装飾あり.xlsx");
        TestTableWorkbookFactory.Create(decorated,
            [new TestSheetSpec
            {
                Name = "表",
                Headers = ["商品", "売上"],
                Rows = [["D", 400]],
                AddChart = true,
                AddImage = true,
            }]);
        inputs.Add(decorated);

        var before = inputs.ToDictionary(path => path, TakeSnapshot);

        foreach (var path in inputs)
        {
            WorkbookAnalyzer.Analyze(path);
        }

        var preview = new MergePlanner().CreatePreview(
            [.. inputs.Select(path => new MergeSourceSelection(path, "表"))],
            new MergeSourceSelection(inputs[0], "表"),
            Options);
        Assert.True(preview.CanExecute);

        var output = dir.File("統合結果.xlsx");
        var result = new TableMerger().Execute(preview, Options, output);
        Assert.True(result.Success, result.Message);
        Assert.Equal(4, result.DataRowCount);

        foreach (var path in inputs)
        {
            Assert.Equal(before[path], TakeSnapshot(path));
        }
    }

    private static MergeExecutionResult Merge(string sourcePath, string outputPath)
    {
        var preview = new MergePlanner().CreatePreview([new(sourcePath, "表")], new(sourcePath, "表"), Options);
        return new TableMerger().Execute(preview, Options, outputPath);
    }

    private static FileSnapshot TakeSnapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var info = new FileInfo(path);
        return new FileSnapshot(hash, info.Length, info.LastWriteTimeUtc);
    }
}
