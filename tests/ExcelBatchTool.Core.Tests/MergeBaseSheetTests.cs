using ExcelBatchTool.App.ViewModels;
using ExcelBatchTool.Core;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 1A.1: 出力データ列の並び順の基準になるシートを明示的に選ぶ挙動。
/// Core 側の検証と、ViewModel 側の基準の保持・自動再選択を確認する。
/// </summary>
public sealed class MergeBaseSheetTests
{
    private static readonly MergeOptions NoMetadataOptions = new()
    {
        IncludeSourceFileColumn = false,
        IncludeSourceSheetColumn = false,
    };

    private static readonly string[] HeadersA = ["商品", "数量", "金額"];
    private static readonly string[] HeadersB = ["金額", "商品", "数量"];

    [Fact]
    public void Preview_WithFirstSheetAsBase_UsesItsHeaderOrder()
    {
        using var dir = new TempDir();
        var (a, b) = CreatePair(dir);

        var preview = new MergePlanner().CreatePreview(
            [new(a, "表"), new(b, "表")], new MergeSourceSelection(a, "表"), NoMetadataOptions);

        Assert.True(preview.CanExecute);
        Assert.Equal(HeadersA, preview.OutputHeaders);
        Assert.Equal("A.xlsx / 表", preview.BaseDisplay);
        Assert.True(preview.Sources[0].IsBase);
        Assert.False(preview.Sources[1].IsBase);
    }

    [Fact]
    public void Preview_WithSecondSheetAsBase_UsesThatHeaderOrderInstead()
    {
        using var dir = new TempDir();
        var (a, b) = CreatePair(dir);

        var preview = new MergePlanner().CreatePreview(
            [new(a, "表"), new(b, "表")], new MergeSourceSelection(b, "表"), NoMetadataOptions);

        Assert.True(preview.CanExecute);
        Assert.Equal(HeadersB, preview.OutputHeaders);
        Assert.Equal("B.xlsx / 表", preview.BaseDisplay);
        Assert.False(preview.Sources[0].IsBase);
        Assert.True(preview.Sources[1].IsBase);

        // 実際の出力も基準シートの列順に並ぶ。
        var output = dir.File("out.xlsx");
        Assert.True(new TableMerger().Execute(preview, NoMetadataOptions, output).Success);
        Assert.Equal(HeadersB, WorksheetTableScanner.Scan(output, "統合結果", CancellationToken.None).Headers);
        Assert.Equal(
            new[]
            {
                new[] { "1000", "架空A", "3" },
                ["2000", "架空B", "5"],
            },
            ReadRowTexts(output));
    }

    [Fact]
    public void Preview_WithBaseNotAmongSelections_IsBlocked()
    {
        using var dir = new TempDir();
        var (a, b) = CreatePair(dir);
        var outsider = dir.File("C.xlsx");
        TestTableWorkbookFactory.CreateTable(outsider, "表", HeadersA, [["架空C", 1, 10]]);

        var preview = new MergePlanner().CreatePreview(
            [new(a, "表"), new(b, "表")], new MergeSourceSelection(outsider, "表"), NoMetadataOptions);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("統合対象に含まれていません"));
    }

    [Fact]
    public void Preview_WithBaseSheetNameThatDoesNotMatch_IsBlocked()
    {
        using var dir = new TempDir();
        var (a, b) = CreatePair(dir);

        var preview = new MergePlanner().CreatePreview(
            [new(a, "表"), new(b, "表")], new MergeSourceSelection(a, "存在しないシート"), NoMetadataOptions);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("統合対象に含まれていません"));
    }

    [Fact]
    public void Preview_WithoutBase_IsBlocked()
    {
        using var dir = new TempDir();
        var (a, b) = CreatePair(dir);

        var preview = new MergePlanner().CreatePreview([new(a, "表"), new(b, "表")], null, NoMetadataOptions);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("基準シートが指定されていません"));
        Assert.Null(preview.BaseSelection);
    }

    [Fact]
    public void Preview_WhenBaseSheetItselfHasHeaderProblem_IsBlocked()
    {
        using var dir = new TempDir();
        var (_, b) = CreatePair(dir);
        var broken = dir.File("重複ヘッダー.xlsx");
        TestTableWorkbookFactory.CreateTable(broken, "表", ["商品", "商品"], [["A", "B"]]);

        var preview = new MergePlanner().CreatePreview(
            [new(broken, "表"), new(b, "表")], new MergeSourceSelection(broken, "表"), NoMetadataOptions);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("ヘッダー「商品」"));
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("基準シートに統合できない問題があるため"));
    }

    [Fact]
    public async Task ViewModel_ChangingBase_MakesPreviewStaleAndReordersOutput()
    {
        using var dir = new TempDir();
        var (a, b) = CreatePair(dir);
        var merge = CreateViewModel(dir, a, b);
        merge.IncludeSourceFileColumn = false;
        merge.IncludeSourceSheetColumn = false;

        await merge.RefreshPreviewAsync();
        Assert.False(merge.IsPreviewStale);
        Assert.Equal(HeadersA, merge.OutputHeaders);
        Assert.Equal("A.xlsx / 表", merge.PreviewBaseText);

        merge.Sources[1].IsBase = true;

        Assert.True(merge.IsPreviewStale);
        Assert.False(merge.CanCreate);
        Assert.True(merge.Sources[1].IsBase);
        Assert.False(merge.Sources[0].IsBase);

        await merge.RefreshPreviewAsync();
        Assert.False(merge.IsPreviewStale);
        Assert.Equal(HeadersB, merge.OutputHeaders);
        Assert.Equal("B.xlsx / 表", merge.PreviewBaseText);
    }

    [Fact]
    public async Task ViewModel_ExcludingTheBaseSource_MovesBaseToTheFirstRemainingSource()
    {
        using var dir = new TempDir();
        var (a, b) = CreatePair(dir);
        var merge = CreateViewModel(dir, a, b);

        Assert.True(merge.Sources[0].IsBase);

        await merge.RefreshPreviewAsync();
        Assert.False(merge.IsPreviewStale);

        merge.Sources[0].IsIncluded = false;

        Assert.True(merge.IsPreviewStale);
        Assert.False(merge.Sources[0].IsBase);
        Assert.True(merge.Sources[1].IsBase);
        Assert.Same(merge.Sources[1], merge.BaseSource);
    }

    [Fact]
    public void ViewModel_WithASingleIncludedSource_AlwaysKeepsItAsBase()
    {
        using var dir = new TempDir();
        var (a, b) = CreatePair(dir);
        var merge = CreateViewModel(dir, a, b);

        merge.Sources[1].IsIncluded = false;

        Assert.Same(merge.Sources[0], merge.BaseSource);
        Assert.Single(merge.Sources.Where(source => source.IsBase));

        // 基準を対象外にすると、残った 1 件へ基準が移る。
        merge.Sources[1].IsIncluded = true;
        merge.Sources[0].IsIncluded = false;

        Assert.Same(merge.Sources[1], merge.BaseSource);
        Assert.Single(merge.Sources.Where(source => source.IsBase));
    }

    [Fact]
    public void ViewModel_WithNoIncludedSource_HasNoBaseAndCannotRun()
    {
        using var dir = new TempDir();
        var (a, b) = CreatePair(dir);
        var merge = CreateViewModel(dir, a, b);

        foreach (var source in merge.Sources)
        {
            source.IsIncluded = false;
        }

        Assert.Null(merge.BaseSource);
        Assert.Equal("未選択", merge.BaseDisplayText);
        Assert.DoesNotContain(merge.Sources, source => source.IsBase);
        Assert.False(merge.CanCreate);
        Assert.False(merge.RefreshPreviewCommand.CanExecute(null));
    }

    [Fact]
    public async Task ViewModel_ChangingTheSheetOfTheBaseSource_KeepsItAsBaseAndGoesStale()
    {
        using var dir = new TempDir();
        var multi = dir.File("複数シート.xlsx");
        TestTableWorkbookFactory.Create(multi,
        [
            new TestSheetSpec { Name = "売上", Headers = HeadersA, Rows = [["架空A", 3, 1000]] },
            new TestSheetSpec { Name = "月報", Headers = HeadersB, Rows = [[2000, "架空B", 5]] },
        ]);
        var other = dir.File("B.xlsx");
        TestTableWorkbookFactory.CreateTable(other, "表", HeadersB, [[3000, "架空C", 7]]);

        var merge = CreateViewModel(dir, multi, other);
        merge.IncludeSourceFileColumn = false;
        merge.IncludeSourceSheetColumn = false;

        Assert.True(merge.Sources[0].IsBase);
        Assert.Equal("売上", merge.Sources[0].SelectedSheetName);

        await merge.RefreshPreviewAsync();
        Assert.Equal(HeadersA, merge.OutputHeaders);
        Assert.Equal("複数シート.xlsx / 売上", merge.PreviewBaseText);

        merge.Sources[0].SelectedSheetName = "月報";

        Assert.True(merge.IsPreviewStale);
        Assert.True(merge.Sources[0].IsBase);
        Assert.Equal("複数シート.xlsx / 月報", merge.BaseDisplayText);

        await merge.RefreshPreviewAsync();
        Assert.Equal(HeadersB, merge.OutputHeaders);
        Assert.Equal("複数シート.xlsx / 月報", merge.PreviewBaseText);
    }

    [Fact]
    public void ChangingBaseAndMerging_DoesNotModifyInputWorkbooks()
    {
        using var dir = new TempDir();
        var (a, b) = CreatePair(dir);
        var before = new[] { a, b }.ToDictionary(path => path, Snapshot);

        foreach (var basePath in new[] { a, b })
        {
            var preview = new MergePlanner().CreatePreview(
                [new(a, "表"), new(b, "表")], new MergeSourceSelection(basePath, "表"), NoMetadataOptions);
            Assert.True(preview.CanExecute);

            var output = dir.File($"out-{Path.GetFileNameWithoutExtension(basePath)}.xlsx");
            Assert.True(new TableMerger().Execute(preview, NoMetadataOptions, output).Success);
        }

        foreach (var path in new[] { a, b })
        {
            Assert.Equal(before[path], Snapshot(path));
        }
    }

    // --- helpers -------------------------------------------------------

    private static (string A, string B) CreatePair(TempDir dir)
    {
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        TestTableWorkbookFactory.CreateTable(a, "表", HeadersA, [["架空A", 3, 1000]]);
        TestTableWorkbookFactory.CreateTable(b, "表", HeadersB, [[2000, "架空B", 5]]);
        return (a, b);
    }

    /// <summary>解析済みの状態を作ってから MergeViewModel を組み立てる。</summary>
    private static MergeViewModel CreateViewModel(TempDir dir, params string[] paths)
    {
        var merge = new MergeViewModel(_ => dir.File("out.xlsx"));
        var items = new List<WorkbookItemViewModel>();
        foreach (var path in paths)
        {
            var item = new WorkbookItemViewModel(path);
            item.Apply(WorkbookAnalyzer.Analyze(path));
            items.Add(item);
        }

        merge.Sync(items);
        return merge;
    }

    private static string[][] ReadRowTexts(string path)
    {
        var headerCount = WorksheetTableScanner.Scan(path, "統合結果", CancellationToken.None).Headers.Count;
        return
        [
            .. WorksheetTableScanner.ReadDataRows(path, "統合結果", headerCount)
                .Select(row => row.Select(value => value.ToDisplayString()).ToArray()),
        ];
    }

    private static (string Hash, long Length, DateTime LastWriteUtc) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }
}
