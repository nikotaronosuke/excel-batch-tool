using ExcelBatchTool.App.ViewModels;
using ExcelBatchTool.Core;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2A の画面側。指定が変わったらプレビューを無効(stale)にして、
/// 古い内容のまま実行できないことを確かめる。
/// </summary>
public sealed class CellMutationViewModelTests
{
    [Fact]
    public async Task Preview_WithASelectedSheet_BecomesExecutable()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        Assert.True(viewModel.HasPreview);
        Assert.True(viewModel.CanExecute);
        Assert.True(viewModel.ExecuteCommand.CanExecute(null));
    }

    [Fact]
    public async Task ChangingTheCellAddress_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        viewModel.CellReference = "C3";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheNewValue_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        viewModel.NewValueText = "別の値";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheValueKind_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        viewModel.IsBlankKind = true;

        Assert.False(viewModel.IsValueInputEnabled);
        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheSheetSelection_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        viewModel.Workbooks[0].Sheets[0].IsSelected = false;

        AssertStale(viewModel);
        Assert.False(viewModel.RefreshPreviewCommand.CanExecute(null));
    }

    [Fact]
    public async Task ChangingTheOutputSuffix_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        viewModel.OutputSuffix = "_確認";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ReloadingTheFileList_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        viewModel.Sync([]);

        AssertStale(viewModel);
        Assert.False(viewModel.HasWorkbooks);
    }

    [Fact]
    public async Task EveryTargetIsNoOp_LeavesExecuteDisabled()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir, newValue: "未確認");

        Assert.True(viewModel.HasPreview);
        Assert.False(viewModel.CanExecute);
        Assert.False(viewModel.ExecuteCommand.CanExecute(null));
        Assert.Equal(1, viewModel.Preview!.NoOpCount);
    }

    [Fact]
    public async Task ExecutingWritesTheOutputAndThenRequiresANewPreview()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        await viewModel.ExecuteCommandAsync();

        Assert.True(File.Exists(dir.File("大阪_変更済み.xlsx")));
        Assert.Contains("大阪_変更済み.xlsx", viewModel.ResultText);
        AssertStale(viewModel);
    }

    [Fact]
    public async Task WorkbookWithAFormula_IsNotSelectable()
    {
        using var dir = new TempDir();
        var path = dir.File("数式あり.xlsx");
        TestMutationWorkbookFactory.Create(path,
        [
            new MutationTestSheet
            {
                Name = "月報",
                Cells = [new MutationTestCell("A1", "項目")],
                FormulaCell = "B2",
            },
        ]);

        var viewModel = new CellMutationViewModel();
        viewModel.Sync([await AnalyzedAsync(path)]);

        var workbook = Assert.Single(viewModel.Workbooks);
        Assert.False(workbook.CanSelect);
        Assert.Contains("数式", workbook.UnavailableReason);
    }

    private static void AssertStale(CellMutationViewModel viewModel)
    {
        Assert.True(viewModel.IsPreviewStale);
        Assert.False(viewModel.HasPreview);
        Assert.False(viewModel.CanExecute);
        Assert.False(viewModel.ExecuteCommand.CanExecute(null));
    }

    private static async Task<CellMutationViewModel> BuildAsync(TempDir dir, string newValue = "確認済み")
    {
        var path = dir.File("大阪.xlsx");
        TestMutationWorkbookFactory.Create(path,
        [
            new MutationTestSheet
            {
                Name = "月報",
                Cells = [new MutationTestCell("B2", "未確認")],
            },
        ]);

        var viewModel = new CellMutationViewModel();
        viewModel.Sync([await AnalyzedAsync(path)]);
        viewModel.Workbooks[0].Sheets[0].IsSelected = true;
        viewModel.CellReference = "B2";
        viewModel.NewValueText = newValue;

        await viewModel.RefreshPreviewAsync();
        return viewModel;
    }

    private static async Task<WorkbookItemViewModel> AnalyzedAsync(string path)
    {
        var item = new WorkbookItemViewModel(path);
        item.Apply(await Task.Run(() => WorkbookAnalyzer.Analyze(path)));
        return item;
    }
}

/// <summary>テストから実行コマンドの本体を待てるようにする。</summary>
internal static class CellMutationViewModelTestExtensions
{
    public static async Task ExecuteCommandAsync(this CellMutationViewModel viewModel)
    {
        viewModel.ExecuteCommand.Execute(null);

        // コマンドは fire-and-forget なので、完了(busy 解除)を待つ。
        for (var i = 0; i < 200 && (viewModel.IsBusy || viewModel.ResultText is null); i++)
        {
            await Task.Delay(25);
        }
    }
}
