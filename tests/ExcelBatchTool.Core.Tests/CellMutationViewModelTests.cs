using ExcelBatchTool.App.ViewModels;
using ExcelBatchTool.Core;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2A / 2B の画面側。入力セット(複数行)のどこかが変わったらプレビューを
/// 無効(stale)にして、古い内容のまま実行できないことを確かめる。
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

        viewModel.Operations[0].CellReference = "C3";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheNewValue_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        viewModel.Operations[0].ValueText = "別の値";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheValueKind_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        viewModel.Operations[0].KindDisplay = MutationOperationViewModel.KindBlank;

        Assert.False(viewModel.Operations[0].IsValueEnabled);
        AssertStale(viewModel);
    }

    [Fact]
    public async Task AddingAnOperation_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        viewModel.AddOperationCommand.Execute(null);

        Assert.Equal(2, viewModel.Operations.Count);
        Assert.Same(viewModel.Operations[1], viewModel.SelectedOperation);
        AssertStale(viewModel);
    }

    [Fact]
    public async Task RemovingAnOperation_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        viewModel.SelectedOperation = viewModel.Operations[0];
        viewModel.RemoveOperationCommand.Execute(null);

        Assert.Empty(viewModel.Operations);
        AssertStale(viewModel);
    }

    [Fact]
    public async Task RemoveCommand_RequiresASelectedRow()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        viewModel.SelectedOperation = null;
        Assert.False(viewModel.RemoveOperationCommand.CanExecute(null));

        viewModel.SelectedOperation = viewModel.Operations[0];
        Assert.True(viewModel.RemoveOperationCommand.CanExecute(null));
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
    public async Task MultipleOperations_AreAllApplied()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        // 2 行目: A1 を数値へ、3 行目: A2 を空欄へ。
        viewModel.AddOperationCommand.Execute(null);
        viewModel.Operations[1].CellReference = "A1";
        viewModel.Operations[1].KindDisplay = MutationOperationViewModel.KindNumber;
        viewModel.Operations[1].ValueText = "42";

        viewModel.AddOperationCommand.Execute(null);
        viewModel.Operations[2].CellReference = "A2";
        viewModel.Operations[2].KindDisplay = MutationOperationViewModel.KindBlank;

        await viewModel.RefreshPreviewAsync();
        Assert.True(viewModel.CanExecute);
        Assert.Equal(3, viewModel.Preview!.ChangeCount);

        await viewModel.ExecuteCommandAsync();
        Assert.True(viewModel.LastRunSucceeded);
        Assert.True(File.Exists(dir.File("大阪_変更済み.xlsx")));
    }

    [Fact]
    public async Task ExecutingWritesTheOutputAndThenRequiresANewPreview()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        await viewModel.ExecuteCommandAsync();

        Assert.True(File.Exists(dir.File("大阪_変更済み.xlsx")));
        Assert.Contains("大阪_変更済み.xlsx", viewModel.ResultText);
        Assert.True(viewModel.LastRunSucceeded);
        AssertStale(viewModel);
    }

    [Fact]
    public async Task AFailedRunIsNotShownAsASuccess()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        // 控えファイルの置き場所をフォルダーで塞ぎ、確定の途中で失敗させる。
        Directory.CreateDirectory(dir.File("大阪_変更済み.xlsx.audit.json"));

        await viewModel.ExecuteCommandAsync();

        Assert.False(viewModel.LastRunSucceeded);
        Assert.Contains("取り消しました", viewModel.ResultText);
        Assert.DoesNotContain("作成していません", viewModel.ResultText);
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

    // ── 表からの貼り付け ─────────────────────────────────

    [Fact]
    public void ParsePaste_ThreeValidRows_AreAccepted()
    {
        var text = "B2\t文字\t確認済み\r\nF8\t数値\t1500\r\nH10\t空欄\t\r\n";

        Assert.True(CellMutationViewModel.TryParsePastedOperations(text, out var rows, out var error), error);
        Assert.Equal(3, rows.Count);
        Assert.Equal(("B2", "文字", "確認済み"), rows[0]);
        Assert.Equal(("F8", "数値", "1500"), rows[1]);
        Assert.Equal(("H10", "空欄", ""), rows[2]);
    }

    [Fact]
    public void ParsePaste_BlankKindMayOmitTheValueColumn()
    {
        Assert.True(CellMutationViewModel.TryParsePastedOperations("H10\t空欄", out var rows, out _));
        Assert.Equal(("H10", "空欄", ""), Assert.Single(rows));
    }

    [Fact]
    public void ParsePaste_UnknownKind_RejectsTheWholePaste()
    {
        var text = "B2\t文字\t確認済み\nD5\t日付\t2026-01-01";

        Assert.False(CellMutationViewModel.TryParsePastedOperations(text, out var rows, out var error));
        Assert.Empty(rows); // 一部だけ追加しない。
        Assert.Contains("2 行目", error);
        Assert.Contains("日付", error);
    }

    [Fact]
    public void ParsePaste_BlankKindWithAValue_RejectsTheWholePaste()
    {
        var text = "B2\t文字\t確認済み\nH10\t空欄\t残っている値";

        Assert.False(CellMutationViewModel.TryParsePastedOperations(text, out var rows, out var error));
        Assert.Empty(rows);
        Assert.Contains("空欄", error);
    }

    [Fact]
    public void ParsePaste_WrongColumnCount_RejectsTheWholePaste()
    {
        Assert.False(CellMutationViewModel.TryParsePastedOperations("B2", out _, out var error));
        Assert.Contains("3 列", error);

        Assert.False(CellMutationViewModel.TryParsePastedOperations(
            "B2\t文字\t値\t余分", out _, out var error2));
        Assert.Contains("3 列", error2);
    }

    [Fact]
    public async Task PasteCommand_AppendsRowsAndMakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir, clipboard: () => "A1\t文字\tメモ\nA2\t空欄\t");

        viewModel.PasteOperationsCommand.Execute(null);

        Assert.Equal(3, viewModel.Operations.Count);
        Assert.Equal("A1", viewModel.Operations[1].CellReference);
        Assert.Equal(MutationOperationViewModel.KindBlank, viewModel.Operations[2].KindDisplay);
        AssertStale(viewModel);
    }

    [Fact]
    public async Task PasteCommand_InvalidText_AddsNothing()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir, clipboard: () => "A1\t不明\tメモ");

        viewModel.PasteOperationsCommand.Execute(null);

        Assert.Single(viewModel.Operations);
        Assert.Contains("不明", viewModel.StatusText);
        // 何も追加していないので、プレビューは有効なまま。
        Assert.True(viewModel.HasPreview);
    }

    private static void AssertStale(CellMutationViewModel viewModel)
    {
        Assert.True(viewModel.IsPreviewStale);
        Assert.False(viewModel.HasPreview);
        Assert.False(viewModel.CanExecute);
        Assert.False(viewModel.ExecuteCommand.CanExecute(null));
    }

    private static async Task<CellMutationViewModel> BuildAsync(
        TempDir dir, string newValue = "確認済み", Func<string?>? clipboard = null)
    {
        var path = dir.File("大阪.xlsx");
        TestMutationWorkbookFactory.Create(path,
        [
            new MutationTestSheet
            {
                Name = "月報",
                Cells =
                [
                    new MutationTestCell("A1", "項目"),
                    new MutationTestCell("A2", "架空A"),
                    new MutationTestCell("B2", "未確認"),
                ],
            },
        ]);

        var viewModel = new CellMutationViewModel(clipboard ?? (() => null));
        viewModel.Sync([await AnalyzedAsync(path)]);
        viewModel.Workbooks[0].Sheets[0].IsSelected = true;
        viewModel.Operations[0].CellReference = "B2";
        viewModel.Operations[0].ValueText = newValue;

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
