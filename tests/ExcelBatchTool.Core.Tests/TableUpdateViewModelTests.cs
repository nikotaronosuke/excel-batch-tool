using ExcelBatchTool.App.ViewModels;
using ExcelBatchTool.Core;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2C2 の画面側。データ元・転記先の表・対応付けのどこかが変わったら
/// プレビューを無効(stale)にして、古い内容のまま実行できないことを確かめる。
/// </summary>
public sealed class TableUpdateViewModelTests
{
    [Fact]
    public async Task LoadingBothSides_FillsColumnsAndKeys()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        Assert.Equal(["商品コード", "単価", "在庫"], viewModel.SourceColumns);
        Assert.Equal(["商品コード", "単価", "在庫", "備考"], viewModel.TargetColumns);
        Assert.Equal("商品コード", viewModel.SourceKeyColumn);
        Assert.Equal("商品コード", viewModel.TargetKeyColumn);
    }

    [Fact]
    public async Task ReferenceSheetChoices_FollowTheSelectedSheets()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        Assert.Equal(["マスタ.xlsx / 商品一覧"], viewModel.ReferenceSheetChoices);

        // 選択を外すと基準の候補も消え、基準は選び直しになる。
        viewModel.Workbooks[0].Sheets[0].IsSelected = false;

        Assert.Empty(viewModel.ReferenceSheetChoices);
        Assert.Null(viewModel.ReferenceSheetDisplay);
    }

    [Fact]
    public async Task Preview_WithMappings_BecomesExecutable()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        Assert.True(viewModel.HasPreview);
        Assert.True(viewModel.CanExecute);
        Assert.Equal(2, viewModel.Preview!.Mutation.ChangeCount);
        Assert.Contains("一致 1 件", viewModel.MatchSummaryText);
    }

    [Fact]
    public async Task ChangingTheSourceSide_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.SourceHeaderRowText = "2";
        AssertStale(viewModel);

        await viewModel.RefreshPreviewAsync(); // 失敗してもよい。stale 判定だけ見る。
        viewModel.SourceKeyColumn = "単価";
        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheTargetSide_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.TargetHeaderRowText = "2";
        AssertStale(viewModel);

        await viewModel.RefreshPreviewAsync();
        viewModel.TargetKeyColumn = "備考";
        AssertStale(viewModel);
    }

    [Fact]
    public async Task EditingMappings_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.Mappings[0].TargetColumn = "備考";
        AssertStale(viewModel);

        await viewModel.RefreshPreviewAsync();
        viewModel.AddMappingCommand.Execute(null);
        AssertStale(viewModel);

        viewModel.SelectedMapping = viewModel.Mappings[^1];
        viewModel.RemoveMappingCommand.Execute(null);
        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheOutputSuffix_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.OutputSuffix = "_突合済み";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ExecutingWritesTheOutputAndThenRequiresANewPreview()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        await viewModel.ExecuteCommandAsync();

        Assert.True(viewModel.LastRunSucceeded, viewModel.ResultText);
        Assert.True(File.Exists(dir.File("マスタ_更新済み.xlsx")));
        AssertStale(viewModel);
    }

    [Fact]
    public async Task AFailedRunIsNotShownAsASuccess()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        Directory.CreateDirectory(dir.File("マスタ_更新済み.xlsx.audit.json"));

        await viewModel.ExecuteCommandAsync();

        Assert.False(viewModel.LastRunSucceeded);
        Assert.Contains("取り消しました", viewModel.ResultText);
        Assert.DoesNotContain("作成していません", viewModel.ResultText);
    }

    [Fact]
    public async Task ReloadingSourceColumns_ClearsMappingsThatNoLongerExist()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        var other = dir.File("別の元データ.csv");
        TestSourceTableFactory.CreateCsv(other, ["商品コード,重量", "A001,5"]);

        viewModel.SetSourceFile(other);
        viewModel.LoadSourceColumns();

        Assert.Equal(string.Empty, viewModel.Mappings[0].SourceColumn);
        Assert.Equal(["商品コード", "重量"], viewModel.Mappings[0].SourceColumns);
    }

    private static void AssertStale(TableUpdateViewModel viewModel)
    {
        Assert.True(viewModel.IsPreviewStale);
        Assert.False(viewModel.HasPreview);
        Assert.False(viewModel.CanExecute);
        Assert.False(viewModel.ExecuteCommand.CanExecute(null));
    }

    private static async Task<TableUpdateViewModel> BuildAsync(TempDir dir)
    {
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["商品コード,単価,在庫", "A001,1200,10"]);

        var target = dir.File("マスタ.xlsx");
        TestMutationWorkbookFactory.Create(target,
        [
            new MutationTestSheet
            {
                Name = "商品一覧",
                Cells =
                [
                    new MutationTestCell("A1", "商品コード"),
                    new MutationTestCell("B1", "単価"),
                    new MutationTestCell("C1", "在庫"),
                    new MutationTestCell("D1", "備考"),
                    new MutationTestCell("A2", "A001"),
                    new MutationTestCell("B2", 1100),
                    new MutationTestCell("C2", 5),
                    new MutationTestCell("D2", "残す"),
                ],
            },
        ]);

        var viewModel = new TableUpdateViewModel(() => source);
        viewModel.Sync([await AnalyzedAsync(target)]);
        viewModel.Workbooks[0].Sheets[0].IsSelected = true;

        viewModel.SelectSourceCommand.Execute(null);
        viewModel.LoadSourceColumns();

        viewModel.ReferenceSheetDisplay = viewModel.ReferenceSheetChoices[0];
        viewModel.LoadTargetColumns();
        return viewModel;
    }

    private static async Task<TableUpdateViewModel> BuildPreviewAsync(TempDir dir)
    {
        var viewModel = await BuildAsync(dir);

        viewModel.AddMappingCommand.Execute(null);
        viewModel.Mappings[0].SourceColumn = "単価";
        viewModel.Mappings[0].TargetColumn = "単価";
        viewModel.Mappings[0].KindDisplay = TableColumnMappingRowViewModel.KindNumber;

        viewModel.AddMappingCommand.Execute(null);
        viewModel.Mappings[1].SourceColumn = "在庫";
        viewModel.Mappings[1].TargetColumn = "在庫";
        viewModel.Mappings[1].KindDisplay = TableColumnMappingRowViewModel.KindNumber;

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
internal static class TableUpdateViewModelTestExtensions
{
    public static async Task ExecuteCommandAsync(this TableUpdateViewModel viewModel)
    {
        viewModel.ExecuteCommand.Execute(null);

        for (var i = 0; i < 200 && (viewModel.IsBusy || viewModel.ResultText is null); i++)
        {
            await Task.Delay(25);
        }
    }
}
