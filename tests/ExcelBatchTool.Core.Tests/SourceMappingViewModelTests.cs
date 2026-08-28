using ExcelBatchTool.App.ViewModels;
using ExcelBatchTool.Core;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2C1 の画面側。データ元・対応付け・転記先のどこかが変わったらプレビューを
/// 無効(stale)にして、古い内容のまま実行できないことを確かめる。
/// </summary>
public sealed class SourceMappingViewModelTests
{
    [Fact]
    public async Task LoadingColumns_FillsTheSourceHeaders()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        Assert.Equal(["店舗コード", "担当者", "売上"], viewModel.SourceColumns);
        Assert.Equal("店舗コード", viewModel.KeyColumn);
        Assert.Contains("UTF-8", viewModel.SourceInfoText);
    }

    [Fact]
    public async Task LoadingColumnsFromABrokenSource_ShowsTheReason()
    {
        using var dir = new TempDir();
        var viewModel = await BuildAsync(dir);

        var broken = dir.File("重複.csv");
        TestSourceTableFactory.CreateCsv(broken, ["店舗コード,担当者,担当者", "OSAKA,佐藤,鈴木"]);

        viewModel.SetSourceFile(broken);
        viewModel.LoadColumnsCommand.Execute(null);

        Assert.Empty(viewModel.SourceColumns);
        Assert.Contains("重複", viewModel.StatusText);
    }

    [Fact]
    public async Task Preview_WithAMappingAndASelectedSheet_BecomesExecutable()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        Assert.True(viewModel.HasPreview);
        Assert.True(viewModel.CanExecute);
        Assert.True(viewModel.ExecuteCommand.CanExecute(null));
        Assert.Equal(2, viewModel.Preview!.ChangeCount);
    }

    [Fact]
    public async Task ChangingTheSourceFile_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        var other = dir.File("別の元データ.csv");
        TestSourceTableFactory.CreateCsv(other, ["店舗コード,担当者,売上", "OSAKA,鈴木,900"]);
        viewModel.SetSourceFile(other);

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheHeaderRow_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.HeaderRowText = "3";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheKeyColumn_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.KeyColumn = "担当者";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheTargetKeyCell_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.TargetKeyCell = "B1";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task EditingAMapping_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.Mappings[0].TargetCell = "H10";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingAMappingKind_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.Mappings[0].KindDisplay = SourceMappingRowViewModel.KindNumber;

        AssertStale(viewModel);
    }

    [Fact]
    public async Task AddingAndRemovingMappings_MakeThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.AddMappingCommand.Execute(null);
        Assert.Equal(3, viewModel.Mappings.Count);
        AssertStale(viewModel);

        await viewModel.RefreshPreviewAsync();
        viewModel.SelectedMapping = viewModel.Mappings[^1];
        viewModel.RemoveMappingCommand.Execute(null);

        Assert.Equal(2, viewModel.Mappings.Count);
        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheTargetSheetSelection_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.Workbooks[0].Sheets[0].IsSelected = false;

        AssertStale(viewModel);
        Assert.False(viewModel.RefreshPreviewCommand.CanExecute(null));
    }

    [Fact]
    public async Task ChangingTheOutputSuffix_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.OutputSuffix = "_反映済み";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ReloadingTheFileList_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.Sync([]);

        AssertStale(viewModel);
        Assert.False(viewModel.HasWorkbooks);
    }

    [Fact]
    public async Task ExecutingWritesTheOutputAndThenRequiresANewPreview()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        await viewModel.ExecuteCommandAsync();

        Assert.True(viewModel.LastRunSucceeded, viewModel.ResultText);
        Assert.True(File.Exists(dir.File("大阪_転記済み.xlsx")));
        AssertStale(viewModel);
    }

    [Fact]
    public async Task AFailedRunIsNotShownAsASuccess()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        Directory.CreateDirectory(dir.File("大阪_転記済み.xlsx.audit.json"));

        await viewModel.ExecuteCommandAsync();

        Assert.False(viewModel.LastRunSucceeded);
        Assert.Contains("取り消しました", viewModel.ResultText);
        Assert.DoesNotContain("作成していません", viewModel.ResultText);
    }

    [Fact]
    public async Task ReloadingColumns_ClearsMappingsThatNoLongerExist()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        // 項目名が変わったデータ元へ切り替えて読み直す。
        var other = dir.File("別の元データ.csv");
        TestSourceTableFactory.CreateCsv(other, ["店舗コード,部署", "OSAKA,営業"]);

        viewModel.SetSourceFile(other);
        viewModel.LoadColumnsCommand.Execute(null);

        // 消えた項目は選び直してもらう(古い名前のまま実行させない)。
        Assert.Equal(string.Empty, viewModel.Mappings[0].SourceColumn);
        Assert.Equal(["店舗コード", "部署"], viewModel.Mappings[0].AvailableColumns);
    }

    private static void AssertStale(SourceMappingViewModel viewModel)
    {
        Assert.True(viewModel.IsPreviewStale);
        Assert.False(viewModel.HasPreview);
        Assert.False(viewModel.CanExecute);
        Assert.False(viewModel.ExecuteCommand.CanExecute(null));
    }

    private static async Task<SourceMappingViewModel> BuildAsync(TempDir dir)
    {
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["店舗コード,担当者,売上", "OSAKA,佐藤,1500"]);

        var target = dir.File("大阪.xlsx");
        TestMutationWorkbookFactory.Create(target,
        [
            new MutationTestSheet
            {
                Name = "月報",
                Cells =
                [
                    new MutationTestCell("A1", "OSAKA"),
                    new MutationTestCell("D5", "旧"),
                    new MutationTestCell("F8", 0),
                ],
            },
        ]);

        var viewModel = new SourceMappingViewModel(() => source);
        viewModel.Sync([await AnalyzedAsync(target)]);
        viewModel.SelectSourceCommand.Execute(null);
        viewModel.LoadColumnsCommand.Execute(null);
        return viewModel;
    }

    private static async Task<SourceMappingViewModel> BuildPreviewAsync(TempDir dir)
    {
        var viewModel = await BuildAsync(dir);

        viewModel.Workbooks[0].Sheets[0].IsSelected = true;

        viewModel.AddMappingCommand.Execute(null);
        viewModel.Mappings[0].SourceColumn = "担当者";
        viewModel.Mappings[0].TargetCell = "D5";

        viewModel.AddMappingCommand.Execute(null);
        viewModel.Mappings[1].SourceColumn = "売上";
        viewModel.Mappings[1].TargetCell = "F8";
        viewModel.Mappings[1].KindDisplay = SourceMappingRowViewModel.KindNumber;

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
internal static class SourceMappingViewModelTestExtensions
{
    public static async Task ExecuteCommandAsync(this SourceMappingViewModel viewModel)
    {
        viewModel.ExecuteCommand.Execute(null);

        for (var i = 0; i < 200 && (viewModel.IsBusy || viewModel.ResultText is null); i++)
        {
            await Task.Delay(25);
        }
    }
}
