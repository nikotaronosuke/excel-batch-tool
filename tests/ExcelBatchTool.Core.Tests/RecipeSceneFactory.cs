using ExcelBatchTool.App.ViewModels;
using ExcelBatchTool.Core.Recipes;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// レシピ関連のテストで使う、実行できる状態まで組み立てた画面。すべて架空データ。
/// </summary>
internal static class RecipeSceneFactory
{
    public static RecipeStore StoreIn(TempDir dir) => new(dir.File("recipes.json"));

    /// <summary>タブ 4:「大阪.xlsx」の 月報!B2 を書き換えるところまで用意する。</summary>
    public static async Task<CellMutationViewModel> MutationAsync(
        TempDir dir, RecipeStore store, string newValue = "確認済み")
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

        var viewModel = new CellMutationViewModel(() => null, store, _ => true);
        viewModel.Recipes.Reload();
        viewModel.Sync([await AnalyzedAsync(path)]);
        viewModel.Workbooks[0].Sheets[0].IsSelected = true;
        viewModel.Operations[0].CellReference = "B2";
        viewModel.Operations[0].ValueText = newValue;

        await viewModel.RefreshPreviewAsync();
        return viewModel;
    }

    /// <summary>タブ 5: CSV の 1 行を「大阪.xlsx」の決まったセルへ転記するところまで。</summary>
    public static async Task<SourceMappingViewModel> MappingAsync(TempDir dir, RecipeStore store)
    {
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者,売上", "OSAKA,架空 太郎,1500"]);

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

        var viewModel = new SourceMappingViewModel(() => source, store, _ => true);
        viewModel.Recipes.Reload();
        viewModel.Sync([await AnalyzedAsync(target)]);
        viewModel.SelectSourceCommand.Execute(null);
        viewModel.LoadColumnsCommand.Execute(null);
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

    /// <summary>タブ 6: CSV と「マスタ.xlsx」を商品コードで突合するところまで。</summary>
    public static async Task<TableUpdateViewModel> TableAsync(TempDir dir, RecipeStore store)
    {
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["SKU,単価,在庫", "A001,1200,10"]);

        var target = dir.File("マスタ.xlsx");
        TestMutationWorkbookFactory.Create(target,
        [
            new MutationTestSheet
            {
                Name = "商品一覧",
                Cells =
                [
                    new MutationTestCell("A1", "商品コード"),
                    new MutationTestCell("B1", "販売単価"),
                    new MutationTestCell("C1", "在庫数"),
                    new MutationTestCell("A2", "A001"),
                    new MutationTestCell("B2", 1100),
                    new MutationTestCell("C2", 5),
                ],
            },
        ]);

        var viewModel = new TableUpdateViewModel(() => source, store, _ => true);
        viewModel.Recipes.Reload();
        viewModel.Sync([await AnalyzedAsync(target)]);
        viewModel.Workbooks[0].Sheets[0].IsSelected = true;

        viewModel.SelectSourceCommand.Execute(null);
        viewModel.LoadSourceColumns();
        viewModel.ReferenceSheetDisplay = viewModel.ReferenceSheetChoices[0];
        viewModel.LoadTargetColumns();

        viewModel.AddMappingCommand.Execute(null);
        viewModel.Mappings[0].SourceColumn = "単価";
        viewModel.Mappings[0].TargetColumn = "販売単価";
        viewModel.Mappings[0].KindDisplay = TableColumnMappingRowViewModel.KindNumber;

        viewModel.AddMappingCommand.Execute(null);
        viewModel.Mappings[1].SourceColumn = "在庫";
        viewModel.Mappings[1].TargetColumn = "在庫数";
        viewModel.Mappings[1].KindDisplay = TableColumnMappingRowViewModel.KindNumber;

        await viewModel.RefreshPreviewAsync();
        return viewModel;
    }

    public static async Task<WorkbookItemViewModel> AnalyzedAsync(string path)
    {
        var item = new WorkbookItemViewModel(path);
        item.Apply(await Task.Run(() => WorkbookAnalyzer.Analyze(path)));
        return item;
    }
}
