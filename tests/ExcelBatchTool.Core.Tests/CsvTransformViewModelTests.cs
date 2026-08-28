using ExcelBatchTool.App.ViewModels;
using ExcelBatchTool.Core.CsvTransform;
using ExcelBatchTool.Core.Mapping;
using ExcelBatchTool.Core.Recipes;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2E の画面側。項目の追加・削除・並べ替え・文字コード・引用符のどれが変わっても
/// プレビューを無効(stale)にし、レシピにファイルの情報を残さないことを確かめる。
/// すべて架空データ。
/// </summary>
public sealed class CsvTransformViewModelTests
{
    // ── 出力する項目の編集 ───────────────────────────────

    [Fact]
    public void AddingAndRemovingColumns_ChangesTheList()
    {
        using var dir = new TempDir();
        var viewModel = Build(dir);

        viewModel.AddColumnCommand.Execute(null);
        Assert.Single(viewModel.Columns);
        Assert.Same(viewModel.Columns[0], viewModel.SelectedColumn);

        viewModel.RemoveColumnCommand.Execute(null);
        Assert.Empty(viewModel.Columns);
        Assert.Null(viewModel.SelectedColumn);
        Assert.False(viewModel.RemoveColumnCommand.CanExecute(null));
    }

    [Fact]
    public void AddingEverySourceColumn_KeepsTheNamesAndOrder()
    {
        using var dir = new TempDir();
        var viewModel = Build(dir);

        viewModel.AddAllSourceColumnsCommand.Execute(null);

        Assert.Equal(
            ["商品コード", "商品名", "価格", "内部メモ"],
            viewModel.Columns.Select(column => column.OutputName));
        Assert.All(viewModel.Columns, column => Assert.True(column.IsSourceColumn));
    }

    [Fact]
    public async Task ReorderingColumns_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.SelectedColumn = viewModel.Columns[1];
        Assert.True(viewModel.MoveUpCommand.CanExecute(null));
        viewModel.MoveUpCommand.Execute(null);

        Assert.Equal("名称", viewModel.Columns[0].OutputName);
        AssertStale(viewModel);

        await viewModel.RefreshPreviewAsync();
        viewModel.SelectedColumn = viewModel.Columns[0];
        Assert.False(viewModel.MoveUpCommand.CanExecute(null));
        viewModel.MoveDownCommand.Execute(null);

        Assert.Equal("商品番号", viewModel.Columns[0].OutputName);
        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheEncoding_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.EncodingDisplay = "Shift_JIS";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingTheQuoteMode_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.QuoteDisplay = "すべての項目";

        AssertStale(viewModel);
    }

    [Fact]
    public async Task ChangingAnyOtherSetting_MakesThePreviewStale()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        viewModel.OutputSuffix = "_別名";
        AssertStale(viewModel);

        await viewModel.RefreshPreviewAsync();
        viewModel.Columns[0].OutputName = "コード";
        AssertStale(viewModel);

        await viewModel.RefreshPreviewAsync();
        viewModel.Columns[0].KindDisplay = CsvOutputColumnRowViewModel.KindFixed;
        AssertStale(viewModel);

        await viewModel.RefreshPreviewAsync();
        viewModel.Columns[0].FixedValue = "1";
        AssertStale(viewModel);

        await viewModel.RefreshPreviewAsync();
        viewModel.HeaderRowText = "2";
        AssertStale(viewModel);
    }

    [Fact]
    public async Task ThePreviewShowsWhatWillBeCreated()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        Assert.True(viewModel.HasPreview);
        Assert.True(viewModel.CanExecute);
        Assert.Contains("入力 4 列 2 行", viewModel.RowSummaryText);
        Assert.Contains("出力 3 列 2 行", viewModel.RowSummaryText);
        Assert.Contains("1 商品番号", viewModel.ColumnSummaryText);
        Assert.Contains("元データ_変換済み.csv", viewModel.OutputSummaryText);
        Assert.Equal(2, viewModel.SampleRows.Count);
        Assert.Equal(["A001", "商品A", "1"], viewModel.SampleRows[0].Values);
    }

    [Fact]
    public async Task ExecutingWritesTheCsvAndThenRequiresANewPreview()
    {
        using var dir = new TempDir();
        var viewModel = await BuildPreviewAsync(dir);

        await viewModel.ExecuteAndWaitAsync();

        Assert.True(viewModel.LastRunSucceeded);
        Assert.True(File.Exists(dir.File("元データ_変換済み.csv")));
        Assert.True(File.Exists(dir.File("元データ_変換済み.csv.audit.json")));
        AssertStale(viewModel);
    }

    // ── レシピ ───────────────────────────────────────────

    [Fact]
    public async Task TheSettingsComeBackFromARecipe()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var saved = await BuildPreviewAsync(dir, store);
        saved.EncodingDisplay = "Shift_JIS";
        saved.QuoteDisplay = "すべての項目";
        await saved.RefreshPreviewAsync();

        Save(saved.Recipes, "商品 CSV 変換");

        var loaded = NewViewModel(dir, store);
        Load(loaded.Recipes, "商品 CSV 変換");

        Assert.Equal(3, loaded.Columns.Count);
        Assert.Equal("商品番号", loaded.Columns[0].OutputName);
        Assert.Equal("商品コード", loaded.Columns[0].SourceColumn);
        Assert.Equal(CsvOutputColumnRowViewModel.KindFixed, loaded.Columns[2].KindDisplay);
        Assert.Equal("1", loaded.Columns[2].FixedValue);
        Assert.Equal("Shift_JIS", loaded.EncodingDisplay);
        Assert.Equal("すべての項目", loaded.QuoteDisplay);
        Assert.Equal("1", loaded.HeaderRowText);
        Assert.Equal(CsvTransformDefaults.OutputSuffix, loaded.OutputSuffix);
    }

    [Fact]
    public async Task TheRecipeHoldsNoFileInformation()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await BuildPreviewAsync(dir, store);

        Save(viewModel.Recipes, "ファイルを持たない設定");

        var json = File.ReadAllText(store.FilePath);
        Assert.DoesNotContain(dir.Root, json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(":\\", json, StringComparison.Ordinal);
        Assert.DoesNotContain("元データ.csv", json, StringComparison.Ordinal);
        Assert.DoesNotContain("元データ_変換済み", json, StringComparison.Ordinal);
        Assert.DoesNotContain("sha", json, StringComparison.OrdinalIgnoreCase);

        // 種類と値は将来も変わらない固定文字列で入る。
        Assert.Contains("\"csv-transform\"", json, StringComparison.Ordinal);
        Assert.Contains("\"source-column\"", json, StringComparison.Ordinal);
        Assert.Contains("\"fixed-text\"", json, StringComparison.Ordinal);
        Assert.Contains("\"utf-8-bom\"", json, StringComparison.Ordinal);
        Assert.Contains("\"minimal\"", json, StringComparison.Ordinal);
    }

    [Fact]
    public async Task LoadingARecipeMakesThePreviewStaleAndRunsNothing()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await BuildPreviewAsync(dir, store);
        Save(viewModel.Recipes, "読み込む設定");

        Assert.True(viewModel.HasPreview);
        Load(viewModel.Recipes, "読み込む設定");

        Assert.True(viewModel.IsPreviewStale);
        Assert.False(viewModel.CanExecute);
        Assert.False(viewModel.ExecuteCommand.CanExecute(null));
        Assert.False(File.Exists(dir.File("元データ_変換済み.csv")));
    }

    [Fact]
    public async Task ASuccessfulRunLetsTheSameSettingsBeSaved()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await BuildPreviewAsync(dir, store);

        await viewModel.ExecuteAndWaitAsync();
        Assert.True(viewModel.LastRunSucceeded);

        // 作った CSV が残っているので、取り直したプレビューは同名衝突で止まる。
        await viewModel.RefreshPreviewAsync();
        Assert.True(viewModel.Preview!.HasBlocks);

        Save(viewModel.Recipes, "実行後に保存した設定");

        Assert.Contains("保存しました", viewModel.Recipes.MessageText);
        Assert.Single(store.Load().Recipes);
    }

    [Fact]
    public async Task ChangingTheSettingsAfterTheRun_StopsTheSave()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await BuildPreviewAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        viewModel.QuoteDisplay = "すべての項目";
        Save(viewModel.Recipes, "変えたあとの設定");
        Assert.Contains("プレビューを更新", viewModel.Recipes.MessageText);

        // 元へ戻せばまた保存できる。
        viewModel.QuoteDisplay = "必要なときだけ";
        Save(viewModel.Recipes, "戻したあとの設定");
        Assert.Contains("保存しました", viewModel.Recipes.MessageText);

        Assert.Single(store.Load().Recipes);
    }

    [Fact]
    public async Task MissingSavedColumns_AreReportedInsteadOfGuessed()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var saved = await BuildPreviewAsync(dir, store);
        Save(saved.Recipes, "商品コードの設定");

        // 項目名が違うデータ元に切り替える。
        var other = dir.File("翌月.csv");
        TestSourceTableFactory.CreateCsv(other, ["SKU,名前", "A001,商品A"]);

        var loaded = NewViewModel(dir, store);
        loaded.SetSourceFile(other);
        loaded.LoadColumns();
        Load(loaded.Recipes, "商品コードの設定");

        Assert.Equal(string.Empty, loaded.Columns[0].SourceColumn);
        Assert.Contains("商品コード", loaded.Recipes.MessageText);
        Assert.Contains("選び直して", loaded.Recipes.MessageText);
    }

    [Fact]
    public async Task TheOtherRecipeTypesStillWorkAlongside()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);

        Save((await RecipeSceneFactory.MutationAsync(dir, store)).Recipes, "入力セット");
        Save((await RecipeSceneFactory.MappingAsync(dir, store)).Recipes, "固定セル転記");
        Save((await RecipeSceneFactory.TableAsync(dir, store)).Recipes, "表の突合");

        using var csvDir = new TempDir();
        var csvStore = new RecipeStore(store.FilePath);
        Save((await BuildPreviewAsync(csvDir, csvStore)).Recipes, "CSV 変換");

        var loaded = store.Load();
        Assert.True(loaded.IsSuccess, loaded.Error);
        Assert.Equal(4, loaded.Recipes.Count);
        Assert.Equal(
            [RecipeType.CellInputSet, RecipeType.SourceToFixedCells,
                RecipeType.SourceTableToTargetTable, RecipeType.CsvTransform],
            loaded.Recipes.OrderBy(recipe => recipe.Type).Select(recipe => recipe.Type));

        // 一覧はタブごとに種類で分かれる。
        Assert.Equal(["CSV 変換"], NewViewModel(csvDir, csvStore).Recipes.Recipes
            .Select(item => item.Name));
    }

    [Fact]
    public void ARecipeFileFromBeforeThisPhase_StillReads()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);

        // Phase 2D の時点で書かれた内容(csv-transform は入っていない)。
        File.WriteAllText(store.FilePath, """
            {
              "schemaVersion": 1,
              "recipes": [
                {
                  "id": "11111111-1111-1111-1111-111111111111",
                  "name": "前からある設定",
                  "type": "cell-input-set",
                  "createdAt": "2026-08-28T00:00:00.0000000+09:00",
                  "updatedAt": "2026-08-28T00:00:00.0000000+09:00",
                  "cellInputSet": {
                    "operations": [{ "cell": "B2", "kind": "text", "value": "確認済み" }],
                    "outputSuffix": "_変更済み"
                  }
                }
              ]
            }
            """);

        var loaded = store.Load();

        Assert.True(loaded.IsSuccess, loaded.Error);
        var recipe = Assert.Single(loaded.Recipes);
        Assert.Equal(RecipeType.CellInputSet, recipe.Type);
        Assert.Equal("確認済み", recipe.CellInputSet!.Operations[0].Value);
    }

    // ── 補助 ─────────────────────────────────────────────

    private static void AssertStale(CsvTransformViewModel viewModel)
    {
        Assert.True(viewModel.IsPreviewStale);
        Assert.False(viewModel.HasPreview);
        Assert.False(viewModel.CanExecute);
        Assert.False(viewModel.ExecuteCommand.CanExecute(null));
    }

    private static void Save(RecipeAreaViewModel area, string name)
    {
        area.NameText = name;
        area.SaveCommand.Execute(null);
    }

    private static void Load(RecipeAreaViewModel area, string name)
    {
        area.SelectedRecipe = area.Recipes.Single(item => item.Name == name);
        area.LoadCommand.Execute(null);
    }

    private static CsvTransformViewModel NewViewModel(TempDir dir, RecipeStore store)
    {
        var viewModel = new CsvTransformViewModel(() => dir.File("元データ.csv"), store, _ => true);
        viewModel.Recipes.Reload();
        return viewModel;
    }

    /// <summary>データ元を読み込んだところまで。</summary>
    private static CsvTransformViewModel Build(TempDir dir, RecipeStore? store = null)
    {
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
        [
            "商品コード,商品名,価格,内部メモ",
            "A001,商品A,1200,確認済",
            "A002,商品B,1500,確認済",
        ]);

        var viewModel = NewViewModel(dir, store ?? RecipeSceneFactory.StoreIn(dir));
        viewModel.SelectSourceCommand.Execute(null);
        viewModel.LoadColumns();
        return viewModel;
    }

    /// <summary>出力する 3 列を決めてプレビューまで。</summary>
    private static async Task<CsvTransformViewModel> BuildPreviewAsync(
        TempDir dir, RecipeStore? store = null)
    {
        var viewModel = Build(dir, store);

        viewModel.AddColumnCommand.Execute(null);
        viewModel.Columns[0].OutputName = "商品番号";
        viewModel.Columns[0].SourceColumn = "商品コード";

        viewModel.AddColumnCommand.Execute(null);
        viewModel.Columns[1].OutputName = "名称";
        viewModel.Columns[1].SourceColumn = "商品名";

        viewModel.AddColumnCommand.Execute(null);
        viewModel.Columns[2].OutputName = "公開状態";
        viewModel.Columns[2].KindDisplay = CsvOutputColumnRowViewModel.KindFixed;
        viewModel.Columns[2].FixedValue = "1";

        await viewModel.RefreshPreviewAsync();
        return viewModel;
    }
}

/// <summary>テストから実行コマンドの本体を待てるようにする。</summary>
internal static class CsvTransformViewModelTestExtensions
{
    public static async Task ExecuteAndWaitAsync(this CsvTransformViewModel viewModel)
    {
        viewModel.ExecuteCommand.Execute(null);

        for (var i = 0; i < 200 && (viewModel.IsBusy || viewModel.ResultText is null); i++)
        {
            await Task.Delay(25);
        }
    }
}
