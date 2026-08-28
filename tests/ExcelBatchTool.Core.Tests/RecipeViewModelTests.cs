using ExcelBatchTool.App.ViewModels;
using ExcelBatchTool.Core.Mapping;
using ExcelBatchTool.Core.Mutation;
using ExcelBatchTool.Core.Recipes;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2D の画面側。処理設定を保存・読み込みしても、
/// 「今回使うファイル」はレシピに入れず、読み込んだだけでは何も実行しないことを確かめる。
/// すべて架空データ。
/// </summary>
public sealed class RecipeViewModelTests
{
    // ── 保存する内容(ファイルの情報を持たない)──────────────

    [Fact]
    public async Task CellInputSetRecipe_HoldsNoFilePath()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store);

        Save(viewModel.Recipes, "月末確認入力");

        AssertNoFileInformation(store, dir);
    }

    [Fact]
    public async Task SourceToFixedCellsRecipe_HoldsNoSourcePath()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMappingAsync(dir, store);

        Save(viewModel.Recipes, "店舗別月報");

        AssertNoFileInformation(store, dir);
        Assert.DoesNotContain("元データ", Json(store), StringComparison.Ordinal);
    }

    [Fact]
    public async Task TableUpdateRecipe_HoldsNoTargetPath()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildTableAsync(dir, store);

        Save(viewModel.Recipes, "商品マスタ月次更新");

        AssertNoFileInformation(store, dir);
        Assert.DoesNotContain("マスタ.xlsx", Json(store), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Recipes_HoldNoRecordOfWhatWasRun()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildTableAsync(dir, store);

        Save(viewModel.Recipes, "実行前に保存した設定");
        var before = Json(store);

        await viewModel.ExecuteAndWaitAsync();
        Assert.True(viewModel.LastRunSucceeded);

        // 実行しても、処理設定のファイルには何も書き足さない。
        Assert.Equal(before, Json(store));

        var json = Json(store);
        Assert.DoesNotContain("audit", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("changes", json, StringComparison.OrdinalIgnoreCase);

        // 出力名の指定(接尾辞)は設定なので残るが、作ったファイル名は残さない。
        Assert.Contains("_更新済み", json, StringComparison.Ordinal);
        Assert.DoesNotContain("マスタ_更新済み.xlsx", json, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Recipes_HoldNoFileHashes()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildTableAsync(dir, store);

        Save(viewModel.Recipes, "控えを持たない設定");

        var json = Json(store);
        Assert.DoesNotContain("sha", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("hash", json, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Recipes_HoldNoActualFileNames()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);

        Save((await BuildMutationAsync(dir, store)).Recipes, "入力セット");
        Save((await BuildMappingAsync(dir, store)).Recipes, "固定セル転記");
        Save((await BuildTableAsync(dir, store)).Recipes, "表の突合");

        var json = Json(store);
        foreach (var name in new[] { ".xlsx", ".csv", "大阪", "元データ", "マスタ" })
        {
            Assert.DoesNotContain(name, json, StringComparison.Ordinal);
        }
    }

    // ── 4. セルをまとめて変更 ─────────────────────────────

    [Fact]
    public async Task CellInputSet_TextOperation_ComesBack()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMutationAsync(dir, store);
        Save(saved.Recipes, "文字の入力");

        var loaded = await BuildEmptyMutationAsync(dir, store);
        Load(loaded.Recipes, "文字の入力");

        var operation = Assert.Single(loaded.Operations);
        Assert.Equal("B2", operation.CellReference);
        Assert.Equal(MutationOperationViewModel.KindText, operation.KindDisplay);
        Assert.Equal("確認済み", operation.ValueText);
    }

    [Fact]
    public async Task CellInputSet_NumberOperation_ComesBack()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMutationAsync(dir, store);
        saved.Operations[0].KindDisplay = MutationOperationViewModel.KindNumber;
        saved.Operations[0].ValueText = "1500";
        await saved.RefreshPreviewAsync();
        Save(saved.Recipes, "数値の入力");

        var loaded = await BuildEmptyMutationAsync(dir, store);
        Load(loaded.Recipes, "数値の入力");

        Assert.Equal(MutationOperationViewModel.KindNumber, loaded.Operations[0].KindDisplay);
        Assert.Equal("1500", loaded.Operations[0].ValueText);
    }

    [Fact]
    public async Task CellInputSet_BlankOperation_ComesBackWithNoValue()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMutationAsync(dir, store);
        saved.Operations[0].KindDisplay = MutationOperationViewModel.KindBlank;
        await saved.RefreshPreviewAsync();
        Save(saved.Recipes, "空欄にする");

        var loaded = await BuildEmptyMutationAsync(dir, store);
        Load(loaded.Recipes, "空欄にする");

        Assert.Equal(MutationOperationViewModel.KindBlank, loaded.Operations[0].KindDisplay);
        Assert.Equal(string.Empty, loaded.Operations[0].ValueText);
        Assert.False(loaded.Operations[0].IsValueEnabled);
    }

    [Fact]
    public async Task CellInputSet_KeepsEveryRowAndItsOrder()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMutationAsync(dir, store);

        saved.AddOperationCommand.Execute(null);
        saved.Operations[1].CellReference = "A1";
        saved.Operations[1].KindDisplay = MutationOperationViewModel.KindNumber;
        saved.Operations[1].ValueText = "42";

        saved.AddOperationCommand.Execute(null);
        saved.Operations[2].CellReference = "A2";
        saved.Operations[2].KindDisplay = MutationOperationViewModel.KindBlank;

        await saved.RefreshPreviewAsync();
        Save(saved.Recipes, "3 行の入力セット");

        var loaded = await BuildEmptyMutationAsync(dir, store);
        Load(loaded.Recipes, "3 行の入力セット");

        Assert.Equal(["B2", "A1", "A2"], loaded.Operations.Select(item => item.CellReference));
        Assert.Equal(["42"], loaded.Operations.Skip(1).Take(1).Select(item => item.ValueText));
    }

    [Fact]
    public async Task CellInputSet_KeepsTheOutputSuffix()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMutationAsync(dir, store);
        saved.OutputSuffix = "_確認済み";
        await saved.RefreshPreviewAsync();
        Save(saved.Recipes, "出力名つき");

        var loaded = await BuildEmptyMutationAsync(dir, store);
        Load(loaded.Recipes, "出力名つき");

        Assert.Equal("_確認済み", loaded.OutputSuffix);
    }

    [Fact]
    public async Task CellInputSet_LoadingMakesThePreviewStale()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store);
        Save(viewModel.Recipes, "読み込み後の状態");

        Assert.True(viewModel.HasPreview);
        Load(viewModel.Recipes, "読み込み後の状態");

        Assert.True(viewModel.IsPreviewStale);
        Assert.False(viewModel.HasPreview);
    }

    [Fact]
    public async Task CellInputSet_LoadingLeavesExecuteDisabledAndWritesNothing()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store);
        Save(viewModel.Recipes, "自動実行しない");

        Load(viewModel.Recipes, "自動実行しない");

        Assert.False(viewModel.CanExecute);
        Assert.False(viewModel.ExecuteCommand.CanExecute(null));
        Assert.False(File.Exists(dir.File("大阪_変更済み.xlsx")));

        // 対象ファイルの選択は消さない(選び直しの手間を増やさない)。
        Assert.True(viewModel.Workbooks[0].Sheets[0].IsSelected);
    }

    [Fact]
    public async Task Saving_RequiresAPreviewWithoutBlockingProblems()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store);

        Save(viewModel.Recipes, "問題のない設定");

        Assert.Contains("保存しました", viewModel.Recipes.MessageText);
        Assert.Single(store.Load().Recipes);
    }

    [Fact]
    public async Task Saving_WithAStalePreview_IsRefused()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store);

        viewModel.Operations[0].ValueText = "変更後";
        Save(viewModel.Recipes, "確かめていない設定");

        Assert.Contains("プレビューを更新", viewModel.Recipes.MessageText);
        Assert.True(viewModel.Recipes.IsMessageError);
        Assert.Empty(store.Load().Recipes);
    }

    [Fact]
    public async Task Saving_WithABlockingProblem_IsRefused()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store);

        // 存在しないセルは実行できない問題になる。
        viewModel.Operations[0].CellReference = "Z99";
        await viewModel.RefreshPreviewAsync();
        Assert.True(viewModel.Preview!.BlockCount > 0);

        Save(viewModel.Recipes, "実行できない設定");

        Assert.Contains("実行できない問題", viewModel.Recipes.MessageText);
        Assert.Empty(store.Load().Recipes);
    }

    [Fact]
    public async Task Saving_WithNoChanges_IsStillAllowed()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store, newValue: "未確認");

        Assert.Equal(1, viewModel.Preview!.NoOpCount);
        Assert.Equal(0, viewModel.Preview.ChangeCount);

        Save(viewModel.Recipes, "変更なしでも正しい設定");
        Assert.True(store.Load().Recipes.Count == 1, viewModel.Recipes.MessageText);
    }

    // ── 5. 表から転記 ─────────────────────────────────────

    [Fact]
    public async Task FixedCells_XlsxSource_ComesBack()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMappingAsync(dir, store, useXlsxSource: true);
        Save(saved.Recipes, "Excel のデータ元");

        var recipe = store.Load().Recipes[0].SourceToFixedCells!;
        Assert.Equal(SourceFileKind.Xlsx, recipe.SourceFileKind);
        Assert.Equal("売上一覧", recipe.SourceSheetName);
    }

    [Fact]
    public async Task FixedCells_CsvSource_ComesBackWithoutASheetName()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMappingAsync(dir, store);
        Save(saved.Recipes, "CSV のデータ元");

        var recipe = store.Load().Recipes[0].SourceToFixedCells!;
        Assert.Equal(SourceFileKind.Csv, recipe.SourceFileKind);
        Assert.Null(recipe.SourceSheetName);
    }

    [Fact]
    public async Task FixedCells_KeepsTheHeaderRow()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMappingAsync(dir, store, useXlsxSource: true, headerRow: 3);
        Save(saved.Recipes, "3 行目が項目名");

        var loaded = await BuildEmptyMappingAsync(dir, store);
        Load(loaded.Recipes, "3 行目が項目名");

        Assert.Equal("3", loaded.HeaderRowText);
    }

    [Fact]
    public async Task FixedCells_KeepsTheSourceSheetName()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMappingAsync(dir, store, useXlsxSource: true);
        Save(saved.Recipes, "シート指定つき");

        var loaded = await BuildEmptyMappingAsync(dir, store);
        Load(loaded.Recipes, "シート指定つき");
        loaded.SetSourceFile(dir.File("元データ.xlsx"));

        Assert.Equal("売上一覧", loaded.SourceSheetName);
    }

    [Fact]
    public async Task FixedCells_KeepsTheSourceKey()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMappingAsync(dir, store);
        Save(saved.Recipes, "キーつき");

        var loaded = await BuildEmptyMappingAsync(dir, store);
        Load(loaded.Recipes, "キーつき");

        Assert.Equal("店舗コード", loaded.KeyColumn);
    }

    [Fact]
    public async Task FixedCells_KeepsTheTargetKeyCell()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMappingAsync(dir, store);
        Save(saved.Recipes, "キーのセルつき");

        var loaded = await BuildEmptyMappingAsync(dir, store);
        Load(loaded.Recipes, "キーのセルつき");

        Assert.Equal("A1", loaded.TargetKeyCell);
    }

    [Fact]
    public async Task FixedCells_KeepsEveryMapping()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMappingAsync(dir, store);
        Save(saved.Recipes, "2 件の対応付け");

        var loaded = await BuildEmptyMappingAsync(dir, store);
        Load(loaded.Recipes, "2 件の対応付け");

        Assert.Equal(2, loaded.Mappings.Count);
        Assert.Equal("担当者", loaded.Mappings[0].SourceColumn);
        Assert.Equal("D5", loaded.Mappings[0].TargetCell);
        Assert.Equal(SourceMappingRowViewModel.KindText, loaded.Mappings[0].KindDisplay);
        Assert.Equal("売上", loaded.Mappings[1].SourceColumn);
        Assert.Equal("F8", loaded.Mappings[1].TargetCell);
        Assert.Equal(SourceMappingRowViewModel.KindNumber, loaded.Mappings[1].KindDisplay);
    }

    [Fact]
    public async Task FixedCells_LoadingMakesThePreviewStale()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMappingAsync(dir, store);
        Save(viewModel.Recipes, "読み込み後の状態");

        Assert.True(viewModel.HasPreview);
        Load(viewModel.Recipes, "読み込み後の状態");

        Assert.True(viewModel.IsPreviewStale);
        Assert.False(viewModel.ExecuteCommand.CanExecute(null));
        Assert.False(File.Exists(dir.File("大阪_転記済み.xlsx")));
    }

    [Fact]
    public async Task FixedCells_MissingSavedSheet_IsReportedInsteadOfUsingTheFirstSheet()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMappingAsync(dir, store, useXlsxSource: true);
        Save(saved.Recipes, "売上一覧の設定");

        // 今月のファイルにはそのシートが無い。
        var other = dir.File("今月.xlsx");
        TestSourceTableFactory.CreateXlsx(other, "Sheet1",
            [["店舗コード", "担当者", "売上"], ["OSAKA", "架空", 1500]]);

        var loaded = await BuildEmptyMappingAsync(dir, store);
        Load(loaded.Recipes, "売上一覧の設定");
        loaded.SetSourceFile(other);

        Assert.Null(loaded.SourceSheetName);
        Assert.Contains("売上一覧", loaded.StatusText);
        Assert.Contains("選び直して", loaded.StatusText);
    }

    [Fact]
    public async Task FixedCells_MissingSavedColumn_IsReportedInsteadOfGuessed()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMappingAsync(dir, store);
        Save(saved.Recipes, "店舗コードの設定");

        var other = dir.File("項目が違う.csv");
        TestSourceTableFactory.CreateCsv(other, ["支店コード,部署", "OSAKA,営業"]);

        var loaded = await BuildEmptyMappingAsync(dir, store);
        Load(loaded.Recipes, "店舗コードの設定");
        loaded.SetSourceFile(other);
        loaded.LoadColumnsCommand.Execute(null);

        // 「支店コード」を代わりに使わない。
        Assert.Equal(string.Empty, loaded.KeyColumn);
        Assert.Equal(string.Empty, loaded.Mappings[0].SourceColumn);
        Assert.Contains("店舗コード", loaded.StatusText);
        Assert.Contains("担当者", loaded.StatusText);
    }

    [Fact]
    public async Task FixedCells_SourceKindMismatch_IsReported()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildMappingAsync(dir, store);
        Save(saved.Recipes, "CSV 用の設定");

        // 今回は Excel を選んでいる。
        var other = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(other, "売上一覧",
            [["店舗コード", "担当者", "売上"], ["OSAKA", "架空", 1500]]);

        var loaded = await BuildEmptyMappingAsync(dir, store);
        loaded.SetSourceFile(other);
        Load(loaded.Recipes, "CSV 用の設定");

        Assert.Contains("CSV 用の設定です", loaded.Recipes.MessageText);
        Assert.True(loaded.Recipes.IsMessageError);
    }

    // ── 6. 表を突合して更新 ───────────────────────────────

    [Fact]
    public async Task TableUpdate_KeepsTheSourceHeaderRow()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildTableAsync(dir, store);
        Save(saved.Recipes, "データ元の項目行");

        var loaded = await BuildEmptyTableAsync(dir, store);
        Load(loaded.Recipes, "データ元の項目行");

        Assert.Equal("1", loaded.SourceHeaderRowText);
    }

    [Fact]
    public async Task TableUpdate_KeepsTheTargetHeaderRow()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildTableAsync(dir, store);
        saved.TargetHeaderRowText = "1";
        await saved.RefreshPreviewAsync();
        Save(saved.Recipes, "転記先の項目行");

        var loaded = await BuildEmptyTableAsync(dir, store);
        loaded.TargetHeaderRowText = "5";
        Load(loaded.Recipes, "転記先の項目行");

        Assert.Equal("1", loaded.TargetHeaderRowText);
    }

    [Fact]
    public async Task TableUpdate_KeepsTheSourceKey()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildTableAsync(dir, store);
        Save(saved.Recipes, "データ元のキー");

        var loaded = await BuildEmptyTableAsync(dir, store);
        Load(loaded.Recipes, "データ元のキー");

        Assert.Equal("SKU", loaded.SourceKeyColumn);
    }

    [Fact]
    public async Task TableUpdate_KeepsTheTargetKey()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildTableAsync(dir, store);
        Save(saved.Recipes, "転記先のキー");

        var loaded = await BuildEmptyTableAsync(dir, store);
        Load(loaded.Recipes, "転記先のキー");

        Assert.Equal("商品コード", loaded.TargetKeyColumn);
    }

    [Fact]
    public async Task TableUpdate_KeepsEveryMapping()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildTableAsync(dir, store);
        Save(saved.Recipes, "2 件の列の対応付け");

        var loaded = await BuildEmptyTableAsync(dir, store);
        Load(loaded.Recipes, "2 件の列の対応付け");

        Assert.Equal(2, loaded.Mappings.Count);
        Assert.Equal(TableColumnMappingRowViewModel.KindNumber, loaded.Mappings[0].KindDisplay);
        Assert.Equal(TableColumnMappingRowViewModel.KindNumber, loaded.Mappings[1].KindDisplay);
    }

    [Fact]
    public async Task TableUpdate_KeepsColumnNamesThatDifferBetweenTheTwoSides()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildTableAsync(dir, store);
        Save(saved.Recipes, "名前が違う列");

        var recipe = store.Load().Recipes[0].SourceTableToTargetTable!;
        Assert.Equal("SKU", recipe.SourceKeyColumn);
        Assert.Equal("商品コード", recipe.TargetKeyColumn);
        Assert.Equal("単価", recipe.Mappings[0].SourceColumn);
        Assert.Equal("販売単価", recipe.Mappings[0].TargetColumn);
    }

    [Fact]
    public async Task TableUpdate_KeepsTheOutputSuffix()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildTableAsync(dir, store);
        saved.OutputSuffix = "_突合済み";
        await saved.RefreshPreviewAsync();
        Save(saved.Recipes, "出力名つきの突合");

        var loaded = await BuildEmptyTableAsync(dir, store);
        Load(loaded.Recipes, "出力名つきの突合");

        Assert.Equal("_突合済み", loaded.OutputSuffix);
    }

    [Fact]
    public async Task TableUpdate_HoldsNoTargetWorkbook()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildTableAsync(dir, store);
        Save(saved.Recipes, "転記先を持たない");

        var loaded = await BuildEmptyTableAsync(dir, store);
        Load(loaded.Recipes, "転記先を持たない");

        // 転記先は今回選び直す。レシピからは復元しない。
        Assert.Empty(loaded.Workbooks);
        Assert.Equal(0, loaded.SelectedSheetCount);
    }

    [Fact]
    public async Task TableUpdate_HoldsNoReferenceSheet()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var saved = await BuildTableAsync(dir, store);
        Assert.NotNull(saved.ReferenceSheetDisplay);
        Save(saved.Recipes, "基準シートを持たない");

        Assert.DoesNotContain("商品一覧", Json(store), StringComparison.Ordinal);
        Assert.DoesNotContain("referenceSheet", Json(store), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task TableUpdate_LoadingMakesThePreviewStale()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildTableAsync(dir, store);
        Save(viewModel.Recipes, "読み込み後の状態");

        Assert.True(viewModel.HasPreview);
        Load(viewModel.Recipes, "読み込み後の状態");

        Assert.True(viewModel.IsPreviewStale);
        Assert.False(viewModel.ExecuteCommand.CanExecute(null));
        Assert.False(File.Exists(dir.File("マスタ_更新済み.xlsx")));
    }

    // ── 一覧・更新・削除 ─────────────────────────────────

    [Fact]
    public async Task EachTabListsOnlyItsOwnRecipes()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);

        Save((await BuildMutationAsync(dir, store)).Recipes, "入力セット");
        Save((await BuildMappingAsync(dir, store)).Recipes, "固定セル転記");

        var mutation = await BuildEmptyMutationAsync(dir, store);
        var mapping = await BuildEmptyMappingAsync(dir, store);
        var table = await BuildEmptyTableAsync(dir, store);

        Assert.Equal(["入力セット"], mutation.Recipes.Recipes.Select(item => item.Name));
        Assert.Equal(["固定セル転記"], mapping.Recipes.Recipes.Select(item => item.Name));
        Assert.Empty(table.Recipes.Recipes);
    }

    [Fact]
    public async Task Update_ReplacesTheSelectedRecipeOnly()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store);

        Save(viewModel.Recipes, "残す設定");
        Save(viewModel.Recipes, "置き換える設定");

        viewModel.Recipes.SelectedRecipe = Find(viewModel.Recipes, "置き換える設定");
        viewModel.Operations[0].ValueText = "確認済み(改)";
        await viewModel.RefreshPreviewAsync();
        viewModel.Recipes.UpdateCommand.Execute(null);

        Assert.Contains("更新しました", viewModel.Recipes.MessageText);

        var recipes = store.Load().Recipes;
        Assert.Equal(2, recipes.Count);
        Assert.Equal(
            "確認済み(改)",
            recipes.Single(r => r.Name == "置き換える設定").CellInputSet!.Operations[0].Value);
        Assert.Equal(
            "確認済み",
            recipes.Single(r => r.Name == "残す設定").CellInputSet!.Operations[0].Value);
    }

    [Fact]
    public async Task Update_WithoutConfirmation_ChangesNothing()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store, confirm: false);
        Save(viewModel.Recipes, "そのままの設定");

        viewModel.Recipes.SelectedRecipe = Find(viewModel.Recipes, "そのままの設定");
        viewModel.Operations[0].ValueText = "別の値";
        await viewModel.RefreshPreviewAsync();
        viewModel.Recipes.UpdateCommand.Execute(null);

        Assert.Contains("やめました", viewModel.Recipes.MessageText);
        Assert.Equal("確認済み", store.Load().Recipes[0].CellInputSet!.Operations[0].Value);
    }

    [Fact]
    public async Task Delete_RemovesOnlyTheSelectedRecipe()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store);

        Save(viewModel.Recipes, "消す設定");
        Save(viewModel.Recipes, "残る設定");

        viewModel.Recipes.SelectedRecipe = Find(viewModel.Recipes, "消す設定");
        viewModel.Recipes.DeleteCommand.Execute(null);

        Assert.Contains("削除しました", viewModel.Recipes.MessageText);
        Assert.Equal("残る設定", store.Load().Recipes.Single().Name);
        Assert.Null(viewModel.Recipes.SelectedRecipe);
    }

    [Fact]
    public async Task Delete_WithoutConfirmation_KeepsTheRecipe()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store, confirm: false);
        Save(viewModel.Recipes, "消さない設定");

        viewModel.Recipes.SelectedRecipe = Find(viewModel.Recipes, "消さない設定");
        viewModel.Recipes.DeleteCommand.Execute(null);

        Assert.Contains("やめました", viewModel.Recipes.MessageText);
        Assert.Single(store.Load().Recipes);
    }

    [Fact]
    public async Task DeletingSomethingThatIsGone_ReportsItSafely()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store);
        Save(viewModel.Recipes, "先に消える設定");

        viewModel.Recipes.SelectedRecipe = Find(viewModel.Recipes, "先に消える設定");
        store.Delete(store.Load().Recipes[0].Id); // 別の場所で消えた状態にする。

        viewModel.Recipes.DeleteCommand.Execute(null);

        Assert.Contains("見つかりません", viewModel.Recipes.MessageText);
        Assert.True(viewModel.Recipes.IsMessageError);
        Assert.Empty(store.Load().Recipes);
    }

    [Fact]
    public async Task SavingWithAnExistingName_DoesNotOverwrite()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildMutationAsync(dir, store);
        Save(viewModel.Recipes, "同じ名前");

        viewModel.Operations[0].ValueText = "別の値";
        await viewModel.RefreshPreviewAsync();
        Save(viewModel.Recipes, "同じ名前");

        Assert.Contains("同じ名前のレシピがあります", viewModel.Recipes.MessageText);
        var recipe = Assert.Single(store.Load().Recipes);
        Assert.Equal("確認済み", recipe.CellInputSet!.Operations[0].Value);
    }

    [Fact]
    public async Task ARecipeFileThatCannotBeRead_DoesNotStopTheNormalWork()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        File.WriteAllText(store.FilePath, "壊れた内容");

        var viewModel = await BuildMutationAsync(dir, store);
        viewModel.Recipes.Reload();

        Assert.Empty(viewModel.Recipes.Recipes);
        Assert.Contains("読み取れません", viewModel.Recipes.MessageText);

        // Excel の処理そのものはこれまでどおり動く。
        Assert.True(viewModel.CanExecute);
        await viewModel.ExecuteAndWaitAsync();
        Assert.True(viewModel.LastRunSucceeded);
        Assert.True(File.Exists(dir.File("大阪_変更済み.xlsx")));
    }

    // ── これまでの動きが変わっていないこと ─────────────────

    [Fact]
    public async Task Tab4_StillWorksWithoutAnyRecipe()
    {
        using var dir = new TempDir();
        var viewModel = await BuildMutationAsync(dir, StoreIn(dir));

        await viewModel.ExecuteAndWaitAsync();

        Assert.True(viewModel.LastRunSucceeded);
        Assert.True(File.Exists(dir.File("大阪_変更済み.xlsx")));
        Assert.Equal(1, AuditSchemaVersion(dir.File("大阪_変更済み.xlsx.audit.json")));
    }

    [Fact]
    public async Task Tab5_StillWorksWithoutAnyRecipe()
    {
        using var dir = new TempDir();
        var viewModel = await BuildMappingAsync(dir, StoreIn(dir));

        await viewModel.ExecuteAndWaitAsync();

        Assert.True(viewModel.LastRunSucceeded);
        Assert.True(File.Exists(dir.File("大阪_転記済み.xlsx")));
        Assert.Equal(2, AuditSchemaVersion(dir.File("大阪_転記済み.xlsx.audit.json")));
    }

    [Fact]
    public async Task Tab6_StillWorksWithoutAnyRecipe()
    {
        using var dir = new TempDir();
        var viewModel = await BuildTableAsync(dir, StoreIn(dir));

        await viewModel.ExecuteAndWaitAsync();

        Assert.True(viewModel.LastRunSucceeded);
        Assert.True(File.Exists(dir.File("マスタ_更新済み.xlsx")));
        Assert.Equal(3, AuditSchemaVersion(dir.File("マスタ_更新済み.xlsx.audit.json")));
    }

    [Fact]
    public async Task TheAuditStillHoldsNoRecipeInformation()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildTableAsync(dir, store);
        Save(viewModel.Recipes, "控えに出ないはずの設定");

        await viewModel.RefreshPreviewAsync();
        await viewModel.ExecuteAndWaitAsync();

        var audit = File.ReadAllText(dir.File("マスタ_更新済み.xlsx.audit.json"));
        Assert.DoesNotContain("recipe", audit, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("控えに出ないはずの設定", audit, StringComparison.Ordinal);
        Assert.Contains("map-source-table-to-target-table", audit, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AFailedRunStillReportsWhatWasLeftBehind()
    {
        using var dir = new TempDir();
        var viewModel = await BuildMutationAsync(dir, StoreIn(dir));

        Directory.CreateDirectory(dir.File("大阪_変更済み.xlsx.audit.json"));
        await viewModel.ExecuteAndWaitAsync();

        Assert.False(viewModel.LastRunSucceeded);
        Assert.Contains("取り消しました", viewModel.ResultText);
        Assert.DoesNotContain("作成していません", viewModel.ResultText);
    }

    [Fact]
    public async Task TheSourceAndTargetFilesAreStillUnchanged()
    {
        using var dir = new TempDir();
        var store = StoreIn(dir);
        var viewModel = await BuildTableAsync(dir, store);

        var source = dir.File("元データ.csv");
        var target = dir.File("マスタ.xlsx");
        var before = (Source: Fingerprint(source), Target: Fingerprint(target));

        Save(viewModel.Recipes, "実行前に保存した設定");
        await viewModel.RefreshPreviewAsync();
        await viewModel.ExecuteAndWaitAsync();
        Assert.True(viewModel.LastRunSucceeded);

        Assert.Equal(before.Source, Fingerprint(source));
        Assert.Equal(before.Target, Fingerprint(target));
    }

    // ── 補助 ─────────────────────────────────────────────

    private static RecipeStore StoreIn(TempDir dir) => new(dir.File("recipes.json"));

    private static string Json(RecipeStore store) => File.ReadAllText(store.FilePath);

    private static void Save(RecipeAreaViewModel area, string name)
    {
        area.NameText = name;
        area.SaveCommand.Execute(null);
    }

    private static void Load(RecipeAreaViewModel area, string name)
    {
        area.SelectedRecipe = Find(area, name);
        area.LoadCommand.Execute(null);
    }

    private static RecipeItemViewModel Find(RecipeAreaViewModel area, string name)
        => area.Recipes.Single(item => item.Name == name);

    private static void AssertNoFileInformation(RecipeStore store, TempDir dir)
    {
        var json = Json(store);

        Assert.DoesNotContain(dir.Root, json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(":\\", json, StringComparison.Ordinal);
        Assert.DoesNotContain("file:///", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("filePath", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("fileName", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(Environment.UserName, json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(Environment.MachineName, json, StringComparison.OrdinalIgnoreCase);
    }

    private static int AuditSchemaVersion(string path)
    {
        using var document = System.Text.Json.JsonDocument.Parse(File.ReadAllText(path));
        return document.RootElement.GetProperty("schemaVersion").GetInt32();
    }

    private static (long Length, DateTime Written) Fingerprint(string path)
    {
        var info = new FileInfo(path);
        return (info.Length, info.LastWriteTimeUtc);
    }

    private static async Task<CellMutationViewModel> BuildMutationAsync(
        TempDir dir, RecipeStore store, string newValue = "確認済み", bool confirm = true)
    {
        var viewModel = await BuildEmptyMutationAsync(dir, store, confirm);

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

        viewModel.Sync([await AnalyzedAsync(path)]);
        viewModel.Workbooks[0].Sheets[0].IsSelected = true;
        viewModel.Operations[0].CellReference = "B2";
        viewModel.Operations[0].ValueText = newValue;

        await viewModel.RefreshPreviewAsync();
        return viewModel;
    }

    private static Task<CellMutationViewModel> BuildEmptyMutationAsync(
        TempDir dir, RecipeStore store, bool confirm = true)
    {
        var viewModel = new CellMutationViewModel(() => null, store, _ => confirm);
        viewModel.Recipes.Reload();
        return Task.FromResult(viewModel);
    }

    private static async Task<SourceMappingViewModel> BuildMappingAsync(
        TempDir dir, RecipeStore store, bool useXlsxSource = false, int headerRow = 1)
    {
        string source;
        if (useXlsxSource)
        {
            source = dir.File("元データ.xlsx");
            SourceTestCell[][] rows =
            [
                ["店舗コード", "担当者", "売上"],
                ["OSAKA", "架空 太郎", 1500],
            ];
            TestSourceTableFactory.CreateXlsx(source, "売上一覧", rows, headerRow);
        }
        else
        {
            source = dir.File("元データ.csv");
            TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者,売上", "OSAKA,架空 太郎,1500"]);
        }

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
        viewModel.HeaderRowText = headerRow.ToString();
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

    private static Task<SourceMappingViewModel> BuildEmptyMappingAsync(TempDir dir, RecipeStore store)
    {
        var viewModel = new SourceMappingViewModel(() => null, store, _ => true);
        viewModel.Recipes.Reload();
        return Task.FromResult(viewModel);
    }

    private static async Task<TableUpdateViewModel> BuildTableAsync(TempDir dir, RecipeStore store)
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

    private static Task<TableUpdateViewModel> BuildEmptyTableAsync(TempDir dir, RecipeStore store)
    {
        var viewModel = new TableUpdateViewModel(() => null, store, _ => true);
        viewModel.Recipes.Reload();
        return Task.FromResult(viewModel);
    }

    private static async Task<WorkbookItemViewModel> AnalyzedAsync(string path)
    {
        var item = new WorkbookItemViewModel(path);
        item.Apply(await Task.Run(() => WorkbookAnalyzer.Analyze(path)));
        return item;
    }
}

/// <summary>テストから実行コマンドの本体を待てるようにする。</summary>
internal static class RecipeViewModelTestExtensions
{
    public static Task ExecuteAndWaitAsync(this CellMutationViewModel viewModel)
        => WaitAsync(() => viewModel.ExecuteCommand.Execute(null), () => viewModel.IsBusy,
            () => viewModel.ResultText);

    public static Task ExecuteAndWaitAsync(this SourceMappingViewModel viewModel)
        => WaitAsync(() => viewModel.ExecuteCommand.Execute(null), () => viewModel.IsBusy,
            () => viewModel.ResultText);

    public static Task ExecuteAndWaitAsync(this TableUpdateViewModel viewModel)
        => WaitAsync(() => viewModel.ExecuteCommand.Execute(null), () => viewModel.IsBusy,
            () => viewModel.ResultText);

    private static async Task WaitAsync(Action execute, Func<bool> isBusy, Func<string?> result)
    {
        execute();

        // コマンドは fire-and-forget なので、完了(busy 解除)を待つ。
        for (var i = 0; i < 200 && (isBusy() || result() is null); i++)
        {
            await Task.Delay(25);
        }
    }
}
