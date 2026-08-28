using ExcelBatchTool.App.ViewModels;
using ExcelBatchTool.Core.Mutation;
using ExcelBatchTool.Core.Recipes;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2D.1。正常に実行できた設定は、作った出力が同じ名前で残っていても
/// 処理設定として保存できる。ただしその実行で使った設定と完全に同じ場合だけ。
/// すべて架空データ。
/// </summary>
public sealed class RecipeCheckpointTests
{
    // ── 4. セルをまとめて変更 ─────────────────────────────

    [Fact]
    public async Task Tab4_AfterASuccessfulRun_TheSameSettingsCanStillBeSaved()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);

        await viewModel.ExecuteAndWaitAsync();
        Assert.True(viewModel.LastRunSucceeded);
        Assert.True(File.Exists(dir.File("大阪_変更済み.xlsx")));

        // 作った出力が残っているので、取り直したプレビューは同名衝突で止まる。
        await viewModel.RefreshPreviewAsync();
        Assert.True(viewModel.Preview!.BlockCount > 0);
        Assert.False(viewModel.CanExecute);

        Save(viewModel.Recipes, "月末確認入力");

        Assert.Contains("保存しました", viewModel.Recipes.MessageText);
        Assert.Single(store.Load().Recipes);
    }

    [Fact]
    public async Task Tab4_ChangingAValueAfterTheRun_StopsTheSaveAgain()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        viewModel.Operations[0].ValueText = "完了";
        Save(viewModel.Recipes, "変えたあとの設定");

        Assert.Contains("プレビューを更新", viewModel.Recipes.MessageText);
        Assert.Empty(store.Load().Recipes);
    }

    [Fact]
    public async Task Tab4_PuttingTheValueBack_AllowsTheSaveAgain()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        viewModel.Operations[0].ValueText = "完了";
        viewModel.Operations[0].ValueText = "確認済み";

        // 「変更されたか」ではなく「今の内容が同じか」で判断する。
        Save(viewModel.Recipes, "戻した設定");

        Assert.Contains("保存しました", viewModel.Recipes.MessageText);
        Assert.Single(store.Load().Recipes);
    }

    [Fact]
    public async Task Tab4_ChangingTheOutputSuffixAfterTheRun_StopsTheSave()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        viewModel.OutputSuffix = "_確認済み";
        Save(viewModel.Recipes, "出力名を変えた設定");

        Assert.Contains("プレビューを更新", viewModel.Recipes.MessageText);
        Assert.Empty(store.Load().Recipes);
    }

    [Fact]
    public async Task Tab4_AFailedRun_LeavesNothingToSaveFrom()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);

        // 控えファイルの置き場所を塞いで、確定の途中で失敗させる。
        Directory.CreateDirectory(dir.File("大阪_変更済み.xlsx.audit.json"));
        await viewModel.ExecuteAndWaitAsync();

        Assert.False(viewModel.LastRunSucceeded);
        Assert.Contains("取り消しました", viewModel.ResultText);

        // 失敗した実行は「確認済み」にならない。設定を変えれば当然保存できない。
        viewModel.Operations[0].ValueText = "完了";
        Save(viewModel.Recipes, "失敗したあとの設定");

        Assert.Contains("プレビューを更新", viewModel.Recipes.MessageText);
        Assert.Empty(store.Load().Recipes);
    }

    [Fact]
    public async Task ACancelledRun_IsNotASuccess()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);

        // 画面には中止のボタンがまだ無いので、中止は処理側の経路で起こす。
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var result = new CellMutator().Execute(viewModel.Preview!, null, cancellation.Token);

        // 控えを残すのは result.Success のときだけなので、中止は残らない。
        Assert.False(result.Success);
        Assert.Contains("中止しました", result.Message);
        Assert.False(File.Exists(dir.File("大阪_変更済み.xlsx")));
    }

    [Fact]
    public async Task Tab4_AfterTheRun_ANewRecipeIsSavedInFull()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        Save(viewModel.Recipes, "実行後に付けた名前");

        var recipe = Assert.Single(store.Load().Recipes);
        Assert.Equal("実行後に付けた名前", recipe.Name);
        Assert.Equal(RecipeType.CellInputSet, recipe.Type);
        Assert.Equal("B2", recipe.CellInputSet!.Operations[0].Cell);
        Assert.Equal("確認済み", recipe.CellInputSet.Operations[0].Value);
        Assert.Equal("_変更済み", recipe.CellInputSet.OutputSuffix);
    }

    [Fact]
    public async Task Tab4_AfterTheRun_AnExistingRecipeCanBeUpdated()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);

        Save(viewModel.Recipes, "前からある設定");
        await viewModel.ExecuteAndWaitAsync();
        Assert.True(viewModel.LastRunSucceeded);

        viewModel.Recipes.SelectedRecipe = Find(viewModel.Recipes, "前からある設定");
        viewModel.Recipes.UpdateCommand.Execute(null);

        Assert.Contains("更新しました", viewModel.Recipes.MessageText);
        Assert.Equal("確認済み", store.Load().Recipes[0].CellInputSet!.Operations[0].Value);
    }

    // ── 5. 表から転記 ─────────────────────────────────────

    [Fact]
    public async Task Tab5_AfterASuccessfulRun_TheSameSettingsCanStillBeSaved()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MappingAsync(dir, store);

        await viewModel.ExecuteAndWaitAsync();
        Assert.True(viewModel.LastRunSucceeded);
        Assert.True(File.Exists(dir.File("大阪_転記済み.xlsx")));

        Save(viewModel.Recipes, "月次の転記");

        Assert.Contains("保存しました", viewModel.Recipes.MessageText);
        Assert.Single(store.Load().Recipes);
    }

    [Fact]
    public async Task Tab5_ChangingAMappingAfterTheRun_StopsTheSave()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MappingAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        viewModel.Mappings[0].TargetCell = "D6";
        Save(viewModel.Recipes, "対応付けを変えた設定");

        Assert.Contains("プレビューを更新", viewModel.Recipes.MessageText);
        Assert.Empty(store.Load().Recipes);
    }

    [Fact]
    public async Task Tab5_ChangingTheSourceKeyAfterTheRun_StopsTheSave()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MappingAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        viewModel.KeyColumn = "担当者";
        Save(viewModel.Recipes, "キーを変えた設定");

        Assert.Contains("プレビューを更新", viewModel.Recipes.MessageText);
        Assert.Empty(store.Load().Recipes);
    }

    [Fact]
    public async Task Tab5_ChangingOnlyTheSourceFile_StillMatches()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MappingAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        // 項目のならびが同じ別の CSV。レシピに入る内容は 1 つも変わらない。
        var other = dir.File("翌月.csv");
        TestSourceTableFactory.CreateCsv(other, ["店舗コード,担当者,売上", "OSAKA,架空 花子,1600"]);
        viewModel.SetSourceFile(other);

        Save(viewModel.Recipes, "ファイルだけ替えた設定");

        Assert.Contains("保存しました", viewModel.Recipes.MessageText);
        Assert.Single(store.Load().Recipes);
    }

    [Fact]
    public async Task Tab5_SwitchingFromCsvToXlsx_NoLongerMatches()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MappingAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        var other = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(other, "売上一覧",
            [["店舗コード", "担当者", "売上"], ["OSAKA", "架空 太郎", 1500]]);
        viewModel.SetSourceFile(other);

        Save(viewModel.Recipes, "Excel に替えた設定");

        Assert.Contains("プレビューを更新", viewModel.Recipes.MessageText);
        Assert.Empty(store.Load().Recipes);
    }

    // ── 6. 表を突合して更新 ───────────────────────────────

    [Fact]
    public async Task Tab6_AfterASuccessfulRun_TheSameSettingsCanStillBeSaved()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.TableAsync(dir, store);

        await viewModel.ExecuteAndWaitAsync();
        Assert.True(viewModel.LastRunSucceeded);
        Assert.True(File.Exists(dir.File("マスタ_更新済み.xlsx")));

        Save(viewModel.Recipes, "商品マスタ月次更新");

        Assert.Contains("保存しました", viewModel.Recipes.MessageText);
        Assert.Single(store.Load().Recipes);
    }

    [Fact]
    public async Task Tab6_ChangingAColumnMappingAfterTheRun_StopsTheSave()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.TableAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        viewModel.Mappings[0].SourceColumn = "在庫";
        Save(viewModel.Recipes, "データ元の列を変えた設定");
        Assert.Contains("プレビューを更新", viewModel.Recipes.MessageText);

        viewModel.Mappings[0].SourceColumn = "単価";
        viewModel.Mappings[0].TargetColumn = "在庫数";
        Save(viewModel.Recipes, "転記先の列を変えた設定");
        Assert.Contains("プレビューを更新", viewModel.Recipes.MessageText);

        Assert.Empty(store.Load().Recipes);
    }

    [Fact]
    public async Task Tab6_ChangingTheTargetKeyAfterTheRun_StopsTheSave()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.TableAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        viewModel.TargetKeyColumn = "販売単価";
        Save(viewModel.Recipes, "転記先のキーを変えた設定");

        Assert.Contains("プレビューを更新", viewModel.Recipes.MessageText);
        Assert.Empty(store.Load().Recipes);
    }

    [Fact]
    public async Task Tab6_ChangingOnlyTheTargetSelection_StillMatches()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.TableAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        // 転記先の選択はレシピに入らないので、外しても内容は変わらない。
        viewModel.Workbooks[0].Sheets[0].IsSelected = false;
        Assert.Null(viewModel.ReferenceSheetDisplay);

        Save(viewModel.Recipes, "選択を外した設定");

        Assert.Contains("保存しました", viewModel.Recipes.MessageText);
        Assert.Single(store.Load().Recipes);
    }

    // ── 安全性 ───────────────────────────────────────────

    [Fact]
    public async Task TheCheckpointDoesNotMakeTheRunItselfPossible()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.TableAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        await viewModel.RefreshPreviewAsync();

        // 同名の出力があるので、実行は従来どおりできない。保存だけができる。
        Assert.True(viewModel.Preview!.Mutation.BlockCount > 0);
        Assert.False(viewModel.CanExecute);
        Assert.False(viewModel.ExecuteCommand.CanExecute(null));

        await viewModel.ExecuteAndWaitAsync();
        Assert.Single(Directory.GetFiles(dir.Root, "マスタ_更新済み.xlsx"));

        Save(viewModel.Recipes, "実行はできないが保存はできる設定");
        Assert.Single(store.Load().Recipes);
    }

    [Fact]
    public async Task LoadingARecipeStillMakesThePreviewStaleAndRunsNothing()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();
        Save(viewModel.Recipes, "読み込み直す設定");

        var outputs = Directory.GetFiles(dir.Root, "大阪_変更済み.xlsx*").Length;

        viewModel.Recipes.SelectedRecipe = Find(viewModel.Recipes, "読み込み直す設定");
        viewModel.Recipes.LoadCommand.Execute(null);

        Assert.True(viewModel.IsPreviewStale);
        Assert.False(viewModel.CanExecute);
        Assert.False(viewModel.ExecuteCommand.CanExecute(null));
        Assert.Equal(outputs, Directory.GetFiles(dir.Root, "大阪_変更済み.xlsx*").Length);
    }

    [Fact]
    public async Task SavingFromTheCheckpointStillGoesThroughTheUsualChecks()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        // 名前の決まりは変わらない。
        Save(viewModel.Recipes, "   ");
        Assert.Contains("名前を入力", viewModel.Recipes.MessageText);

        Save(viewModel.Recipes, new string('あ', 61));
        Assert.Contains("60 文字以内", viewModel.Recipes.MessageText);

        Save(viewModel.Recipes, "実行後の設定");
        Assert.Contains("保存しました", viewModel.Recipes.MessageText);

        // 保存されたのは中身のそろった 1 件で、読み直せる。
        var recipe = Assert.Single(store.Load().Recipes);
        Assert.NotEmpty(recipe.Id);
        Assert.NotEmpty(recipe.CreatedAt);
        Assert.NotNull(recipe.CellInputSet);
    }

    [Fact]
    public async Task ADuplicateNameIsStillRefusedAfterASuccessfulRun()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);

        Save(viewModel.Recipes, "同じ名前");
        await viewModel.ExecuteAndWaitAsync();

        Save(viewModel.Recipes, "同じ名前");

        Assert.Contains("同じ名前のレシピがあります", viewModel.Recipes.MessageText);
        Assert.Single(store.Load().Recipes);
    }

    [Fact]
    public async Task AnUnreadableRecipeFileIsStillNotOverwrittenAfterASuccessfulRun()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        const string Broken = "{ これは JSON ではない";
        File.WriteAllText(store.FilePath, Broken);

        Save(viewModel.Recipes, "書き込めないはずの設定");

        Assert.Contains("読み取れません", viewModel.Recipes.MessageText);
        Assert.Equal(Broken, File.ReadAllText(store.FilePath));
    }

    [Fact]
    public async Task TheAuditIsUnchangedByTheCheckpoint()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);

        var mutation = await RecipeSceneFactory.MutationAsync(dir, store);
        await mutation.ExecuteAndWaitAsync();
        Save(mutation.Recipes, "入力セット");
        Assert.Equal(1, AuditSchemaVersion(dir.File("大阪_変更済み.xlsx.audit.json")));

        using var dir5 = new TempDir();
        var mapping = await RecipeSceneFactory.MappingAsync(dir5, RecipeSceneFactory.StoreIn(dir5));
        await mapping.ExecuteAndWaitAsync();
        Save(mapping.Recipes, "固定セル転記");
        Assert.Equal(2, AuditSchemaVersion(dir5.File("大阪_転記済み.xlsx.audit.json")));

        using var dir6 = new TempDir();
        var table = await RecipeSceneFactory.TableAsync(dir6, RecipeSceneFactory.StoreIn(dir6));
        await table.ExecuteAndWaitAsync();
        Save(table.Recipes, "表の突合");
        var audit = File.ReadAllText(dir6.File("マスタ_更新済み.xlsx.audit.json"));
        Assert.Equal(3, AuditSchemaVersion(dir6.File("マスタ_更新済み.xlsx.audit.json")));
        Assert.DoesNotContain("recipe", audit, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("表の突合", audit, StringComparison.Ordinal);
    }

    [Fact]
    public async Task TheSourceAndTargetFilesStayUnchanged()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.TableAsync(dir, store);

        var source = dir.File("元データ.csv");
        var target = dir.File("マスタ.xlsx");
        var before = (Source: Fingerprint(source), Target: Fingerprint(target));

        await viewModel.ExecuteAndWaitAsync();
        Save(viewModel.Recipes, "実行後に保存した設定");

        Assert.Contains("保存しました", viewModel.Recipes.MessageText);
        Assert.Equal(before.Source, Fingerprint(source));
        Assert.Equal(before.Target, Fingerprint(target));
    }

    // ── 記録しないもの ───────────────────────────────────

    [Fact]
    public async Task TheCheckpointIsNotWrittenToTheRecipeFile()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.TableAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        Save(viewModel.Recipes, "実行後の設定");

        var json = File.ReadAllText(store.FilePath);
        foreach (var word in new[]
        {
            "lastSuccessful", "checkpoint", "executed", "run", "sha", "audit",
        })
        {
            Assert.DoesNotContain(word, json, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public async Task TheCheckpointStoresNoFileInformation()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.TableAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();

        Save(viewModel.Recipes, "実行後の設定");

        var json = File.ReadAllText(store.FilePath);
        Assert.DoesNotContain(dir.Root, json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(":\\", json, StringComparison.Ordinal);
        Assert.DoesNotContain("マスタ.xlsx", json, StringComparison.Ordinal);
        Assert.DoesNotContain("元データ.csv", json, StringComparison.Ordinal);
        Assert.DoesNotContain("マスタ_更新済み", json, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AFreshScreenHasNothingToSaveFrom()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);
        await viewModel.ExecuteAndWaitAsync();
        Assert.True(viewModel.LastRunSucceeded);

        // 別の画面(アプリを開き直したのと同じ状態)には控えが無い。
        var fresh = new CellMutationViewModel(() => null, store, _ => true);
        fresh.Recipes.Reload();
        fresh.Sync([await RecipeSceneFactory.AnalyzedAsync(dir.File("大阪.xlsx"))]);
        fresh.Workbooks[0].Sheets[0].IsSelected = true;
        fresh.Operations[0].CellReference = "B2";
        fresh.Operations[0].ValueText = "確認済み";

        Save(fresh.Recipes, "開き直したあとの設定");

        Assert.Contains("プレビューを更新", fresh.Recipes.MessageText);
        Assert.Empty(store.Load().Recipes);
    }

    [Fact]
    public async Task TheHintAppearsOnlyAfterASuccessfulRun()
    {
        using var dir = new TempDir();
        var store = RecipeSceneFactory.StoreIn(dir);
        var viewModel = await RecipeSceneFactory.MutationAsync(dir, store);

        Assert.False(viewModel.Recipes.HasMessage);

        Directory.CreateDirectory(dir.File("大阪_変更済み.xlsx.audit.json"));
        await viewModel.ExecuteAndWaitAsync();
        Assert.False(viewModel.LastRunSucceeded);
        Assert.False(viewModel.Recipes.HasMessage);

        Directory.Delete(dir.File("大阪_変更済み.xlsx.audit.json"));
        await viewModel.RefreshPreviewAsync();
        await viewModel.ExecuteAndWaitAsync();

        Assert.True(viewModel.LastRunSucceeded);
        Assert.Contains("処理設定として保存できます", viewModel.Recipes.MessageText);
        Assert.False(viewModel.Recipes.IsMessageError);
    }

    // ── 設定の比較そのもの ───────────────────────────────

    [Fact]
    public void TheComparisonIgnoresTheNameAndTheTimestamps()
    {
        var left = RecipeStoreTests.TableUpdate("名前 A") with
        {
            Id = "1", CreatedAt = "2026-01-01T00:00:00.0000000+09:00", UpdatedAt = "x",
        };

        var right = RecipeStoreTests.TableUpdate("名前 B") with
        {
            Id = "2", CreatedAt = "2026-08-28T00:00:00.0000000+09:00", UpdatedAt = "y",
        };

        Assert.True(RecipeConfiguration.AreSame(left, right));
        Assert.False(RecipeConfiguration.AreSame(left, null));
        Assert.False(RecipeConfiguration.AreSame(null, null));
        Assert.False(RecipeConfiguration.AreSame(left, RecipeStoreTests.CellInput("別の種類")));
    }

    [Fact]
    public void TheComparisonNoticesEveryStoredField()
    {
        var baseline = RecipeStoreTests.TableUpdate("設定");
        var payload = baseline.SourceTableToTargetTable!;

        Assert.False(Same(baseline, payload with { SourceHeaderRow = 2 }));
        Assert.False(Same(baseline, payload with { TargetHeaderRow = 1 }));
        Assert.False(Same(baseline, payload with { SourceKeyColumn = "コード" }));
        Assert.False(Same(baseline, payload with { TargetKeyColumn = "コード" }));
        Assert.False(Same(baseline, payload with { OutputSuffix = "_済み" }));
        Assert.False(Same(baseline, payload with { SourceSheetName = "売上一覧" }));
        Assert.False(Same(baseline, payload with { Mappings = [] }));
        Assert.False(Same(baseline, payload with
        {
            Mappings =
            [
                new RecipeColumnMapping
                {
                    SourceColumn = "単価", TargetColumn = "販売単価", Kind = CellWriteKind.Text,
                },
            ],
        }));

        // 前後の空白・大文字小文字も区別する。
        Assert.False(Same(baseline, payload with { SourceKeyColumn = " SKU" }));
        Assert.False(Same(baseline, payload with { SourceKeyColumn = "sku" }));

        Assert.True(Same(baseline, payload with { }));

        static bool Same(SavedRecipe baseline, SourceTableToTargetTableRecipe changed)
            => RecipeConfiguration.AreSame(
                baseline, baseline with { SourceTableToTargetTable = changed });
    }

    // ── 補助 ─────────────────────────────────────────────

    private static void Save(RecipeAreaViewModel area, string name)
    {
        area.NameText = name;
        area.SaveCommand.Execute(null);
    }

    private static RecipeItemViewModel Find(RecipeAreaViewModel area, string name)
        => area.Recipes.Single(item => item.Name == name);

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
}
