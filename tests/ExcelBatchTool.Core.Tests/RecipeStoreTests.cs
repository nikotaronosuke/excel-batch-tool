using System.Text;
using System.Text.Json;
using ExcelBatchTool.Core.Mapping;
using ExcelBatchTool.Core.Mutation;
using ExcelBatchTool.Core.Recipes;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2D の保存側。処理設定はこの PC の中の 1 ファイルにだけ置き、
/// 書き換えは「別名で書いてから置き換える」方式で、失敗しても元の内容を壊さない。
/// すべて架空データ。
/// </summary>
public sealed class RecipeStoreTests
{
    [Fact]
    public void Load_WithNoFile_ReturnsEmpty()
    {
        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));

        var loaded = store.Load();

        Assert.True(loaded.IsSuccess);
        Assert.Empty(loaded.Recipes);
        Assert.False(File.Exists(store.FilePath));
    }

    [Fact]
    public void Add_ThenLoad_ReturnsTheSavedRecipe()
    {
        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));

        var added = store.Add(CellInput("月末確認入力"));
        Assert.True(added.IsSuccess);

        var loaded = store.Load();
        var recipe = Assert.Single(loaded.Recipes);

        Assert.Equal("月末確認入力", recipe.Name);
        Assert.Equal(RecipeType.CellInputSet, recipe.Type);
        Assert.NotEmpty(recipe.Id);
        Assert.NotEmpty(recipe.CreatedAt);
        Assert.Equal("_変更済み", recipe.CellInputSet!.OutputSuffix);
        Assert.Equal("B2", recipe.CellInputSet.Operations[0].Cell);
    }

    [Fact]
    public void AllThreeTypes_SurviveASaveAndLoad()
    {
        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));

        Assert.True(store.Add(CellInput("入力セット")).IsSuccess);
        Assert.True(store.Add(FixedCells("固定セル転記")).IsSuccess);
        Assert.True(store.Add(TableUpdate("表の突合")).IsSuccess);

        var loaded = store.Load();
        Assert.Equal(3, loaded.Recipes.Count);

        var cells = Single(loaded, RecipeType.SourceToFixedCells).SourceToFixedCells!;
        Assert.Equal(SourceFileKind.Xlsx, cells.SourceFileKind);
        Assert.Equal("売上一覧", cells.SourceSheetName);
        Assert.Equal(2, cells.HeaderRow);
        Assert.Equal("店舗コード", cells.SourceKeyColumn);
        Assert.Equal("B1", cells.TargetKeyCell);
        Assert.Equal("D5", cells.Mappings[0].TargetCell);
        Assert.Equal(CellWriteKind.Number, cells.Mappings[0].Kind);

        var table = Single(loaded, RecipeType.SourceTableToTargetTable).SourceTableToTargetTable!;
        Assert.Equal(SourceFileKind.Csv, table.SourceFileKind);
        Assert.Null(table.SourceSheetName);
        Assert.Equal(1, table.SourceHeaderRow);
        Assert.Equal(3, table.TargetHeaderRow);
        Assert.Equal("SKU", table.SourceKeyColumn);
        Assert.Equal("商品コード", table.TargetKeyColumn);
        Assert.Equal("販売単価", table.Mappings[0].TargetColumn);

        var input = Single(loaded, RecipeType.CellInputSet).CellInputSet!;
        Assert.Equal(CellWriteKind.Text, input.Operations[0].Kind);
    }

    [Fact]
    public void JapaneseTextAndValues_AreStoredReadably()
    {
        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));

        store.Add(CellInput("月次の確認入力"));

        var json = File.ReadAllText(store.FilePath, Encoding.UTF8);
        Assert.Contains("月次の確認入力", json, StringComparison.Ordinal);
        Assert.Contains("確認済み", json, StringComparison.Ordinal);

        // 記号にエスケープされていない(そのまま読める)ことを確かめる。
        Assert.DoesNotContain("\\u6708", json, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("月次の確認入力", store.Load().Recipes[0].Name);
    }

    [Fact]
    public void UnknownSchemaVersion_IsReportedInsteadOfGuessed()
    {
        using var dir = new TempDir();
        var path = dir.File("recipes.json");
        File.WriteAllText(path, """{"schemaVersion": 99, "recipes": []}""");

        var loaded = new RecipeStore(path).Load();

        Assert.False(loaded.IsSuccess);
        Assert.Contains("新しい版", loaded.Error);
        Assert.Empty(loaded.Recipes);
    }

    [Fact]
    public void UnknownRecipeType_IsReportedInsteadOfGuessed()
    {
        using var dir = new TempDir();
        var path = dir.File("recipes.json");
        File.WriteAllText(path, """
            {"schemaVersion": 1, "recipes": [
              {"id": "1", "name": "未知", "type": "delete-everything"}]}
            """);

        var loaded = new RecipeStore(path).Load();

        Assert.False(loaded.IsSuccess);
        Assert.Empty(loaded.Recipes);
    }

    [Fact]
    public void MalformedJson_IsReportedInsteadOfSilentlyEmptied()
    {
        using var dir = new TempDir();
        var path = dir.File("recipes.json");
        File.WriteAllText(path, "{ これは JSON ではない");

        var loaded = new RecipeStore(path).Load();

        Assert.False(loaded.IsSuccess);
        Assert.Contains("読み取れません", loaded.Error);
        Assert.Empty(loaded.Recipes);
    }

    [Fact]
    public void MalformedJson_IsNotOverwrittenBySaving()
    {
        using var dir = new TempDir();
        var path = dir.File("recipes.json");
        const string Broken = "{ これは JSON ではない";
        File.WriteAllText(path, Broken);

        var store = new RecipeStore(path);
        var added = store.Add(CellInput("新しい設定"));
        var deleted = store.Delete("なにか");

        Assert.False(added.IsSuccess);
        Assert.False(deleted.IsSuccess);
        Assert.Equal(Broken, File.ReadAllText(path));
        Assert.False(File.Exists(path + ".tmp"));
    }

    [Fact]
    public void Saving_LeavesNoWorkingFileBehind()
    {
        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));

        store.Add(CellInput("1 件目"));
        store.Add(CellInput("2 件目"));

        Assert.True(File.Exists(store.FilePath));
        Assert.False(File.Exists(store.FilePath + ".tmp"));

        // 保存後のファイルはそのまま読み直せる。
        Assert.Equal(2, store.Load().Recipes.Count);
    }

    [Fact]
    public void WhenTheSaveFails_TheOldFileIsKept()
    {
        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));
        store.Add(CellInput("残るはずの設定"));
        var before = File.ReadAllText(store.FilePath);

        // 一時ファイルの場所をふさいで、書き込みを失敗させる。
        Directory.CreateDirectory(store.FilePath + ".tmp");

        var result = store.Add(CellInput("保存できない設定"));

        Assert.False(result.IsSuccess);
        Assert.Contains("保存できませんでした", result.Error);
        Assert.Equal(before, File.ReadAllText(store.FilePath));

        Directory.Delete(store.FilePath + ".tmp");
        Assert.Equal("残るはずの設定", store.Load().Recipes[0].Name);
    }

    [Fact]
    public void PreviousContent_IsKeptAsABackup()
    {
        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));

        store.Add(CellInput("1 件目"));
        Assert.False(File.Exists(store.BackupFilePath));

        store.Add(CellInput("2 件目"));

        Assert.True(File.Exists(store.BackupFilePath));
        var backup = File.ReadAllText(store.BackupFilePath);
        Assert.Contains("1 件目", backup, StringComparison.Ordinal);
        Assert.DoesNotContain("2 件目", backup, StringComparison.Ordinal);

        // 本体が壊れていても、控えは黙って上書きされない。
        File.WriteAllText(store.FilePath, "壊れた内容");
        Assert.False(store.Add(CellInput("3 件目")).IsSuccess);
        Assert.Equal(backup, File.ReadAllText(store.BackupFilePath));

        var loaded = store.Load();
        Assert.False(loaded.IsSuccess);
        Assert.True(loaded.HasBackup);
    }

    [Fact]
    public void BlankName_IsRejected()
    {
        Assert.False(RecipeName.TryNormalize("   ", out _, out var error));
        Assert.Contains("名前を入力", error);

        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));
        Assert.False(store.Add(CellInput("  ")).IsSuccess);
        Assert.False(File.Exists(store.FilePath));
    }

    [Fact]
    public void TooLongName_IsRejected()
    {
        Assert.True(RecipeName.TryNormalize(new string('あ', 60), out _, out _));
        Assert.False(RecipeName.TryNormalize(new string('あ', 61), out _, out var error));
        Assert.Contains("60 文字以内", error);
    }

    [Fact]
    public void SameName_IsNotOverwritten()
    {
        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));

        Assert.True(store.Add(CellInput("商品マスタ月次更新")).IsSuccess);
        var again = store.Add(FixedCells("商品マスタ月次更新"));

        Assert.False(again.IsSuccess);
        Assert.Contains("同じ名前のレシピがあります", again.Error);

        var recipe = Assert.Single(store.Load().Recipes);
        Assert.Equal(RecipeType.CellInputSet, recipe.Type);
    }

    [Fact]
    public void NamesDifferingOnlyInCase_AreTreatedAsTheSame()
    {
        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));

        Assert.True(store.Add(CellInput("Monthly")).IsSuccess);
        var again = store.Add(CellInput("MONTHLY"));

        Assert.False(again.IsSuccess);
        Assert.Single(store.Load().Recipes);
    }

    [Fact]
    public void JapaneseName_IsAccepted()
    {
        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));

        Assert.True(store.Add(CellInput("店舗別 月報(確認済み)")).IsSuccess);
        Assert.Equal("店舗別 月報(確認済み)", store.Load().Recipes[0].Name);
    }

    [Fact]
    public void ControlCharactersInName_AreRejected()
    {
        Assert.False(RecipeName.TryNormalize("商品\nマスタ", out _, out var error));
        Assert.Contains("改行や特殊な文字", error);

        Assert.False(RecipeName.TryNormalize("商品\tマスタ", out _, out _));
        Assert.False(RecipeName.TryNormalize("商品マスタ", out _, out _));

        // 前後の空白・改行は落としてから見る。
        Assert.True(RecipeName.TryNormalize("  商品マスタ\r\n", out var normalized, out _));
        Assert.Equal("商品マスタ", normalized);
    }

    [Fact]
    public void Recipes_AreListedByName()
    {
        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));

        store.Add(CellInput("さ行の設定"));
        store.Add(CellInput("あ行の設定"));
        store.Add(CellInput("か行の設定"));

        Assert.Equal(
            ["あ行の設定", "か行の設定", "さ行の設定"],
            store.Load().Recipes.Select(recipe => recipe.Name));
    }

    [Fact]
    public void SavedJson_HoldsOnlyTheStableFieldNames()
    {
        using var dir = new TempDir();
        var store = new RecipeStore(dir.File("recipes.json"));
        store.Add(TableUpdate("表の突合"));

        using var document = JsonDocument.Parse(File.ReadAllText(store.FilePath));
        var root = document.RootElement;

        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());

        var recipe = root.GetProperty("recipes")[0];
        Assert.Equal("source-table-to-target-table", recipe.GetProperty("type").GetString());

        var payload = recipe.GetProperty("sourceTableToTargetTable");
        Assert.Equal("csv", payload.GetProperty("sourceFileKind").GetString());
        Assert.Equal("number", payload.GetProperty("mappings")[0].GetProperty("kind").GetString());

        // 型に合わない payload は書かない。
        Assert.False(recipe.TryGetProperty("cellInputSet", out _));
        Assert.False(payload.TryGetProperty("sourceSheetName", out _));
    }

    private static SavedRecipe Single(RecipeLoadResult loaded, RecipeType type)
        => loaded.Recipes.Single(recipe => recipe.Type == type);

    internal static SavedRecipe CellInput(string name) => new()
    {
        Name = name,
        Type = RecipeType.CellInputSet,
        CellInputSet = new CellInputSetRecipe
        {
            Operations =
            [
                new RecipeOperation { Cell = "B2", Kind = CellWriteKind.Text, Value = "確認済み" },
                new RecipeOperation { Cell = "F8", Kind = CellWriteKind.Number, Value = "1500" },
                new RecipeOperation { Cell = "C3", Kind = CellWriteKind.Blank },
            ],
            OutputSuffix = "_変更済み",
        },
    };

    internal static SavedRecipe FixedCells(string name) => new()
    {
        Name = name,
        Type = RecipeType.SourceToFixedCells,
        SourceToFixedCells = new SourceToFixedCellsRecipe
        {
            SourceFileKind = SourceFileKind.Xlsx,
            SourceSheetName = "売上一覧",
            HeaderRow = 2,
            SourceKeyColumn = "店舗コード",
            TargetKeyCell = "B1",
            Mappings =
            [
                new RecipeCellMapping
                {
                    SourceColumn = "売上", TargetCell = "D5", Kind = CellWriteKind.Number,
                },
            ],
            OutputSuffix = "_転記済み",
        },
    };

    internal static SavedRecipe TableUpdate(string name) => new()
    {
        Name = name,
        Type = RecipeType.SourceTableToTargetTable,
        SourceTableToTargetTable = new SourceTableToTargetTableRecipe
        {
            SourceFileKind = SourceFileKind.Csv,
            SourceHeaderRow = 1,
            SourceKeyColumn = "SKU",
            TargetHeaderRow = 3,
            TargetKeyColumn = "商品コード",
            Mappings =
            [
                new RecipeColumnMapping
                {
                    SourceColumn = "単価", TargetColumn = "販売単価", Kind = CellWriteKind.Number,
                },
            ],
            OutputSuffix = "_更新済み",
        },
    };
}
