using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

namespace ExcelBatchTool.Core.Recipes;

/// <summary>レシピの読み込み結果。</summary>
public sealed record RecipeLoadResult
{
    public required bool IsSuccess { get; init; }

    /// <summary>名前順に並べたレシピ。読めなかったときは空。</summary>
    public IReadOnlyList<SavedRecipe> Recipes { get; init; } = [];

    public string? Error { get; init; }

    /// <summary>読めなかったとき、1 つ前の内容の控えが残っているか。</summary>
    public bool HasBackup { get; init; }
}

/// <summary>レシピの保存・更新・削除の結果。</summary>
public sealed record RecipeSaveResult
{
    public required bool IsSuccess { get; init; }

    public IReadOnlyList<SavedRecipe> Recipes { get; init; } = [];

    public string? Error { get; init; }

    /// <summary>保存・更新したレシピ(削除・失敗時は null)。</summary>
    public SavedRecipe? Recipe { get; init; }
}

/// <summary>
/// 処理設定(レシピ)をこの PC の中だけに保存する。
/// 置き場所は %LOCALAPPDATA%\ExcelBatchTool\recipes.json で、外部へは一切送らない。
/// 保存は「同じフォルダーに書いてから置き換える」方式で、途中で落ちても元の内容を壊さない。
/// </summary>
public sealed class RecipeStore
{
    /// <summary>現在の版。読めない版のファイルは勝手に解釈しない。</summary>
    public const int CurrentSchemaVersion = 1;

    public const string FileName = "recipes.json";

    private const string CorruptMessage =
        "保存済みレシピを読み取れません。ファイルが壊れている可能性があります。"
        + "内容を消さないよう、この画面からの保存・更新は行いません。";

    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        // 日本語のレシピ名・項目名をそのまま読めるようにする。
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
        Converters =
        {
            new RecipeTypeConverter(),
            new CellWriteKindConverter(),
            new SourceFileKindConverter(),
            new CsvValueSourceKindConverter(),
            new CsvOutputEncodingConverter(),
            new CsvQuoteModeConverter(),
        },
    };

    public RecipeStore()
        : this(DefaultFilePath())
    {
    }

    /// <summary>置き場所を指定する(テストと、既定以外に置きたいとき用)。</summary>
    public RecipeStore(string filePath)
    {
        FilePath = filePath;
        BackupFilePath = filePath + ".bak";
        TempFilePath = filePath + ".tmp";
    }

    public string FilePath { get; }

    public string BackupFilePath { get; }

    private string TempFilePath { get; }

    /// <summary>既定の置き場所(%LOCALAPPDATA%\ExcelBatchTool\recipes.json)。</summary>
    public static string DefaultFilePath()
        => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ExcelBatchTool",
            FileName);

    /// <summary>保存済みのレシピを読み込む。ファイルが無ければ空として扱う。</summary>
    public RecipeLoadResult Load()
    {
        if (!File.Exists(FilePath))
        {
            return new RecipeLoadResult { IsSuccess = true };
        }

        string text;
        try
        {
            text = File.ReadAllText(FilePath);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return new RecipeLoadResult
            {
                IsSuccess = false,
                Error = $"保存済みレシピを開けません: {ex.Message}",
                HasBackup = File.Exists(BackupFilePath),
            };
        }

        return Parse(text);
    }

    /// <summary>新しいレシピを追加する。同じ名前があるときは上書きせずに知らせる。</summary>
    public RecipeSaveResult Add(SavedRecipe recipe)
    {
        var loaded = Load();
        if (!loaded.IsSuccess)
        {
            return Failed(loaded.Error!);
        }

        if (!RecipeName.TryNormalize(recipe.Name, out var name, out var nameError))
        {
            return Failed(nameError!, loaded.Recipes);
        }

        if (loaded.Recipes.Any(item => RecipeName.AreSame(item.Name, name)))
        {
            return Failed($"同じ名前のレシピがあります:「{name}」", loaded.Recipes);
        }

        var timestamp = RecipeValidation.Timestamp();
        var added = Copy(recipe) with
        {
            Id = string.IsNullOrEmpty(recipe.Id) ? Guid.NewGuid().ToString("d") : recipe.Id,
            Name = name,
            CreatedAt = timestamp,
            UpdatedAt = timestamp,
        };

        if (RecipeValidation.Validate(added) is { } invalid)
        {
            return Failed(invalid, loaded.Recipes);
        }

        return Write([.. loaded.Recipes, added], added);
    }

    /// <summary>既存のレシピを今の設定で置き換える。</summary>
    public RecipeSaveResult Update(string id, SavedRecipe recipe)
    {
        var loaded = Load();
        if (!loaded.IsSuccess)
        {
            return Failed(loaded.Error!);
        }

        var existing = loaded.Recipes.FirstOrDefault(item => item.Id == id);
        if (existing is null)
        {
            return Failed("更新するレシピが見つかりません。一覧を選び直してください。", loaded.Recipes);
        }

        if (!RecipeName.TryNormalize(recipe.Name, out var name, out var nameError))
        {
            return Failed(nameError!, loaded.Recipes);
        }

        if (loaded.Recipes.Any(item => item.Id != id && RecipeName.AreSame(item.Name, name)))
        {
            return Failed($"同じ名前のレシピがあります:「{name}」", loaded.Recipes);
        }

        var updated = Copy(recipe) with
        {
            Id = existing.Id,
            Name = name,
            CreatedAt = existing.CreatedAt,
            UpdatedAt = RecipeValidation.Timestamp(),
        };

        if (RecipeValidation.Validate(updated) is { } invalid)
        {
            return Failed(invalid, loaded.Recipes);
        }

        return Write(
            [.. loaded.Recipes.Select(item => item.Id == id ? updated : item)], updated);
    }

    /// <summary>レシピを 1 件削除する。</summary>
    public RecipeSaveResult Delete(string id)
    {
        var loaded = Load();
        if (!loaded.IsSuccess)
        {
            return Failed(loaded.Error!);
        }

        if (loaded.Recipes.All(item => item.Id != id))
        {
            return Failed("削除するレシピが見つかりません。一覧を選び直してください。", loaded.Recipes);
        }

        return Write([.. loaded.Recipes.Where(item => item.Id != id)], null);
    }

    private RecipeLoadResult Parse(string text)
    {
        RecipeDocument? document;
        try
        {
            document = JsonSerializer.Deserialize<RecipeDocument>(text, Options);
        }
        catch (Exception ex) when (ex is JsonException or NotSupportedException)
        {
            return Corrupt();
        }

        if (document is null)
        {
            return Corrupt();
        }

        if (document.SchemaVersion != CurrentSchemaVersion)
        {
            return new RecipeLoadResult
            {
                IsSuccess = false,
                Error = document.SchemaVersion > CurrentSchemaVersion
                    ? "保存済みレシピは新しい版の形式です。このバージョンでは読み込めません。"
                    : "保存済みレシピの形式が正しくありません。",
                HasBackup = File.Exists(BackupFilePath),
            };
        }

        foreach (var recipe in document.Recipes)
        {
            if (RecipeValidation.Validate(recipe) is { } invalid)
            {
                return new RecipeLoadResult
                {
                    IsSuccess = false,
                    Error = $"保存済みレシピを読み取れません: {invalid}",
                    HasBackup = File.Exists(BackupFilePath),
                };
            }
        }

        var ids = new HashSet<string>(StringComparer.Ordinal);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var recipe in document.Recipes)
        {
            if (!ids.Add(recipe.Id) || !names.Add(recipe.Name.Trim()))
            {
                return new RecipeLoadResult
                {
                    IsSuccess = false,
                    Error = "保存済みレシピに同じ名前のものがあります。",
                    HasBackup = File.Exists(BackupFilePath),
                };
            }
        }

        return new RecipeLoadResult { IsSuccess = true, Recipes = Sort(document.Recipes) };
    }

    private RecipeLoadResult Corrupt() => new()
    {
        IsSuccess = false,
        Error = CorruptMessage,
        HasBackup = File.Exists(BackupFilePath),
    };

    /// <summary>
    /// 一時ファイルへ書いてから置き換える。既存を消してから書く方式は使わない。
    /// 途中で失敗したときは元の recipes.json をそのまま残す。
    /// </summary>
    private RecipeSaveResult Write(IReadOnlyList<SavedRecipe> recipes, SavedRecipe? recipe)
    {
        var sorted = Sort(recipes);
        var document = new RecipeDocument { SchemaVersion = CurrentSchemaVersion, Recipes = sorted };

        try
        {
            if (Path.GetDirectoryName(FilePath) is { Length: > 0 } directory)
            {
                Directory.CreateDirectory(directory);
            }

            using (var stream = new FileStream(
                TempFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                JsonSerializer.Serialize(stream, document, Options);
                stream.Flush(flushToDisk: true);
            }

            if (File.Exists(FilePath))
            {
                // 置き換えと同時に 1 つ前の内容を .bak として残す。
                File.Replace(TempFilePath, FilePath, BackupFilePath, ignoreMetadataErrors: true);
            }
            else
            {
                File.Move(TempFilePath, FilePath);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            DeleteTempQuietly();
            return Failed($"レシピを保存できませんでした: {ex.Message}", Load() is { IsSuccess: true } current
                ? current.Recipes
                : []);
        }

        return new RecipeSaveResult { IsSuccess = true, Recipes = sorted, Recipe = recipe };
    }

    private void DeleteTempQuietly()
    {
        try
        {
            if (File.Exists(TempFilePath))
            {
                File.Delete(TempFilePath);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // 消せなくても元のファイルは無事なので、そのまま進める。
        }
    }

    /// <summary>探しやすいよう名前順に並べる(大文字小文字は区別しない)。</summary>
    private static IReadOnlyList<SavedRecipe> Sort(IEnumerable<SavedRecipe> recipes)
        => [.. recipes.OrderBy(recipe => recipe.Name, StringComparer.OrdinalIgnoreCase)];

    /// <summary>保存対象の中身だけを取り出す(呼び出し側が入れた Id・日時は使わない)。</summary>
    private static SavedRecipe Copy(SavedRecipe recipe) => new()
    {
        Name = recipe.Name,
        Type = recipe.Type,
        CellInputSet = recipe.Type == RecipeType.CellInputSet ? recipe.CellInputSet : null,
        SourceToFixedCells = recipe.Type == RecipeType.SourceToFixedCells
            ? recipe.SourceToFixedCells
            : null,
        SourceTableToTargetTable = recipe.Type == RecipeType.SourceTableToTargetTable
            ? recipe.SourceTableToTargetTable
            : null,
        CsvTransform = recipe.Type == RecipeType.CsvTransform ? recipe.CsvTransform : null,
        PdfRead = recipe.Type == RecipeType.PdfRead ? recipe.PdfRead : null,
    };

    private static RecipeSaveResult Failed(string error, IReadOnlyList<SavedRecipe>? recipes = null)
        => new() { IsSuccess = false, Error = error, Recipes = recipes ?? [] };
}
