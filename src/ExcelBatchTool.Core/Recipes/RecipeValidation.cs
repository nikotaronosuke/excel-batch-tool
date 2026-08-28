using System.Globalization;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Recipes;

/// <summary>レシピ名の決まり(前後の空白を落として 1〜60 文字、改行・制御文字なし)。</summary>
public static class RecipeName
{
    public const int MaxLength = 60;

    /// <summary>入力されたレシピ名を整える。使えない名前は理由を返す。</summary>
    public static bool TryNormalize(string? name, out string normalized, out string? error)
    {
        normalized = (name ?? string.Empty).Trim();

        if (normalized.Length == 0)
        {
            error = "レシピの名前を入力してください。";
            return false;
        }

        if (normalized.Length > MaxLength)
        {
            error = $"レシピの名前は {MaxLength} 文字以内にしてください。";
            return false;
        }

        // 前後の空白は落としてから見る(貼り付けの末尾改行は許し、途中の改行は許さない)。
        if (normalized.Any(char.IsControl))
        {
            error = "レシピの名前に改行や特殊な文字は使えません。";
            return false;
        }

        error = null;
        return true;
    }

    /// <summary>同じ名前かどうか(大文字小文字は区別しない)。</summary>
    public static bool AreSame(string left, string right)
        => string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
}

/// <summary>読み込んだレシピの中身が使える形かどうかを調べる。</summary>
internal static class RecipeValidation
{
    /// <summary>1 件のレシピを検証する。問題があれば理由を返す。</summary>
    public static string? Validate(SavedRecipe recipe)
    {
        if (string.IsNullOrWhiteSpace(recipe.Id))
        {
            return "識別子のないレシピがあります。";
        }

        if (!RecipeName.TryNormalize(recipe.Name, out _, out var nameError))
        {
            return nameError;
        }

        return recipe.Type switch
        {
            RecipeType.CellInputSet => ValidatePayload(recipe.Name, recipe.CellInputSet, ValidateCellInputSet),
            RecipeType.SourceToFixedCells
                => ValidatePayload(recipe.Name, recipe.SourceToFixedCells, ValidateSourceToFixedCells),
            _ => ValidatePayload(
                recipe.Name, recipe.SourceTableToTargetTable, ValidateSourceTableToTargetTable),
        };
    }

    private static string? ValidatePayload<T>(string name, T? payload, Func<T, string?> validate)
        where T : class
        => payload is null ? $"「{name}」の設定が入っていません。" : validate(payload);

    private static string? ValidateCellInputSet(CellInputSetRecipe recipe)
    {
        if (recipe.Operations.Count == 0)
        {
            return "変更するセルが 1 つも入っていないレシピがあります。";
        }

        foreach (var operation in recipe.Operations)
        {
            if (string.IsNullOrWhiteSpace(operation.Cell))
            {
                return "セルの位置が空のレシピがあります。";
            }

            if (operation.Kind == CellWriteKind.Blank && !string.IsNullOrEmpty(operation.Value))
            {
                return "「空欄」なのに値が入っているレシピがあります。";
            }
        }

        return null;
    }

    private static string? ValidateSourceToFixedCells(SourceToFixedCellsRecipe recipe)
    {
        if (recipe.HeaderRow < 1)
        {
            return "項目名の行が正しくないレシピがあります。";
        }

        if (string.IsNullOrWhiteSpace(recipe.SourceKeyColumn))
        {
            return "データ元のキーが空のレシピがあります。";
        }

        if (string.IsNullOrWhiteSpace(recipe.TargetKeyCell))
        {
            return "キーのセルが空のレシピがあります。";
        }

        return ValidateMappings(
            recipe.Mappings.Count,
            recipe.Mappings.Select(mapping => (mapping.SourceColumn, mapping.TargetCell, mapping.Kind)));
    }

    private static string? ValidateSourceTableToTargetTable(SourceTableToTargetTableRecipe recipe)
    {
        if (recipe.SourceHeaderRow < 1 || recipe.TargetHeaderRow < 1)
        {
            return "項目名の行が正しくないレシピがあります。";
        }

        if (string.IsNullOrWhiteSpace(recipe.SourceKeyColumn)
            || string.IsNullOrWhiteSpace(recipe.TargetKeyColumn))
        {
            return "キーの項目が空のレシピがあります。";
        }

        return ValidateMappings(
            recipe.Mappings.Count,
            recipe.Mappings.Select(mapping => (mapping.SourceColumn, mapping.TargetColumn, mapping.Kind)));
    }

    private static string? ValidateMappings(
        int count, IEnumerable<(string Source, string Target, CellWriteKind Kind)> mappings)
    {
        if (count == 0)
        {
            return "対応付けが 1 つも入っていないレシピがあります。";
        }

        foreach (var (source, target, kind) in mappings)
        {
            if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(target))
            {
                return "対応付けの項目が空のレシピがあります。";
            }

            if (kind == CellWriteKind.Blank)
            {
                // データ元から値を供給するので「空欄」は使わない。
                return "データ元からの転記に「空欄」は使えません。";
            }
        }

        return null;
    }

    /// <summary>現在時刻を ISO 8601 で記録する(表示用。処理の判断には使わない)。</summary>
    public static string Timestamp()
        => DateTimeOffset.Now.ToString("o", CultureInfo.InvariantCulture);
}
