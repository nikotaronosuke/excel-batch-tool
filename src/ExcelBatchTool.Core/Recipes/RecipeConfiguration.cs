using ExcelBatchTool.Core.CsvTransform;

namespace ExcelBatchTool.Core.Recipes;

/// <summary>
/// 2 つのレシピの「処理ルール」が同じかどうかを調べる。
/// 名前・識別子・日時は比べない(実行したときにはまだ名前が無いこともあるため)。
/// ファイルの場所や名前はそもそもレシピに入っていないので、比較対象にもならない。
/// </summary>
public static class RecipeConfiguration
{
    /// <summary>種類と中身が完全に同じか。文字は完全一致で比べる。</summary>
    public static bool AreSame(SavedRecipe? left, SavedRecipe? right)
    {
        if (left is null || right is null)
        {
            return false;
        }

        if (left.Type != right.Type)
        {
            return false;
        }

        return left.Type switch
        {
            RecipeType.CellInputSet => AreSame(left.CellInputSet, right.CellInputSet),
            RecipeType.SourceToFixedCells => AreSame(left.SourceToFixedCells, right.SourceToFixedCells),
            RecipeType.SourceTableToTargetTable
                => AreSame(left.SourceTableToTargetTable, right.SourceTableToTargetTable),
            _ => AreSame(left.CsvTransform, right.CsvTransform),
        };
    }

    private static bool AreSame(CellInputSetRecipe? left, CellInputSetRecipe? right)
        => left is not null
            && right is not null
            && Same(left.OutputSuffix, right.OutputSuffix)
            && SameList(
                left.Operations,
                right.Operations,
                (a, b) => Same(a.Cell, b.Cell) && a.Kind == b.Kind && Same(a.Value, b.Value));

    private static bool AreSame(SourceToFixedCellsRecipe? left, SourceToFixedCellsRecipe? right)
        => left is not null
            && right is not null
            && left.SourceFileKind == right.SourceFileKind
            && Same(left.SourceSheetName, right.SourceSheetName)
            && left.HeaderRow == right.HeaderRow
            && Same(left.SourceKeyColumn, right.SourceKeyColumn)
            && Same(left.TargetKeyCell, right.TargetKeyCell)
            && Same(left.OutputSuffix, right.OutputSuffix)
            && SameList(
                left.Mappings,
                right.Mappings,
                (a, b) => Same(a.SourceColumn, b.SourceColumn)
                    && Same(a.TargetCell, b.TargetCell)
                    && a.Kind == b.Kind);

    private static bool AreSame(
        SourceTableToTargetTableRecipe? left, SourceTableToTargetTableRecipe? right)
        => left is not null
            && right is not null
            && left.SourceFileKind == right.SourceFileKind
            && Same(left.SourceSheetName, right.SourceSheetName)
            && left.SourceHeaderRow == right.SourceHeaderRow
            && Same(left.SourceKeyColumn, right.SourceKeyColumn)
            && left.TargetHeaderRow == right.TargetHeaderRow
            && Same(left.TargetKeyColumn, right.TargetKeyColumn)
            && Same(left.OutputSuffix, right.OutputSuffix)
            && SameList(
                left.Mappings,
                right.Mappings,
                (a, b) => Same(a.SourceColumn, b.SourceColumn)
                    && Same(a.TargetColumn, b.TargetColumn)
                    && a.Kind == b.Kind);

    private static bool AreSame(CsvTransformRecipe? left, CsvTransformRecipe? right)
        => left is not null
            && right is not null
            && left.SourceFileKind == right.SourceFileKind
            && Same(left.SourceSheetName, right.SourceSheetName)
            && left.HeaderRow == right.HeaderRow
            && left.Encoding == right.Encoding
            && left.QuoteMode == right.QuoteMode
            && Same(left.OutputSuffix, right.OutputSuffix)
            && SameList(
                left.OutputColumns,
                right.OutputColumns,
                (a, b) => Same(a.OutputName, b.OutputName)
                    && a.ValueSourceKind == b.ValueSourceKind
                    && Same(a.SourceColumn, b.SourceColumn)
                    && Same(a.FixedValue, b.FixedValue));

    /// <summary>並び順も含めて同じか。</summary>
    private static bool SameList<T>(
        IReadOnlyList<T> left, IReadOnlyList<T> right, Func<T, T, bool> same)
    {
        if (left.Count != right.Count)
        {
            return false;
        }

        for (var index = 0; index < left.Count; index++)
        {
            if (!same(left[index], right[index]))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>文字は完全一致で比べる(前後の空白・大文字小文字も区別する)。</summary>
    private static bool Same(string? left, string? right)
        => string.Equals(left, right, StringComparison.Ordinal);
}
