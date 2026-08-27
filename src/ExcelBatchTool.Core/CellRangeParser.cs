namespace ExcelBatchTool.Core;

/// <summary>A1 形式のセル参照・範囲参照のパーサー。</summary>
public static class CellRangeParser
{
    public readonly record struct CellRange(int FirstColumn, int FirstRow, int LastColumn, int LastRow);

    /// <summary>"A1:F120" または "A1" 形式の範囲をパースする(列・行とも 1 始まり)。</summary>
    public static bool TryParseRange(string reference, out CellRange range)
    {
        range = default;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return false;
        }

        var parts = reference.Split(':');
        if (parts.Length is not (1 or 2))
        {
            return false;
        }

        if (!TryParseCell(parts[0], out var firstColumn, out var firstRow))
        {
            return false;
        }

        var lastColumn = firstColumn;
        var lastRow = firstRow;
        if (parts.Length == 2 && !TryParseCell(parts[1], out lastColumn, out lastRow))
        {
            return false;
        }

        range = new CellRange(
            Math.Min(firstColumn, lastColumn),
            Math.Min(firstRow, lastRow),
            Math.Max(firstColumn, lastColumn),
            Math.Max(firstRow, lastRow));
        return true;
    }

    /// <summary>"F120" のようなセル参照をパースする(列・行とも 1 始まり)。</summary>
    public static bool TryParseCell(string reference, out int column, out int row)
    {
        column = 0;
        row = 0;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return false;
        }

        var index = 0;
        while (index < reference.Length && char.IsAsciiLetter(reference[index]))
        {
            column = column * 26 + (char.ToUpperInvariant(reference[index]) - 'A' + 1);
            index++;
        }

        if (index == 0 || index == reference.Length)
        {
            return false;
        }

        return int.TryParse(reference[index..], out row) && row > 0 && column > 0;
    }

    /// <summary>1 始まりの列番号を "A" / "AB" 形式へ変換する。</summary>
    public static string ColumnIndexToLetters(int columnIndex)
    {
        if (columnIndex < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(columnIndex));
        }

        var letters = string.Empty;
        while (columnIndex > 0)
        {
            columnIndex--;
            letters = (char)('A' + columnIndex % 26) + letters;
            columnIndex /= 26;
        }

        return letters;
    }
}
