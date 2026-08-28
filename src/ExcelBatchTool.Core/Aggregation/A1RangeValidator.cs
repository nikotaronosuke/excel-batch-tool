namespace ExcelBatchTool.Core.Aggregation;

/// <summary>
/// A1 形式のセル・範囲参照が Excel のワークシート上限に収まっているかを確かめる。
/// 参照の解釈自体は既存の <see cref="CellRangeParser"/> を使い、上限判定だけをここに集約する。
/// </summary>
internal static class A1RangeValidator
{
    /// <summary>Excel の最大列(XFD)。</summary>
    public const int MaxColumn = 16_384;

    /// <summary>Excel の最大行。</summary>
    public const int MaxRow = 1_048_576;

    /// <summary>A1 / A1:B5 形式で、かつシートの上限内か($ 付きも許容)。</summary>
    public static bool IsValidRange(string? reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            return false;
        }

        if (reference.Contains("#REF!", StringComparison.Ordinal))
        {
            return false;
        }

        var normalized = reference.Replace("$", string.Empty, StringComparison.Ordinal);
        return CellRangeParser.TryParseRange(normalized, out var range)
            && range.LastColumn <= MaxColumn
            && range.LastRow <= MaxRow;
    }

    /// <summary>空白区切りの範囲リスト(sqref など)がすべて有効か。</summary>
    public static bool IsValidRangeList(string? references, out string? invalidToken)
    {
        invalidToken = null;

        if (string.IsNullOrWhiteSpace(references))
        {
            invalidToken = string.Empty;
            return false;
        }

        var tokens = references.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length == 0)
        {
            invalidToken = string.Empty;
            return false;
        }

        foreach (var token in tokens)
        {
            if (!IsValidRange(token))
            {
                invalidToken = token;
                return false;
            }
        }

        return true;
    }
}
