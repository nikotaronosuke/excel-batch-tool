using ExcelBatchTool.Core.Aggregation;

namespace ExcelBatchTool.Core.Mutation;

/// <summary>
/// 一括変更で指定できるセル位置。A1 形式の単一セルのみを受け付ける。
/// 範囲(A1:B5)・シート名付き・名前定義は今回扱わない。
/// </summary>
internal readonly record struct TargetCellAddress(string Reference, int Column, int Row)
{
    /// <summary>利用者の入力を解釈する。解釈できない場合は理由を返す。</summary>
    public static bool TryParse(string? input, out TargetCellAddress address, out string? error)
    {
        address = default;
        error = null;

        var text = input?.Trim();
        if (string.IsNullOrEmpty(text))
        {
            error = "変更するセルの位置を入力してください(例: B2)。";
            return false;
        }

        if (text.Contains(':', StringComparison.Ordinal))
        {
            error = "セルの範囲(例: A1:B5)は指定できません。1 つのセルだけを指定してください。";
            return false;
        }

        if (text.Contains('!', StringComparison.Ordinal))
        {
            error = "シート名付きの指定はできません。セルの位置だけを入力してください(例: B2)。";
            return false;
        }

        // $B$2 のような絶対参照も受け取り、位置として同じものとして扱う。
        var normalized = text.Replace("$", string.Empty, StringComparison.Ordinal).ToUpperInvariant();

        if (!CellRangeParser.TryParseCell(normalized, out var column, out var row))
        {
            error = $"セルの位置「{text}」を解釈できません(例: B2)。";
            return false;
        }

        if (column > A1RangeValidator.MaxColumn || row > A1RangeValidator.MaxRow)
        {
            error = $"セルの位置「{text}」は Excel の範囲を超えています"
                + $"(最大 {CellRangeParser.ColumnIndexToLetters(A1RangeValidator.MaxColumn)}"
                + $"{A1RangeValidator.MaxRow})。";
            return false;
        }

        address = new TargetCellAddress(normalized, column, row);
        return true;
    }
}
