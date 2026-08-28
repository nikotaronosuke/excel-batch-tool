namespace ExcelBatchTool.Core.Aggregation;

/// <summary>
/// 出力 Workbook のシート名を決める。Excel のシート名制約(31 文字・使用不可文字・
/// Workbook 内で重複不可)に合わせる。利用者が入力した名前は勝手に置き換えず、
/// 問題があれば理由を返して Block する。
/// </summary>
public static class OutputSheetNameResolver
{
    /// <summary>Excel のシート名の最大文字数。</summary>
    public const int MaxLength = 31;

    private static readonly char[] InvalidCharacters = [':', '\\', '/', '?', '*', '[', ']'];

    /// <summary>Excel が予約しているシート名。</summary>
    private const string ReservedName = "History";

    /// <summary>シート名として使えるか調べる。使えない場合は理由を返す。</summary>
    public static string? Validate(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return "出力シート名が空です。";
        }

        if (name.Length > MaxLength)
        {
            return $"出力シート名は {MaxLength} 文字までです(現在 {name.Length} 文字)。";
        }

        if (name.IndexOfAny(InvalidCharacters) >= 0)
        {
            return @"出力シート名に使えない文字が含まれています( : \ / ? * [ ] )。";
        }

        if (name.StartsWith('\'') || name.EndsWith('\''))
        {
            return "出力シート名の先頭と末尾にアポストロフィ(')は使えません。";
        }

        if (string.Equals(name, ReservedName, StringComparison.OrdinalIgnoreCase))
        {
            return $"「{ReservedName}」は Excel が予約しているシート名です。";
        }

        return null;
    }

    /// <summary>
    /// 元シート名から、まだ使われていない出力シート名を決定的に作る。
    /// 重複する場合は「名前 (2)」「名前 (3)」…とし、31 文字を超えないよう元の名前側を短くする。
    /// </summary>
    public static string Propose(string sourceSheetName, IEnumerable<string> usedNames)
    {
        var used = new HashSet<string>(usedNames, StringComparer.OrdinalIgnoreCase);
        var baseName = Truncate(SanitizeForProposal(sourceSheetName), MaxLength);

        if (baseName.Length == 0)
        {
            baseName = "Sheet";
        }

        if (!used.Contains(baseName))
        {
            return baseName;
        }

        for (var number = 2; number < 10_000; number++)
        {
            var suffix = $" ({number})";
            var candidate = Truncate(baseName, MaxLength - suffix.Length) + suffix;
            if (!used.Contains(candidate))
            {
                return candidate;
            }
        }

        // 実務上ここには到達しない。到達した場合は検証側で重複として Block される。
        return baseName;
    }

    /// <summary>
    /// 元シート名から提案を作るときだけ使う整形。使用不可文字を「_」に置き換える。
    /// (Excel が作った .xlsx のシート名は本来この文字を含まないが、外部ツール製の
    /// ファイルに備えた保険。利用者が入力した名前にはこの整形を適用しない。)
    /// </summary>
    private static string SanitizeForProposal(string name)
    {
        var trimmed = name.Trim().Trim('\'');
        var builder = new System.Text.StringBuilder(trimmed.Length);
        foreach (var c in trimmed)
        {
            builder.Append(Array.IndexOf(InvalidCharacters, c) >= 0 ? '_' : c);
        }

        return builder.ToString();
    }

    private static string Truncate(string value, int maxLength)
        => value.Length <= maxLength ? value : value[..Math.Max(0, maxLength)];
}
