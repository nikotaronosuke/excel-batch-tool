using System.Text;
using System.Text.RegularExpressions;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>印刷範囲・印刷タイトルのどちらか。</summary>
internal enum PrintDefinedNameKind
{
    PrintArea,
    PrintTitles,
}

/// <summary>
/// 印刷範囲 (_xlnm.Print_Area) と印刷タイトル (_xlnm.Print_Titles) の参照文字列を扱う。
/// Excel の数式全般は解釈せず、この 2 つで実際に使われる限定文法だけを決定的に扱う。
/// 解釈できないものは書き換えず、呼び出し側が Block できるように理由を返す。
/// </summary>
internal static partial class PrintDefinedNameParser
{
    public const string PrintAreaName = "_xlnm.Print_Area";
    public const string PrintTitlesName = "_xlnm.Print_Titles";

    /// <summary>$A$1 または $A$1:$F$100(印刷範囲)。</summary>
    [GeneratedRegex(@"^\$[A-Za-z]{1,3}\$[0-9]{1,7}(:\$[A-Za-z]{1,3}\$[0-9]{1,7})?$")]
    private static partial Regex AreaPattern();

    /// <summary>$1:$3(印刷タイトルの行範囲)。</summary>
    [GeneratedRegex(@"^\$[0-9]{1,7}:\$[0-9]{1,7}$")]
    private static partial Regex RowRangePattern();

    /// <summary>$A:$C(印刷タイトルの列範囲)。</summary>
    [GeneratedRegex(@"^\$[A-Za-z]{1,3}:\$[A-Za-z]{1,3}$")]
    private static partial Regex ColumnRangePattern();

    public static string NameOf(PrintDefinedNameKind kind)
        => kind == PrintDefinedNameKind.PrintArea ? PrintAreaName : PrintTitlesName;

    public static string DisplayNameOf(PrintDefinedNameKind kind)
        => kind == PrintDefinedNameKind.PrintArea ? "印刷範囲" : "印刷タイトル";

    /// <summary>
    /// 参照文字列を解析し、対象シート自身への絶対参照だけで構成されていれば、
    /// シート名を除いた範囲部分の一覧を返す。
    /// </summary>
    public static bool TryParse(
        string? reference,
        string expectedSheetName,
        PrintDefinedNameKind kind,
        out IReadOnlyList<string> ranges,
        out string? error)
    {
        ranges = Array.Empty<string>();
        error = null;

        if (string.IsNullOrWhiteSpace(reference))
        {
            error = $"{DisplayNameOf(kind)}の参照が空です。";
            return false;
        }

        if (reference.Contains("#REF!", StringComparison.Ordinal))
        {
            error = $"{DisplayNameOf(kind)}の参照が壊れています(#REF!)。";
            return false;
        }

        var parsed = new List<string>();
        foreach (var part in SplitTopLevel(reference))
        {
            var text = part.Trim();
            if (text.Length == 0)
            {
                error = $"{DisplayNameOf(kind)}の参照を解釈できません({reference})。";
                return false;
            }

            if (!TrySplitSheetAndRange(text, kind, out var sheetName, out var range, out error))
            {
                return false;
            }

            if (sheetName.Contains('[') || sheetName.Contains(']'))
            {
                error = $"{DisplayNameOf(kind)}が他のブックを参照しているため、現在のバージョンでは集約できません。";
                return false;
            }

            if (!string.Equals(sheetName, expectedSheetName, StringComparison.Ordinal))
            {
                error = $"{DisplayNameOf(kind)}が他のシート「{sheetName}」を参照しているため、"
                    + "現在のバージョンでは集約できません。";
                return false;
            }

            if (!IsSupportedRange(range, kind))
            {
                error = $"{DisplayNameOf(kind)}に対応していない参照形式があります({range})。";
                return false;
            }

            parsed.Add(range);
        }

        if (parsed.Count == 0)
        {
            error = $"{DisplayNameOf(kind)}の参照を解釈できません({reference})。";
            return false;
        }

        ranges = parsed;
        return true;
    }

    /// <summary>出力シート名を使って参照文字列を作り直す。シート名は常に引用符で囲む。</summary>
    public static string Format(string outputSheetName, IEnumerable<string> ranges)
        => string.Join(",", ranges.Select(range => $"{QuoteSheetName(outputSheetName)}!{range}"));

    /// <summary>シート名を参照内で使える形にする(共通処理へ委譲)。</summary>
    public static string QuoteSheetName(string sheetName) => SheetReferenceSyntax.Quote(sheetName);

    /// <summary>引用符の外側にあるカンマだけで区切る。</summary>
    private static List<string> SplitTopLevel(string reference)
    {
        var parts = new List<string>();
        var builder = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < reference.Length; i++)
        {
            var c = reference[i];
            if (c == '\'')
            {
                // '' は引用符内のエスケープ。
                if (inQuotes && i + 1 < reference.Length && reference[i + 1] == '\'')
                {
                    builder.Append("''");
                    i++;
                    continue;
                }

                inQuotes = !inQuotes;
                builder.Append(c);
                continue;
            }

            if (c == ',' && !inQuotes)
            {
                parts.Add(builder.ToString());
                builder.Clear();
                continue;
            }

            builder.Append(c);
        }

        parts.Add(builder.ToString());
        return parts;
    }

    private static bool TrySplitSheetAndRange(
        string text,
        PrintDefinedNameKind kind,
        out string sheetName,
        out string range,
        out string? error)
    {
        error = null;

        if (SheetReferenceSyntax.TrySplit(text, out sheetName, out range, out var problem))
        {
            return true;
        }

        error = problem switch
        {
            SheetReferenceProblem.ThreeDimensional =>
                $"{DisplayNameOf(kind)}が複数シートにまたがって参照しているため、"
                    + "現在のバージョンでは集約できません。",
            _ when !SheetReferenceSyntax.HasSheetName(text) =>
                $"{DisplayNameOf(kind)}の参照にシート名がありません({text})。",
            _ => $"{DisplayNameOf(kind)}の参照を解釈できません({text})。",
        };

        return false;
    }

    private static bool IsSupportedRange(string range, PrintDefinedNameKind kind) => kind switch
    {
        PrintDefinedNameKind.PrintArea => AreaPattern().IsMatch(range),
        _ => RowRangePattern().IsMatch(range) || ColumnRangePattern().IsMatch(range),
    };
}
