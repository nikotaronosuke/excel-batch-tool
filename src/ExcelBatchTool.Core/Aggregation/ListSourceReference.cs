using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>リスト入力規則の参照元を解釈できなかった理由。</summary>
internal enum ListSourceProblem
{
    None = 0,
    Empty,
    BrokenReference,
    ExternalWorkbook,
    ThreeDimensional,
    FormulaLike,
    MissingSheetName,
    NotAbsolute,
    TwoDimensional,
    OutOfRange,
}

/// <summary>リスト入力規則の参照元(シート名と範囲)。</summary>
internal sealed record ListRangeReference(string? SheetName, string Range);

/// <summary>
/// リスト入力規則の参照元として使える「1 行または 1 列の完全絶対 A1 範囲」だけを解釈する。
/// Excel の数式パーサーは自作せず、この限定文法から外れたものは理由付きで拒否する。
/// </summary>
internal static partial class ListSourceParser
{
    /// <summary>$A$2:$A$50 のような完全絶対参照。</summary>
    [GeneratedRegex(@"^\$[A-Za-z]{1,3}\$[0-9]{1,7}(:\$[A-Za-z]{1,3}\$[0-9]{1,7})?$")]
    private static partial Regex AbsoluteRangePattern();

    public static bool TryParse(
        string? text,
        bool requireSheetName,
        out ListRangeReference? result,
        out ListSourceProblem problem)
    {
        result = null;
        problem = ListSourceProblem.None;

        var trimmed = text?.Trim();
        if (string.IsNullOrEmpty(trimmed))
        {
            problem = ListSourceProblem.Empty;
            return false;
        }

        if (trimmed.Contains("#REF!", StringComparison.Ordinal))
        {
            problem = ListSourceProblem.BrokenReference;
            return false;
        }

        // 関数呼び出しを含むものは、範囲参照ではなく数式として扱う。
        if (trimmed.Contains('(') || trimmed.StartsWith('='))
        {
            problem = ListSourceProblem.FormulaLike;
            return false;
        }

        string? sheetName = null;
        var rangeText = trimmed;

        if (SheetReferenceSyntax.HasSheetName(trimmed))
        {
            if (!SheetReferenceSyntax.TrySplit(trimmed, out var parsedSheet, out var parsedRange, out var syntax))
            {
                problem = syntax switch
                {
                    SheetReferenceProblem.ThreeDimensional => ListSourceProblem.ThreeDimensional,
                    _ => ListSourceProblem.FormulaLike,
                };
                return false;
            }

            if (parsedSheet.Contains('[') || parsedSheet.Contains(']'))
            {
                problem = ListSourceProblem.ExternalWorkbook;
                return false;
            }

            sheetName = parsedSheet;
            rangeText = parsedRange;
        }
        else if (requireSheetName)
        {
            problem = ListSourceProblem.MissingSheetName;
            return false;
        }

        if (!AbsoluteRangePattern().IsMatch(rangeText))
        {
            problem = ListSourceProblem.NotAbsolute;
            return false;
        }

        if (!A1RangeValidator.IsValidRange(rangeText))
        {
            problem = ListSourceProblem.OutOfRange;
            return false;
        }

        // 候補一覧は 1 行または 1 列。縦横に広がる範囲は今回扱わない。
        var normalized = rangeText.Replace("$", string.Empty, StringComparison.Ordinal);
        if (!CellRangeParser.TryParseRange(normalized, out var range))
        {
            problem = ListSourceProblem.NotAbsolute;
            return false;
        }

        if (range.FirstColumn != range.LastColumn && range.FirstRow != range.LastRow)
        {
            problem = ListSourceProblem.TwoDimensional;
            return false;
        }

        result = new ListRangeReference(sheetName, rangeText);
        return true;
    }

    /// <summary>利用者向けの理由文にする。</summary>
    public static string Describe(ListSourceProblem problem, string context, string? reference) => problem switch
    {
        ListSourceProblem.BrokenReference => $"{context}の参照が壊れています(#REF!)。",
        ListSourceProblem.ExternalWorkbook => $"{context}が他のブックを参照しているため、"
            + "現在のバージョンでは安全に集約できません。",
        ListSourceProblem.ThreeDimensional => $"{context}が複数シートにまたがって参照しているため、"
            + "現在のバージョンでは安全に集約できません。",
        ListSourceProblem.MissingSheetName => $"{context}にシート名がありません({reference})。",
        ListSourceProblem.TwoDimensional => $"{context}が縦横に広がる範囲({reference})を参照しているため、"
            + "現在のバージョンでは安全に集約できません。",
        ListSourceProblem.OutOfRange => $"{context}の範囲({reference})が Excel の上限を超えています。",
        ListSourceProblem.Empty => $"{context}が空です。",
        _ => $"{context}({reference})は現在のバージョンでは安全に集約できません。",
    };
}

/// <summary>候補一覧として使える、ブック全体に対する名前定義。</summary>
internal sealed record ListSourceDefinedName(string Name, string TargetSheetName, string Range);

/// <summary>
/// Workbook の名前定義から、リスト入力規則の参照元として安全に使えるものだけを引く。
/// 対象は「ブック全体を対象とする、1 行/1 列の完全絶対範囲」だけ。
/// </summary>
internal sealed class WorkbookDefinedNameIndex
{
    private readonly List<DefinedName> _definedNames;

    private WorkbookDefinedNameIndex(List<DefinedName> definedNames) => _definedNames = definedNames;

    public static WorkbookDefinedNameIndex Create(WorkbookPart workbookPart)
        => new(workbookPart.Workbook?.DefinedNames?.Elements<DefinedName>().ToList() ?? []);

    /// <summary>名前が 1 つでもあるか(名前らしき文字列の判定に使う)。</summary>
    public bool HasAnyName => _definedNames.Count > 0;

    /// <summary>
    /// 名前を引く。Excel の名前は大文字小文字を区別しないので、その規則で照合する。
    /// シート固有の名前・複雑な参照・想定外の属性を持つ名前は拒否する。
    /// </summary>
    public bool TryResolve(
        string name,
        out ListSourceDefinedName? resolved,
        out string? error)
    {
        resolved = null;
        error = null;

        var matches = _definedNames
            .Where(item => string.Equals(item.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (matches.Count == 0)
        {
            error = $"名前定義「{name}」が見つかりません。";
            return false;
        }

        if (matches.Any(item => item.LocalSheetId is not null))
        {
            error = $"名前定義「{name}」はシート固有の名前のため、現在のバージョンでは安全に集約できません。";
            return false;
        }

        if (matches.Count > 1)
        {
            error = $"名前定義「{name}」が複数あります。";
            return false;
        }

        var definedName = matches[0];
        if (HasUnsupportedMetadata(definedName))
        {
            error = $"名前定義「{name}」に現在のバージョンでは扱えない設定が付いています。";
            return false;
        }

        if (!ListSourceParser.TryParse(definedName.Text, requireSheetName: true, out var reference, out var problem))
        {
            error = ListSourceParser.Describe(problem, $"名前定義「{name}」", definedName.Text);
            return false;
        }

        resolved = new ListSourceDefinedName(definedName.Name!.Value!, reference!.SheetName!, reference.Range);
        return true;
    }

    /// <summary>普通のユーザー作成の範囲名以外(関数名・マクロ・公開設定など)を弾く。</summary>
    private static bool HasUnsupportedMetadata(DefinedName definedName)
        => definedName.Hidden is not null
            || definedName.Function is not null
            || definedName.VbProcedure is not null
            || definedName.Xlm is not null
            || definedName.FunctionGroupId is not null
            || definedName.ShortcutKey is not null
            || definedName.PublishToServer is not null
            || definedName.WorkbookParameter is not null
            || definedName.CustomMenu is not null
            || definedName.Description is not null
            || definedName.Help is not null
            || definedName.StatusBar is not null
            || definedName.Comment is not null;
}
