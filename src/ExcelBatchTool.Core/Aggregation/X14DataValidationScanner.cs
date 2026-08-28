using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>
/// 走査で取り出した 1 件の x14 リスト入力規則。
/// 参照先シートの解決(出力シート名への書き換え)は Planner が行う。
/// </summary>
internal sealed record X14ListValidationInfo
{
    /// <summary>元の要素(参照元はまだ書き換えていない)。Block 時は null。</summary>
    public X14.DataValidation? Element { get; init; }

    /// <summary>候補一覧の範囲(シート名を除いた部分)。</summary>
    public string Range { get; init; } = string.Empty;

    /// <summary>候補一覧が置かれている元シート名。シート名なしの参照では null。</summary>
    public string? TargetSheetName { get; init; }

    /// <summary>候補一覧が名前定義で指定されている場合のその名前。</summary>
    public string? DefinedName { get; init; }

    public string Sqref { get; init; } = string.Empty;

    public string? BlockReason { get; init; }
}

/// <summary>
/// Office 2010 以降の拡張入力規則(x14)のうち、type="list" だけを扱う。
/// 構造は characterization で確認した形(x14:formula1 &gt; xm:f、xm:sqref)に限定する。
/// </summary>
internal static class X14DataValidationScanner
{
    /// <summary>Office 2010 の入力規則拡張の URI。</summary>
    public const string ExtensionUri = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";

    /// <summary>リビジョン識別子 (xr:uid) の名前空間。出力では付けない。</summary>
    public const string RevisionNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";

    /// <summary>x14 入力規則 1 件を解析する。</summary>
    public static X14ListValidationInfo Scan(
        X14.DataValidation validation,
        WorkbookDefinedNameIndex definedNames)
    {
        var sqref = validation.ReferenceSequence?.Text ?? string.Empty;

        var type = validation.Type?.InnerText ?? "none";
        if (!string.Equals(type, "list", StringComparison.Ordinal))
        {
            return Blocked(sqref,
                $"セル {sqref} の新しい形式の入力規則(種類「{type}」)は、"
                    + "現在のバージョンでは安全に集約できません。");
        }

        // xr:uid 以外の想定外の属性・子要素は意味を保証できない。
        foreach (var attribute in validation.ExtendedAttributes)
        {
            if (!string.Equals(attribute.NamespaceUri, RevisionNamespace, StringComparison.Ordinal))
            {
                return Blocked(sqref, $"セル {sqref} の入力規則に対応していない設定が含まれています。");
            }
        }

        foreach (var child in validation.ChildElements)
        {
            if (child is not (X14.DataValidationForumla1 or X14.DataValidationForumla2 or Xm.ReferenceSequence))
            {
                return Blocked(sqref, $"セル {sqref} の入力規則に対応していない内容が含まれています。");
            }
        }

        if (validation.DataValidationForumla2 is not null)
        {
            return Blocked(sqref, $"セル {sqref} のリスト入力規則の構造が想定と異なります。");
        }

        if (validation.Operator is not null)
        {
            return Blocked(sqref, $"セル {sqref} の入力規則に、意味を持たない条件設定が含まれています。");
        }

        if (!A1RangeValidator.IsValidRangeList(sqref, out var invalidToken))
        {
            return Blocked(sqref, $"入力規則の適用範囲「{invalidToken}」を解釈できません。");
        }

        var formulas = validation.DataValidationForumla1?.Elements<Xm.Formula>().ToList() ?? [];
        if (validation.DataValidationForumla1 is null || formulas.Count != 1)
        {
            return Blocked(sqref, $"セル {sqref} のリスト入力規則の参照元を解釈できません。");
        }

        var source = formulas[0].Text?.Trim();
        if (string.IsNullOrEmpty(source))
        {
            return Blocked(sqref, $"セル {sqref} のリスト入力規則に選択肢がありません。");
        }

        // 直接指定("A,B,C")は標準形式と同じ扱いで、そのまま保持する。
        if (source.StartsWith('"'))
        {
            return source.Length >= 2 && source.EndsWith('"') && !source[1..^1].Contains('"')
                ? Supported(validation, sqref, range: source)
                : Blocked(sqref, $"セル {sqref} のリスト入力規則の選択肢を解釈できません。");
        }

        // 範囲参照。シート名付きなら参照先シートの解決を Planner に任せる。
        if (ListSourceParser.TryParse(source, requireSheetName: false, out var reference, out var problem))
        {
            return Supported(validation, sqref, reference!.Range, reference.SheetName);
        }

        // 範囲として読めない場合は、名前定義として引けるか試す。
        if (LooksLikeName(source) && definedNames.HasAnyName)
        {
            if (definedNames.TryResolve(source, out var resolvedName, out var nameError))
            {
                return new X14ListValidationInfo
                {
                    Element = (X14.DataValidation)validation.CloneNode(true),
                    Range = resolvedName!.Range,
                    TargetSheetName = resolvedName.TargetSheetName,
                    DefinedName = resolvedName.Name,
                    Sqref = sqref,
                };
            }

            return Blocked(sqref, $"セル {sqref} の入力規則: {nameError}");
        }

        return Blocked(sqref, ListSourceParser.Describe(problem, $"セル {sqref} の入力規則の参照元", source));
    }

    /// <summary>名前定義として引けそうな形か(関数・演算子・参照記号を含まない)。</summary>
    public static bool LooksLikeName(string text)
    {
        if (text.Length == 0 || char.IsDigit(text[0]))
        {
            return false;
        }

        foreach (var c in text)
        {
            if (char.IsLetterOrDigit(c) || c is '_' or '.' or '\\')
            {
                continue;
            }

            return false;
        }

        return true;
    }

    private static X14ListValidationInfo Supported(
        X14.DataValidation validation,
        string sqref,
        string range,
        string? targetSheetName = null)
        => new()
        {
            Element = (X14.DataValidation)validation.CloneNode(true),
            Range = range,
            TargetSheetName = targetSheetName,
            Sqref = sqref,
        };

    private static X14ListValidationInfo Blocked(string sqref, string reason)
        => new() { BlockReason = reason, Sqref = sqref };

    /// <summary>
    /// 出力用に x14 入力規則を作り直す。参照先シートの書き換えと、
    /// リビジョン識別子 (xr:uid) の除去を行う。
    /// </summary>
    public static X14.DataValidation BuildOutputElement(X14.DataValidation source, string listSource)
    {
        // clone すると元ブックの名前空間宣言(xr など)まで付いてくるので、
        // 対応することが分かっている項目だけを写して新しく組み立てる。
        // これにより、リビジョン識別子 (xr:uid) も持ち込まれない。
        var element = new X14.DataValidation
        {
            Type = source.Type is null ? null : new EnumValue<DataValidationValues>(source.Type),
            AllowBlank = source.AllowBlank,
            ShowDropDown = source.ShowDropDown,
            ShowInputMessage = source.ShowInputMessage,
            ShowErrorMessage = source.ShowErrorMessage,
            ErrorStyle = source.ErrorStyle is null
                ? null
                : new EnumValue<DataValidationErrorStyleValues>(source.ErrorStyle),
            ImeMode = source.ImeMode is null
                ? null
                : new EnumValue<DataValidationImeModeValues>(source.ImeMode),
            ErrorTitle = source.ErrorTitle,
            Error = source.Error,
            PromptTitle = source.PromptTitle,
            Prompt = source.Prompt,
        };

        element.Append(new X14.DataValidationForumla1(new Xm.Formula(listSource)));
        element.Append(new Xm.ReferenceSequence(source.ReferenceSequence?.Text ?? string.Empty));

        return element;
    }

    /// <summary>x14 入力規則を収める extLst を組み立てる。</summary>
    public static WorksheetExtensionList BuildExtensionList(IReadOnlyList<X14.DataValidation> validations)
    {
        var container = new X14.DataValidations { Count = (uint)validations.Count };
        foreach (var validation in validations)
        {
            container.Append(validation.CloneNode(true));
        }

        var extension = new WorksheetExtension(container) { Uri = ExtensionUri };
        extension.AddNamespaceDeclaration("x14", DataValidationScanner.X14Namespace);

        return new WorksheetExtensionList(extension);
    }
}
