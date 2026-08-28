using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>
/// 走査で取り出した 1 件の入力規則。安全に移植できる場合だけ
/// <see cref="Element"/> に出力用の要素が入る。
/// </summary>
internal sealed record DataValidationInfo
{
    /// <summary>出力へ書き込む要素(必要な変換を済ませたもの)。Block 時は null。</summary>
    public DataValidation? Element { get; init; }

    /// <summary>安全に移植できない理由。null なら移植できる。</summary>
    public string? BlockReason { get; init; }

    public string Sqref { get; init; } = string.Empty;
}

/// <summary>
/// 標準の入力規則(x:dataValidation)のうち、意味を決定的に維持できるものだけを扱う。
/// Excel の数式パーサーは自作せず、対応する限定形式以外は理由を付けて Block する。
/// </summary>
internal static partial class DataValidationScanner
{
    /// <summary>Office 2010 以降の拡張入力規則の名前空間。</summary>
    public const string X14Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    /// <summary>Excel の入力規則リスト(直接指定)の文字数上限。</summary>
    private const int MaxListLiteralLength = 255;

    private static readonly string[] SupportedTypes =
        ["none", "list", "whole", "decimal", "date", "time", "textLength"];

    private static readonly string[] KnownOperators =
    [
        "between", "notBetween", "equal", "notEqual",
        "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual",
    ];

    private static readonly string[] RangeOperators = ["between", "notBetween"];

    /// <summary>$B$2:$B$10 のような、シート名を含まない完全絶対参照。</summary>
    [GeneratedRegex(@"^\$[A-Za-z]{1,3}\$[0-9]{1,7}(:\$[A-Za-z]{1,3}\$[0-9]{1,7})?$")]
    private static partial Regex AbsoluteRangePattern();

    /// <summary>入力規則 1 件を解析する。</summary>
    public static DataValidationInfo Scan(DataValidation validation, bool sourceIsDate1904)
    {
        var sqref = validation.SequenceOfReferences?.InnerText ?? string.Empty;

        if (validation.ExtendedAttributes.Any())
        {
            return Blocked(sqref, "対応していない設定を含む入力規則があります。");
        }

        foreach (var child in validation.ChildElements)
        {
            if (child is not (Formula1 or Formula2))
            {
                return Blocked(sqref, "対応していない内容を含む入力規則があります。");
            }
        }

        if (!A1RangeValidator.IsValidRangeList(sqref, out var invalidToken))
        {
            return Blocked(sqref, $"入力規則の適用範囲「{invalidToken}」を解釈できません。");
        }

        var type = validation.Type?.InnerText ?? "none";
        if (string.Equals(type, "custom", StringComparison.Ordinal))
        {
            return Blocked(sqref, $"セル {sqref} の入力規則はユーザー設定の数式を使っているため、"
                + "現在のバージョンでは安全に集約できません。");
        }

        if (!SupportedTypes.Contains(type, StringComparer.Ordinal))
        {
            return Blocked(sqref, $"セル {sqref} の入力規則の種類「{type}」には対応していません。");
        }

        var hasOperator = validation.Operator is not null;
        var op = validation.Operator?.InnerText ?? "between";
        if (hasOperator && !KnownOperators.Contains(op, StringComparer.Ordinal))
        {
            return Blocked(sqref, $"セル {sqref} の入力規則の条件「{op}」には対応していません。");
        }

        var formula1 = validation.Formula1?.Text;
        var formula2 = validation.Formula2?.Text;

        if (type is "none" or "list" && hasOperator)
        {
            // list / none では operator は意味を持たない。勝手に消して正常化しない。
            return Blocked(sqref, $"セル {sqref} の入力規則に、意味を持たない条件設定が含まれています。");
        }

        if (type == "none")
        {
            return formula1 is null && formula2 is null
                ? Supported(validation, sqref)
                : Blocked(sqref, $"セル {sqref} の入力規則の構造が想定と異なります。");
        }

        if (type == "list")
        {
            if (formula2 is not null)
            {
                return Blocked(sqref, $"セル {sqref} のリスト入力規則の構造が想定と異なります。");
            }

            return ValidateListSource(formula1, sqref) is { } listError
                ? Blocked(sqref, listError)
                : Supported(validation, sqref);
        }

        // whole / decimal / date / time / textLength
        if (formula1 is null)
        {
            return Blocked(sqref, $"セル {sqref} の入力規則に条件値がありません。");
        }

        var needsSecond = RangeOperators.Contains(op, StringComparer.Ordinal);
        if (needsSecond && formula2 is null)
        {
            return Blocked(sqref, $"セル {sqref} の入力規則に 2 つ目の条件値がありません。");
        }

        if (!needsSecond && formula2 is not null)
        {
            return Blocked(sqref, $"セル {sqref} の入力規則の条件値の数が想定と異なります。");
        }

        if (!TryParseConstant(formula1, out var value1))
        {
            return Blocked(sqref, DescribeUnsupportedFormula(formula1, sqref));
        }

        double? value2 = null;
        if (formula2 is not null)
        {
            if (!TryParseConstant(formula2, out var parsed2))
            {
                return Blocked(sqref, DescribeUnsupportedFormula(formula2, sqref));
            }

            value2 = parsed2;
        }

        // 日付だけは 1904 date system の Workbook から来た値を 1900 系へそろえる。
        if (type == "date" && sourceIsDate1904)
        {
            var element = (DataValidation)validation.CloneNode(true);
            element.Formula1 = new Formula1(FormatSerial(
                MergeCellValue.NormalizeSerialTo1900(value1, sourceIsDate1904: true)));

            if (value2 is { } second)
            {
                element.Formula2 = new Formula2(FormatSerial(
                    MergeCellValue.NormalizeSerialTo1900(second, sourceIsDate1904: true)));
            }

            return new DataValidationInfo { Element = element, Sqref = sqref };
        }

        return Supported(validation, sqref);
    }

    /// <summary>リストの参照元が、直接指定か同じシート内の単純な絶対参照かを確かめる。</summary>
    private static string? ValidateListSource(string? formula1, string sqref)
    {
        if (string.IsNullOrWhiteSpace(formula1))
        {
            return $"セル {sqref} のリスト入力規則に選択肢がありません。";
        }

        var text = formula1.Trim();

        // A. 直接指定("赤,青,緑")。引用符ごとそのまま保持するので中身は作り直さない。
        if (text.StartsWith('"'))
        {
            if (text.Length < 2 || !text.EndsWith('"'))
            {
                return $"セル {sqref} のリスト入力規則の選択肢を解釈できません。";
            }

            var inner = text[1..^1];
            if (inner.Contains('"'))
            {
                return $"セル {sqref} のリスト入力規則の選択肢を解釈できません。";
            }

            return text.Length > MaxListLiteralLength
                ? $"セル {sqref} のリスト入力規則の選択肢が Excel の上限({MaxListLiteralLength} 文字)を超えています。"
                : null;
        }

        // B. 同じシート内の完全絶対参照($B$2:$B$10)。位置は出力でも変わらない。
        if (AbsoluteRangePattern().IsMatch(text) && A1RangeValidator.IsValidRange(text))
        {
            return null;
        }

        return DescribeUnsupportedFormula(text, sqref);
    }

    /// <summary>数値定数として解釈できるか(InvariantCulture・有限数のみ)。</summary>
    private static bool TryParseConstant(string? formula, out double value)
    {
        value = 0;
        var text = formula?.Trim();
        if (string.IsNullOrEmpty(text))
        {
            return false;
        }

        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
            && double.IsFinite(value);
    }

    private static string FormatSerial(double serial) => serial.ToString(CultureInfo.InvariantCulture);

    /// <summary>対応していない数式について、利用者に伝わる理由を作る。</summary>
    private static string DescribeUnsupportedFormula(string? formula, string sqref)
    {
        var text = (formula ?? string.Empty).Trim();

        if (text.Contains("#REF!", StringComparison.Ordinal))
        {
            return $"セル {sqref} の入力規則の参照が壊れています(#REF!)。";
        }

        if (text.Contains('[') || text.Contains(']'))
        {
            return $"セル {sqref} の入力規則が他のブックや表を参照しているため、"
                + "現在のバージョンでは安全に集約できません。";
        }

        if (text.Contains('!'))
        {
            return $"セル {sqref} の入力規則が他のシートを参照しているため、"
                + "現在のバージョンでは安全に集約できません。";
        }

        if (text.Contains('(') || text.StartsWith('='))
        {
            return $"セル {sqref} の入力規則が関数や数式({text})を使っているため、"
                + "現在のバージョンでは安全に集約できません。";
        }

        return $"セル {sqref} の入力規則が名前定義または解釈できない参照({text})を使っているため、"
            + "現在のバージョンでは安全に集約できません。";
    }

    private static DataValidationInfo Supported(DataValidation validation, string sqref)
        => new() { Element = (DataValidation)validation.CloneNode(true), Sqref = sqref };

    private static DataValidationInfo Blocked(string sqref, string reason)
        => new() { BlockReason = reason, Sqref = sqref };
}
