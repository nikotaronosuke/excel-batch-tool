using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>走査で取り出した条件付き書式のルール 1 件。</summary>
internal sealed record CfRuleInfo
{
    /// <summary>元の要素(dxfId はまだ元ブックのまま)。Block 時は null。</summary>
    public ConditionalFormattingRule? Element { get; init; }

    /// <summary>元ブックでの書式(dxf)の位置。</summary>
    public uint SourceDxfId { get; init; }

    /// <summary>元ブックの書式そのもの(出力へ写す)。</summary>
    public DifferentialFormat? Dxf { get; init; }

    /// <summary>書式に含まれる項目("font,fill" など)。出力後の照合に使う。</summary>
    public string DxfChildren { get; init; } = string.Empty;

    /// <summary>書式の表示形式(formatCode)。指定が無ければ null。</summary>
    public string? DxfNumberFormatCode { get; init; }

    public string? BlockReason { get; init; }
}

/// <summary>走査で取り出した条件付き書式(1 つの範囲に対するルール群)。</summary>
internal sealed record ConditionalFormattingInfo
{
    public string Sqref { get; init; } = string.Empty;

    public IReadOnlyList<CfRuleInfo> Rules { get; init; } = Array.Empty<CfRuleInfo>();

    public string? BlockReason { get; init; }
}

/// <summary>
/// 標準の条件付き書式のうち、数式を使わず dxf で書式を指定する基本ルールだけを扱う。
/// 数式条件・カラースケール・データバー・アイコンセット・x14 拡張は対象外。
/// </summary>
internal static class ConditionalFormattingScanner
{
    /// <summary>Office 2010 以降の条件付き書式拡張(x14)を表す extLst の URI。</summary>
    public const string X14ExtensionUri = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";

    /// <summary>今回対応するルールの種類。</summary>
    private static readonly string[] SupportedTypes =
        ["duplicateValues", "uniqueValues", "top10", "aboveAverage"];

    /// <summary>rank の有効範囲(Microsoft の仕様: 個数指定は 1〜1000、割合指定は 0〜100)。</summary>
    private const uint MaxRankCount = 1000;

    private const uint MaxRankPercent = 100;

    /// <summary>条件付き書式 1 件(1 つの適用範囲)を解析する。</summary>
    public static ConditionalFormattingInfo Scan(
        ConditionalFormatting conditionalFormatting,
        IReadOnlyList<DifferentialFormat> differentialFormats)
    {
        var sqref = conditionalFormatting.SequenceOfReferences?.InnerText ?? string.Empty;

        if (conditionalFormatting.ExtendedAttributes.Any())
        {
            return Blocked(sqref, "対応していない設定を含む条件付き書式があります。");
        }

        if (conditionalFormatting.Pivot?.Value == true)
        {
            return Blocked(sqref, "ピボットテーブル用の条件付き書式は、現在のバージョンでは集約できません。");
        }

        if (!A1RangeValidator.IsValidRangeList(sqref, out var invalidToken))
        {
            return Blocked(sqref, $"条件付き書式の適用範囲「{invalidToken}」を解釈できません。");
        }

        foreach (var child in conditionalFormatting.ChildElements)
        {
            if (child is not ConditionalFormattingRule)
            {
                return Blocked(sqref, "対応していない内容を含む条件付き書式があります。");
            }
        }

        var rules = new List<CfRuleInfo>();
        foreach (var rule in conditionalFormatting.Elements<ConditionalFormattingRule>())
        {
            rules.Add(ScanRule(rule, sqref, differentialFormats));
        }

        if (rules.Count == 0)
        {
            return Blocked(sqref, $"条件付き書式({sqref})にルールがありません。");
        }

        return new ConditionalFormattingInfo { Sqref = sqref, Rules = rules };
    }

    private static CfRuleInfo ScanRule(
        ConditionalFormattingRule rule,
        string sqref,
        IReadOnlyList<DifferentialFormat> differentialFormats)
    {
        var type = rule.Type?.InnerText ?? string.Empty;

        if (!SupportedTypes.Contains(type, StringComparer.Ordinal))
        {
            return BlockedRule(DescribeUnsupportedType(type, sqref));
        }

        if (rule.ExtendedAttributes.Any())
        {
            return BlockedRule($"{sqref} の条件付き書式に対応していない設定が含まれています。");
        }

        // 数式を持つルールは、参照の解釈が必要になるため今回は扱わない。
        if (rule.ChildElements.Count > 0)
        {
            return BlockedRule(
                $"{sqref} の条件付き書式が数式など対応していない内容を含むため、"
                    + "現在のバージョンでは安全に集約できません。");
        }

        if (rule.Priority?.Value is not { } priority || priority <= 0)
        {
            return BlockedRule($"{sqref} の条件付き書式の優先順位が正しくありません。");
        }

        if (ValidateTypeSpecificAttributes(rule, type, sqref) is { } attributeError)
        {
            return BlockedRule(attributeError);
        }

        if (rule.FormatId?.Value is not { } dxfId)
        {
            return BlockedRule($"{sqref} の条件付き書式に書式の指定がありません。");
        }

        if (dxfId >= (uint)differentialFormats.Count)
        {
            return BlockedRule($"{sqref} の条件付き書式の書式情報が壊れているため、安全に集約できません。");
        }

        var dxf = differentialFormats[(int)dxfId];
        if (ValidateDifferentialFormat(dxf, sqref) is { } dxfError)
        {
            return BlockedRule(dxfError);
        }

        return new CfRuleInfo
        {
            Element = (ConditionalFormattingRule)rule.CloneNode(true),
            SourceDxfId = dxfId,
            Dxf = (DifferentialFormat)dxf.CloneNode(true),
            DxfChildren = DescribeDxfChildren(dxf),
            DxfNumberFormatCode = dxf.NumberingFormat?.FormatCode?.Value,
        };
    }

    /// <summary>種類ごとに、意味を持つ属性だけが付いているか確かめる。</summary>
    private static string? ValidateTypeSpecificAttributes(
        ConditionalFormattingRule rule,
        string type,
        string sqref)
    {
        // どの対応 type でも使わない属性。付いていれば構造が想定と異なる。
        if (rule.Operator is not null || rule.Text is not null || rule.TimePeriod is not null)
        {
            return $"{sqref} の条件付き書式の構造が想定と異なります。";
        }

        var isTop10 = string.Equals(type, "top10", StringComparison.Ordinal);
        var isAboveAverage = string.Equals(type, "aboveAverage", StringComparison.Ordinal);

        if (!isTop10 && (rule.Rank is not null || rule.Percent is not null || rule.Bottom is not null))
        {
            return $"{sqref} の条件付き書式の構造が想定と異なります。";
        }

        if (!isAboveAverage
            && (rule.AboveAverage is not null || rule.EqualAverage is not null || rule.StdDev is not null))
        {
            return $"{sqref} の条件付き書式の構造が想定と異なります。";
        }

        if (isTop10)
        {
            if (rule.Rank?.Value is not { } rank)
            {
                return $"{sqref} の条件付き書式に上位/下位の件数が指定されていません。";
            }

            var isPercent = rule.Percent?.Value == true;
            var valid = isPercent
                ? rank <= MaxRankPercent
                : rank >= 1 && rank <= MaxRankCount;

            if (!valid)
            {
                return $"{sqref} の条件付き書式の上位/下位の指定({rank})が有効な範囲を超えています。";
            }
        }

        if (isAboveAverage)
        {
            if (rule.EqualAverage?.Value == true && rule.StdDev is not null)
            {
                // 平均を含める指定と標準偏差の指定は同時に成立しない。
                return $"{sqref} の条件付き書式の平均条件の指定が想定と異なります。";
            }

            if (rule.StdDev?.Value is { } standardDeviation && standardDeviation < 1)
            {
                return $"{sqref} の条件付き書式の平均条件の指定({standardDeviation})が想定と異なります。";
            }
        }

        return null;
    }

    /// <summary>
    /// 書式(dxf)が出力へそのまま写せるか確かめる。テーマ色は出力ブックのテーマ次第で
    /// 見え方が変わるため、黙って移さず Block する。
    /// </summary>
    private static string? ValidateDifferentialFormat(DifferentialFormat dxf, string sqref)
    {
        if (dxf.ExtendedAttributes.Any())
        {
            return $"{sqref} の条件付き書式に対応していない書式設定が含まれています。";
        }

        foreach (var child in dxf.ChildElements)
        {
            if (child is not (Font or NumberingFormat or Fill or Alignment or Border or Protection))
            {
                return $"{sqref} の条件付き書式に対応していない書式設定が含まれています。";
            }
        }

        if (dxf.Descendants<ColorType>().Any(color => color.Theme is not null))
        {
            return $"{sqref} の条件付き書式がテーマの色を使っているため、"
                + "集約後に色が変わる可能性があります。現在のバージョンでは安全に集約できません。";
        }

        // dxf の numFmt では formatCode が必須(ID によらない)。
        // 欠けているものは表示形式が元ブックの定義に依存するため、引き継げない。
        if (dxf.NumberingFormat is { } numberFormat
            && string.IsNullOrEmpty(numberFormat.FormatCode?.Value))
        {
            return $"{sqref} の条件付き書式の表示形式を安全に引き継げません。";
        }

        return null;
    }

    private static string DescribeUnsupportedType(string type, string sqref) => type switch
    {
        "expression" or "cellIs" =>
            $"{sqref} の条件付き書式が数式を使っているため、現在のバージョンでは安全に集約できません。",
        "colorScale" =>
            $"{sqref} のカラースケールは、現在のバージョンでは集約できません。",
        "dataBar" =>
            $"{sqref} のデータバーは、現在のバージョンでは集約できません。",
        "iconSet" =>
            $"{sqref} のアイコンセットは、現在のバージョンでは集約できません。",
        _ => $"{sqref} の条件付き書式(種類「{type}」)は、現在のバージョンでは集約できません。",
    };

    /// <summary>
    /// 出力用のルール要素を、既知の属性だけから組み立て直す。
    /// 元ファイルに無い属性は書かない(Excel の見え方が変わらないようにするため)。
    /// dxfId は出力ブックでの書式の位置に差し替える。
    /// </summary>
    public static ConditionalFormattingRule BuildOutputRule(CfRuleInfo rule, uint outputDxfId)
    {
        var source = rule.Element
            ?? throw new InvalidOperationException("移植できない条件付き書式を出力しようとしました。");

        var output = new ConditionalFormattingRule
        {
            Type = source.Type!.Value,
            Priority = source.Priority!.Value,
            FormatId = outputDxfId,
        };

        if (source.StopIfTrue?.Value is { } stopIfTrue)
        {
            output.StopIfTrue = stopIfTrue;
        }

        if (source.Rank?.Value is { } rank)
        {
            output.Rank = rank;
        }

        if (source.Percent?.Value is { } percent)
        {
            output.Percent = percent;
        }

        if (source.Bottom?.Value is { } bottom)
        {
            output.Bottom = bottom;
        }

        if (source.AboveAverage?.Value is { } aboveAverage)
        {
            output.AboveAverage = aboveAverage;
        }

        if (source.EqualAverage?.Value is { } equalAverage)
        {
            output.EqualAverage = equalAverage;
        }

        if (source.StdDev?.Value is { } standardDeviation)
        {
            output.StdDev = standardDeviation;
        }

        return output;
    }

    /// <summary>適用範囲は元のまま。ルールは出力用に組み立て済みのものを受け取る。</summary>
    public static ConditionalFormatting BuildOutputFormatting(
        string sqref,
        IEnumerable<ConditionalFormattingRule> rules)
        => new(rules) { SequenceOfReferences = new ListValue<StringValue> { InnerText = sqref } };

    /// <summary>出力後の照合に使う概要を作る。</summary>
    public static ConditionalFormattingSummary Summarize(ConditionalFormattingInfo info) => new()
    {
        Sqref = info.Sqref,
        Rules = [.. info.Rules
            .Where(rule => rule.Element is not null)
            .Select(rule => new ConditionalFormattingRuleSummary
            {
                Type = rule.Element!.Type?.InnerText ?? string.Empty,
                Priority = rule.Element.Priority?.Value ?? 0,
                StopIfTrue = rule.Element.StopIfTrue?.Value,
                Rank = rule.Element.Rank?.Value,
                Percent = rule.Element.Percent?.Value,
                Bottom = rule.Element.Bottom?.Value,
                AboveAverage = rule.Element.AboveAverage?.Value,
                EqualAverage = rule.Element.EqualAverage?.Value,
                StandardDeviation = rule.Element.StdDev?.Value,
                FormatChildren = rule.DxfChildren,
                FormatNumberCode = rule.DxfNumberFormatCode,
            })],
    };

    /// <summary>書式(dxf)に含まれる項目の並びを、要素名の一覧として書き出す。</summary>
    public static string DescribeDxfChildren(OpenXmlElement dxf)
        => string.Join(",", dxf.ChildElements.Select(child => child.LocalName));

    private static ConditionalFormattingInfo Blocked(string sqref, string reason)
        => new() { Sqref = sqref, BlockReason = reason };

    private static CfRuleInfo BlockedRule(string reason) => new() { BlockReason = reason };

    /// <summary>Workbook の書式(dxf)一覧を読む。</summary>
    public static IReadOnlyList<DifferentialFormat> ReadDifferentialFormats(WorkbookPart workbookPart)
        => workbookPart.WorkbookStylesPart?.Stylesheet?.DifferentialFormats?
            .Elements<DifferentialFormat>().ToList() ?? [];

    /// <summary>同じシート内で優先順位が重複していないか確かめる。</summary>
    public static string? FindDuplicatePriority(IReadOnlyList<ConditionalFormattingInfo> formattings)
    {
        var seen = new HashSet<int>();
        foreach (var rule in formattings
            .SelectMany(formatting => formatting.Rules)
            .Select(rule => rule.Element?.Priority?.Value)
            .OfType<int>())
        {
            if (!seen.Add(rule))
            {
                return $"条件付き書式の優先順位({rule})が重複しています。";
            }
        }

        return null;
    }
}
