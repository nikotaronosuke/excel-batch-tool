using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>ハイパーリンクの種類。</summary>
internal enum HyperlinkKind
{
    /// <summary>Web / メールへの外部リンク。</summary>
    External,

    /// <summary>同じシート内へのリンク。</summary>
    InternalSameSheet,

    /// <summary>同じブックの別シートへのリンク。</summary>
    InternalOtherSheet,
}

/// <summary>
/// 走査で取り出した 1 件のハイパーリンク。
/// 別シート宛かどうかの解決は Planner が行う(選択状況を知る必要があるため)。
/// </summary>
internal sealed record HyperlinkInfo
{
    public required string Reference { get; init; }

    public required HyperlinkKind Kind { get; init; }

    /// <summary>外部リンクの絶対 URI(External のときのみ)。</summary>
    public string? ExternalTarget { get; init; }

    /// <summary>外部リンクの場合は対象文書内のアンカー、内部リンクの場合はセル参照部分。</summary>
    public string? Location { get; init; }

    /// <summary>内部リンクの参照先シート名(InternalOtherSheet のときのみ)。</summary>
    public string? TargetSheetName { get; init; }

    public string? Tooltip { get; init; }

    public string? Display { get; init; }

    /// <summary>安全に移植できない理由。null なら移植できる。</summary>
    public string? BlockReason { get; init; }
}

/// <summary>
/// Worksheet のハイパーリンクを、意味を保ったまま移植できるか調べる。
/// リンク先へは一切アクセスしない(文字列と relationship の情報だけを読む)。
/// </summary>
internal static class HyperlinkScanner
{
    private static readonly string[] AllowedSchemes = ["http", "https", "mailto"];

    /// <summary>ハイパーリンク 1 件を解析する。</summary>
    public static HyperlinkInfo Scan(Hyperlink hyperlink, string sheetName, WorksheetPart worksheetPart)
    {
        var reference = hyperlink.Reference?.Value;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return Blocked("(位置不明)", "リンクの位置が指定されていません。");
        }

        if (!A1RangeValidator.IsValidRange(reference))
        {
            return Blocked(reference, $"リンクの位置「{reference}」を解釈できません。");
        }

        // 想定していない属性・子要素があるものは、意味を保証できないので移植しない。
        if (hyperlink.ExtendedAttributes.Any() || hyperlink.HasChildren)
        {
            return Blocked(reference, "対応していない設定を含むリンクがあります。");
        }

        var tooltip = hyperlink.Tooltip?.Value;
        var display = hyperlink.Display?.Value;
        var location = hyperlink.Location?.Value;

        if (hyperlink.Id?.Value is { } relationshipId)
        {
            return ScanExternal(reference, relationshipId, location, tooltip, display, worksheetPart);
        }

        if (string.IsNullOrWhiteSpace(location))
        {
            return Blocked(reference, "リンク先が指定されていません。");
        }

        return ScanInternal(reference, location, sheetName, tooltip, display);
    }

    private static HyperlinkInfo ScanExternal(
        string reference,
        string relationshipId,
        string? location,
        string? tooltip,
        string? display,
        WorksheetPart worksheetPart)
    {
        HyperlinkRelationship? relationship;
        try
        {
            relationship = worksheetPart.HyperlinkRelationships
                .FirstOrDefault(item => string.Equals(item.Id, relationshipId, StringComparison.Ordinal));
        }
        catch (Exception)
        {
            relationship = null;
        }

        if (relationship is null)
        {
            return Blocked(reference, "リンク先の情報が見つかりません(ファイル内の参照が壊れています)。");
        }

        if (!relationship.IsExternal)
        {
            return Blocked(reference, "対応していない形式のリンクです。");
        }

        var uri = relationship.Uri;
        if (uri is null || !uri.IsAbsoluteUri)
        {
            return Blocked(
                reference,
                "ローカルファイルへのリンクは、保存場所が変わるとリンク先が変わる可能性があるため、"
                    + "現在のバージョンでは安全に集約できません。");
        }

        if (!AllowedSchemes.Contains(uri.Scheme, StringComparer.OrdinalIgnoreCase))
        {
            return Blocked(
                reference,
                uri.IsFile
                    ? "ローカルファイルへのリンクは、保存場所が変わるとリンク先が変わる可能性があるため、"
                        + "現在のバージョンでは安全に集約できません。"
                    : $"「{uri.Scheme}」形式のリンクは現在のバージョンでは対応していません。");
        }

        return new HyperlinkInfo
        {
            Reference = reference,
            Kind = HyperlinkKind.External,
            ExternalTarget = uri.OriginalString,
            Location = location,
            Tooltip = tooltip,
            Display = display,
        };
    }

    private static HyperlinkInfo ScanInternal(
        string reference,
        string location,
        string sheetName,
        string? tooltip,
        string? display)
    {
        if (location.Contains("#REF!", StringComparison.Ordinal))
        {
            return Blocked(reference, "リンク先が壊れています(#REF!)。");
        }

        if (!SheetReferenceSyntax.HasSheetName(location))
        {
            // シート名なし = 同じシート内。セル参照として解釈できることだけ確かめる。
            return IsCellReference(location)
                ? new HyperlinkInfo
                {
                    Reference = reference,
                    Kind = HyperlinkKind.InternalSameSheet,
                    Location = location,
                    Tooltip = tooltip,
                    Display = display,
                }
                : Blocked(
                    reference,
                    $"リンク先「{location}」は名前定義か、現在のバージョンでは解釈できない形式です。");
        }

        if (!SheetReferenceSyntax.TrySplit(location, out var targetSheet, out var cell, out var problem))
        {
            return Blocked(reference, problem switch
            {
                SheetReferenceProblem.ThreeDimensional =>
                    "複数シートにまたがるリンクは現在のバージョンでは対応していません。",
                _ => $"リンク先「{location}」を解釈できません。",
            });
        }

        if (targetSheet.Contains('[') || targetSheet.Contains(']'))
        {
            return Blocked(reference, "他のブックへのリンクは現在のバージョンでは対応していません。");
        }

        if (!IsCellReference(cell))
        {
            return Blocked(reference, $"リンク先「{location}」のセル位置を解釈できません。");
        }

        // シート名が自分自身でも、出力シート名へ書き直す必要があるので同じ扱いにする。
        return new HyperlinkInfo
        {
            Reference = reference,
            Kind = HyperlinkKind.InternalOtherSheet,
            Location = cell,
            TargetSheetName = targetSheet,
            Tooltip = tooltip,
            Display = display,
        };
    }

    /// <summary>A1 / A1:B5 形式で、シートの上限内か($ 付きも許容)。</summary>
    private static bool IsCellReference(string text) => A1RangeValidator.IsValidRange(text);

    private static HyperlinkInfo Blocked(string reference, string reason) => new()
    {
        Reference = reference,
        Kind = HyperlinkKind.External,
        BlockReason = reason,
    };
}
