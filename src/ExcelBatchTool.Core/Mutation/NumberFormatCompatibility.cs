using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBatchTool.Core.Mutation;

/// <summary>
/// セルの表示形式が、書き込もうとしている値の種類と噛み合うかを判定する。
///
/// 書式(StyleIndex)は変えずに値だけ差し替えるため、たとえば日付書式のセルへ
/// 数値を入れると別の日付として表示される。%(0%)のセルへ 50 を入れると 5000% になる。
/// こうした「気づけない意味の変化」を避けるため、安全と言い切れる表示形式だけを許可し、
/// 判断が曖昧なものは Block する。
/// </summary>
internal sealed class NumberFormatCompatibility
{
    /// <summary>組み込み表示形式のうち General。</summary>
    private const uint General = 0;

    /// <summary>組み込み表示形式のうち文字列(@)。</summary>
    private const uint TextFormat = 49;

    /// <summary>ユーザー定義表示形式の開始 ID。</summary>
    private const uint FirstCustomFormatId = 164;

    /// <summary>
    /// 数値をそのまま書いてよい組み込み表示形式。
    /// 0 = General、1 = 0、2 = 0.00、3 = #,##0、4 = #,##0.00。
    /// 通貨・%・分数・指数・日付・時刻はいずれも意味が変わりうるので入れない。
    /// </summary>
    private static readonly uint[] NumericFormats = [General, 1, 2, 3, 4];

    /// <summary>文字列をそのまま書いてよい組み込み表示形式(General と文字列)。</summary>
    private static readonly uint[] TextFormats = [General, TextFormat];

    /// <summary>StyleIndex → numFmtId。</summary>
    private readonly uint[] _formatIdByStyle;

    private NumberFormatCompatibility(uint[] formatIdByStyle) => _formatIdByStyle = formatIdByStyle;

    public static NumberFormatCompatibility Create(WorkbookPart workbookPart)
    {
        var formats = workbookPart.WorkbookStylesPart?.Stylesheet?.CellFormats?
            .Elements<CellFormat>()
            .Select(format => format.NumberFormatId?.Value ?? General)
            .ToArray() ?? [];

        return new NumberFormatCompatibility(formats);
    }

    /// <summary>
    /// この書式のセルへ指定の種類の値を書けるか。書けない場合は理由を返す。
    /// 空欄にするだけなら表示形式によらず安全(値が無ければ表示形式は効かない)。
    /// </summary>
    public string? Check(uint? styleIndex, CellWriteKind kind, string cellReference)
    {
        if (kind == CellWriteKind.Blank)
        {
            return null;
        }

        var formatId = ResolveFormatId(styleIndex);
        var allowed = kind == CellWriteKind.Number ? NumericFormats : TextFormats;

        if (allowed.Contains(formatId))
        {
            return null;
        }

        return Describe(formatId, kind, cellReference);
    }

    /// <summary>
    /// この書式の数値を、意味を変えずにそのまま読み取れるか。
    /// 転記元(Phase 2C)で「表示 15% / 生の値 0.15」のような食い違いを避けるために使う。
    /// </summary>
    public bool IsPlainNumber(uint? styleIndex) => NumericFormats.Contains(ResolveFormatId(styleIndex));

    private static string Describe(uint formatId, CellWriteKind kind, string cellReference)
    {

        var what = kind == CellWriteKind.Number ? "数値" : "文字";
        return formatId >= FirstCustomFormatId
            ? $"{cellReference} にはユーザー設定の表示形式が設定されているため、"
                + $"{what}を入れると表示が変わる可能性があります。現在のバージョンでは変更できません。"
            : $"{cellReference} には日付・時刻・通貨などの表示形式が設定されているため、"
                + $"{what}を入れると表示が変わる可能性があります。現在のバージョンでは変更できません。";
    }

    /// <summary>StyleIndex が無い・範囲外なら General 扱い。</summary>
    private uint ResolveFormatId(uint? styleIndex)
        => styleIndex is { } index && index < _formatIdByStyle.Length
            ? _formatIdByStyle[index]
            : General;
}
