namespace ExcelBatchTool.Core.Ocr;

/// <summary>スキャンされたページの種類。</summary>
public enum ScanPageKind
{
    /// <summary>文章。行ごとに取り出す。</summary>
    Prose = 0,

    /// <summary>罫線のある表。罫線の格子から行・列を作る。</summary>
    RuledTable,

    /// <summary>罫線のない表。文字の位置から列を作る。</summary>
    BorderlessTable,

    /// <summary>記入欄の並んだ帳票。項目を指定して読む。</summary>
    FixedForm,

    /// <summary>どれとも決められない。</summary>
    Unknown,
}

/// <summary>種類を決めるために測った値。</summary>
public sealed record ScanPageMetrics(
    int HorizontalRulings,
    int VerticalRulings,
    int LineCount,
    int AlignedRowCount,
    int ColumnCount,
    int UnderlineCount);

/// <summary>
/// スキャンされたページが、文章・罫線表・罫線なし表・帳票のどれかを決める。
/// 利用者に「これは表ですか」と聞かない。ただし決め手が無ければ推測せず
/// <see cref="ScanPageKind.Unknown"/> にして、候補を見せる側に判断を渡す。
///
/// Phase 2F-A で「記入欄の下線が並ぶ帳票を表と誤判定した」ことがあるので、
/// **横線だけでは表とみなさない**。縦線と組み合わさって格子になっていること、
/// あるいは文字の位置が列として揃っていることを条件にする。
/// </summary>
public static class ScanPageClassifier
{
    /// <summary>格子とみなすのに必要な罫線の本数。</summary>
    public const int GridLines = 3;

    /// <summary>列が揃っているとみなす、行全体に対する割合。</summary>
    public const double AlignedRatio = 0.6;

    /// <summary>帳票とみなすのに必要な下線の本数。</summary>
    public const int FormUnderlines = 4;

    public static ScanPageKind Classify(ScanPageMetrics metrics)
    {
        if (metrics.LineCount == 0)
        {
            return ScanPageKind.Unknown;
        }

        // 縦横が組み合わさって格子になっているなら罫線表。
        if (metrics.HorizontalRulings >= GridLines && metrics.VerticalRulings >= GridLines)
        {
            return ScanPageKind.RuledTable;
        }

        // 横線だけが並ぶのは記入欄の下線。表ではなく帳票として扱う。
        if (metrics.UnderlineCount >= FormUnderlines && metrics.VerticalRulings < GridLines)
        {
            return ScanPageKind.FixedForm;
        }

        // 罫線が無くても、列の位置が揃った行が続くなら表。
        if (metrics.ColumnCount >= ScanTableBuilder.MinimumColumns
            && metrics.AlignedRowCount >= ScanTableBuilder.MinimumRows
            && metrics.AlignedRowCount >= metrics.LineCount * AlignedRatio)
        {
            return ScanPageKind.BorderlessTable;
        }

        return ScanPageKind.Prose;
    }

    public static string Display(ScanPageKind kind) => kind switch
    {
        ScanPageKind.Prose => "文章",
        ScanPageKind.RuledTable => "罫線のある表",
        ScanPageKind.BorderlessTable => "罫線のない表",
        ScanPageKind.FixedForm => "記入欄のある帳票",
        _ => "判定できない",
    };
}
