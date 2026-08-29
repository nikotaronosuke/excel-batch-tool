namespace ExcelBatchTool.Core.Ocr;

/// <summary>
/// 表の 1 つの列の中で、その列だけ形が違うセルを自動確定から外す。
///
/// なぜ要るか: 二重読みの「両方が一致した」という根拠は、
/// 2 つのモデルへ**同じ間違った入力**が入ったときには成り立たない。
/// 切り出しが上下逆のまま両モデルへ渡ると、両方が同じように読み、
/// 自信も高いまま一致する。
///
/// 実測(Phase 2F-B2 / PdfScanBench table)。誤って自動確定した 2 件は
/// どちらも上下逆に読まれたものだった:
///
/// | 位置 | 読み | 正 | 自信 |
/// |---|---|---|---|
/// | 罫線あり p5 r16c0 | 9600 | A0096 | 99.8% |
/// | 罫線なし p4 r9c0  | 6900 | A0069 | 99.7% |
///
/// 自信でも一致でも捕まえられない。捕まえられるのは**列**である。
/// 同じ列の他のセルは「英字 + 数字」なのに、この 2 件だけ「数字」だけになる。
///
/// 判定は粗くする。文字の並びまで一致を求めると、金額の「1,234」と「12,345」が
/// 別扱いになって正しいセルまで人へ回してしまう。
/// **どの種類の文字が含まれるか**の集合だけを見る。
///
/// いちばん多い形が全体の 4 割に満たないときは、形がばらけていて
/// 「その列らしい形」が決まらないので何もしない。
///
/// 4 割にした根拠: 実測で誤確定が残ったページの列は
/// 「英字+数字 11 / 数字 9 / 日本語 1」の 21 件で、いちばん多い形が 52% だった。
/// 6 割を求めると、この列では何も判断できずに誤確定がそのまま残る。
/// この判定が外れたときの損は**人が確認する手間が増えるだけ**で、
/// 間違ったまま出力されることは無い。だから厳しすぎるほうが害が大きい。
/// </summary>
public static class ColumnShapeGuard
{
    /// <summary>いちばん多い形がこの割合に満たなければ、列の形を判断しない。</summary>
    public const double MajorityRatio = 0.4;

    /// <summary>多数派を決めるのに要る、中身のあるセルの数。</summary>
    public const int MinimumSamples = 5;

    [Flags]
    public enum Shape
    {
        None = 0,
        Digit = 1,
        Latin = 2,
        Japanese = 4,
    }

    /// <summary>
    /// 記号と空白は数えない。桁区切りや単位は書き方の違いであって、
    /// 文字の種類の違いではないため。これを数えると、金額の列で
    /// 「999」と「1,234」が別の形になり、正しいセルまで人へ回してしまう。
    /// </summary>
    public static Shape ShapeOf(string text)
    {
        var shape = Shape.None;
        foreach (var c in text)
        {
            shape |= char.IsAsciiDigit(c) ? Shape.Digit
                : char.IsAsciiLetter(c) ? Shape.Latin
                : c > (char)0x7F && char.IsLetter(c) ? Shape.Japanese
                : Shape.None;
        }

        return shape;
    }

    /// <summary>
    /// 列の多数派の形を返す。はっきりしなければ <see cref="Shape.None"/>。
    /// </summary>
    public static Shape MajorityShape(IEnumerable<string> columnTexts)
    {
        var shapes = columnTexts
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .Select(ShapeOf)
            .Where(shape => shape != Shape.None)
            .ToList();

        if (shapes.Count < MinimumSamples)
        {
            return Shape.None;
        }

        var best = shapes.GroupBy(shape => shape)
            .OrderByDescending(group => group.Count())
            .First();

        return (double)best.Count() / shapes.Count >= MajorityRatio ? best.Key : Shape.None;
    }

    /// <summary>
    /// この列でこのセルを自動確定してよいか。
    /// 列の形が決まらないとき、セルが空のとき、記号だけのときは判断しない。
    /// </summary>
    public static bool CanAutoAccept(Shape majority, string text)
        => majority == Shape.None
            || string.IsNullOrWhiteSpace(text)
            || ShapeOf(text) is Shape.None
            || ShapeOf(text) == majority;

    public const string Reason =
        "同じ列の他の行と文字の種類が違います(上下逆に読まれていることがあります)。"
        + "元のページと見比べてください。";
}
