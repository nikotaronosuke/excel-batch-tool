namespace ExcelBatchTool.Core.Ocr;

/// <summary>
/// 傾きを直したページと、元のページのあいだの座標の対応。
///
/// 傾いたページは、まっすぐに直してから OCR にかける。そのままだと読み取り位置は
/// 「直した画像の座標」になり、確認画面で元の PDF を出したときに枠がずれる。
/// そこで**必ずこの変換で元の座標へ戻してから**読み取り項目に持たせる。
/// 角度が 0 のときは何もしない(<see cref="None"/>)。
/// </summary>
public sealed record DeskewTransform(double AngleDegrees, double CenterX, double CenterY)
{
    /// <summary>傾きを直していない(座標はそのまま)。</summary>
    public static readonly DeskewTransform None = new(0, 0, 0);

    public bool IsIdentity => AngleDegrees == 0;

    /// <summary>直した画像の座標 → 元のページの座標。</summary>
    public OcrBox ToOriginal(OcrBox box) => Map(box, AngleDegrees);

    /// <summary>元のページの座標 → 直した画像の座標。</summary>
    public OcrBox ToDeskewed(OcrBox box) => Map(box, -AngleDegrees);

    /// <summary>
    /// 4 隅を回してから、それを囲む軸に沿った矩形にする。
    /// 枠を出すのが目的なので、回転した矩形そのものではなく外接矩形でよい。
    /// </summary>
    private OcrBox Map(OcrBox box, double degrees)
    {
        if (degrees == 0)
        {
            return box;
        }

        var radians = degrees * Math.PI / 180;
        var cos = Math.Cos(radians);
        var sin = Math.Sin(radians);

        double minX = double.MaxValue, minY = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue;

        foreach (var (x, y) in new[]
        {
            (box.X, box.Y),
            (box.Right, box.Y),
            (box.X, box.Bottom),
            (box.Right, box.Bottom),
        })
        {
            var dx = x - CenterX;
            var dy = y - CenterY;
            var rx = CenterX + (dx * cos) - (dy * sin);
            var ry = CenterY + (dx * sin) + (dy * cos);

            minX = Math.Min(minX, rx);
            minY = Math.Min(minY, ry);
            maxX = Math.Max(maxX, rx);
            maxY = Math.Max(maxY, ry);
        }

        return new OcrBox(minX, minY, maxX - minX, maxY - minY);
    }
}

/// <summary>傾きを直すかどうかの決め方。</summary>
public static class DeskewPolicy
{
    /// <summary>
    /// これ以下の傾きは直さない。わずかな傾きのために回すと、
    /// 補間で文字がぼやけて読み取りがかえって悪くなる。
    /// </summary>
    public const double MinimumAngle = 0.3;

    /// <summary>
    /// これを超える傾きは、この段階では扱わない。
    /// 大きく傾いた紙は撮り直したほうが速く、無理に直すと端が切れる。
    /// </summary>
    public const double MaximumAngle = 6.0;

    /// <summary>直すべき傾きか。</summary>
    public static bool ShouldDeskew(double angleDegrees, bool reliable)
        => reliable
            && Math.Abs(angleDegrees) >= MinimumAngle
            && Math.Abs(angleDegrees) <= MaximumAngle;

    /// <summary>
    /// 大きすぎて扱えない傾きか。
    ///
    /// 角度を測れなかったページは**回さずにそのまま読む**。止めない。
    /// 測れないのは「行らしい塊が少ない」ページ(見出しだけのページ、
    /// 図が主のページなど)で、その多くは実際には傾いていない。
    /// 測れないというだけで止めると、そういうページが軒並み扱えなくなる。
    /// 読み取った結果は確認の画面を必ず通るので、そこが安全網になる。
    /// </summary>
    public static bool IsTooTilted(double angleDegrees, bool reliable)
        => reliable && Math.Abs(angleDegrees) > MaximumAngle;
}
