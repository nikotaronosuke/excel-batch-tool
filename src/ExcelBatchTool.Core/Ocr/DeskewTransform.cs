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

    /// <summary>
    /// ページを <paramref name="degrees"/> 度回してまっすぐにしたときの対応を作る。
    ///
    /// **ここで角度の向きを変えてはいけない。** 画像を回す側は OpenCV の
    /// getRotationMatrix2D(中心, degrees, 1) を使っていて、元の点 (x, y) の中身は
    ///
    ///   x' = cx + dx*cos(degrees) + dy*sin(degrees)
    ///   y' = cy - dx*sin(degrees) + dy*cos(degrees)   (dx = x - cx, dy = y - cy)
    ///
    /// へ移る。これは <see cref="Map"/> に -degrees を渡したものと同じなので、
    /// <see cref="ToDeskewed"/> が -AngleDegrees を使う今の作りに合わせるには、
    /// AngleDegrees は**回した角度そのまま**でなければならない。
    ///
    /// 実際、ここを -degrees にしていたために <see cref="ToOriginal"/> と
    /// <see cref="ToDeskewed"/> の働きが入れ替わり、戻すどころか傾き 2 つぶん回って
    /// いた。1240 × 1754 のページを 2 度回した実測では、(330, 520) にある文字に対して
    /// 枠が (305.6, 540.8) に出ていた(約 32 画素。中心から遠い行ほど大きくなる)。
    /// </summary>
    public static DeskewTransform FromRotation(double degrees, double centerX, double centerY)
        => degrees == 0 ? None : new DeskewTransform(degrees, centerX, centerY);

    public bool IsIdentity => AngleDegrees == 0;

    /// <summary>直した画像の座標 → 元のページの座標(確認画面はこちらを使う)。</summary>
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
