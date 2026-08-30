using ExcelBatchTool.Core.Ocr;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// 傾きを直したページで、確認画面に出す画像と読み取り位置(赤い枠)が
/// 同じ座標で語られていることを確かめる。
///
/// この取り違えは実際に起きた。読み取りは「直した画像」で行うのに、確認画面へ出すのは
/// 「元のページ」なので、枠は必ず元の座標へ戻してから渡さなければならない。戻す向きを
/// 逆にしていたため、枠は文字から傾き 2 つぶん離れた場所に出ていた。値そのものは
/// 正しかったので、出力を見ているだけでは気付けない。
///
/// **ここでの要は <see cref="RotatedInk"/> を製品コードで作らないこと。**
/// 製品の変換で期待値を作ると、向きを取り違えたときに期待値も一緒に反転して素通りする。
/// 画素が実際にどう動くかを外から書き下したものだけを正解として使う。
/// </summary>
public class DeskewReviewCoordinateTests
{
    /// <summary>確認画面が使うページの大きさ (A4 300dpi 相当)。</summary>
    private const double CenterX = FakeOcrEngine.PageCenterX;

    /// <inheritdoc cref="CenterX"/>
    private const double CenterY = FakeOcrEngine.PageCenterY;

    /// <summary>
    /// 元のページの点が、<paramref name="degrees"/> 度回した画像のどこへ来るか。
    ///
    /// 製品は OpenCV の getRotationMatrix2D(中心, degrees, 1) で画像を回している。
    /// その行列は [[cos, sin, ...], [-sin, cos, ...]] なので、中心からの差 (dx, dy) は
    ///
    ///   x' = cx + dx*cos + dy*sin
    ///   y' = cy - dx*sin + dy*cos
    ///
    /// へ移る。これを製品コードに頼らず、ここへ直接書いておく。
    /// この式が本物の OpenCV と一致していることは
    /// <see cref="TheFormulaMatchesWhatTheRealRotationDidToThePixels"/> で押さえる。
    /// </summary>
    private static OcrBox RotatedInk(OcrBox box, double degrees)
    {
        var radians = degrees * Math.PI / 180;
        var cos = Math.Cos(radians);
        var sin = Math.Sin(radians);

        double minX = double.MaxValue, minY = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue;

        foreach (var (x, y) in new[]
        {
            (box.X, box.Y), (box.Right, box.Y), (box.X, box.Bottom), (box.Right, box.Bottom),
        })
        {
            var dx = x - CenterX;
            var dy = y - CenterY;
            var rx = CenterX + (dx * cos) + (dy * sin);
            var ry = CenterY - (dx * sin) + (dy * cos);

            minX = Math.Min(minX, rx);
            minY = Math.Min(minY, ry);
            maxX = Math.Max(maxX, rx);
            maxY = Math.Max(maxY, ry);
        }

        return new OcrBox(minX, minY, maxX - minX, maxY - minY);
    }

    /// <summary>
    /// 上の式が、実際に画素が動いた先と合っていること。
    ///
    /// 1240 × 1754 の白紙に (300, 500, 60, 40) の黒い四角を置き、製品と同じ回し方
    /// (OpenCV) で回してから、回った画像の中で黒い四角を測った実測値が下の数字。
    /// 回転した矩形を囲む枠は元より少し大きくなるので、中心どうしで比べる。
    /// </summary>
    [Theory]
    [InlineData(2.0, 317.5, 530.0)]
    [InlineData(-2.0, 342.5, 510.0)]
    public void TheFormulaMatchesWhatTheRealRotationDidToThePixels(
        double degrees, double measuredCenterX, double measuredCenterY)
    {
        var moved = RotatedInk(new OcrBox(300, 500, 60, 40), degrees);

        // 実測は整数の枠から測っているので、1 画素未満のずれは許す。
        Assert.Equal(measuredCenterX, moved.CenterX, 1.0);
        Assert.Equal(measuredCenterY, moved.CenterY, 1.0);
    }

    /// <summary>
    /// 直した画像で読んだ位置を元のページへ戻すと、文字が本当にある場所に戻ること。
    /// 向きを逆にすると、戻すどころか傾きが 2 倍になって離れていく。
    /// </summary>
    [Theory]
    [InlineData(0.5)]
    [InlineData(2.0)]
    [InlineData(-2.0)]
    [InlineData(3.5)]
    [InlineData(-5.0)]
    public void TheReadingGoesBackToTheInkOnTheOriginalPage(double degrees)
    {
        var ink = new OcrBox(300, 500, 240, 40);
        var seenWhileReading = RotatedInk(ink, degrees);

        var transform = DeskewTransform.FromRotation(degrees, CenterX, CenterY);
        var forReview = transform.ToOriginal(seenWhileReading);

        Assert.Equal(ink.CenterX, forReview.CenterX, 0.01);
        Assert.Equal(ink.CenterY, forReview.CenterY, 0.01);
    }

    /// <summary>
    /// 指定した領域(元のページの座標)を、読み取りに使う「直した画像」へ移すと、
    /// その画像で文字が実際にある場所になること。帳票の項目はこの向きで探す。
    /// </summary>
    [Theory]
    [InlineData(2.0)]
    [InlineData(-2.0)]
    public void TheDeclaredAreaLandsOnTheInkOfTheStraightenedImage(double degrees)
    {
        var area = new OcrBox(300, 500, 240, 40);

        var transform = DeskewTransform.FromRotation(degrees, CenterX, CenterY);
        var whereToLook = transform.ToDeskewed(area);

        var ink = RotatedInk(area, degrees);
        Assert.Equal(ink.CenterX, whereToLook.CenterX, 0.01);
        Assert.Equal(ink.CenterY, whereToLook.CenterY, 0.01);
    }

    /// <summary>
    /// 読み取りから確認画面まで通したときに、赤い枠が原文の上へ乗ること。
    ///
    /// 確認画面は「元のページを描いた画像」を出し、その上へ
    /// <see cref="OcrBoxMapper.ToDisplay"/> で枠を置く。ここが揃っていないと、
    /// 人が原文と見比べるという安全網そのものが働かない。
    /// </summary>
    [Theory]
    [InlineData(2.0)]
    [InlineData(-2.0)]
    public void TheRedBoxSitsOnTheInkOfThePageTheReviewShows(double degrees)
    {
        using var dir = new TempDir();
        var pdf = dir.File("傾き.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        // 元のページで文字がある場所。読み取りは直した画像で行うので、
        // 検出結果はそこで見える位置(RotatedInk)で返す。
        var ink = new OcrBox(300, 500, 240, 40);
        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.At("架空商事", 0.99, RotatedInk(ink, degrees)))
            .Probe(1, skew: degrees, horizontal: 0, vertical: 0);

        var reading = new PdfScanReader().Read(engine, pdf, [1]);
        var item = Assert.Single(reading.Items);

        // 確認画面が実際に出す画像と、その上に置く枠。
        using var source = engine.Open(pdf);
        var image = source.RenderPage(1, 150, CancellationToken.None);
        const double Zoom = 1.5;
        var rect = OcrBoxMapper.ToDisplay(item.BoundingBox, image, Zoom);

        // 同じ画像・同じ倍率で測った「文字のあるところ」。
        var scale = image.ScaleFromOcr * Zoom;
        var inkLeft = ink.X * scale;
        var inkTop = ink.Y * scale;
        var inkRight = ink.Right * scale;
        var inkBottom = ink.Bottom * scale;

        // 中心が合っていること。向きを取り違えると、ここが傾き 2 つぶん離れる。
        Assert.Equal(ink.CenterX * scale, rect.Left + (rect.Width / 2), 0.5);
        Assert.Equal(ink.CenterY * scale, rect.Top + (rect.Height / 2), 0.5);

        // 枠が文字を囲んでいること(これが確認画面でやりたいこと)。
        Assert.True(rect.Left <= inkLeft, $"枠の左が文字より内側: {rect.Left} > {inkLeft}");
        Assert.True(rect.Top <= inkTop, $"枠の上が文字より内側: {rect.Top} > {inkTop}");
        Assert.True(
            rect.Left + rect.Width >= inkRight,
            $"枠の右が文字より内側: {rect.Left + rect.Width} < {inkRight}");
        Assert.True(
            rect.Top + rect.Height >= inkBottom,
            $"枠の下が文字より内側: {rect.Top + rect.Height} < {inkBottom}");

        // ただし「大きく囲めば当たる」では意味がないので、広がってよい量にも上限を置く。
        //
        // 枠は軸に沿った矩形のまま持ち回るので、回すたびに外接矩形のぶんだけ広がる。
        // 読むときに 1 回、確認へ戻すときにもう 1 回なので、横は縦の、縦は横の
        // sin(傾き) ぶんが 2 回ぶん足される。それを超えて広がっていたら、
        // 位置合わせではなく「とりあえず大きく囲んだ」ことになる。
        var slack = 1 * scale;
        var spread = 2 * Math.Abs(Math.Sin(degrees * Math.PI / 180));
        Assert.True(
            rect.Width <= (inkRight - inkLeft) + ((inkBottom - inkTop) * spread) + slack,
            $"枠が横に広すぎる: {rect.Width}");
        Assert.True(
            rect.Height <= (inkBottom - inkTop) + ((inkRight - inkLeft) * spread) + slack,
            $"枠が縦に広すぎる: {rect.Height}");
    }

    /// <summary>傾きを直さないページは、これまでどおり座標に手を入れないこと。</summary>
    [Theory]
    [InlineData(0)]
    [InlineData(0.2)]
    public void AStraightPageKeepsTheBoxItWasRead(double skew)
    {
        using var dir = new TempDir();
        var pdf = dir.File("まっすぐ.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var ink = new OcrBox(300, 500, 240, 40);
        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.At("架空商事", 0.99, ink))
            .Probe(1, skew: skew, horizontal: 0, vertical: 0);

        var item = Assert.Single(new PdfScanReader().Read(engine, pdf, [1]).Items);

        Assert.Equal(ink.X, item.BoundingBox.X, 6);
        Assert.Equal(ink.Y, item.BoundingBox.Y, 6);
        Assert.Equal(ink.Width, item.BoundingBox.Width, 6);
        Assert.Equal(ink.Height, item.BoundingBox.Height, 6);
    }
}
