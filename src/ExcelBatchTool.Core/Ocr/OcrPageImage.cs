namespace ExcelBatchTool.Core.Ocr;

/// <summary>
/// 確認用に描いた 1 ページの画像。
///
/// OCR は 300dpi で読むが、画面に出すのはもっと粗くてよい。
/// <see cref="ScaleFromOcr"/> は「OCR の座標 → この画像の座標」の倍率で、
/// 読み取り位置の枠を画像の上へ重ねるときに使う。
/// </summary>
public sealed record OcrPageImage(int Page, byte[] Png, int Width, int Height, double ScaleFromOcr);

/// <summary>画面に置く矩形(左上からの位置と大きさ)。</summary>
public readonly record struct OcrDisplayRect(double Left, double Top, double Width, double Height);

/// <summary>
/// 読み取り位置を、画面に出している画像の上のどこへ描くか。
///
/// 座標は「OCR の画素 → 画像の画素 → 画面の画素」と 2 段階で縮む。
/// 画面へ出すのは <see cref="OcrPageImage.Width"/> に拡大率を掛けた大きさなので、
/// 倍率はその 2 つを掛けたものになる。
/// </summary>
public static class OcrBoxMapper
{
    /// <summary>枠が細すぎて見えなくならないように、最低限の太さを持たせる。</summary>
    private const double MinimumSize = 4;

    public static OcrDisplayRect ToDisplay(OcrBox box, OcrPageImage image, double zoom)
    {
        var scale = image.ScaleFromOcr * zoom;

        return new OcrDisplayRect(
            box.X * scale,
            box.Y * scale,
            Math.Max(box.Width * scale, MinimumSize),
            Math.Max(box.Height * scale, MinimumSize));
    }

    /// <summary>画像全体が入る拡大率。</summary>
    public static double FitZoom(OcrPageImage image, double viewportWidth, double viewportHeight)
    {
        if (image.Width <= 0 || image.Height <= 0 || viewportWidth <= 0 || viewportHeight <= 0)
        {
            return 1;
        }

        return Math.Min(viewportWidth / image.Width, viewportHeight / image.Height);
    }

    /// <summary>枠が見えるところへ来るよう、表示位置(スクロール量)を決める。</summary>
    public static (double Left, double Top) ScrollToShow(
        OcrDisplayRect rect, double viewportWidth, double viewportHeight)
    {
        // 枠を真ん中に置く。行き過ぎは呼び出し側(ScrollViewer)が丸める。
        return (
            rect.Left + (rect.Width / 2) - (viewportWidth / 2),
            rect.Top + (rect.Height / 2) - (viewportHeight / 2));
    }
}
