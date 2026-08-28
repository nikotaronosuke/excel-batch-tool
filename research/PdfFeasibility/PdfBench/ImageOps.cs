using SkiaSharp;

namespace PdfBench;

/// <summary>スキャン劣化の再現(ぼかし・傾き・ノイズ・かすれ)。</summary>
public static class ImageOps
{
    public static SKBitmap Degrade(SKBitmap source, Random random)
    {
        var result = new SKBitmap(source.Width, source.Height, SKColorType.Bgra8888, SKAlphaType.Premul);
        using (var canvas = new SKCanvas(result))
        {
            canvas.Clear(SKColors.White);

            // 1〜3 度の傾き。
            var angle = (float)(1 + random.NextDouble() * 2) * (random.Next(2) == 0 ? 1 : -1);
            canvas.RotateDegrees(angle, source.Width / 2f, source.Height / 2f);

            // かすれ(薄い文字)+ ぼかし。
            using var paint = new SKPaint
            {
                Color = SKColors.White.WithAlpha(215), // 全体を少し薄くする
                ImageFilter = SKImageFilter.CreateBlur(1.1f, 1.1f),
            };

            canvas.DrawBitmap(source, 0, 0, paint);
        }

        // ごま塩ノイズ。
        var pixels = result.Pixels;
        var count = pixels.Length / 220;
        for (var i = 0; i < count; i++)
        {
            var index = random.Next(pixels.Length);
            pixels[index] = random.Next(2) == 0 ? new SKColor(60, 60, 60) : SKColors.White;
        }

        result.Pixels = pixels;
        return result;
    }

    /// <summary>PDF ポイント座標を、指定 dpi で描画した画像のピクセル矩形へ変換する。</summary>
    public static SKRectI ToPixels(float x, float y, float width, float height, int dpi, int imageHeight)
    {
        var scale = dpi / 72f;
        return new SKRectI(
            (int)(x * scale),
            (int)(y * scale),
            (int)((x + width) * scale),
            (int)((y + height) * scale));
    }
}
