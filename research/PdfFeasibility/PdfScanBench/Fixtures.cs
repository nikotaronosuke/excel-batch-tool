using System.Globalization;
using SkiaSharp;

/// <summary>Phase 2F-B2 の測定に使う架空 PDF を作る。第三者のデータは使わない。</summary>
internal static class Fixtures
{
    public const float PageWidth = 595.28f;
    public const float PageHeight = 841.89f;

    private static SKTypeface Japanese()
        => SKFontManager.Default.MatchCharacter('あ') ?? SKTypeface.Default;

    /// <summary>文章のページを、指定した角度だけ傾けたスキャン PDF。</summary>
    public static void TiltedText(string path, double[] degrees, int dpi = 300)
    {
        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);

        foreach (var angle in degrees)
        {
            using var flat = RenderText();
            using var tilted = Tilt(flat, angle);
            DrawImagePage(document, tilted);
        }

        document.Close();
    }

    /// <summary>罫線のある表のスキャン PDF。</summary>
    public static void RuledTable(
        string path, int pages, int rowsPerPage, double tiltDegrees = 0, bool degraded = false)
        => TablePdf(path, pages, rowsPerPage, lined: true, tiltDegrees, degraded);

    /// <summary>罫線のない表のスキャン PDF。</summary>
    public static void BorderlessTable(
        string path, int pages, int rowsPerPage, double tiltDegrees = 0)
        => TablePdf(path, pages, rowsPerPage, lined: false, tiltDegrees, degraded: false);

    /// <summary>同じ様式の帳票が続くスキャン PDF。ページごとにずれ・傾きを混ぜられる。</summary>
    public static void Forms(
        string path,
        int pages,
        out List<Dictionary<string, string>> truth,
        double shift = 0,
        double tiltDegrees = 0,
        double scale = 1.0)
    {
        truth = [];
        var random = new Random(2026);

        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);

        for (var page = 1; page <= pages; page++)
        {
            var fields = new Dictionary<string, string>
            {
                ["店舗コード"] = $"S{page:D3}-{random.Next(10, 99)}",
                ["担当者"] = new[] { "架空 太郎", "架空 花子", "架空 一郎" }[page % 3],
                ["日付"] = $"2026/{(page % 12) + 1:D2}/{(page % 27) + 1:D2}",
                ["売上"] = (1000 + (page * 137)).ToString("N0", CultureInfo.InvariantCulture),
            };

            truth.Add(fields);

            var dx = shift == 0 ? 0 : ((page % 3) - 1) * shift;
            var dy = shift == 0 ? 0 : ((page % 2) - 0.5) * shift * 2;
            var tilt = tiltDegrees == 0 ? 0 : ((page % 3) - 1) * tiltDegrees;

            using var flat = RenderForm(fields, dx, dy, scale, page);
            using var final = tilt == 0 ? flat.Copy() : Tilt(flat, tilt);
            DrawImagePage(document, final);
        }

        document.Close();
    }

    /// <summary>選択肢に印の付いた帳票。</summary>
    public static void Marks(
        string path, int pages, out List<string> truth, bool degraded = false, double tilt = 0)
    {
        truth = [];
        string[] options = ["はい", "いいえ", "未回答"];
        string[] styles = ["check", "fill", "circle", "cross", "none"];

        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);

        for (var page = 1; page <= pages; page++)
        {
            var style = styles[page % styles.Length];
            var chosen = style == "none" ? string.Empty : options[page % options.Length];
            truth.Add(chosen);

            using var flat = RenderMarks(options, chosen, style);
            using var tilted = tilt == 0 ? flat.Copy() : Tilt(flat, tilt);
            using var final = degraded ? Degrade(tilted) : tilted.Copy();
            DrawImagePage(document, final);
        }

        document.Close();
    }

    /// <summary>選択肢の位置(OCR 300dpi の座標)。</summary>
    public static (string Label, float X, float Y, float Size)[] MarkBoxes()
    {
        const float scale = 300f / 72f;
        return
        [
            ("はい", 80 * scale, 200 * scale, 12 * scale),
            ("いいえ", 200 * scale, 200 * scale, 12 * scale),
            ("未回答", 330 * scale, 200 * scale, 12 * scale),
        ];
    }

    /// <summary>帳票の項目の位置(OCR 300dpi の座標)。</summary>
    public static (string Name, float X, float Y, float Width, float Height)[] FormAreas()
    {
        const float scale = 300f / 72f;
        return
        [
            ("店舗コード", 170 * scale, 108 * scale, 180 * scale, 20 * scale),
            ("担当者", 170 * scale, 138 * scale, 180 * scale, 20 * scale),
            ("日付", 170 * scale, 168 * scale, 180 * scale, 20 * scale),
            ("売上", 170 * scale, 198 * scale, 180 * scale, 20 * scale),
        ];
    }

    public static (string Text, float X, float Y, float Width, float Height) AnchorArea()
    {
        const float scale = 300f / 72f;
        return ("店舗コード", 60 * scale, 108 * scale, 100 * scale, 20 * scale);
    }

    // ── 描画 ─────────────────────────────────────────

    private static SKBitmap RenderText()
    {
        var bitmap = NewPage(out var canvas, out var font, out var paint);
        using (canvas)
        {
            var y = 120f;
            foreach (var line in new[]
            {
                "月次報告書(架空の見本)",
                "会社名: 架空商事株式会社",
                "金額: 4,917,087 円",
                "電話: 000-5072-4291",
                "日付: 2026/02/23",
                "本書は動作確認のために生成した架空の文書です。",
            })
            {
                canvas.DrawText(line, 60 * Scale, y * Scale, SKTextAlign.Left, font, paint);
                y += 30;
            }
        }

        font.Dispose();
        paint.Dispose();
        return bitmap;
    }

    private static SKBitmap RenderForm(
        Dictionary<string, string> fields, double dx, double dy, double scale, int page)
    {
        var bitmap = NewPage(out var canvas, out var font, out var paint);
        using (canvas)
        {
            using var line = new SKPaint
            {
                Color = SKColors.Black, IsStroke = true, StrokeWidth = 0.8f,
            };

            canvas.Save();
            canvas.Translate((float)(dx * Scale), (float)(dy * Scale));
            if (scale != 1.0)
            {
                canvas.Scale((float)scale, (float)scale, PageWidth * Scale / 2, PageHeight * Scale / 2);
            }

            canvas.DrawText(
                $"売上報告書(架空 {page})", 60 * Scale, 70 * Scale, SKTextAlign.Left, font, paint);

            var y = 120f;
            foreach (var (label, value) in fields)
            {
                canvas.DrawText(label + ":", 60 * Scale, y * Scale, SKTextAlign.Left, font, paint);
                canvas.DrawText(value, 170 * Scale, y * Scale, SKTextAlign.Left, font, paint);

                // 記入欄の下線(表の罫線ではない)。
                canvas.DrawLine(
                    165 * Scale, (y + 4) * Scale, 400 * Scale, (y + 4) * Scale, line);
                y += 30;
            }

            canvas.Restore();
        }

        font.Dispose();
        paint.Dispose();
        return bitmap;
    }

    private static SKBitmap RenderMarks(string[] options, string chosen, string style)
    {
        var bitmap = NewPage(out var canvas, out var font, out var paint);
        using (canvas)
        {
            using var stroke = new SKPaint
            {
                Color = SKColors.Black, IsStroke = true, StrokeWidth = 1.2f, IsAntialias = true,
            };
            using var fill = new SKPaint { Color = SKColors.Black, IsAntialias = true };

            canvas.DrawText("回答:", 60 * Scale, 195 * Scale, SKTextAlign.Left, font, paint);

            foreach (var (label, x, y, size) in MarkBoxes())
            {
                var boxX = x / Scale;
                var boxY = y / Scale;
                var boxSize = size / Scale;

                canvas.DrawRect(
                    new SKRect(boxX * Scale, boxY * Scale,
                        (boxX + boxSize) * Scale, (boxY + boxSize) * Scale), stroke);
                canvas.DrawText(
                    label, (boxX + boxSize + 4) * Scale, (boxY + boxSize - 1) * Scale,
                    SKTextAlign.Left, font, paint);

                if (label != chosen)
                {
                    continue;
                }

                switch (style)
                {
                    case "check":
                        using (var thick = new SKPaint
                        {
                            Color = SKColors.Black, IsStroke = true, StrokeWidth = PenWidth,
                            IsAntialias = true,
                        })
                        {
                            canvas.DrawLine(
                                (boxX + 2) * Scale, (boxY + boxSize / 2) * Scale,
                                (boxX + boxSize / 2) * Scale, (boxY + boxSize - 2) * Scale, thick);
                            canvas.DrawLine(
                                (boxX + boxSize / 2) * Scale, (boxY + boxSize - 2) * Scale,
                                (boxX + boxSize - 1) * Scale, (boxY + 1) * Scale, thick);
                        }

                        break;

                    case "fill":
                        canvas.DrawRect(
                            new SKRect((boxX + 1) * Scale, (boxY + 1) * Scale,
                                (boxX + boxSize - 1) * Scale, (boxY + boxSize - 1) * Scale), fill);
                        break;

                    case "cross":
                        using (var thick = new SKPaint
                        {
                            Color = SKColors.Black, IsStroke = true, StrokeWidth = PenWidth,
                            IsAntialias = true,
                        })
                        {
                            canvas.DrawLine(
                                (boxX + 1) * Scale, (boxY + 1) * Scale,
                                (boxX + boxSize - 1) * Scale, (boxY + boxSize - 1) * Scale, thick);
                            canvas.DrawLine(
                                (boxX + boxSize - 1) * Scale, (boxY + 1) * Scale,
                                (boxX + 1) * Scale, (boxY + boxSize - 1) * Scale, thick);
                        }

                        break;

                    case "circle":
                        using (var ring = new SKPaint
                        {
                            Color = SKColors.Black, IsStroke = true, StrokeWidth = PenWidth,
                            IsAntialias = true,
                        })
                        {
                            canvas.DrawOval(
                                new SKRect((boxX + boxSize + 1) * Scale, (boxY - 4) * Scale,
                                    (boxX + boxSize + 48) * Scale, (boxY + boxSize + 4) * Scale),
                                ring);
                        }

                        break;
                }
            }
        }

        font.Dispose();
        paint.Dispose();
        return bitmap;
    }

    private static void TablePdf(
        string path, int pages, int rowsPerPage, bool lined, double tiltDegrees, bool degraded)
    {
        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);

        for (var page = 1; page <= pages; page++)
        {
            using var flat = RenderTable(page, rowsPerPage, lined);
            using var tilted = tiltDegrees == 0 ? flat.Copy() : Tilt(flat, tiltDegrees);
            using var final = degraded ? Degrade(tilted) : tilted.Copy();
            DrawImagePage(document, final);
        }

        document.Close();
    }

    private static SKBitmap RenderTable(int page, int rows, bool lined)
    {
        var bitmap = NewPage(out var canvas, out var font, out var paint);
        using (canvas)
        {
            using var line = new SKPaint
            {
                Color = SKColors.Black, IsStroke = true, StrokeWidth = 0.9f,
            };

            float[] columns = [60, 170, 380, 470];
            const float right = 540;
            const float top = 110;
            const float rowHeight = 18;

            string[] names = ["架空りんご", "架空みかん", "架空ぶどう", "架空の緑茶", "架空ノート"];
            var all = new List<string[]> { new[] { "商品コード", "商品名", "単価", "在庫" } };

            for (var index = 1; index <= rows; index++)
            {
                var serial = ((page - 1) * rows) + index;
                all.Add(new[]
                {
                    $"A{serial:D4}",
                    // 5 行に 1 つは空欄のセルを混ぜる。
                    serial % 5 == 0 ? string.Empty : names[serial % names.Length],
                    (1000 + (serial * 37)).ToString("N0", CultureInfo.InvariantCulture),
                    (serial * 3).ToString(CultureInfo.InvariantCulture),
                });
            }

            for (var r = 0; r < all.Count; r++)
            {
                var y = top + (r * rowHeight);
                for (var c = 0; c < 4; c++)
                {
                    canvas.DrawText(
                        all[r][c], (columns[c] + 3) * Scale, (y + 12) * Scale,
                        SKTextAlign.Left, font, paint);
                }

                if (lined)
                {
                    canvas.DrawLine(columns[0] * Scale, y * Scale, right * Scale, y * Scale, line);
                }
            }

            if (lined)
            {
                var bottom = top + (all.Count * rowHeight);
                canvas.DrawLine(
                    columns[0] * Scale, bottom * Scale, right * Scale, bottom * Scale, line);
                foreach (var x in columns.Append(right))
                {
                    canvas.DrawLine(x * Scale, top * Scale, x * Scale, bottom * Scale, line);
                }
            }
        }

        font.Dispose();
        paint.Dispose();
        return bitmap;
    }

    /// <summary>表の正解(ページごとの行)。</summary>
    public static List<string[]> TableTruth(int page, int rows)
    {
        string[] names = ["架空りんご", "架空みかん", "架空ぶどう", "架空の緑茶", "架空ノート"];
        var all = new List<string[]> { new[] { "商品コード", "商品名", "単価", "在庫" } };

        for (var index = 1; index <= rows; index++)
        {
            var serial = ((page - 1) * rows) + index;
            all.Add(new[]
            {
                $"A{serial:D4}",
                serial % 5 == 0 ? string.Empty : names[serial % names.Length],
                (1000 + (serial * 37)).ToString("N0", CultureInfo.InvariantCulture),
                (serial * 3).ToString(CultureInfo.InvariantCulture),
            });
        }

        return all;
    }

    // ── 画像の下ごしらえ ──────────────────────────────

    private const float Scale = 300f / 72f;

    /// <summary>手書きの印の線の太さ。ボールペン約 0.5mm 相当(300dpi で約 6 画素)。</summary>
    private const float PenWidth = 6f;

    private static SKBitmap NewPage(out SKCanvas canvas, out SKFont font, out SKPaint paint)
    {
        var bitmap = new SKBitmap((int)(PageWidth * Scale), (int)(PageHeight * Scale));
        canvas = new SKCanvas(bitmap);
        canvas.Clear(SKColors.White);
        font = new SKFont(Japanese(), 10 * Scale);
        paint = new SKPaint { Color = SKColors.Black, IsAntialias = true };
        return bitmap;
    }

    /// <summary>指定した角度だけ傾ける(スキャンで曲がって入った状態を作る)。</summary>
    private static SKBitmap Tilt(SKBitmap source, double degrees)
    {
        var result = new SKBitmap(source.Width, source.Height);
        using var canvas = new SKCanvas(result);
        canvas.Clear(SKColors.White);
        canvas.RotateDegrees((float)degrees, source.Width / 2f, source.Height / 2f);
        canvas.DrawBitmap(source, 0, 0, new SKSamplingOptions(SKFilterMode.Linear));
        return result;
    }

    /// <summary>解像度を落として少しぼかす(取り込みの粗い状態を作る)。</summary>
    private static SKBitmap Degrade(SKBitmap source)
    {
        var half = new SKBitmap(source.Width / 2, source.Height / 2);
        using (var canvas = new SKCanvas(half))
        {
            canvas.Clear(SKColors.White);
            canvas.DrawBitmap(
                source, new SKRect(0, 0, half.Width, half.Height),
                new SKSamplingOptions(SKFilterMode.Linear));
        }

        var result = new SKBitmap(source.Width, source.Height);
        using (var canvas = new SKCanvas(result))
        {
            canvas.Clear(SKColors.White);
            canvas.DrawBitmap(
                half, new SKRect(0, 0, result.Width, result.Height),
                new SKSamplingOptions(SKFilterMode.Linear));
        }

        half.Dispose();
        return result;
    }

    private static void DrawImagePage(SKDocument document, SKBitmap bitmap)
    {
        using var image = SKImage.FromBitmap(bitmap);
        var canvas = document.BeginPage(PageWidth, PageHeight);
        canvas.Clear(SKColors.White);
        canvas.DrawImage(
            image, new SKRect(0, 0, PageWidth, PageHeight),
            new SKSamplingOptions(SKFilterMode.Linear));
        document.EndPage();
    }
}
