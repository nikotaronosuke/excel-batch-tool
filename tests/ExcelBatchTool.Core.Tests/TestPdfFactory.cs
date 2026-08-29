using SkiaSharp;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// テスト用の PDF を架空データだけで生成する。第三者の PDF は使わない。
/// 埋め込みテキストの PDF は SkiaSharp の PDF バックエンドで作り、
/// 画像だけ(スキャン相当)のページは、描いた内容を画像にして貼り直して作る。
/// </summary>
internal static class TestPdfFactory
{
    public const float PageWidth = 595.28f;
    public const float PageHeight = 841.89f;

    /// <summary>日本語のグリフを実際に持つフォントを選ぶ。</summary>
    private static SKTypeface Japanese()
    {
        foreach (var name in new[] { "Yu Gothic UI", "Yu Gothic", "Meiryo", "MS Gothic" })
        {
            var face = SKTypeface.FromFamilyName(name);
            if (face is not null && new SKFont(face).GetGlyph('あ') != 0)
            {
                return face;
            }
        }

        return SKFontManager.Default.MatchCharacter('あ') ?? SKTypeface.Default;
    }

    /// <summary>1 ページに数行だけ書いた、通常の文字 PDF。</summary>
    public static void CreateText(string path, IReadOnlyList<IReadOnlyList<string>> pages)
    {
        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);
        using var face = Japanese();
        using var font = new SKFont(face, 12);
        using var paint = new SKPaint { Color = SKColors.Black, IsAntialias = true };

        foreach (var lines in pages)
        {
            var canvas = document.BeginPage(PageWidth, PageHeight);
            canvas.Clear(SKColors.White);

            var y = 100f;
            foreach (var line in lines)
            {
                canvas.DrawText(line, 60, y, SKTextAlign.Left, font, paint);
                y += 24;
            }

            document.EndPage();
        }

        document.Close();
    }

    /// <summary>2 段組の文字 PDF(左の段をすべて読んでから右の段、にはならない配置)。</summary>
    public static void CreateTwoColumnText(
        string path, IReadOnlyList<string> left, IReadOnlyList<string> right)
    {
        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);
        using var face = Japanese();
        using var font = new SKFont(face, 12);
        using var paint = new SKPaint { Color = SKColors.Black, IsAntialias = true };

        var canvas = document.BeginPage(PageWidth, PageHeight);
        canvas.Clear(SKColors.White);

        var y = 100f;
        for (var index = 0; index < Math.Max(left.Count, right.Count); index++)
        {
            if (index < left.Count)
            {
                canvas.DrawText(left[index], 60, y, SKTextAlign.Left, font, paint);
            }

            if (index < right.Count)
            {
                canvas.DrawText(right[index], 320, y, SKTextAlign.Left, font, paint);
            }

            y += 22;
        }

        document.EndPage();
        document.Close();
    }

    /// <summary>表の PDF。<paramref name="lined"/> が false なら罫線を引かない。</summary>
    public static void CreateTable(
        string path,
        IReadOnlyList<IReadOnlyList<string[]>> pages,
        bool lined = true,
        float[]? columnX = null)
    {
        var columns = columnX ?? [60, 170, 380, 470];
        const float right = 540;
        const float top = 110;
        const float rowHeight = 18;

        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);
        using var face = Japanese();
        using var font = new SKFont(face, 10);
        using var text = new SKPaint { Color = SKColors.Black, IsAntialias = true };
        using var line = new SKPaint
        {
            Color = SKColors.Black, IsAntialias = false, IsStroke = true, StrokeWidth = 0.7f,
        };

        foreach (var rows in pages)
        {
            var canvas = document.BeginPage(PageWidth, PageHeight);
            canvas.Clear(SKColors.White);

            for (var r = 0; r < rows.Count; r++)
            {
                var y = top + r * rowHeight;
                for (var c = 0; c < rows[r].Length && c < columns.Length; c++)
                {
                    canvas.DrawText(rows[r][c], columns[c] + 3, y + 12, SKTextAlign.Left, font, text);
                }

                if (lined)
                {
                    canvas.DrawLine(columns[0], y, right, y, line);
                }
            }

            if (lined)
            {
                var bottom = top + rows.Count * rowHeight;
                canvas.DrawLine(columns[0], bottom, right, bottom, line);
                foreach (var x in columns.Append(right))
                {
                    canvas.DrawLine(x, top, x, bottom, line);
                }
            }

            document.EndPage();
        }

        document.Close();
    }

    /// <summary>画像だけのページを持つ PDF(スキャン相当)。文字情報は入らない。</summary>
    public static void CreateImageOnly(string path, int pages)
        => CreateMixed(path, textPages: 0, imagePages: pages);

    /// <summary>文字のページと画像だけのページが混ざった PDF。</summary>
    public static void CreateMixed(string path, int textPages, int imagePages)
    {
        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);
        using var face = Japanese();
        using var font = new SKFont(face, 12);
        using var paint = new SKPaint { Color = SKColors.Black, IsAntialias = true };

        for (var page = 1; page <= textPages; page++)
        {
            var canvas = document.BeginPage(PageWidth, PageHeight);
            canvas.Clear(SKColors.White);
            canvas.DrawText($"文字のページ {page}(架空の見本)", 60, 100, SKTextAlign.Left, font, paint);
            canvas.DrawText("金額: 1,200 円", 60, 130, SKTextAlign.Left, font, paint);
            document.EndPage();
        }

        for (var page = 1; page <= imagePages; page++)
        {
            using var bitmap = new SKBitmap((int)PageWidth, (int)PageHeight);
            using (var bitmapCanvas = new SKCanvas(bitmap))
            {
                bitmapCanvas.Clear(SKColors.White);
                bitmapCanvas.DrawText(
                    $"画像のページ {page}(架空の見本)", 60, 100, SKTextAlign.Left, font, paint);
            }

            using var image = SKImage.FromBitmap(bitmap);
            var canvas = document.BeginPage(PageWidth, PageHeight);
            canvas.Clear(SKColors.White);
            canvas.DrawImage(
                image, new SKRect(0, 0, PageWidth, PageHeight), new SKSamplingOptions(SKFilterMode.Linear));
            document.EndPage();
        }

        document.Close();
    }

    /// <summary>罫線のある表のページ + 画像だけのページ(構造を揃えられない混在)。</summary>
    public static void CreateTableThenImage(string path, IReadOnlyList<string[]> rows)
    {
        var table = path + ".table.tmp.pdf";
        CreateTable(table, [rows]);

        using (var stream = File.Create(path))
        using (var document = SKDocument.CreatePdf(stream))
        {
            using var face = Japanese();
            using var font = new SKFont(face, 10);
            using var text = new SKPaint { Color = SKColors.Black, IsAntialias = true };
            using var line = new SKPaint
            {
                Color = SKColors.Black, IsStroke = true, StrokeWidth = 0.7f,
            };

            float[] columns = [60, 170, 380, 470];
            const float right = 540;
            const float top = 110;
            const float rowHeight = 18;

            var canvas = document.BeginPage(PageWidth, PageHeight);
            canvas.Clear(SKColors.White);
            for (var index = 0; index < rows.Count; index++)
            {
                var y = top + (index * rowHeight);
                for (var column = 0; column < columns.Length; column++)
                {
                    canvas.DrawText(
                        rows[index].ElementAtOrDefault(column) ?? string.Empty,
                        columns[column] + 3, y + 12, SKTextAlign.Left, font, text);
                }

                canvas.DrawLine(columns[0], y, right, y, line);
            }

            var bottom = top + (rows.Count * rowHeight);
            canvas.DrawLine(columns[0], bottom, right, bottom, line);
            foreach (var x in columns.Append(right))
            {
                canvas.DrawLine(x, top, x, bottom, line);
            }

            document.EndPage();

            using var bitmap = new SKBitmap((int)PageWidth, (int)PageHeight);
            using (var bitmapCanvas = new SKCanvas(bitmap))
            {
                bitmapCanvas.Clear(SKColors.White);
                bitmapCanvas.DrawText("画像のページ(架空の見本)", 60, 100, SKTextAlign.Left, font, text);
            }

            using var image = SKImage.FromBitmap(bitmap);
            var imageCanvas = document.BeginPage(PageWidth, PageHeight);
            imageCanvas.Clear(SKColors.White);
            imageCanvas.DrawImage(
                image, new SKRect(0, 0, PageWidth, PageHeight), new SKSamplingOptions(SKFilterMode.Linear));
            document.EndPage();
            document.Close();
        }

        File.Delete(table);
    }

    /// <summary>ページを 1 枚も持たない PDF。</summary>
    public static void CreateEmpty(string path)
    {
        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);
        document.Close();
    }

    /// <summary>壊れた PDF(先頭だけ PDF に見えて、中身が壊れている)。</summary>
    public static void CreateCorrupted(string path)
        => File.WriteAllBytes(path, System.Text.Encoding.ASCII.GetBytes(
            "%PDF-1.7\n1 0 obj\n<< /Type /Catalog >>\nendobj\ntrailer\n<< /Root 9 0 R >>\n%%EOF\n"));
}
