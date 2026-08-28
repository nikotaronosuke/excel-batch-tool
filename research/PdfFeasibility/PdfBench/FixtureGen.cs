using SkiaSharp;

namespace PdfBench;

/// <summary>
/// 架空データだけでテスト用 PDF を生成する。
/// born-digital(埋め込みテキスト)は SkiaSharp の PDF バックエンドで作る。
/// スキャン PDF は、ページを画像化 → 劣化 → 画像だけの PDF に包んで作る。
/// </summary>
public static class FixtureGen
{
    private static SKTypeface Jp(string? preferred = null)
    {
        foreach (var name in new[]
        {
            preferred, "UD デジタル 教科書体 N-R", "Yu Gothic UI", "Yu Gothic", "Meiryo", "MS Gothic",
        })
        {
            if (name is null)
            {
                continue;
            }

            var face = SKTypeface.FromFamilyName(name);
            if (face is not null && face.GetGlyph('あ') != 0)
            {
                return face;
            }
        }

        // 最後の手段: 「あ」を描けるフォントをシステムから探す。
        return SKFontManager.Default.MatchCharacter('あ') ?? SKTypeface.Default;
    }

    public static void GenerateAll(string dir)
    {
        Directory.CreateDirectory(dir);
        Directory.CreateDirectory(Path.Combine(dir, "gt"));

        GenerateTextPdf(dir, pages: 10);
        GenerateTablePdf(dir, "table-lined.pdf", "table.json", lined: true, pages: 10);
        GenerateTablePdf(dir, "table-borderless.pdf", "table-borderless.json", lined: false, pages: 10);
        GenerateForms(dir, pages: 120);
        GenerateHandwritingLike(dir, pages: 5);

        // スキャン版: きれい(300dpi)と劣化(150dpi + ぼかし + 傾き + ノイズ + JPEG)。
        Scanify(dir, "form.pdf", "form-scan-clean.pdf", 300, degrade: false, maxPages: 120);
        Scanify(dir, "form.pdf", "form-scan-degraded.pdf", 150, degrade: true, maxPages: 30);
        Scanify(dir, "table-lined.pdf", "table-scan-clean.pdf", 300, degrade: false, maxPages: 10);
        Scanify(dir, "table-lined.pdf", "table-scan-degraded.pdf", 150, degrade: true, maxPages: 5);
        Scanify(dir, "text.pdf", "text-scan-clean.pdf", 300, degrade: false, maxPages: 5);
        Scanify(dir, "text.pdf", "text-scan-degraded.pdf", 150, degrade: true, maxPages: 5);
        Scanify(dir, "handwriting.pdf", "handwriting-scan.pdf", 300, degrade: false, maxPages: 5);

        Console.WriteLine("fixtures done: " + string.Join(", ",
            Directory.GetFiles(dir, "*.pdf").Select(Path.GetFileName).Order()));
    }

    /// <summary>A. born-digital の文章 + フィールド(会社名・金額・電話・日付・複数段)。</summary>
    private static void GenerateTextPdf(string dir, int pages)
    {
        var random = Layouts.NewRandom(101);
        var gt = new List<TextPageGt>();

        using var stream = File.Create(Path.Combine(dir, "text.pdf"));
        using var document = SKDocument.CreatePdf(stream);
        using var face = Jp();
        using var font = new SKFont(face, 11);
        using var titleFont = new SKFont(face, 16);
        using var paint = new SKPaint { Color = SKColors.Black, IsAntialias = true };

        for (var page = 1; page <= pages; page++)
        {
            var canvas = document.BeginPage(Layouts.PageWidth, Layouts.PageHeight);
            canvas.Clear(SKColors.White);

            var fields = new Dictionary<string, string>
            {
                ["会社名"] = $"架空商事株式会社 第{page}支店",
                ["金額"] = Layouts.Money(random) + " 円",
                ["電話"] = Layouts.Phone(random),
                ["日付"] = Layouts.DateText(random),
            };

            canvas.DrawText($"月次報告書(見本 {page})", 60, 80, SKTextAlign.Left, titleFont, paint);

            var y = 130f;
            foreach (var (name, value) in fields)
            {
                canvas.DrawText($"{name}: {value}", 60, y, SKTextAlign.Left, font, paint);
                y += 24;
            }

            // 2 段組の本文(改行あり)。
            string[] body =
            [
                "本書は動作確認のために生成した架空の文書です。",
                "記載の数値・氏名・連絡先はすべて実在しません。",
                "第一段落では商品の入荷と検品について記載します。",
                "第二段落では翌月の予定を記載します。",
            ];

            y = 280;
            foreach (var line in body)
            {
                canvas.DrawText(line, 60, y, SKTextAlign.Left, font, paint);
                canvas.DrawText(line, 320, y, SKTextAlign.Left, font, paint);
                y += 20;
            }

            document.EndPage();
            gt.Add(new TextPageGt(page, fields));
        }

        document.Close();
        Json.Save(Path.Combine(dir, "gt", "text.json"), gt);
    }

    /// <summary>B/C. born-digital の表(罫線あり / なし)。</summary>
    private static void GenerateTablePdf(string dir, string fileName, string gtName, bool lined, int pages)
    {
        var random = Layouts.NewRandom(lined ? 202 : 203);
        var gt = new List<TablePageGt>();

        using var stream = File.Create(Path.Combine(dir, fileName));
        using var document = SKDocument.CreatePdf(stream);
        using var face = Jp();
        using var font = new SKFont(face, 10);
        using var text = new SKPaint { Color = SKColors.Black, IsAntialias = true };
        using var line = new SKPaint
        {
            Color = SKColors.Black, IsAntialias = false, IsStroke = true, StrokeWidth = 0.7f,
        };

        var rowIndex = 0;
        for (var page = 1; page <= pages; page++)
        {
            var canvas = document.BeginPage(Layouts.PageWidth, Layouts.PageHeight);
            canvas.Clear(SKColors.White);

            var rows = new List<string[]> { Layouts.TableHeaders };
            for (var i = 0; i < Layouts.TableRowsPerPage; i++)
            {
                rowIndex++;
                rows.Add(
                [
                    Layouts.ProductCode(rowIndex),
                    Layouts.ProductName(random),
                    random.Next(100, 99999).ToString("N0"),
                    random.Next(0, 500).ToString(),
                ]);
            }

            for (var r = 0; r < rows.Count; r++)
            {
                var y = Layouts.TableTop + r * Layouts.TableRowHeight;
                for (var c = 0; c < 4; c++)
                {
                    canvas.DrawText(rows[r][c], Layouts.TableColumnX[c] + 3, y + 12,
                        SKTextAlign.Left, font, text);
                }

                if (lined)
                {
                    canvas.DrawLine(Layouts.TableColumnX[0], y, Layouts.TableRight, y, line);
                }
            }

            if (lined)
            {
                var bottom = Layouts.TableTop + rows.Count * Layouts.TableRowHeight;
                canvas.DrawLine(Layouts.TableColumnX[0], bottom, Layouts.TableRight, bottom, line);
                foreach (var x in Layouts.TableColumnX.Append(Layouts.TableRight))
                {
                    canvas.DrawLine(x, Layouts.TableTop, x, bottom, line);
                }
            }

            document.EndPage();
            gt.Add(new TablePageGt(page, rows));
        }

        document.Close();
        Json.Save(Path.Combine(dir, "gt", gtName), gt);
    }

    /// <summary>F/G. 同一レイアウトの固定帳票 + チェックボックス。</summary>
    private static void GenerateForms(string dir, int pages)
    {
        var random = Layouts.NewRandom(303);
        var gt = new List<FormPageGt>();

        using var stream = File.Create(Path.Combine(dir, "form.pdf"));
        using var document = SKDocument.CreatePdf(stream);
        using var face = Jp();
        using var font = new SKFont(face, 12);
        using var titleFont = new SKFont(face, 15);
        using var paint = new SKPaint { Color = SKColors.Black, IsAntialias = true };
        using var box = new SKPaint
        {
            Color = SKColors.Black, IsStroke = true, StrokeWidth = 1.2f, IsAntialias = false,
        };
        using var mark = new SKPaint
        {
            Color = SKColors.Black, IsStroke = true, StrokeWidth = 2.0f, IsAntialias = true,
        };

        for (var page = 1; page <= pages; page++)
        {
            var canvas = document.BeginPage(Layouts.PageWidth, Layouts.PageHeight);
            canvas.Clear(SKColors.White);

            canvas.DrawText("店舗日次報告(架空の見本帳票)", 60, 90, SKTextAlign.Left, titleFont, paint);

            var fields = new Dictionary<string, string>
            {
                ["店舗コード"] = Layouts.StoreCode(random),
                ["担当者"] = Layouts.Person(random),
                ["日付"] = Layouts.DateText(random),
                ["売上"] = Layouts.Money(random),
                ["Q1"] = random.Next(1, 6).ToString(),
                ["Q2"] = random.Next(1, 6).ToString(),
                ["備考"] = Layouts.Remark(random),
            };

            foreach (var (name, _, valueX, y, width) in Layouts.FormFields)
            {
                canvas.DrawText(name + ":", 60, y, SKTextAlign.Left, font, paint);
                canvas.DrawLine(valueX, y + 4, valueX + width, y + 4, box);
                if (fields[name].Length > 0)
                {
                    canvas.DrawText(fields[name], valueX + 4, y, SKTextAlign.Left, font, paint);
                }
            }

            // チェックボックス。マークは チェック / 塗りつぶし / 丸囲み の 3 種類 + 無印。
            var selected = random.Next(4); // 0..2 = 選択肢、3 = 無印
            var markKind = random.Next(3);
            canvas.DrawText("Q3:", 60, Layouts.CheckY + Layouts.CheckBoxSize - 2,
                SKTextAlign.Left, font, paint);

            for (var option = 0; option < Layouts.CheckOptions.Length; option++)
            {
                var (label, boxX) = Layouts.CheckOptions[option];
                var rect = new SKRect(
                    boxX, Layouts.CheckY, boxX + Layouts.CheckBoxSize, Layouts.CheckY + Layouts.CheckBoxSize);
                canvas.DrawRect(rect, box);
                canvas.DrawText(label, boxX + Layouts.CheckBoxSize + 5,
                    Layouts.CheckY + Layouts.CheckBoxSize - 2, SKTextAlign.Left, font, paint);

                if (option != selected)
                {
                    continue;
                }

                switch (markKind)
                {
                    case 0: // チェック
                        canvas.DrawLine(rect.Left + 2, rect.MidY, rect.MidX - 1, rect.Bottom - 3, mark);
                        canvas.DrawLine(rect.MidX - 1, rect.Bottom - 3, rect.Right - 1, rect.Top + 2, mark);
                        break;
                    case 1: // 塗りつぶし
                        using (var fill = new SKPaint { Color = SKColors.Black })
                        {
                            canvas.DrawRect(SKRect.Inflate(rect, -2.5f, -2.5f), fill);
                        }

                        break;
                    default: // ラベルの丸囲み
                        canvas.DrawOval(new SKRect(
                            boxX + Layouts.CheckBoxSize + 1, Layouts.CheckY - 4,
                            boxX + Layouts.CheckBoxSize + 48, Layouts.CheckY + Layouts.CheckBoxSize + 4), mark);
                        break;
                }
            }

            document.EndPage();
            gt.Add(new FormPageGt(page, fields,
                selected == 3 ? "none" : Layouts.CheckOptions[selected].Label));
        }

        document.Close();
        Json.Save(Path.Combine(dir, "gt", "form.json"), gt);
    }

    /// <summary>I. 手書き風(教科書体 + 文字ごとの揺らぎ)。characterization のみ。</summary>
    private static void GenerateHandwritingLike(string dir, int pages)
    {
        var random = Layouts.NewRandom(404);
        var gt = new List<TextPageGt>();

        using var stream = File.Create(Path.Combine(dir, "handwriting.pdf"));
        using var document = SKDocument.CreatePdf(stream);
        using var face = Jp("UD Digi Kyokasho N-R");
        using var font = new SKFont(face, 14);
        using var paint = new SKPaint { Color = new SKColor(20, 20, 40), IsAntialias = true };

        for (var page = 1; page <= pages; page++)
        {
            var canvas = document.BeginPage(Layouts.PageWidth, Layouts.PageHeight);
            canvas.Clear(SKColors.White);

            var fields = new Dictionary<string, string>
            {
                ["氏名"] = Layouts.Person(random),
                ["金額"] = Layouts.Money(random),
                ["メモ"] = "あすは 雨のち くもり",
            };

            var y = 150f;
            foreach (var (name, value) in fields)
            {
                DrawJittered(canvas, font, paint, random, $"{name}: {value}", 70, y);
                y += 60;
            }

            document.EndPage();
            gt.Add(new TextPageGt(page, fields));
        }

        document.Close();
        Json.Save(Path.Combine(dir, "gt", "handwriting.json"), gt);
    }

    /// <summary>文字ごとに位置と角度を揺らして、手書きらしい乱れを作る。</summary>
    private static void DrawJittered(
        SKCanvas canvas, SKFont font, SKPaint paint, Random random, string text, float x, float y)
    {
        foreach (var character in text)
        {
            var s = character.ToString();
            canvas.Save();
            canvas.RotateDegrees(
                (float)(random.NextDouble() * 8 - 4), x, y);
            canvas.DrawText(s, x, y + (float)(random.NextDouble() * 3 - 1.5), SKTextAlign.Left, font, paint);
            canvas.Restore();
            x += font.MeasureText(s) + (float)(random.NextDouble() * 2);
        }
    }

    /// <summary>D/E/H. born-digital の PDF を画像化して、画像だけの PDF(スキャン相当)にする。</summary>
    private static void Scanify(
        string dir, string sourceName, string outputName, int dpi, bool degrade, int maxPages)
    {
        var sourcePath = Path.Combine(dir, sourceName);
        var random = Layouts.NewRandom(505);

        using var output = File.Create(Path.Combine(dir, outputName));
        using var document = SKDocument.CreatePdf(output);

        var bytes = File.ReadAllBytes(sourcePath);
        var pageCount = Math.Min(PDFtoImage.Conversion.GetPageCount(bytes), maxPages);

        for (var page = 0; page < pageCount; page++)
        {
            using var rendered = PDFtoImage.Conversion.ToImage(
                bytes, page: page, options: new PDFtoImage.RenderOptions(Dpi: dpi));

            using var final = degrade ? ImageOps.Degrade(rendered, random) : rendered.Copy();

            // JPEG にしてから画像だけのページとして貼る(スキャナー出力の再現)。
            using var image = SKImage.FromBitmap(final);
            using var jpeg = image.Encode(SKEncodedImageFormat.Jpeg, degrade ? 40 : 80);
            using var decoded = SKImage.FromEncodedData(jpeg);

            var canvas = document.BeginPage(Layouts.PageWidth, Layouts.PageHeight);
            canvas.Clear(SKColors.White);
            canvas.DrawImage(decoded, new SKRect(0, 0, Layouts.PageWidth, Layouts.PageHeight));
            document.EndPage();
        }

        document.Close();
    }
}
