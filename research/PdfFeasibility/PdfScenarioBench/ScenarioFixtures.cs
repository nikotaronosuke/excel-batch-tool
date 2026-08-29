using System.Globalization;
using SkiaSharp;

namespace PdfScenarioBench;

/// <summary>
/// 実案件に近い形の PDF を作る。**すべて架空**で、生成コードから作る。
/// 第三者の実データ・実案件のファイル・個人情報は一切使わない。
///
/// 参考にしたのは、クラウドソーシングで人力発注されている
/// 「PDF から Excel へ」の作業の型(商品一覧・見積明細、アンケート、
/// 定型業務帳票、契約・申込書、混在 PDF、取り込みの粗い PDF)。
/// 募集文や発注者を特定できるものは持ち込まない。
/// </summary>
internal static class ScenarioFixtures
{
    private const float PageWidth = 595f;
    private const float PageHeight = 842f;
    private const float Scale = 300f / 72f;
    private const float PenWidth = 6f;

    private static SKTypeface Japanese()
        => SKFontManager.Default.MatchCharacter('あ') ?? SKTypeface.Default;

    // ── A: 商品一覧 / 見積明細(表 → Excel) ───────────────

    /// <summary>罫線あり / なしを混ぜた明細表。ページ見出しは毎ページ繰り返す。</summary>
    public static void ItemTable(
        string path, int pages, int rowsPerPage, out List<string[]> truth, bool ruled)
    {
        truth = [];
        string[] header = ["商品コード", "商品名", "区分", "数量", "単価", "金額"];
        truth.Add(header);

        var random = new Random(3001);
        string[] names =
        [
            "架空りんご", "架空みかん", "架空ぶどう", "架空の緑茶", "架空パン",
            "架空ノート", "架空ペン", "架空の封筒", "架空クリップ", "架空テープ",
        ];
        string[] kinds = ["食品", "文具", "雑貨"];

        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);

        var serial = 1;
        for (var page = 1; page <= pages; page++)
        {
            var rows = new List<string[]>();
            for (var row = 0; row < rowsPerPage; row++)
            {
                // 一部の行はわざと空欄を混ぜる(実務の表にはよくある)。
                var blank = serial % 17 == 0;
                var quantity = random.Next(1, 99);
                var unit = 100 + (serial * 37 % 9000);
                rows.Add(
                [
                    $"A{serial:D4}",
                    names[serial % names.Length],
                    blank ? string.Empty : kinds[serial % kinds.Length],
                    blank ? string.Empty : quantity.ToString(CultureInfo.InvariantCulture),
                    unit.ToString("N0", CultureInfo.InvariantCulture),
                    blank
                        ? string.Empty
                        : (quantity * unit).ToString("N0", CultureInfo.InvariantCulture),
                ]);

                serial++;
            }

            truth.AddRange(rows);

            using var bitmap = RenderTable(header, rows, ruled);
            DrawImagePage(document, bitmap);
        }

        document.Close();
    }

    // ── B: アンケート(選択 + 数値 + 短い印字自由記述) ──────────

    public static void Survey(
        string path, int pages, out List<Dictionary<string, string>> truth, int choiceCount)
    {
        truth = [];
        var random = new Random(3002);
        string[] answers = ["はい", "いいえ", "どちらでもない"];
        string[] comments =
        [
            "架空の意見です", "架空の要望です", "特にありません", "架空の感想",
        ];

        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);

        for (var page = 1; page <= pages; page++)
        {
            var fields = new Dictionary<string, string>
            {
                ["整理番号"] = $"Q{page:D5}",
                ["年齢"] = random.Next(18, 79).ToString(CultureInfo.InvariantCulture),
                ["自由記述"] = comments[page % comments.Length],
            };

            var chosen = new string[choiceCount];
            for (var q = 0; q < choiceCount; q++)
            {
                chosen[q] = answers[(page + q) % answers.Length];
                fields[$"設問{q + 1}"] = chosen[q];
            }

            truth.Add(fields);

            using var bitmap = RenderSurvey(fields, answers, chosen);
            DrawImagePage(document, bitmap);
        }

        document.Close();
    }

    // ── C: 定型業務帳票(項目 + 印) ──────────────────────

    public static void BusinessForm(
        string path, int pages, out List<Dictionary<string, string>> truth)
    {
        truth = [];
        var random = new Random(3003);
        string[] people = ["架空 太郎", "架空 花子", "架空 一郎", "架空 二郎"];
        string[] statuses = ["承認", "保留", "差戻"];

        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);

        for (var page = 1; page <= pages; page++)
        {
            var fields = new Dictionary<string, string>
            {
                ["店舗コード"] = $"S{page:D3}-{10 + (page % 89)}",
                ["担当者"] = people[page % people.Length],
                ["日付"] = $"2026/{(page % 12) + 1:D2}/{(page % 27) + 1:D2}",
                ["金額"] = (10000 + (page * 371)).ToString("N0", CultureInfo.InvariantCulture),
                ["数量"] = random.Next(1, 999).ToString(CultureInfo.InvariantCulture),
                ["備考"] = "架空の記録",
                ["状態"] = statuses[page % statuses.Length],
            };

            truth.Add(fields);

            // 位置ずれ・倍率・軽い傾きを混ぜる(実際の取り込みに近づける)。
            var dx = ((page % 3) - 1) * 6.0;
            var dy = ((page % 2) - 0.5) * 8.0;
            var zoom = page % 5 == 0 ? 1.015 : 1.0;
            var tilt = page % 4 == 0 ? 1.5 : 0;

            using var flat = RenderBusinessForm(fields, statuses, dx, dy, zoom);
            using var final = tilt == 0 ? flat.Copy() : Tilt(flat, tilt);
            DrawImagePage(document, final);
        }

        document.Close();
    }

    // ── D: 契約 / 申込書 ────────────────────────────

    public static void Contract(
        string path, int pages, out List<Dictionary<string, string>> truth)
    {
        truth = [];
        string[] plans = ["標準プラン", "拡張プラン", "試用プラン"];

        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);

        for (var page = 1; page <= pages; page++)
        {
            var fields = new Dictionary<string, string>
            {
                ["契約番号"] = $"K-{2026}{page:D4}",
                ["契約日"] = $"2026/{(page % 12) + 1:D2}/{(page % 27) + 1:D2}",
                ["法人名"] = $"架空第{page}商事株式会社",
                ["契約金額"] = (500000 + (page * 1234)).ToString(
                    "N0", CultureInfo.InvariantCulture),
                ["プラン"] = plans[page % plans.Length],
            };

            truth.Add(fields);

            using var bitmap = RenderContract(fields, plans);
            DrawImagePage(document, bitmap);
        }

        document.Close();
    }

    // ── E: 混在 PDF ────────────────────────────────

    /// <summary>
    /// 文字情報のあるページと画像だけのページを混ぜる。
    /// どのページがどれなのかを truth として返す。
    /// </summary>
    public static void Mixed(string path, int pages, out List<string> kinds)
    {
        kinds = [];
        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);

        var font = new SKFont(Japanese(), 11);
        var paint = new SKPaint { Color = SKColors.Black, IsAntialias = true };

        for (var page = 1; page <= pages; page++)
        {
            var kind = (page % 4) switch
            {
                0 => "born-digital 文章",
                1 => "born-digital 表",
                2 => "スキャン 文章",
                _ => "スキャン 帳票",
            };

            kinds.Add(kind);

            if (kind.StartsWith("born-digital", StringComparison.Ordinal))
            {
                // 文字情報を持つページ(画像にしない)。
                var canvas = document.BeginPage(PageWidth, PageHeight);
                canvas.Clear(SKColors.White);

                if (kind.EndsWith("表", StringComparison.Ordinal))
                {
                    var y = 90f;
                    canvas.DrawText(
                        $"明細({page} ページ・架空)", 60, 60, SKTextAlign.Left, font, paint);
                    for (var row = 0; row < 12; row++)
                    {
                        canvas.DrawText($"A{page:D2}{row:D2}", 60, y, SKTextAlign.Left, font, paint);
                        canvas.DrawText("架空の品目", 160, y, SKTextAlign.Left, font, paint);
                        canvas.DrawText(
                            (1000 + (row * 13)).ToString("N0", CultureInfo.InvariantCulture),
                            320, y, SKTextAlign.Left, font, paint);
                        y += 22;
                    }
                }
                else
                {
                    var y = 90f;
                    foreach (var line in new[]
                    {
                        $"架空の報告 第 {page} 節",
                        "本書は動作確認のために生成した架空の文書です。",
                        "金額や名称に実在のものは含まれません。",
                    })
                    {
                        canvas.DrawText(line, 60, y, SKTextAlign.Left, font, paint);
                        y += 26;
                    }
                }

                document.EndPage();
                continue;
            }

            // 画像だけのページ。
            if (kind.EndsWith("文章", StringComparison.Ordinal))
            {
                using var bitmap = RenderScanProse(page);
                DrawImagePage(document, bitmap);
            }
            else
            {
                var fields = new Dictionary<string, string>
                {
                    ["店舗コード"] = $"S{page:D3}-42",
                    ["担当者"] = "架空 太郎",
                    ["日付"] = "2026/02/10",
                    ["金額"] = "1,234,567",
                    ["数量"] = "12",
                    ["備考"] = "架空の記録",
                    ["状態"] = "承認",
                };

                using var bitmap = RenderBusinessForm(fields, ["承認", "保留", "差戻"], 0, 0, 1.0);
                DrawImagePage(document, bitmap);
            }
        }

        font.Dispose();
        paint.Dispose();
        document.Close();
    }

    // ── F: 悪条件 ──────────────────────────────────

    public static void Rough(
        string path, int pages, out List<Dictionary<string, string>> truth)
    {
        truth = [];
        string[] people = ["架空 太郎", "架空 花子", "架空 一郎"];

        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);

        for (var page = 1; page <= pages; page++)
        {
            var fields = new Dictionary<string, string>
            {
                ["店舗コード"] = $"S{page:D3}-{10 + (page % 89)}",
                ["担当者"] = people[page % people.Length],
                ["日付"] = $"2026/{(page % 12) + 1:D2}/{(page % 27) + 1:D2}",
                ["金額"] = (10000 + (page * 371)).ToString("N0", CultureInfo.InvariantCulture),
                ["数量"] = (page * 7 % 900).ToString(CultureInfo.InvariantCulture),
                ["備考"] = "架空の記録",
                ["状態"] = "承認",
            };

            truth.Add(fields);

            var tilt = ((page % 3) - 1) * 2.5;
            using var flat = RenderBusinessForm(fields, ["承認", "保留", "差戻"], 0, 0, 1.0);
            using var tilted = tilt == 0 ? flat.Copy() : Tilt(flat, tilt);
            using var rough = Rough(tilted, page);
            DrawImagePage(document, rough);
        }

        document.Close();
    }

    // ── 描画 ─────────────────────────────────────────

    private static SKBitmap NewPage(out SKCanvas canvas, out SKFont font, out SKPaint paint)
    {
        var bitmap = new SKBitmap((int)(PageWidth * Scale), (int)(PageHeight * Scale));
        canvas = new SKCanvas(bitmap);
        canvas.Clear(SKColors.White);
        font = new SKFont(Japanese(), 10 * Scale);
        paint = new SKPaint { Color = SKColors.Black, IsAntialias = true };
        return bitmap;
    }

    private static SKBitmap RenderTable(
        string[] header, List<string[]> rows, bool ruled)
    {
        var bitmap = NewPage(out var canvas, out var font, out var paint);
        using (canvas)
        {
            using var line = new SKPaint
            {
                Color = SKColors.Black, IsStroke = true, StrokeWidth = 1.4f,
            };

            float[] columnX = [40, 110, 250, 320, 380, 460, 555];
            var top = 70f;
            var rowHeight = 17f;
            var count = rows.Count + 1;

            if (ruled)
            {
                for (var row = 0; row <= count; row++)
                {
                    var y = (top + (row * rowHeight)) * Scale;
                    canvas.DrawLine(columnX[0] * Scale, y, columnX[^1] * Scale, y, line);
                }

                foreach (var x in columnX)
                {
                    canvas.DrawLine(
                        x * Scale, top * Scale,
                        x * Scale, (top + (count * rowHeight)) * Scale, line);
                }
            }

            void DrawRow(string[] values, int index)
            {
                var y = (top + (index * rowHeight) + 12) * Scale;
                for (var column = 0; column < values.Length; column++)
                {
                    canvas.DrawText(
                        values[column], (columnX[column] + 4) * Scale, y,
                        SKTextAlign.Left, font, paint);
                }
            }

            DrawRow(header, 0);
            for (var row = 0; row < rows.Count; row++)
            {
                DrawRow(rows[row], row + 1);
            }
        }

        font.Dispose();
        paint.Dispose();
        return bitmap;
    }

    private static SKBitmap RenderSurvey(
        Dictionary<string, string> fields, string[] answers, string[] chosen)
    {
        var bitmap = NewPage(out var canvas, out var font, out var paint);
        using (canvas)
        {
            using var stroke = new SKPaint
            {
                Color = SKColors.Black, IsStroke = true, StrokeWidth = PenWidth,
                IsAntialias = true,
            };
            using var thin = new SKPaint
            {
                Color = SKColors.Black, IsStroke = true, StrokeWidth = 1.2f,
            };

            canvas.DrawText("架空アンケート", 60 * Scale, 50 * Scale, SKTextAlign.Left, font, paint);
            canvas.DrawText(
                $"整理番号: {fields["整理番号"]}", 60 * Scale, 78 * Scale,
                SKTextAlign.Left, font, paint);
            canvas.DrawText(
                $"年齢: {fields["年齢"]}", 300 * Scale, 78 * Scale, SKTextAlign.Left, font, paint);

            var y = 110f;
            for (var q = 0; q < chosen.Length; q++)
            {
                canvas.DrawText(
                    $"設問{q + 1}", 50 * Scale, (y + 10) * Scale, SKTextAlign.Left, font, paint);

                for (var option = 0; option < answers.Length; option++)
                {
                    var x = 120 + (option * 130);
                    var box = new SKRect(
                        x * Scale, y * Scale, (x + 12) * Scale, (y + 12) * Scale);
                    canvas.DrawRect(box, thin);
                    canvas.DrawText(
                        answers[option], (x + 18) * Scale, (y + 10) * Scale,
                        SKTextAlign.Left, font, paint);

                    if (answers[option] == chosen[q])
                    {
                        // チェックの線(実際のペン相当の太さ)。
                        canvas.DrawLine(
                            box.Left + (2 * Scale), box.MidY,
                            box.MidX, box.Bottom - (2 * Scale), stroke);
                        canvas.DrawLine(
                            box.MidX, box.Bottom - (2 * Scale),
                            box.Right - (1 * Scale), box.Top + (1 * Scale), stroke);
                    }
                }

                y += 26;
            }

            canvas.DrawText(
                $"自由記述: {fields["自由記述"]}", 50 * Scale, (y + 16) * Scale,
                SKTextAlign.Left, font, paint);
        }

        font.Dispose();
        paint.Dispose();
        return bitmap;
    }

    private static SKBitmap RenderBusinessForm(
        Dictionary<string, string> fields, string[] statuses, double dx, double dy, double zoom)
    {
        var bitmap = NewPage(out var canvas, out var font, out var paint);
        using (canvas)
        {
            using var thin = new SKPaint
            {
                Color = SKColors.Black, IsStroke = true, StrokeWidth = 1.2f,
            };
            using var stroke = new SKPaint
            {
                Color = SKColors.Black, IsStroke = true, StrokeWidth = PenWidth,
                IsAntialias = true,
            };

            canvas.Save();
            canvas.Translate((float)(dx * Scale), (float)(dy * Scale));
            canvas.Scale((float)zoom, (float)zoom, bitmap.Width / 2f, bitmap.Height / 2f);

            canvas.DrawText("架空 業務報告書", 60 * Scale, 50 * Scale, SKTextAlign.Left, font, paint);

            var y = 108f;
            foreach (var name in new[] { "店舗コード", "担当者", "日付", "金額", "数量", "備考" })
            {
                canvas.DrawText(
                    $"{name}:", 60 * Scale, (y + 14) * Scale, SKTextAlign.Left, font, paint);
                canvas.DrawText(
                    fields[name], 175 * Scale, (y + 14) * Scale, SKTextAlign.Left, font, paint);
                canvas.DrawLine(
                    170 * Scale, (y + 18) * Scale, 400 * Scale, (y + 18) * Scale, thin);
                y += 30;
            }

            // 状態のチェック欄。
            canvas.DrawText("状態:", 60 * Scale, (y + 14) * Scale, SKTextAlign.Left, font, paint);
            for (var index = 0; index < statuses.Length; index++)
            {
                var x = 130 + (index * 110);
                var box = new SKRect(x * Scale, y * Scale, (x + 12) * Scale, (y + 12) * Scale);
                canvas.DrawRect(box, thin);
                canvas.DrawText(
                    statuses[index], (x + 18) * Scale, (y + 11) * Scale,
                    SKTextAlign.Left, font, paint);

                if (statuses[index] == fields["状態"])
                {
                    canvas.DrawLine(
                        box.Left + (2 * Scale), box.MidY,
                        box.MidX, box.Bottom - (2 * Scale), stroke);
                    canvas.DrawLine(
                        box.MidX, box.Bottom - (2 * Scale),
                        box.Right - (1 * Scale), box.Top + (1 * Scale), stroke);
                }
            }

            canvas.Restore();
        }

        font.Dispose();
        paint.Dispose();
        return bitmap;
    }

    private static SKBitmap RenderContract(Dictionary<string, string> fields, string[] plans)
    {
        var bitmap = NewPage(out var canvas, out var font, out var paint);
        using (canvas)
        {
            using var thin = new SKPaint
            {
                Color = SKColors.Black, IsStroke = true, StrokeWidth = 1.2f,
            };
            using var stroke = new SKPaint
            {
                Color = SKColors.Black, IsStroke = true, StrokeWidth = PenWidth,
                IsAntialias = true,
            };

            canvas.DrawText("架空 申込書", 60 * Scale, 50 * Scale, SKTextAlign.Left, font, paint);

            var y = 108f;
            foreach (var name in new[] { "契約番号", "契約日", "法人名", "契約金額" })
            {
                canvas.DrawText(
                    $"{name}:", 60 * Scale, (y + 14) * Scale, SKTextAlign.Left, font, paint);
                canvas.DrawText(
                    fields[name], 175 * Scale, (y + 14) * Scale, SKTextAlign.Left, font, paint);
                canvas.DrawLine(
                    170 * Scale, (y + 18) * Scale, 430 * Scale, (y + 18) * Scale, thin);
                y += 30;
            }

            canvas.DrawText("プラン:", 60 * Scale, (y + 14) * Scale, SKTextAlign.Left, font, paint);
            for (var index = 0; index < plans.Length; index++)
            {
                var x = 140 + (index * 130);
                var box = new SKRect(x * Scale, y * Scale, (x + 12) * Scale, (y + 12) * Scale);
                canvas.DrawRect(box, thin);
                canvas.DrawText(
                    plans[index], (x + 18) * Scale, (y + 11) * Scale,
                    SKTextAlign.Left, font, paint);

                if (plans[index] == fields["プラン"])
                {
                    canvas.DrawLine(
                        box.Left + (2 * Scale), box.MidY,
                        box.MidX, box.Bottom - (2 * Scale), stroke);
                    canvas.DrawLine(
                        box.MidX, box.Bottom - (2 * Scale),
                        box.Right - (1 * Scale), box.Top + (1 * Scale), stroke);
                }
            }

            canvas.DrawText(
                "本書は動作確認のために生成した架空の書面です。",
                60 * Scale, (y + 70) * Scale, SKTextAlign.Left, font, paint);
        }

        font.Dispose();
        paint.Dispose();
        return bitmap;
    }

    private static SKBitmap RenderScanProse(int page)
    {
        var bitmap = NewPage(out var canvas, out var font, out var paint);
        using (canvas)
        {
            var y = 90f;
            foreach (var line in new[]
            {
                $"架空の連絡文 第 {page} 号",
                "会社名: 架空商事株式会社",
                "金額: 4,917,087 円",
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

    private static SKBitmap Tilt(SKBitmap source, double degrees)
    {
        var result = new SKBitmap(source.Width, source.Height);
        using var canvas = new SKCanvas(result);
        canvas.Clear(SKColors.White);
        canvas.RotateDegrees((float)degrees, source.Width / 2f, source.Height / 2f);
        canvas.DrawBitmap(source, 0, 0, new SKSamplingOptions(SKFilterMode.Linear));
        return result;
    }

    /// <summary>
    /// 取り込みの粗い状態にする。150dpi 相当まで落とし、ぼかし・かすれ・
    /// 粒状のノイズを混ぜる(JPEG 圧縮相当の劣化も含む)。
    /// </summary>
    private static SKBitmap Rough(SKBitmap source, int seed)
    {
        // 150dpi 相当へ落として戻す。
        var half = new SKBitmap(source.Width / 2, source.Height / 2);
        using (var canvas = new SKCanvas(half))
        {
            canvas.Clear(SKColors.White);
            canvas.DrawBitmap(
                source, new SKRect(0, 0, half.Width, half.Height),
                new SKSamplingOptions(SKFilterMode.Linear));
        }

        // JPEG 圧縮を通す。
        using var image = SKImage.FromBitmap(half);
        using var jpeg = image.Encode(SKEncodedImageFormat.Jpeg, 35);
        using var decoded = SKBitmap.Decode(jpeg);
        half.Dispose();

        var result = new SKBitmap(source.Width, source.Height);
        using (var canvas = new SKCanvas(result))
        {
            canvas.Clear(SKColors.White);
            using var blur = new SKPaint { ImageFilter = SKImageFilter.CreateBlur(1.2f, 1.2f) };
            canvas.DrawBitmap(
                decoded, new SKRect(0, 0, result.Width, result.Height),
                new SKSamplingOptions(SKFilterMode.Linear), blur);
        }

        // かすれと粒状のノイズ。
        var random = new Random(seed * 7919);
        for (var i = 0; i < result.Width * result.Height / 120; i++)
        {
            var x = random.Next(result.Width);
            var y = random.Next(result.Height);
            var dark = random.Next(2) == 0;
            result.SetPixel(x, y, dark ? new SKColor(70, 70, 70) : new SKColor(238, 238, 238));
        }

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

    // ── 領域(OCR 300dpi の座標) ────────────────────────

    public static (string Name, float X, float Y, float Width, float Height)[] BusinessFormAreas()
    {
        var areas = new List<(string, float, float, float, float)>();
        var y = 108f;
        foreach (var name in new[] { "店舗コード", "担当者", "日付", "金額", "数量", "備考" })
        {
            areas.Add((name, 170 * Scale, y * Scale, 240 * Scale, 20 * Scale));
            y += 30;
        }

        return [.. areas];
    }

    public static (string Label, float X, float Y, float Size)[] BusinessFormStatusBoxes()
    {
        var y = 108f + (6 * 30);
        return
        [
            ("承認", 130 * Scale, y * Scale, 12 * Scale),
            ("保留", 240 * Scale, y * Scale, 12 * Scale),
            ("差戻", 350 * Scale, y * Scale, 12 * Scale),
        ];
    }

    public static (string Name, float X, float Y, float Width, float Height)[] ContractAreas()
    {
        var areas = new List<(string, float, float, float, float)>();
        var y = 108f;
        foreach (var name in new[] { "契約番号", "契約日", "法人名", "契約金額" })
        {
            areas.Add((name, 170 * Scale, y * Scale, 270 * Scale, 20 * Scale));
            y += 30;
        }

        return [.. areas];
    }

    public static (string Label, float X, float Y, float Size)[] ContractPlanBoxes()
    {
        var y = 108f + (4 * 30);
        return
        [
            ("標準プラン", 140 * Scale, y * Scale, 12 * Scale),
            ("拡張プラン", 270 * Scale, y * Scale, 12 * Scale),
            ("試用プラン", 400 * Scale, y * Scale, 12 * Scale),
        ];
    }

    public static (string Name, float X, float Y, float Width, float Height)[] SurveyAreas(
        int choiceCount)
    {
        var areas = new List<(string, float, float, float, float)>
        {
            ("整理番号", 60 * Scale, 62 * Scale, 200 * Scale, 22 * Scale),
            ("年齢", 300 * Scale, 62 * Scale, 160 * Scale, 22 * Scale),
        };

        var y = 110f + (choiceCount * 26);
        areas.Add(("自由記述", 50 * Scale, (y + 2) * Scale, 420 * Scale, 22 * Scale));
        return [.. areas];
    }

    public static (string Label, float X, float Y, float Size)[] SurveyBoxes(
        int question, string[] answers)
    {
        var y = 110f + (question * 26);
        var boxes = new List<(string, float, float, float)>();
        for (var option = 0; option < answers.Length; option++)
        {
            var x = 120 + (option * 130);
            boxes.Add((answers[option], x * Scale, y * Scale, 12 * Scale));
        }

        return [.. boxes];
    }
}
