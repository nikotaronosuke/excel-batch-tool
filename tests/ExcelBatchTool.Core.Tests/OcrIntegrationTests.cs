using ExcelBatchTool.Core.Ocr;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2F-B3。全方式統合で足した判断。すべて架空データ。
/// </summary>
public sealed class OcrIntegrationTests
{
    // ── 上下逆でも数として読めてしまう値 ─────────────────

    [Theory]
    [InlineData("90", true)]      // 上下逆で「06」。どちらも数として正しい
    [InlineData("99", true)]      // 上下逆で「66」
    [InlineData("0168", true)]
    [InlineData("34", false)]     // 3 と 4 は上下逆では数にならない
    [InlineData("1,234", false)]  // 桁区切りが入れば上下逆では読めない
    [InlineData("", false)]
    public void ValuesThatStayNumbersUpsideDownAreNotAutoAccepted(string text, bool expected)
        => Assert.Equal(expected, FieldAutoAcceptPolicy.IsUpsideDownAmbiguous(text));

    [Theory]
    [InlineData("099DQ66")]
    [InlineData("006DQ969")]
    public void TheTraceOfARotatedThousandsSeparatorIsCaught(string text)
    {
        // 実測: 「99,660」が「099DQ66」として自信 99.0% で自動確定していた。
        // カンマを上下逆に読むとアポストロフィになる。数字の並びだけを見ると
        // 列の中で浮かないので、記号のほうで気づく。
        Assert.True(FieldAutoAcceptPolicy.IsUpsideDownAmbiguous(text.Replace("DQ", "'")));
    }

    // ── 途中で失敗したとき ────────────────────────────

    [Fact]
    public void AFailureInTheMiddleLeavesNothingBehindAndTheSourceUntouched()
    {
        // OCR は本体と同じプロセスで動く(別プロセスの補助を使わないので、
        // 取り残しのプロセスも無い)。途中で native 側の失敗が起きても、
        // 出力・控え・作業用ファイルを残さず、元の PDF も変えない。
        using var dir = new TempDir();
        var pdf = dir.File("途中で失敗.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 3);
        var before = Snapshot(pdf);

        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.Agreed("1 ページ目", 0.99))
            .Page(2, FakeOcrEngine.Agreed("2 ページ目", 0.99))
            .Page(3, FakeOcrEngine.Agreed("3 ページ目", 0.99));

        engine.OnRead = page =>
        {
            if (page == 2)
            {
                throw new InvalidOperationException("読み取りに失敗しました(見本)");
            }
        };

        Assert.Throws<InvalidOperationException>(
            () => new PdfScanReader().Read(engine, pdf, [1, 2, 3]));

        Assert.Equal(before, Snapshot(pdf));
        Assert.Empty(Directory.GetFiles(dir.Root, "*.xlsx"));
        Assert.Empty(Directory.GetFiles(dir.Root, "*.csv"));
        Assert.Empty(Directory.GetFiles(dir.Root, "*.audit.json"));
        Assert.Empty(Directory.GetFiles(dir.Root, "*.tmp"));
    }

    private static (string Sha, long Length, DateTime Modified) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }

    // ── 「項目名: 値」と 1 行に刷られた帳票 ──────────────────

    [Theory]
    [InlineData("年齢", "年齢: 58", "58")]
    [InlineData("年齢", "年齢:58", "58")]
    [InlineData("年齢", "年齢：58", "58")]
    [InlineData("自由記述", "自由記述: 架空の要望です", "架空の要望です")]
    public void TheFieldLabelPrintedOnTheSameLineIsRemoved(
        string field, string text, string expected)
        => Assert.Equal(expected, FormFieldExtractor.StripLabel(field, text));

    [Theory]
    [InlineData("年齢", "58")]                     // 見出しが無ければ何もしない
    [InlineData("年齢", "生年月日: 2000/01/01")]     // 前が項目名でなければ何もしない
    [InlineData("備考", "備考です")]                 // 区切りが無ければ何もしない
    [InlineData("時刻", "時刻")]
    public void AnythingThatIsNotTheFieldLabelIsLeftAlone(string field, string text)
        => Assert.Equal(text, FormFieldExtractor.StripLabel(field, text));

    [Fact]
    public void AColonInsideTheValueIsNotUsedAsASeparator()
    {
        // 区切りより後ろを機械的に採ると、値の中の「:」で切ってしまう。
        Assert.Equal("10:30 から 12:00", FormFieldExtractor.StripLabel("時間", "時間: 10:30 から 12:00"));
    }

    // ── 項目 × 全ページで形を学ぶ(自動確定の回復) ────────────

    [Fact]
    public void AFieldShapeIsLearnedFromEveryPage()
    {
        // 店舗コードが全ページ「英字 + 数字 3 桁 - 数字 2 桁」なら、その形を学ぶ。
        var readings = Enumerable.Range(1, 20).Select(page => $"S{page:D3}-24");

        Assert.Equal("A999-99", FieldShapePattern.Learn(readings));
    }

    [Fact]
    public void AReadingThatBreaksTheLearnedShapeIsNotAutoAccepted()
    {
        var pattern = FieldShapePattern.Learn(
            Enumerable.Range(1, 20).Select(page => $"S{page:D3}-24"));

        // 正しく読めたページは形どおり。
        Assert.True(FieldShapePattern.Matches(pattern, "S001-24"));

        // 0 を O と読んだページは形が変わる(英字が 1 つ増える)。
        Assert.False(FieldShapePattern.Matches(pattern, "SO01-24"));
    }

    [Fact]
    public void ShapeIsNotLearnedFromTooFewOrTooMixedPages()
    {
        Assert.Null(FieldShapePattern.Learn(["S001-24", "S002-24", "S003-24"]));

        // 様式が揃っていない項目(自由記述)では形を決めない。
        Assert.Null(FieldShapePattern.Learn(
            ["架空の備考", "1,234", "AB-99", "架空", "2026/01/01", "x", "yy", "zzz", "9"]));
    }

    [Fact]
    public void TheLearnedShapeNeverRewritesTheReading()
    {
        // 形は「自動確定してよいか」の判断にだけ使う。値は触らない。
        var pattern = FieldShapePattern.Learn(
            Enumerable.Range(1, 20).Select(page => $"S{page:D3}-24"));

        Assert.False(FieldShapePattern.Matches(pattern, "SO01-24"));
        Assert.Equal("A999-99", pattern);
    }

    [Fact]
    public void CodeFieldsRecoverAutoAcceptWhenEveryPageHasTheSameShape()
    {
        using var dir = new TempDir();
        var pdf = dir.File("帳票.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 12);

        var engine = new FakeOcrEngine();
        for (var page = 1; page <= 12; page++)
        {
            engine.Page(page,
                    FakeOcrEngine.At($"S{page:D3}-24", 0.99, new OcrBox(200, 100, 120, 30)),
                    FakeOcrEngine.At("1234567", 0.99, new OcrBox(200, 140, 120, 30)))
                .Probe(page, skew: 0, horizontal: 0, vertical: 0);
        }

        var reading = new PdfScanReader().Read(
            engine, pdf, [.. Enumerable.Range(1, 12)],
            new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = CodeTemplate() });

        // 2F-B2 の粗い決まりなら、0 を含むコードは全ページ人へ回っていた。
        var codes = reading.Items.Where(item => item.FieldName == "店舗コード").ToList();
        Assert.Equal(12, codes.Count);
        Assert.All(codes, item => Assert.Equal(OcrItemStatus.AutoAccepted, item.InitialStatus));
    }

    [Fact]
    public void ThePageThatBreaksTheShapeStillGoesToReview()
    {
        using var dir = new TempDir();
        var pdf = dir.File("帳票.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 12);

        var engine = new FakeOcrEngine();
        for (var page = 1; page <= 12; page++)
        {
            // 5 ページ目だけ 0 を O と読んでいる。
            var code = page == 5 ? "SO05-24" : $"S{page:D3}-24";
            engine.Page(page,
                    FakeOcrEngine.At(code, 0.99, new OcrBox(200, 100, 120, 30)),
                    FakeOcrEngine.At("1234567", 0.99, new OcrBox(200, 140, 120, 30)))
                .Probe(page, skew: 0, horizontal: 0, vertical: 0);
        }

        var reading = new PdfScanReader().Read(
            engine, pdf, [.. Enumerable.Range(1, 12)],
            new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = CodeTemplate() });

        var odd = reading.Items.Single(
            item => item.FieldName == "店舗コード" && item.PageNumber == 5);
        Assert.Equal(OcrItemStatus.NeedsReview, odd.InitialStatus);
        Assert.Contains("形が違います", odd.Reason);

        // 値は書き換えない。読んだままを見せて、人が直す。
        Assert.Equal("SO05-24", odd.Text);
    }

    // ── 罫線が落ちて 2 行が 1 区画に入った場合 ──────────────

    [Fact]
    public void TwoRowsThatFellIntoOneBandAreSplitBackApart()
    {
        // 実測で、表の下のほうの細い罫線が落ちて GT の 4 行が 2 区画に収まり、
        // 「A0017A0018」という 1 セルができていた(自信 99.6% で自動確定)。
        var lines = new List<OcrRawLine>
        {
            FakeOcrEngine.At("A0001", 0.99, new OcrBox(60, 105, 60, 20)),
            FakeOcrEngine.At("架空1", 0.99, new OcrBox(210, 105, 60, 20)),
            FakeOcrEngine.At("A0002", 0.99, new OcrBox(60, 145, 60, 20)),
            FakeOcrEngine.At("架空2", 0.99, new OcrBox(210, 145, 60, 20)),

            // ここから下は罫線が落ちていて、1 区画に 2 行ぶん入っている。
            FakeOcrEngine.At("A0003", 0.99, new OcrBox(60, 185, 60, 20)),
            FakeOcrEngine.At("架空3", 0.99, new OcrBox(210, 185, 60, 20)),
            FakeOcrEngine.At("A0004", 0.99, new OcrBox(60, 225, 60, 20)),
            FakeOcrEngine.At("架空4", 0.99, new OcrBox(210, 225, 60, 20)),
        };

        var table = ScanTableBuilder.FromRulings(
            lines, rowLines: [100, 140, 180, 260], columnLines: [50, 200, 350], Read);

        Assert.NotNull(table);

        // 区画を割り直して 4 行に戻る(割らなければ 3 行で、最後が連結される)。
        Assert.Equal(4, table.RowCount);
        var rows = table.ToRows();
        Assert.Equal(["A0003", "架空3"], rows[2]);
        Assert.Equal(["A0004", "架空4"], rows[3]);
    }

    [Fact]
    public void WrappedTextInsideOneCellIsNotSplit()
    {
        // 折り返しは区画の高さを増やさないので、割ってはいけない。
        var lines = new List<OcrRawLine>
        {
            FakeOcrEngine.At("A0001", 0.99, new OcrBox(60, 105, 60, 14)),
            FakeOcrEngine.At("架空の", 0.99, new OcrBox(210, 105, 60, 14)),
            FakeOcrEngine.At("長い商品名", 0.99, new OcrBox(210, 122, 90, 14)),
            FakeOcrEngine.At("A0002", 0.99, new OcrBox(60, 145, 60, 14)),
            FakeOcrEngine.At("架空2", 0.99, new OcrBox(210, 145, 60, 14)),
        };

        var table = ScanTableBuilder.FromRulings(
            lines, rowLines: [100, 140, 180], columnLines: [50, 200, 350], Read);

        Assert.NotNull(table);
        Assert.Equal(2, table.RowCount);
        Assert.Equal("架空の長い商品名", table.ToRows()[0][1]);
    }

    [Fact]
    public void ACellStitchedFromVerticallySeparatedReadingsIsFlagged()
    {
        // 割り直しても解消しなかった連結は、自信が高くても人へ回す。
        Assert.True(ScanTableBuilder.HasVerticalGap(
            [new OcrBox(60, 100, 60, 20), new OcrBox(60, 140, 60, 20)]));

        // 同じ行の続き(重なっている)は連結ではない。
        Assert.False(ScanTableBuilder.HasVerticalGap(
            [new OcrBox(60, 100, 60, 20), new OcrBox(130, 105, 60, 20)]));
    }

    // ── 文字があるのに読めなかった区画 ─────────────────────

    [Fact]
    public void ACellWithInkButNoTextIsSurfacedInsteadOfSilentlyBlank()
    {
        using var dir = new TempDir();
        var pdf = dir.File("読めない欄.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine()
            .Page(1,
                FakeOcrEngine.At("商品コード", 0.99, new OcrBox(60, 110, 90, 20)),
                FakeOcrEngine.At("商品名", 0.99, new OcrBox(210, 110, 90, 20)),
                FakeOcrEngine.At("A0001", 0.99, new OcrBox(60, 150, 90, 20)))
            .Probe(1, skew: 0, horizontal: 3, vertical: 3)
            .Rulings(1, rows: [100, 140, 180], columns: [50, 200, 350])
            // 読めなかった区画に黒い画素がある = 文字はあった。
            .Ink(1, 0.08);

        var reading = new PdfScanReader().Read(
            engine, pdf, [1], new OcrReadOptions { Mode = OcrReadMode.Table });

        var unread = reading.Items.Single(item => item.Status == OcrItemStatus.Unreadable);
        Assert.Equal(1, unread.Row);
        Assert.Equal(1, unread.Column);
        Assert.Contains("文字があるようですが", unread.Reason);
    }

    [Fact]
    public void AGenuinelyBlankCellIsNotTurnedIntoReviewWork()
    {
        using var dir = new TempDir();
        var pdf = dir.File("空欄.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine()
            .Page(1,
                FakeOcrEngine.At("商品コード", 0.99, new OcrBox(60, 110, 90, 20)),
                FakeOcrEngine.At("商品名", 0.99, new OcrBox(210, 110, 90, 20)),
                FakeOcrEngine.At("A0001", 0.99, new OcrBox(60, 150, 90, 20)))
            .Probe(1, skew: 0, horizontal: 3, vertical: 3)
            .Rulings(1, rows: [100, 140, 180], columns: [50, 200, 350])
            .Ink(1, 0.0);

        var reading = new PdfScanReader().Read(
            engine, pdf, [1], new OcrReadOptions { Mode = OcrReadMode.Table });

        Assert.DoesNotContain(reading.Items, item => item.Status == OcrItemStatus.Unreadable);
    }

    private static FormTemplate CodeTemplate() => new()
    {
        Name = "架空の帳票",
        Fields =
        [
            new FormField
            {
                Name = "店舗コード", Area = new OcrBox(200, 100, 120, 30), Kind = FormFieldKind.Code,
            },
            new FormField
            {
                Name = "売上", Area = new OcrBox(200, 140, 120, 30), Kind = FormFieldKind.NumberLike,
            },
        ],
    };

    private static (string Text, double Confidence) Read(OcrRawLine line)
    {
        var result = OcrFusion.Fuse(line);
        return (result.Text, result.Confidence);
    }
}
