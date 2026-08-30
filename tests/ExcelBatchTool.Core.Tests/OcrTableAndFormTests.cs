using ExcelBatchTool.Core.Ocr;
using ExcelBatchTool.Core.Pdf;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2F-B2。傾き補正 / スキャン表 / 大量定型帳票 / 印の判定。
/// すべて架空データ。Offline OCR Pack を置かずに、差し替えた OCR で筋道を確かめる。
/// </summary>
public sealed class OcrTableAndFormTests
{
    // ── 傾き補正 ──────────────────────────────────────

    [Theory]
    [InlineData(0.0, false)]
    [InlineData(0.2, false)]
    [InlineData(0.5, true)]
    [InlineData(1.0, true)]
    [InlineData(2.0, true)]
    [InlineData(3.0, true)]
    [InlineData(5.0, true)]
    [InlineData(-3.0, true)]
    [InlineData(7.0, false)]
    public void OnlyMeaningfulTiltsAreStraightened(double degrees, bool expected)
    {
        // ほぼまっすぐなページを回すと、補間で文字がぼやけてかえって悪くなる。
        // 大きすぎる傾きは直さずに人へ回す。
        Assert.Equal(expected, DeskewPolicy.ShouldDeskew(degrees, reliable: true));
    }

    [Fact]
    public void AnUnreliableAngleMeansReadWithoutRotating()
    {
        // 角度を測れなかったページは回さずそのまま読む。止めはしない
        // (測れないのは行らしい塊が少ないページで、多くは実際には傾いていない)。
        Assert.False(DeskewPolicy.ShouldDeskew(2.0, reliable: false));
        Assert.False(DeskewPolicy.IsTooTilted(2.0, reliable: false));
        Assert.False(DeskewPolicy.IsTooTilted(20.0, reliable: false));
    }

    [Theory]
    [InlineData(0.5)]
    [InlineData(1.0)]
    [InlineData(2.0)]
    [InlineData(3.0)]
    [InlineData(5.0)]
    [InlineData(-2.5)]
    public void ThePositionGoesBackToWhereItWasOnTheOriginalPage(double degrees)
    {
        // 直した画像で読んだ位置を、元のページの座標へ戻せること。
        // これができないと、確認画面の赤い枠が原文とずれる。
        var transform = new DeskewTransform(degrees, 1240, 1754);
        var original = new OcrBox(300, 500, 240, 40);

        var roundTrip = transform.ToOriginal(transform.ToDeskewed(original));

        Assert.Equal(original.CenterX, roundTrip.CenterX, 6);
        Assert.Equal(original.CenterY, roundTrip.CenterY, 6);
    }

    [Fact]
    public void NoTiltMeansTheCoordinatesAreUntouched()
    {
        var box = new OcrBox(10, 20, 30, 40);

        Assert.Equal(box, DeskewTransform.None.ToOriginal(box));
        Assert.Equal(box, DeskewTransform.None.ToDeskewed(box));
    }

    [Fact]
    public void AStraightenedPageKeepsItsReviewPositionOnTheOriginal()
    {
        using var dir = new TempDir();
        var pdf = dir.File("傾き.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.At("架空商事", 0.99, new OcrBox(300, 500, 240, 40)))
            .Probe(1, skew: 2.0, horizontal: 0, vertical: 0);

        var reading = new PdfScanReader().Read(engine, pdf, [1]);
        var item = Assert.Single(reading.Items);

        // 読み取りは直した画像で行うが、確認に使う位置は元のページのもの。
        // 戻す向きが正しいかどうかは DeskewReviewCoordinateTests が画素の実測で見ている。
        // ここでは「元へ戻す変換を通っている」ことだけを見る。
        var expected = DeskewTransform.FromRotation(2.0, 620, 877)
            .ToOriginal(new OcrBox(300, 500, 240, 40));
        Assert.Equal(expected.X, item.BoundingBox.X, 3);
        Assert.Equal(expected.Y, item.BoundingBox.Y, 3);
        Assert.NotEqual(300, item.BoundingBox.X);
    }

    // ── 罫線のある表 ───────────────────────────────────

    [Fact]
    public void ARuledScanTableComesBackAsRowsAndColumns()
    {
        var table = ScanTableBuilder.FromRulings(
            Cells(
                (0, 0, "商品コード"), (0, 1, "商品名"), (0, 2, "単価"), (0, 3, "在庫"),
                (1, 0, "A0001"), (1, 1, "架空りんご"), (1, 2, "1,200"), (1, 3, "15")),
            rowLines: [100, 140, 180],
            columnLines: [50, 200, 350, 500, 650],
            Read);

        Assert.NotNull(table);
        Assert.Equal(2, table.RowCount);
        Assert.Equal(4, table.ColumnCount);
        Assert.True(table.FromRulings);

        var rows = table.ToRows();
        Assert.Equal(["商品コード", "商品名", "単価", "在庫"], rows[0]);
        Assert.Equal(["A0001", "架空りんご", "1,200", "15"], rows[1]);
    }

    [Fact]
    public void AnEmptyCellStaysEmptyInsteadOfShiftingTheOthers()
    {
        var table = ScanTableBuilder.FromRulings(
            Cells(
                (0, 0, "商品コード"), (0, 1, "商品名"), (0, 2, "単価"), (0, 3, "在庫"),
                (1, 0, "A0001"), (1, 2, "1,200"), (1, 3, "15")),
            rowLines: [100, 140, 180],
            columnLines: [50, 200, 350, 500, 650],
            Read);

        Assert.NotNull(table);
        var rows = table.ToRows();

        // 2 列目が空でも、3 列目以降が左へ詰まらない。
        Assert.Equal(["A0001", "", "1,200", "15"], rows[1]);
    }

    [Fact]
    public void TextWrappedInsideACellIsJoined()
    {
        var lines = new List<OcrRawLine>
        {
            FakeOcrEngine.At("商品名", 0.99, new OcrBox(210, 65, 60, 14)),
            FakeOcrEngine.At("在庫", 0.99, new OcrBox(360, 65, 60, 14)),
            FakeOcrEngine.At("架空の", 0.99, new OcrBox(210, 105, 60, 14)),
            FakeOcrEngine.At("長い商品名", 0.99, new OcrBox(210, 122, 90, 14)),
        };

        var table = ScanTableBuilder.FromRulings(
            lines, rowLines: [50, 100, 140], columnLines: [200, 350, 500], Read);

        Assert.NotNull(table);

        // セルの中で行が折り返していても 1 つの値にまとまる。
        Assert.Equal("架空の長い商品名", table.ToRows()[1][0]);
    }

    [Fact]
    public void TooFewRulingsIsNotATable()
    {
        Assert.Null(ScanTableBuilder.FromRulings(
            Cells((0, 0, "何か")), rowLines: [100, 140], columnLines: [50, 200], Read));
    }

    [Fact]
    public void ARuledTableAcrossPagesDropsTheRepeatedHeader()
    {
        using var dir = new TempDir();
        var pdf = dir.File("複数ページ表.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 2);

        var engine = new FakeOcrEngine();
        foreach (var page in new[] { 1, 2 })
        {
            engine.Page(page,
                    FakeOcrEngine.At("商品コード", 0.99, new OcrBox(60, 110, 90, 20)),
                    FakeOcrEngine.At("商品名", 0.99, new OcrBox(210, 110, 90, 20)),
                    FakeOcrEngine.At($"A000{page}", 0.99, new OcrBox(60, 150, 90, 20)),
                    FakeOcrEngine.At("架空の商品", 0.99, new OcrBox(210, 150, 90, 20)))
                .Probe(page, skew: 0, horizontal: 3, vertical: 3)
                .Rulings(page, rows: [100, 140, 180], columns: [50, 200, 350]);
        }

        var reading = new PdfScanReader().Read(engine, pdf, [1, 2]);
        Confirm(reading);

        var issues = new List<Merge.MergeIssue>();
        var rows = PdfReadPlanner.TableRows(reading, issues);

        Assert.Equal(["商品コード", "商品名"], rows[0]);
        Assert.Equal(["A0001", "架空の商品"], rows[1]);
        Assert.Equal(["A0002", "架空の商品"], rows[2]);
        Assert.Equal(3, rows.Count);
        Assert.Contains(issues, issue => issue.Message.Contains("見出しの繰り返し"));
    }

    [Fact]
    public void ATiltedTableIsDeskewedAndThenReadAsATable()
    {
        // Phase 2F-B2 では傾いた表を一律で止めていた。当時はまっすぐな表でも
        // セル一致 61.3% しか無く、傾けると 4.8% まで落ちたため。
        // Paddle Inference 3.3.1 でまっすぐが 94.3% になったので測り直し、
        // 「傾きを直してから行と列へ戻す」ところまで通すようにした。
        using var dir = new TempDir();
        var pdf = dir.File("傾いた表.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine()
            .Page(1,
                FakeOcrEngine.At("商品コード", 0.99, new OcrBox(60, 110, 90, 20)),
                FakeOcrEngine.At("商品名", 0.99, new OcrBox(210, 110, 90, 20)),
                FakeOcrEngine.At("A0001", 0.99, new OcrBox(60, 150, 90, 20)),
                FakeOcrEngine.At("架空の商品", 0.99, new OcrBox(210, 150, 90, 20)))
            .Probe(1, skew: 2.0, horizontal: 3, vertical: 3)
            .Rulings(1, rows: [100, 140, 180], columns: [50, 200, 350]);

        var reading = new PdfScanReader().Read(engine, pdf, [1]);

        // 傾きを直してから読んでいる。
        Assert.Contains(2.0, engine.DeskewAngles);
        Assert.Equal([1], reading.NeedsDeskewPages);

        // 止めずに、行と列のある表として出している。
        Assert.DoesNotContain(reading.Issues, issue => issue.Severity == Merge.MergeIssueSeverity.Block);
        Assert.NotEmpty(reading.Items);
        Assert.All(reading.Items, item => Assert.NotNull(item.Row));
    }

    [Fact]
    public void ATableThatCannotBeRebuiltIsBlockedInsteadOfBeingOutputBroken()
    {
        // 傾きを直しきれないと、区画が噛み合わず「行と列はあるのに中身が空」に
        // なる。この形のまま出すと、表として正しそうに見えて中身が抜け落ちる。
        // 中身のあるセルが半分に満たなければ、表として扱わず理由を示して止める。
        using var dir = new TempDir();
        var pdf = dir.File("戻せない表.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        // 5 行 4 列の区画に対して、文字は 1 か所だけ。
        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.At("A0001", 0.99, new OcrBox(60, 110, 60, 20)))
            .Probe(1, skew: 0, horizontal: 6, vertical: 5)
            .Rulings(1,
                rows: [100, 140, 180, 220, 260, 300],
                columns: [50, 150, 250, 350, 450]);

        var reading = new PdfScanReader().Read(
            engine, pdf, [1], new OcrReadOptions { Mode = OcrReadMode.Table });

        Assert.Contains(reading.Issues, issue =>
            issue.Severity == Merge.MergeIssueSeverity.Block
            && issue.Message.Contains("戻せない", StringComparison.Ordinal));
        Assert.Empty(reading.Items);
    }

    [Fact]
    public void ASparsePageWhoseTiltCannotBeMeasuredIsStillRead()
    {
        // 表でなければ、測れないというだけでは止めない。
        // 見出しだけのページなどは行らしい塊が少なく、測れないのが普通。
        using var dir = new TempDir();
        var pdf = dir.File("見出しだけ.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.Agreed("架空の見出し", 0.99))
            .Probe(1, skew: 0, horizontal: 0, vertical: 0, reliable: false);

        var reading = new PdfScanReader().Read(engine, pdf, [1]);

        Assert.Empty(reading.Issues);
        Assert.Single(reading.Items);
        Assert.Equal([0], engine.DeskewAngles);
    }

    [Fact]
    public void TextOutsideTheDetectedRulingsIsNotThrownAway()
    {
        // 細い外枠の罫線は拾えないことがある(画素の境目に来ると消える)。
        // 実測では 4 列の表が 2 列になり、いちばん左と右の列がまるごと落ちた。
        var extended = ScanTableBuilder.Extend(
            [200.0, 350.0, 500.0],
            [(100.0, 60.0, 140.0), (250.0, 210.0, 290.0), (600.0, 560.0, 640.0)]);

        Assert.Equal(5, extended.Count);
        Assert.True(extended[0] < 100);
        Assert.True(extended[^1] > 600);
    }

    [Fact]
    public void TextThatMerelyOverlapsTheOuterRulingDoesNotAddAColumn()
    {
        // 枠のすぐ内側にある見出しは、端が罫線をまたぐだけ。ここで区切りを増やすと
        // 行や列が 1 つずれる(実測で見出しの行が 1 行下へずれた)。
        var extended = ScanTableBuilder.Extend(
            [200.0, 350.0, 500.0],
            [(250.0, 195.0, 305.0), (400.0, 345.0, 455.0)]);

        Assert.Equal([200.0, 350.0, 500.0], extended);
    }

    // ── 罫線のない表 ───────────────────────────────────

    [Fact]
    public void ABorderlessScanTableIsRebuiltFromTextPositions()
    {
        var lines = new List<OcrRawLine>();
        double[] columns = [60, 210, 380, 500];
        string[] header = ["商品コード", "商品名", "単価", "在庫"];
        string[][] data =
        [
            ["A0001", "架空りんご", "1,200", "15"],
            ["A0002", "架空みかん", "980", "7"],
        ];

        for (var column = 0; column < 4; column++)
        {
            lines.Add(FakeOcrEngine.At(header[column], 0.99, new OcrBox(columns[column], 100, 80, 18)));
        }

        for (var row = 0; row < data.Length; row++)
        {
            for (var column = 0; column < 4; column++)
            {
                lines.Add(FakeOcrEngine.At(
                    data[row][column], 0.99, new OcrBox(columns[column], 130 + (row * 30), 80, 18)));
            }
        }

        var table = ScanTableBuilder.FromAlignment(lines, Read);

        Assert.NotNull(table);
        Assert.False(table.FromRulings);
        Assert.Equal(3, table.RowCount);
        Assert.Equal(4, table.ColumnCount);

        var rows = table.ToRows();
        Assert.Equal(header, rows[0]);
        Assert.Equal(data[0], rows[1]);
        Assert.Equal(data[1], rows[2]);
    }

    [Fact]
    public void ProseIsNotForcedIntoATable()
    {
        // 1 行に 1 かたまりしかない文章は、列が作れないので表にしない。
        var lines = new List<OcrRawLine>
        {
            FakeOcrEngine.At("本書は架空の文書です。", 0.99, new OcrBox(60, 100, 300, 18)),
            FakeOcrEngine.At("記載の数値は実在しません。", 0.99, new OcrBox(60, 130, 300, 18)),
        };

        Assert.Null(ScanTableBuilder.FromAlignment(lines, Read));
    }

    // ── ページの種類の自動判定 ─────────────────────────

    [Theory]
    [InlineData(6, 5, 0, ScanPageKind.RuledTable)]
    [InlineData(8, 0, 8, ScanPageKind.FixedForm)]
    [InlineData(0, 0, 0, ScanPageKind.Prose)]
    public void ThePageKindIsDecidedWithoutAskingTheUser(
        int horizontal, int vertical, int underlines, ScanPageKind expected)
    {
        var metrics = new ScanPageMetrics(horizontal, vertical, 6, 1, 1, underlines);

        Assert.Equal(expected, ScanPageClassifier.Classify(metrics));
    }

    [Fact]
    public void UnderlinesAloneAreNotATable()
    {
        // Phase 2F-A で「記入欄の下線を表と誤判定」した。同じことを起こさない。
        var metrics = new ScanPageMetrics(
            HorizontalRulings: 9, VerticalRulings: 0, LineCount: 8,
            AlignedRowCount: 0, ColumnCount: 1, UnderlineCount: 9);

        Assert.Equal(ScanPageKind.FixedForm, ScanPageClassifier.Classify(metrics));
    }

    [Fact]
    public void AlignedColumnsWithoutRulingsAreABorderlessTable()
    {
        var metrics = new ScanPageMetrics(
            HorizontalRulings: 0, VerticalRulings: 0, LineCount: 5,
            AlignedRowCount: 5, ColumnCount: 4, UnderlineCount: 0);

        Assert.Equal(ScanPageKind.BorderlessTable, ScanPageClassifier.Classify(metrics));
    }

    // ── 定型帳票 ──────────────────────────────────────

    [Fact]
    public void EveryExpectedFieldComesBackEvenWhenNothingWasRead()
    {
        // B1 で 813 項目中 89 件が「項目ごと消えて」いた。これを起こさない。
        var template = Template();
        var readings = FormFieldExtractor.Read(
            template,
            // 店舗コードの場所にだけ読み取りがある。
            [FakeOcrEngine.At("A0001", 0.99, new OcrBox(200, 100, 90, 20))],
            FormOffset.None,
            Read);

        Assert.Equal(template.Fields.Count, readings.Count);
        Assert.Equal(
            template.Fields.Select(field => field.Name),
            readings.Select(reading => reading.Name));

        Assert.True(readings[0].WasFound);
        Assert.All(readings.Skip(1), reading => Assert.False(reading.WasFound));
    }

    [Fact]
    public void AFieldThatCouldNotBeReadShowsUpAsMissing()
    {
        using var dir = new TempDir();
        var pdf = dir.File("帳票.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.At("A0001", 0.99, new OcrBox(200, 100, 90, 20)))
            .Probe(1, skew: 0, horizontal: 0, vertical: 0);

        var reading = new PdfScanReader().Read(
            engine, pdf, [1], new OcrReadOptions
            {
                Mode = OcrReadMode.FixedForm,
                Template = Template(),
            });

        // 読めた 1 件 + 読めなかった 2 件 = 指定した 3 件。項目は消えない。
        Assert.Equal(3, reading.Items.Count);
        Assert.Equal(2, reading.MissingCount);
        Assert.Equal(
            ["店舗コード", "売上", "備考"],
            reading.Items.Select(item => item.FieldName));

        var missing = reading.Items.Single(item => item.FieldName == "売上");
        Assert.True(missing.IsMissing);
        Assert.Equal(OcrItemStatus.Missing, missing.Status);
        Assert.False(missing.IsResolved);

        // 読む場所は分かっているので、元のページのその位置を出せる。
        Assert.Equal(200, missing.BoundingBox.X);
        Assert.Equal(140, missing.BoundingBox.Y);
    }

    [Fact]
    public void MissingFieldsKeepTheOutputBlockedUntilAPersonFillsThemIn()
    {
        using var dir = new TempDir();
        var pdf = dir.File("欠け.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.At("A0001", 0.99, new OcrBox(200, 100, 90, 20)))
            .Probe(1, skew: 0, horizontal: 0, vertical: 0);

        var planner = new PdfReadPlanner();
        var preview = planner.CreatePreview(new PdfReadRequest { SourceFilePath = pdf }, UsablePack);
        var reading = new PdfScanReader().Read(
            engine, pdf, preview.OcrPageNumbers,
            new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = Template() });

        var blocked = planner.CompleteWithOcr(preview, reading);
        Assert.False(blocked.CanExecute);
        Assert.Contains(blocked.Blocks, issue => issue.Message.Contains("見つからない 2 件"));

        // 店舗コードは読めているが、0 と O を取り違えても分からないので
        // 自動確定しない(FieldAutoAcceptPolicy)。人が見て、そのままでよいと決める。
        var code = reading.Items.Single(item => item.FieldName == "店舗コード");
        Assert.Equal(OcrItemStatus.NeedsReview, code.InitialStatus);
        code.Confirm();

        // 人が原文を見て入れれば出せる。
        foreach (var item in reading.Items.Where(item => !item.IsResolved))
        {
            item.Confirm("人が読み取った内容");
        }

        var done = planner.CompleteWithOcr(preview, reading);
        Assert.True(done.CanExecute);

        var rows = done.TableRows;
        Assert.Equal(["ページ", "店舗コード", "売上", "備考"], rows[0]);
        Assert.Equal(["1", "A0001", "人が読み取った内容", "人が読み取った内容"], rows[1]);
    }

    [Fact]
    public void TheNumberOfExpectedFieldsIsAlwaysPreserved()
    {
        using var dir = new TempDir();
        var pdf = dir.File("多ページ帳票.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 20);

        var engine = new FakeOcrEngine();
        for (var page = 1; page <= 20; page++)
        {
            // 3 ページに 1 回、売上だけが読めないページを混ぜる。
            var lines = page % 3 == 0
                ? new[] { FakeOcrEngine.At($"A{page:D4}", 0.99, new OcrBox(200, 100, 90, 20)) }
                : [
                    FakeOcrEngine.At($"A{page:D4}", 0.99, new OcrBox(200, 100, 90, 20)),
                    FakeOcrEngine.At("1,200", 0.99, new OcrBox(200, 140, 90, 20)),
                ];

            engine.Page(page, lines).Probe(page, skew: 0, horizontal: 0, vertical: 0);
        }

        var reading = new PdfScanReader().Read(
            engine, pdf, [.. Enumerable.Range(1, 20)],
            new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = Template() });

        // 指定 3 項目 × 20 ページ = 60 件。1 件も欠けない。
        Assert.Equal(60, reading.Items.Count);
        Assert.Equal(
            60,
            reading.AutoAcceptedCount + reading.NeedsReviewCount
                + reading.UnreadableCount + reading.MissingCount);
    }

    // ── 位置ずれ・大きさ・傾きへの追随 ────────────────────

    [Theory]
    [InlineData(0, 0)]
    [InlineData(10, 8)]
    [InlineData(-12, 6)]
    [InlineData(20, -15)]
    public void ThePageCanShiftAndTheFieldsAreStillFound(double dx, double dy)
    {
        var template = Template() with
        {
            Anchors = [new FormAnchor("店舗コード", new OcrBox(60, 100, 90, 20))],
        };

        var lines = new List<OcrRawLine>
        {
            FakeOcrEngine.At("店舗コード", 0.99, new OcrBox(60 + dx, 100 + dy, 90, 20)),
            FakeOcrEngine.At("A0001", 0.99, new OcrBox(200 + dx, 100 + dy, 90, 20)),
            FakeOcrEngine.At("1,200", 0.99, new OcrBox(200 + dx, 140 + dy, 90, 20)),
        };

        var offset = FormFieldExtractor.FindOffset(template, lines, line => Read(line).Text);
        var readings = FormFieldExtractor.Read(template, lines, offset, Read);

        Assert.Equal(dx, offset.X, 3);
        Assert.Equal(dy, offset.Y, 3);
        Assert.Equal("A0001", readings[0].Text);
        Assert.Equal("1,200", readings[1].Text);
    }

    [Fact]
    public void AnAnchorThatMovedTooFarIsIgnoredRatherThanTrusted()
    {
        var template = Template() with
        {
            Anchors = [new FormAnchor("店舗コード", new OcrBox(60, 100, 90, 20))],
        };

        // まったく別の場所に同じ文字がある場合、そこまで動かすと壊れる。
        var lines = new List<OcrRawLine>
        {
            FakeOcrEngine.At("店舗コード", 0.99, new OcrBox(60, 900, 90, 20)),
        };

        Assert.Equal(FormOffset.None, FormFieldExtractor.FindOffset(
            template, lines, line => Read(line).Text));
    }

    [Fact]
    public void WithoutAnchorsTheTemplateIsNotMoved()
    {
        var lines = new List<OcrRawLine>
        {
            FakeOcrEngine.At("A0001", 0.99, new OcrBox(200, 100, 90, 20)),
        };

        Assert.Equal(FormOffset.None, FormFieldExtractor.FindOffset(
            Template(), lines, line => Read(line).Text));
    }

    // ── 印(チェック / 丸 / ばつ) ────────────────────

    [Theory]
    [InlineData(0.02, 0.01, MarkDecision.None)]
    [InlineData(0.45, 0.01, MarkDecision.Selected)]
    [InlineData(0.90, 0.01, MarkDecision.Selected)]
    [InlineData(0.08, 0.01, MarkDecision.Unclear)]
    public void AMarkIsJudgedFromInkNotFromText(
        double firstInk, double secondInk, MarkDecision expected)
    {
        var result = MarkClassifier.Classify(
        [
            new MarkSample("はい", firstInk, 0),
            new MarkSample("いいえ", secondInk, 0),
        ]);

        Assert.Equal(expected, result.Decision);
    }

    [Fact]
    public void ACircledLabelIsFoundAroundTheLabelNotInsideTheBox()
    {
        var result = MarkClassifier.Classify(
        [
            new MarkSample("はい", 0.01, 0.40),
            new MarkSample("いいえ", 0.01, 0.01),
        ]);

        Assert.Equal(MarkDecision.Selected, result.Decision);
        Assert.Equal("はい", result.Label);
    }

    [Fact]
    public void TwoMarksAtOnceAreNeverDecidedAutomatically()
    {
        var result = MarkClassifier.Classify(
        [
            new MarkSample("はい", 0.42, 0),
            new MarkSample("いいえ", 0.40, 0),
        ]);

        Assert.Equal(MarkDecision.Unclear, result.Decision);
        Assert.Equal(OcrItemStatus.NeedsReview, MarkClassifier.ToStatus(result));
    }

    [Fact]
    public void AFaintMarkGoesToReviewInsteadOfBeingGuessed()
    {
        var result = MarkClassifier.Classify([new MarkSample("はい", 0.09, 0)]);

        Assert.Equal(MarkDecision.Unclear, result.Decision);
        Assert.Equal(OcrItemStatus.NeedsReview, MarkClassifier.ToStatus(result));
    }

    [Fact]
    public void NothingMarkedIsAlsoSomethingAPersonConfirms()
    {
        var result = MarkClassifier.Classify(
        [
            new MarkSample("はい", 0.01, 0),
            new MarkSample("いいえ", 0.01, 0),
        ]);

        Assert.Equal(MarkDecision.None, result.Decision);
        Assert.Equal(OcrItemStatus.NeedsReview, MarkClassifier.ToStatus(result));
        Assert.Equal(string.Empty, result.Text);
    }

    [Fact]
    public void AClearMarkOnAFormIsReadWithoutOcr()
    {
        using var dir = new TempDir();
        var pdf = dir.File("印.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var template = new FormTemplate
        {
            Name = "架空の帳票",
            Fields =
            [
                new FormField
                {
                    Name = "回答",
                    Area = new OcrBox(200, 200, 200, 30),
                    Kind = FormFieldKind.Choice,
                    Choices =
                    [
                        new FormChoice("はい", new OcrBox(200, 200, 20, 20)),
                        new FormChoice("いいえ", new OcrBox(300, 200, 20, 20)),
                    ],
                },
            ],
        };

        // 箱の中 / 丸囲み の順で 2 つずつ。「はい」の箱だけ濃い。
        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.At("回答", 0.99, new OcrBox(60, 200, 60, 20)))
            .Probe(1, skew: 0, horizontal: 0, vertical: 0)
            .Ink(1, 0.55, 0.01, 0.01, 0.01);

        var reading = new PdfScanReader().Read(
            engine, pdf, [1],
            new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = template });

        var item = Assert.Single(reading.Items);
        Assert.Equal("回答", item.FieldName);
        Assert.Equal("はい", item.Text);
        Assert.Equal(OcrItemStatus.AutoAccepted, item.Status);
    }

    // ── 確認と出力 ────────────────────────────────────

    [Fact]
    public void ATableCellCanBeReviewedAgainstTheOriginalPage()
    {
        using var dir = new TempDir();
        var pdf = dir.File("表確認.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine()
            .Page(1,
                FakeOcrEngine.At("商品コード", 0.99, new OcrBox(60, 110, 90, 20)),
                FakeOcrEngine.At("単価", 0.99, new OcrBox(210, 110, 90, 20)),
                FakeOcrEngine.Split("A0001", 0.6, "AOOO1", 0.5, y: 150, x: 60),
                FakeOcrEngine.At("1,200", 0.99, new OcrBox(210, 150, 90, 20)))
            .Probe(1, skew: 0, horizontal: 3, vertical: 3)
            .Rulings(1, rows: [100, 140, 180], columns: [50, 200, 350]);

        var reading = new PdfScanReader().Read(engine, pdf, [1]);

        var uncertain = reading.Items.Single(item => item.Row == 1 && item.Column == 0);
        Assert.Equal(OcrItemStatus.NeedsReview, uncertain.Status);

        // セルにも元のページ上の位置がある。
        Assert.True(uncertain.BoundingBox.Width > 0);
        Assert.Equal(1, uncertain.Row);
        Assert.Equal(0, uncertain.Column);
    }

    [Fact]
    public void ATableStillCannotBeWrittenUntilEveryCellIsResolved()
    {
        using var dir = new TempDir();
        var pdf = dir.File("表未確認.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine()
            .Page(1,
                FakeOcrEngine.At("商品コード", 0.99, new OcrBox(60, 110, 90, 20)),
                FakeOcrEngine.At("単価", 0.99, new OcrBox(210, 110, 90, 20)),
                FakeOcrEngine.Split("A0001", 0.6, "AOOO1", 0.5, y: 150, x: 60),
                FakeOcrEngine.At("1,200", 0.99, new OcrBox(210, 150, 90, 20)))
            .Probe(1, skew: 0, horizontal: 3, vertical: 3)
            .Rulings(1, rows: [100, 140, 180], columns: [50, 200, 350]);

        var planner = new PdfReadPlanner();
        var preview = planner.CreatePreview(new PdfReadRequest { SourceFilePath = pdf }, UsablePack);
        var reading = new PdfScanReader().Read(engine, pdf, preview.OcrPageNumbers);

        Assert.False(planner.CompleteWithOcr(preview, reading).CanExecute);

        Confirm(reading);
        var done = planner.CompleteWithOcr(preview, reading);

        Assert.True(done.CanExecute);
        Assert.True(new PdfReader().Execute(done).Success);
        Assert.Equal(["商品コード", "単価"], done.TableRows[0]);
        Assert.Equal(["A0001", "1,200"], done.TableRows[1]);
    }

    [Fact]
    public void TheSourcePdfIsStillNeverChanged()
    {
        using var dir = new TempDir();
        var pdf = dir.File("元.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);
        var before = Snapshot(pdf);

        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.At("A0001", 0.99, new OcrBox(200, 100, 90, 20)))
            .Probe(1, skew: 2.0, horizontal: 0, vertical: 0);

        var planner = new PdfReadPlanner();
        var preview = planner.CreatePreview(new PdfReadRequest { SourceFilePath = pdf }, UsablePack);
        var reading = new PdfScanReader().Read(
            engine, pdf, preview.OcrPageNumbers,
            new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = Template() });

        Confirm(reading);
        Assert.True(new PdfReader().Execute(planner.CompleteWithOcr(preview, reading)).Success);

        Assert.Equal(before, Snapshot(pdf));
        Assert.Empty(Directory.GetFiles(dir.Root, "*.tmp"));
    }

    [Fact]
    public void CancellingAFormRunLeavesNothingBehind()
    {
        using var dir = new TempDir();
        var pdf = dir.File("中止.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 3);

        using var cancellation = new CancellationTokenSource();
        var engine = new FakeOcrEngine();
        for (var page = 1; page <= 3; page++)
        {
            engine.Page(page, FakeOcrEngine.At("A0001", 0.99, new OcrBox(200, 100, 90, 20)))
                .Probe(page, skew: 0, horizontal: 0, vertical: 0);
        }

        engine.OnRead = page =>
        {
            if (page == 2)
            {
                cancellation.Cancel();
            }
        };

        Assert.Throws<OperationCanceledException>(() => new PdfScanReader().Read(
            engine, pdf, [1, 2, 3],
            new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = Template() },
            null, cancellation.Token));

        Assert.Equal([pdf], Directory.GetFiles(dir.Root));
    }

    [Fact]
    public void TheAuditRecordsWhatWasReadWithoutTheTextItself()
    {
        using var dir = new TempDir();
        var pdf = dir.File("控え.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 2);

        var engine = new FakeOcrEngine();
        foreach (var page in new[] { 1, 2 })
        {
            engine.Page(page,
                    FakeOcrEngine.At($"S{page:D3}", 0.99, new OcrBox(200, 100, 90, 20)))
                .Probe(page, skew: page == 2 ? 2.0 : 0, horizontal: 0, vertical: 0);
        }

        var planner = new PdfReadPlanner();
        var preview = planner.CreatePreview(new PdfReadRequest { SourceFilePath = pdf }, UsablePack);
        var reading = new PdfScanReader().Read(
            engine, pdf, preview.OcrPageNumbers,
            new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = Template() });

        Confirm(reading);
        Assert.True(new PdfReader().Execute(planner.CompleteWithOcr(preview, reading)).Success);

        var text = File.ReadAllText(dir.File("控え_PDF抽出.xlsx.audit.json"));
        var audit = System.Text.Json.JsonDocument.Parse(text).RootElement;
        var ocr = audit.GetProperty("ocr");

        Assert.Equal("fixed-form", ocr.GetProperty("mode").GetString());
        Assert.Equal(1, ocr.GetProperty("deskewedPageCount").GetInt32());
        Assert.Equal(2, ocr.GetProperty("formPageCount").GetInt32());
        Assert.Equal(6, ocr.GetProperty("expectedFieldCount").GetInt32());
        Assert.Equal(6, ocr.GetProperty("itemCount").GetInt32());
        Assert.Equal(4, ocr.GetProperty("missingCount").GetInt32());

        // 読み取った文字そのものは控えへ複製しない。
        Assert.DoesNotContain("S001", text, StringComparison.Ordinal);
        Assert.DoesNotContain(dir.Root, text, StringComparison.Ordinal);
    }

    // ── ヘルパー ────────────────────────────────────

    private static OcrPackStatus UsablePack
        => new(IsPresent: true, IsUsable: true, "OCR Pack を使えます。", "テスト");

    // --- 列の形による自動確定の抑制(Phase 2F-B2) ---
    //
    // 上下逆に切り出されたセルは、両モデルへ同じ間違った画像が入るので
    // 一致も自信も当てにならない(実測 A0096 → 9600 を自信 99.8% で自動確定)。
    // 気づけるのは「同じ列の他の行と文字の種類が違う」ことだけ。

    [Fact]
    public void ColumnShapeIsIgnoredWhenThereAreTooFewSamples()
        => Assert.Equal(
            ColumnShapeGuard.Shape.None,
            ColumnShapeGuard.MajorityShape(["A0001", "A0002", "A0003"]));

    [Fact]
    public void ColumnShapeIsIgnoredWhenNoShapeStandsOut()
    {
        // 3 種類が均等(それぞれ 1/3)なら「その列らしい形」は決められない。
        var majority = ColumnShapeGuard.MajorityShape(
            ["A0001", "B0002", "C0003", "1234", "5678", "9012", "架空", "備考", "見本"]);
        Assert.Equal(ColumnShapeGuard.Shape.None, majority);
    }

    [Fact]
    public void CellThatDiffersFromItsColumnIsNotAutoAccepted()
    {
        // 実測と同じ割れ方(英字+数字 11 / 数字 9 / 日本語 1)。
        var column = Enumerable.Repeat("A0001", 11)
            .Concat(Enumerable.Repeat("9600", 9))
            .Append("架空")
            .ToList();

        var majority = ColumnShapeGuard.MajorityShape(column);
        Assert.Equal(
            ColumnShapeGuard.Shape.Latin | ColumnShapeGuard.Shape.Digit, majority);

        Assert.True(ColumnShapeGuard.CanAutoAccept(majority, "A0096"));
        Assert.False(ColumnShapeGuard.CanAutoAccept(majority, "9600"));

        // 空のセルは形で疑わない(読めなかったことは別に扱う)。
        Assert.True(ColumnShapeGuard.CanAutoAccept(majority, string.Empty));
    }

    [Fact]
    public void AmountsOfDifferentLengthCountAsTheSameShape()
    {
        // 桁区切りを数えると「999」と「1,234」が別物になり、正しいセルまで人へ回る。
        // 記号を数えない決まりは、この失敗を実際に踏んでから入れた。
        var column = new[] { "1,234", "12,345", "999", "1,000,000", "22" };
        var majority = ColumnShapeGuard.MajorityShape(column);

        Assert.NotEqual(ColumnShapeGuard.Shape.None, majority);
        foreach (var text in column)
        {
            Assert.True(ColumnShapeGuard.CanAutoAccept(majority, text));
        }
    }

    [Fact]
    public void UpsideDownCellInATableGoesToReviewWithReason()
    {
        using var dir = new TempDir();
        var pdf = dir.File("逆さ.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        // 1 列目は「英字 + 数字」の並び。1 行だけ数字だけに読めている。
        var lines = new List<OcrRawLine>();
        for (var row = 0; row < 6; row++)
        {
            var text = row == 3 ? "9600" : $"A000{row}";
            lines.Add(FakeOcrEngine.At(text, 0.99, new OcrBox(80, 100 + (row * 40), 90, 20)));
            lines.Add(FakeOcrEngine.At($"架空{row}", 0.99, new OcrBox(240, 100 + (row * 40), 90, 20)));
        }

        var engine = new FakeOcrEngine()
            .Page(1, [.. lines])
            .Probe(1, skew: 0, horizontal: 0, vertical: 0);

        var reading = new PdfScanReader().Read(
            engine, pdf, [1], new OcrReadOptions { Mode = OcrReadMode.Table });

        var odd = reading.Items.Single(item => item.Text == "9600");
        Assert.Equal(OcrItemStatus.NeedsReview, odd.InitialStatus);
        Assert.Contains("上下逆", odd.Reason);

        // 同じ列の他の行は巻き添えにしない。
        var normal = reading.Items.Single(item => item.Text == "A0002");
        Assert.Equal(OcrItemStatus.AutoAccepted, normal.InitialStatus);
    }

    // --- 項目の種類ごとの自動確定の抑制(Phase 2F-B2) ---
    //
    // 実測で誤って自動確定した 4 件はすべて「S001-24」を「SO01-24」と読んだもので、
    // 自信は 98.4〜98.8% あった。2 つのモデルは同じ字形の取り違えを共有するため、
    // 一致も自信も根拠にならない。形で止める。

    [Theory]
    [InlineData("SO01-24", false)]  // O は 0 と取り違えやすい
    [InlineData("S001-24", false)]  // 0 も同じく判別できない。どちらでも止める
    [InlineData("AX-CDEF", true)]   // 取り違えやすい字が無いので自動確定してよい
    [InlineData("AB-CDEF", false)]  // B は 8 と取り違えやすい
    [InlineData("", false)]
    public void CodeFieldAutoAcceptsOnlyWhenNoConfusableGlyph(string text, bool expected)
        => Assert.Equal(expected, FieldAutoAcceptPolicy.CanAutoAccept(FormFieldKind.Code, text));

    [Theory]
    [InlineData("1,234,567", true)]
    [InlineData("-1,234", true)]
    [InlineData("1,O34", false)]     // 数字のはずの場所に英字が出たら形が壊れている
    [InlineData("1234 円", false)]
    public void NumberFieldAutoAcceptsOnlyWhenShapeIsNumeric(string text, bool expected)
        => Assert.Equal(expected, FieldAutoAcceptPolicy.CanAutoAccept(FormFieldKind.NumberLike, text));

    [Theory]
    [InlineData(FormFieldKind.Text)]
    [InlineData(FormFieldKind.Choice)]
    public void ProseFieldsAreNotRestricted(FormFieldKind kind)
        => Assert.True(FieldAutoAcceptPolicy.CanAutoAccept(kind, "架空の担当者"));

    [Fact]
    public void ConfidentButConfusableCodeGoesToReviewWithReason()
    {
        using var dir = new TempDir();
        var pdf = dir.File("コード.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine()
            .Page(1,
                FakeOcrEngine.At("SO01-24", 0.99, new OcrBox(200, 100, 120, 30)),
                FakeOcrEngine.At("1234567", 0.99, new OcrBox(200, 140, 120, 30)),
                FakeOcrEngine.At("架空の備考", 0.99, new OcrBox(200, 180, 300, 30)))
            .Probe(1, skew: 0, horizontal: 0, vertical: 0);

        var planner = new PdfReadPlanner();
        var preview = planner.CreatePreview(new PdfReadRequest { SourceFilePath = pdf }, UsablePack);
        var reading = new PdfScanReader().Read(
            engine, pdf, preview.OcrPageNumbers,
            new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = Template() });

        var code = reading.Items.Single(item => item.FieldName == "店舗コード");
        Assert.Equal(OcrItemStatus.NeedsReview, code.InitialStatus);
        Assert.Contains("取り違えやすい字", code.Reason);
        Assert.Equal("SO01-24", code.Text);

        // 自信で落としたのではないことを明示しておく。
        Assert.True(code.Confidence >= OcrFusion.AutoAcceptThreshold);

        // 他の種類は巻き添えにしない。
        Assert.Equal(OcrItemStatus.AutoAccepted,
            reading.Items.Single(item => item.FieldName == "売上").InitialStatus);
        Assert.Equal(OcrItemStatus.AutoAccepted,
            reading.Items.Single(item => item.FieldName == "備考").InitialStatus);
    }

    private static FormTemplate Template() => new()
    {
        Name = "架空の売上報告",
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
            new FormField
            {
                Name = "備考", Area = new OcrBox(200, 180, 300, 30), Kind = FormFieldKind.Text,
            },
        ],
    };

    private static (string Text, double Confidence) Read(OcrRawLine line)
    {
        var result = OcrFusion.Fuse(line);
        return (result.Text, result.Confidence);
    }

    private static List<OcrRawLine> Cells(params (int Row, int Column, string Text)[] cells)
    {
        double[] rowCenters = [120, 160, 200];
        double[] columnCenters = [120, 270, 420, 570];

        return [.. cells.Select(cell => FakeOcrEngine.At(
            cell.Text, 0.99,
            new OcrBox(columnCenters[cell.Column] - 40, rowCenters[cell.Row] - 9, 80, 18)))];
    }

    private static void Confirm(OcrDocumentReading reading)
    {
        foreach (var item in reading.Items.Where(item => !item.IsResolved))
        {
            item.Confirm();
        }
    }

    private static (string Sha, long Length, DateTime Modified) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }
}
