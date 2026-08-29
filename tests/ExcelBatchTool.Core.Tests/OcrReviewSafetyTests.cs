using ExcelBatchTool.Core.Ocr;
using ExcelBatchTool.Core.Pdf;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2F-B1.1。元のページを見ながら 1 件ずつ確認する、という筋道を守るための決まり。
/// すべて架空データ。
/// </summary>
public sealed class OcrReviewSafetyTests
{
    // ── 選んだ項目とページ画像の対応 ───────────────────────

    [Fact]
    public void SelectingAnItem_PointsAtThePageItWasReadFrom()
    {
        var reading = Reading(
            (1, FakeOcrEngine.Split("1 ページ目", 0.5, "1 ページ目?", 0.4, y: 100)),
            (3, FakeOcrEngine.Split("3 ページ目", 0.5, "3 ページ目?", 0.4, y: 200)));

        var session = new OcrReviewSession(reading);
        session.SelectFirstUnresolved();

        Assert.Equal(1, session.Selected!.PageNumber);

        session.MoveToNextUnresolved();

        Assert.Equal(3, session.Selected!.PageNumber);
    }

    [Fact]
    public void TheReadingPositionMapsOntoTheDisplayedImage()
    {
        // OCR は 300dpi、画面に出すのは 150dpi。つまり画像の座標は半分になる。
        var image = new OcrPageImage(1, [], 1240, 1754, ScaleFromOcr: 0.5);
        var box = new OcrBox(600, 800, 400, 60);

        var atActualSize = OcrBoxMapper.ToDisplay(box, image, zoom: 1);
        var atDoubleSize = OcrBoxMapper.ToDisplay(box, image, zoom: 2);

        Assert.Equal(300, atActualSize.Left);
        Assert.Equal(400, atActualSize.Top);
        Assert.Equal(200, atActualSize.Width);
        Assert.Equal(30, atActualSize.Height);

        Assert.Equal(600, atDoubleSize.Left);
        Assert.Equal(800, atDoubleSize.Top);
        Assert.Equal(400, atDoubleSize.Width);
        Assert.Equal(60, atDoubleSize.Height);
    }

    [Fact]
    public void AVeryThinReadingStillGetsAVisibleFrame()
    {
        var image = new OcrPageImage(1, [], 1240, 1754, ScaleFromOcr: 0.5);

        var rect = OcrBoxMapper.ToDisplay(new OcrBox(10, 10, 2, 1), image, zoom: 1);

        Assert.True(rect.Width >= 4);
        Assert.True(rect.Height >= 4);
    }

    [Fact]
    public void FitZoom_MakesTheWholePageVisible()
    {
        var image = new OcrPageImage(1, [], 1240, 1754, ScaleFromOcr: 0.5);

        var zoom = OcrBoxMapper.FitZoom(image, viewportWidth: 620, viewportHeight: 400);

        // 縦がきつい向きなので、縦に合わせた倍率になる。
        Assert.Equal(400.0 / 1754, zoom, 6);
        Assert.True(image.Width * zoom <= 620);
        Assert.True(image.Height * zoom <= 400);
    }

    [Fact]
    public void ScrollToShow_CentresTheReading()
    {
        var rect = new OcrDisplayRect(500, 300, 100, 40);

        var (left, top) = OcrBoxMapper.ScrollToShow(rect, viewportWidth: 400, viewportHeight: 200);

        Assert.Equal(350, left);
        Assert.Equal(220, top);
    }

    // ── ページ画像を持ちすぎない ──────────────────────────

    [Fact]
    public void OnlyAFewPageImagesAreKept()
    {
        var rendered = new List<int>();
        var cache = new OcrPageImageCache(
            page =>
            {
                rendered.Add(page);
                return new OcrPageImage(page, [], 100, 100, 0.5);
            },
            capacity: 3);

        for (var page = 1; page <= 20; page++)
        {
            cache.Get(page);
        }

        Assert.Equal(20, rendered.Count);
        Assert.Equal(3, cache.Count);
        Assert.Equal([18, 19, 20], cache.Pages);
    }

    [Fact]
    public void ThePageBeingLookedAtIsNotRenderedTwice()
    {
        var cache = new OcrPageImageCache(
            page => new OcrPageImage(page, [], 100, 100, 0.5), capacity: 3);

        cache.Get(5);
        cache.Get(5);
        cache.Get(5);

        Assert.Equal(1, cache.RenderCount);
        Assert.Equal(1, cache.Count);
    }

    [Fact]
    public void GoingBackAndForthDoesNotGrowTheCache()
    {
        var cache = new OcrPageImageCache(
            page => new OcrPageImage(page, [], 100, 100, 0.5), capacity: 3);

        for (var round = 0; round < 40; round++)
        {
            cache.Get(1 + (round % 12));
            cache.Preload(1 + (round % 12), pageCount: 120);
        }

        Assert.True(cache.Count <= 3, $"手元に {cache.Count} 枚残っている");
    }

    [Fact]
    public void ClearingTheCacheReleasesEverything()
    {
        var cache = new OcrPageImageCache(
            page => new OcrPageImage(page, [], 100, 100, 0.5), capacity: 3);
        cache.Get(1);
        cache.Get(2);

        cache.Clear();

        Assert.Equal(0, cache.Count);
        Assert.Empty(cache.Pages);
    }

    [Fact]
    public void PreloadNeverPushesOutThePageBeingLookedAt()
    {
        var cache = new OcrPageImageCache(
            page => new OcrPageImage(page, [], 100, 100, 0.5), capacity: 3);

        cache.Get(10);
        cache.Preload(10, pageCount: 120);

        Assert.Contains(10, cache.Pages);
        Assert.True(cache.Count <= 3);
    }

    // ── まとめて確認済みにはできない ──────────────────────

    [Fact]
    public void ThereIsNoWayToConfirmEverythingAtOnce()
    {
        // 「表示中をすべて確認済みにする」のような操作は置かない。
        // 元のページと見比べずに確認済みにできると、この段階の安全が意味を失う。
        var methods = typeof(OcrReviewSession)
            .GetMethods()
            .Select(method => method.Name)
            .ToList();

        Assert.DoesNotContain(methods, name =>
            name.Contains("All", StringComparison.Ordinal)
            || name.Contains("Bulk", StringComparison.Ordinal));
    }

    [Fact]
    public void ConfirmingOnlyAffectsTheSelectedItem()
    {
        var reading = Reading(
            (1, FakeOcrEngine.Split("1 件目", 0.5, "1 件目?", 0.4, y: 100)),
            (1, FakeOcrEngine.Split("2 件目", 0.5, "2 件目?", 0.4, y: 200)),
            (1, FakeOcrEngine.Split("3 件目", 0.5, "3 件目?", 0.4, y: 300)));

        var session = new OcrReviewSession(reading);
        session.SelectFirstUnresolved();
        session.ConfirmSelectedAndAdvance();

        Assert.Equal(1, reading.UserConfirmedCount);
        Assert.Equal(2, reading.UnresolvedCount);
    }

    // ── 確認して次へ ─────────────────────────────────

    [Fact]
    public void ConfirmingMovesToTheNextUnresolvedItem()
    {
        var reading = Reading(
            (1, FakeOcrEngine.Split("1 件目", 0.5, "1 件目?", 0.4, y: 100)),
            (1, FakeOcrEngine.Agreed("自動で確定する", 0.99, y: 200)),
            (2, FakeOcrEngine.Split("3 件目", 0.5, "3 件目?", 0.4, y: 100)));

        var session = new OcrReviewSession(reading);
        session.SelectFirstUnresolved();

        // 読みが割れた領域は、日本語が混ざるので日本語モデルの読みを見せる。
        Assert.Equal("1 件目?", session.Selected!.Text);

        session.ConfirmSelectedAndAdvance();

        // 自動確定は飛ばして、次の未確認へ。
        Assert.Equal("3 件目?", session.Selected!.Text);
        Assert.Equal(2, session.Selected.PageNumber);
    }

    [Fact]
    public void ConfirmingWithACorrection_KeepsTheOriginalReading()
    {
        var reading = Reading((1, FakeOcrEngine.Split("架空商亊", 0.9, "架空商事", 0.9, y: 100)));

        var session = new OcrReviewSession(reading);
        session.SelectFirstUnresolved();
        session.ConfirmSelectedAndAdvance("架空商事株式会社");

        var item = reading.Items.Single();
        Assert.Equal("架空商事株式会社", item.FinalText);
        Assert.Equal("架空商事", item.Text);
        Assert.True(item.IsUserEdited);
        Assert.Equal(OcrItemStatus.UserConfirmed, item.Status);
    }

    [Fact]
    public void ConfirmingTheOriginal_DoesNotMarkItAsEdited()
    {
        var reading = Reading((1, FakeOcrEngine.Split("1200", 0.9, "12OO", 0.5, y: 100)));

        var session = new OcrReviewSession(reading);
        session.SelectFirstUnresolved();
        session.ConfirmSelectedAndAdvance();

        var item = reading.Items.Single();
        Assert.False(item.IsUserEdited);
        Assert.Equal("1200", item.FinalText);
        Assert.Equal(OcrItemStatus.UserConfirmed, item.Status);
    }

    [Fact]
    public void MovingBackwardsFindsThePreviousUnresolvedItem()
    {
        var reading = Reading(
            (1, FakeOcrEngine.Split("1 件目", 0.5, "1 件目?", 0.4, y: 100)),
            (1, FakeOcrEngine.Split("2 件目", 0.5, "2 件目?", 0.4, y: 200)));

        var session = new OcrReviewSession(reading);
        session.SelectFirstUnresolved();
        session.MoveToNextUnresolved();
        Assert.Equal("2 件目?", session.Selected!.Text);

        Assert.True(session.MoveToPreviousUnresolved());
        Assert.Equal("1 件目?", session.Selected!.Text);
    }

    [Fact]
    public void WhenNothingIsLeftUnresolved_TheSelectionStaysPut()
    {
        var reading = Reading((1, FakeOcrEngine.Split("1 件目", 0.5, "1 件目?", 0.4, y: 100)));

        var session = new OcrReviewSession(reading);
        session.SelectFirstUnresolved();
        var confirmed = session.Selected;

        Assert.False(session.ConfirmSelectedAndAdvance());
        Assert.Same(confirmed, session.Selected);
        Assert.True(session.IsComplete);
    }

    // ── 自動確定も見られる ───────────────────────────

    [Fact]
    public void AutoAcceptedItemsAreHiddenByDefaultButCanBeShown()
    {
        var reading = Reading(
            (1, FakeOcrEngine.Agreed("自動で確定する", 0.99, y: 100)),
            (1, FakeOcrEngine.Split("要確認", 0.5, "要確認?", 0.4, y: 200)));

        var session = new OcrReviewSession(reading);

        Assert.Single(session.Visible);
        Assert.Equal("要確認?", session.Visible[0].Text);

        session.ShowAutoAccepted = true;

        // 自動確定したものも、元のページと見比べられる。
        Assert.Equal(2, session.Visible.Count);
    }

    [Fact]
    public void ConfirmedItemsStayInTheListSoTheyCanBeUndone()
    {
        var reading = Reading((1, FakeOcrEngine.Split("要確認", 0.5, "要確認?", 0.4, y: 100)));

        var session = new OcrReviewSession(reading);
        session.SelectFirstUnresolved();
        session.ConfirmSelectedAndAdvance();

        Assert.Single(session.Visible);
        Assert.Equal(OcrItemStatus.UserConfirmed, session.Visible[0].Status);
    }

    // ── 読取不能も原文を見て直せる ────────────────────

    [Fact]
    public void AnUnreadableItemStillHasAPageAndAPositionToLookAt()
    {
        var reading = Reading((2, FakeOcrEngine.Split(string.Empty, 0, string.Empty, 0, y: 400)));

        var item = reading.Items.Single();

        Assert.Equal(OcrItemStatus.Unreadable, item.Status);
        Assert.Equal(2, item.PageNumber);
        Assert.Equal(400, item.BoundingBox.Y);
    }

    [Fact]
    public void AnUnreadableItemCanBeTypedInByHand()
    {
        var reading = Reading((1, FakeOcrEngine.Split(string.Empty, 0, string.Empty, 0, y: 100)));

        var session = new OcrReviewSession(reading);
        session.SelectFirstUnresolved();
        session.ConfirmSelectedAndAdvance("人が読み取った内容");

        Assert.True(reading.IsComplete);
        Assert.Equal("人が読み取った内容", reading.Items.Single().FinalText);
    }

    [Fact]
    public void AnUnreadableItemLeftAloneKeepsTheOutputBlocked()
    {
        using var dir = new TempDir();
        var pdf = dir.File("読めない.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine().Page(1,
            FakeOcrEngine.Agreed("読める行", 0.99, y: 100),
            FakeOcrEngine.Split(string.Empty, 0, string.Empty, 0, y: 200));

        var planner = new PdfReadPlanner();
        var preview = planner.CreatePreview(new PdfReadRequest { SourceFilePath = pdf }, UsablePack);
        var reading = new PdfScanReader().Read(engine, pdf, preview.OcrPageNumbers);

        var done = planner.CompleteWithOcr(preview, reading);

        Assert.False(done.CanExecute);
        Assert.Contains(done.Blocks, issue => issue.Message.Contains("読取不能 1 件"));
        Assert.False(new PdfReader().Execute(done).Success);
        Assert.Empty(Directory.GetFiles(dir.Root, "*.xlsx"));
    }

    // ── 全部確認すれば出力できる ─────────────────────

    [Fact]
    public void ConfirmingEveryItemOneByOneAllowsTheOutput()
    {
        using var dir = new TempDir();
        var pdf = dir.File("順に確認.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 2);

        var engine = new FakeOcrEngine()
            .Page(1,
                FakeOcrEngine.Split("架空商亊", 0.9, "架空商事", 0.9, y: 100),
                FakeOcrEngine.Agreed("自動で確定する行", 0.99, y: 200))
            .Page(2, FakeOcrEngine.Split(string.Empty, 0, string.Empty, 0, y: 100));

        var planner = new PdfReadPlanner();
        var preview = planner.CreatePreview(new PdfReadRequest { SourceFilePath = pdf }, UsablePack);
        var reading = new PdfScanReader().Read(engine, pdf, preview.OcrPageNumbers);
        var session = new OcrReviewSession(reading);

        Assert.False(planner.CompleteWithOcr(preview, reading).CanExecute);

        session.SelectFirstUnresolved();
        session.ConfirmSelectedAndAdvance("架空商事株式会社");
        Assert.False(planner.CompleteWithOcr(preview, reading).CanExecute);

        session.ConfirmSelectedAndAdvance("人が読み取った内容");

        var done = planner.CompleteWithOcr(preview, reading);
        Assert.True(done.CanExecute);
        Assert.True(new PdfReader().Execute(done).Success);

        var rows = PdfReader.ToRows(done);
        Assert.Equal(["1", "1", "架空商事株式会社"], rows[1]);
        Assert.Equal(["1", "2", "自動で確定する行"], rows[2]);
        Assert.Equal(["2", "1", "人が読み取った内容"], rows[3]);
    }

    // ── データ元は読み取りのみ・後始末 ──────────────────

    [Fact]
    public void LookingAtPageImagesDoesNotChangeTheSourcePdf()
    {
        using var dir = new TempDir();
        var pdf = dir.File("画像確認.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 3);
        var before = Snapshot(pdf);

        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.Agreed("1 ページ目", 0.99))
            .Page(2, FakeOcrEngine.Agreed("2 ページ目", 0.99))
            .Page(3, FakeOcrEngine.Agreed("3 ページ目", 0.99));

        using var source = engine.Open(pdf);
        var cache = new OcrPageImageCache(
            page => source.RenderPage(page, 150, CancellationToken.None), capacity: 3);

        for (var round = 0; round < 10; round++)
        {
            cache.Get(1 + (round % 3));
        }

        Assert.Equal(before, Snapshot(pdf));
        Assert.Empty(Directory.GetFiles(dir.Root, "*.png"));
        Assert.Empty(Directory.GetFiles(dir.Root, "*.tmp"));
    }

    [Fact]
    public void PageImagesAreKeptInMemoryNotOnDisk()
    {
        using var dir = new TempDir();
        var pdf = dir.File("一時ファイルなし.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 2);

        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.Agreed("1 ページ目", 0.99))
            .Page(2, FakeOcrEngine.Agreed("2 ページ目", 0.99));

        using var source = engine.Open(pdf);
        var image = source.RenderPage(1, 150, CancellationToken.None);

        Assert.NotEmpty(image.Png);
        Assert.Equal([pdf], Directory.GetFiles(dir.Root));
    }

    [Fact]
    public void CancellingDuringReadingLeavesNoPageImagesBehind()
    {
        using var dir = new TempDir();
        var pdf = dir.File("中止.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 3);

        using var cancellation = new CancellationTokenSource();
        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.Agreed("1", 0.99))
            .Page(2, FakeOcrEngine.Agreed("2", 0.99))
            .Page(3, FakeOcrEngine.Agreed("3", 0.99));
        engine.OnRead = page =>
        {
            if (page == 2)
            {
                cancellation.Cancel();
            }
        };

        Assert.Throws<OperationCanceledException>(
            () => new PdfScanReader().Read(engine, pdf, [1, 2, 3], null, cancellation.Token));

        Assert.Equal([pdf], Directory.GetFiles(dir.Root));
    }

    // ── 自動確定にしてよい条件 ──────────────────────────

    [Fact]
    public void OnlyIdenticalConfidentReadingsAreEverAutoAccepted()
    {
        // 「自動確定にしたのに間違っていた」を 0 にしている条件そのもの。
        // 架空 fixture での実測(誤確定 0)は、この不変条件が守られている限り成り立つ。
        var texts = new[] { "1,200", "1.200", "架空商事", "架空商亊", "A0001", "AOOO1", string.Empty };
        var scores = new[] { 0.0, 0.5, 0.9, 0.97, 0.9799, 0.98, 0.999, 1.0 };

        foreach (var multi in texts)
        {
            foreach (var japan in texts)
            {
                foreach (var multiScore in scores)
                {
                    foreach (var japanScore in scores)
                    {
                        var result = OcrFusion.Fuse(
                            FakeOcrEngine.Split(multi, multiScore, japan, japanScore));

                        if (result.Status != OcrItemStatus.AutoAccepted)
                        {
                            continue;
                        }

                        Assert.Equal(multi, japan);
                        Assert.NotEqual(string.Empty, multi);
                        Assert.True(
                            Math.Min(multiScore, japanScore) >= OcrFusion.AutoAcceptThreshold,
                            $"自信 {Math.Min(multiScore, japanScore)} で自動確定になった");
                    }
                }
            }
        }
    }

    [Fact]
    public void TwoConfidentButDifferentReadingsAreNeverAutoAccepted()
    {
        // 2 つのモデルが両方間違うこともある。だから一致しても自信が足りなければ止め、
        // 割れていれば自信に関係なく人へ回す。
        var result = OcrFusion.Fuse(FakeOcrEngine.Split("2026/02/10", 1.0, "2026/02/1O", 1.0));

        Assert.Equal(OcrItemStatus.NeedsReview, result.Status);
    }

    [Fact]
    public void TheAutoAcceptThresholdIsTheMeasuredOne()
    {
        // 0.98 / 0.985 / 0.99 / 0.995 / 0.999 を実測し、誤確定はどれも 0 だった。
        // 上げても安全は増えず自動確定だけが減るので、いちばん低い 0.98 を使う。
        Assert.Equal(0.98, OcrFusion.AutoAcceptThreshold);
    }

    // ── ヘルパー ────────────────────────────────────

    private static OcrPackStatus UsablePack
        => new(IsPresent: true, IsUsable: true, "OCR Pack を使えます。", "テスト");

    /// <summary>ページ番号を指定して読み取り結果を組み立てる。</summary>
    private static OcrDocumentReading Reading(params (int Page, OcrRawLine Line)[] lines)
    {
        var items = new List<OcrItem>();
        foreach (var group in lines.GroupBy(entry => entry.Page).OrderBy(group => group.Key))
        {
            items.AddRange(
                PdfScanReader.BuildItems(group.Key, [.. group.Select(entry => entry.Line)]));
        }

        return new OcrDocumentReading
        {
            Items = items,
            OcrPages = [.. lines.Select(entry => entry.Page).Distinct().Order()],
            EngineInfo = new OcrEngineInfo("テスト多言語", "テスト日本語", "テスト", "テスト", 300),
        };
    }

    private static (string Sha, long Length, DateTime Modified) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }
}
