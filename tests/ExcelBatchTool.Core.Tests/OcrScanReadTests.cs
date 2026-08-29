using System.Text;
using System.Text.Json;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Ocr;
using ExcelBatchTool.Core.Pdf;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2F-B1。スキャン PDF の OCR と、確認・修正してから出力するまで。
/// すべて架空データ。Offline OCR Pack を置かずに、差し替えた OCR で筋道を確かめる。
/// </summary>
public sealed class OcrScanReadTests
{
    // ── A. きれいな日本語のスキャン ───────────────────────

    [Fact]
    public void AJapaneseScan_IsReadAndKeepsPageAndLine()
    {
        using var dir = new TempDir();
        var pdf = dir.File("案内.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 2);

        var engine = new FakeOcrEngine()
            .Page(1,
                FakeOcrEngine.Agreed("架空商事株式会社", 0.99, y: 100),
                FakeOcrEngine.Agreed("本書は架空の文書です。", 0.99, y: 200))
            .Page(2, FakeOcrEngine.Agreed("第二段落の内容です。", 0.99, y: 100));

        var reading = Read(engine, pdf, [1, 2]);

        Assert.Equal(3, reading.Items.Count);
        Assert.Equal(3, reading.AutoAcceptedCount);
        Assert.True(reading.IsComplete);

        var lines = LinesOf(reading);
        Assert.Equal((1, 1, "架空商事株式会社"), (lines[0].Page, lines[0].Line, lines[0].Text));
        Assert.Equal((1, 2, "本書は架空の文書です。"), (lines[1].Page, lines[1].Line, lines[1].Text));
        Assert.Equal((2, 1, "第二段落の内容です。"), (lines[2].Page, lines[2].Line, lines[2].Text));
    }

    // ── B. 数字・コード中心のスキャン ─────────────────────

    [Fact]
    public void ACodeScan_PrefersTheMultilingualModelWhenEverythingIsAscii()
    {
        // 英数字だけの領域は多言語モデル、日本語が混ざる領域は日本語モデル。
        // どちらを採ったかは Reason に出す。
        var code = OcrFusion.Fuse(FakeOcrEngine.Split("A0001-7", 0.99, "AOOO1-7", 0.80));
        var japanese = OcrFusion.Fuse(FakeOcrEngine.Split("金额:1,200", 0.99, "金額:1,200円", 0.98));

        Assert.Equal("A0001-7", code.Text);
        Assert.Contains("多言語", code.Reason, StringComparison.Ordinal);

        Assert.Equal("金額:1,200円", japanese.Text);
        Assert.Contains("日本語", japanese.Reason, StringComparison.Ordinal);
    }

    // ── C. 劣化スキャン ───────────────────────────────

    [Fact]
    public void ADegradedScan_SendsEverythingUncertainToReview()
    {
        using var dir = new TempDir();
        var pdf = dir.File("劣化.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine().Page(1,
            FakeOcrEngine.Split("架空商事", 0.71, "架空商亊", 0.68, y: 100),
            FakeOcrEngine.Agreed("1,200 円", 0.72, y: 200));

        var reading = Read(engine, pdf, [1]);

        Assert.Equal(0, reading.AutoAcceptedCount);
        Assert.Equal(2, reading.NeedsReviewCount);
        Assert.False(reading.IsComplete);
    }

    // ── D. 文字ページとスキャンページの混在 ────────────────

    [Fact]
    public void AMixedPdf_ReadsOnlyTheScannedPagesWithOcr()
    {
        using var dir = new TempDir();
        var pdf = dir.File("混在.pdf");
        TestPdfFactory.CreateMixed(pdf, textPages: 2, imagePages: 1);

        var preview = Preview(pdf, UsablePack);

        Assert.Equal(PdfReadStage.NeedsOcr, preview.Stage);
        Assert.Equal([3], preview.OcrPageNumbers);
        Assert.Equal(
            [PdfPageRoute.BornDigitalText, PdfPageRoute.BornDigitalText, PdfPageRoute.Scan],
            preview.PagePlans.Select(plan => plan.Route));

        var engine = new FakeOcrEngine().Page(3, FakeOcrEngine.Agreed("画像のページの内容", 0.99));
        var reading = new PdfScanReader().Read(engine, pdf, preview.OcrPageNumbers);

        // OCR にかけるのは画像のページだけ。文字情報のあるページは通さない。
        Assert.Equal([3], engine.ReadPages);

        var done = new PdfReadPlanner().CompleteWithOcr(preview, reading);

        Assert.Equal(PdfReadStage.Ready, done.Stage);
        Assert.True(done.CanExecute);
        Assert.Contains(done.Lines, line => line.Page == 1);
        Assert.Contains(done.Lines, line => line.Page == 2);
        Assert.Contains(done.Lines, line => line is { Page: 3, Text: "画像のページの内容" });
    }

    [Fact]
    public void AMixedPdfWithATablePage_IsBlockedInsteadOfBeingForcedIntoOneTable()
    {
        using var dir = new TempDir();
        var pdf = dir.File("表と画像.pdf");
        TestPdfFactory.CreateTableThenImage(pdf, SampleRows(4));

        var preview = Preview(pdf, UsablePack);

        Assert.Equal(PdfReadStage.Blocked, preview.Stage);
        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("表のページとスキャンのページ", StringComparison.Ordinal));
    }

    // ── E. 白紙のスキャン ─────────────────────────────

    [Fact]
    public void ABlankScan_IsBlockedInsteadOfProducingAnEmptyFile()
    {
        using var dir = new TempDir();
        var pdf = dir.File("白紙.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine().Page(1);
        var reading = Read(engine, pdf, [1]);

        Assert.Empty(reading.Items);
        Assert.Contains(reading.Issues, issue => issue.Severity == MergeIssueSeverity.Block);

        var done = new PdfReadPlanner().CompleteWithOcr(Preview(pdf, UsablePack), reading);
        Assert.False(done.CanExecute);
    }

    // ── F / G. 自信の低い読み・自信が高いのに違う読み ────────

    [Fact]
    public void LowConfidenceAgreement_IsNotAutoAccepted()
    {
        var result = OcrFusion.Fuse(FakeOcrEngine.Agreed("架空商事", 0.97));

        Assert.Equal(OcrItemStatus.NeedsReview, result.Status);
    }

    [Fact]
    public void ConfidentButDisagreeingModels_AreNeverAutoAccepted()
    {
        // どちらのモデルも自信満々でも、読みが割れているなら人が見る。
        // これが「自動確定にしたのに間違っていた」を防ぐいちばんの働き。
        var result = OcrFusion.Fuse(FakeOcrEngine.Split("本動作確の生成架空の文。", 0.99, "本書は動作確認のために生成した架空の文書です。", 0.99));

        Assert.Equal(OcrItemStatus.NeedsReview, result.Status);
        Assert.Equal("本書は動作確認のために生成した架空の文書です。", result.Text);
    }

    [Fact]
    public void ANonFiniteScore_IsTreatedAsNoConfidence()
    {
        // 認識器が NaN / Infinity を返すことがある(実測)。
        // そのまま比べると閾値を通ってしまうので 0 に倒す。
        var infinity = OcrFusion.Fuse(new OcrRawLine(
            new OcrBox(0, 0, 100, 30), "1200", double.PositiveInfinity, "1200", double.PositiveInfinity));
        var notANumber = OcrFusion.Fuse(new OcrRawLine(
            new OcrBox(0, 0, 100, 30), "1200", double.NaN, "1200", double.NaN));

        Assert.Equal(OcrItemStatus.NeedsReview, infinity.Status);
        Assert.Equal(OcrItemStatus.NeedsReview, notANumber.Status);
        Assert.Equal(0, infinity.Confidence);
    }

    [Fact]
    public void AnEmptyReading_IsUnreadable()
    {
        var result = OcrFusion.Fuse(FakeOcrEngine.Split(string.Empty, 0.1, string.Empty, 0.1));

        Assert.Equal(OcrItemStatus.Unreadable, result.Status);
        Assert.Equal("読取不能", OcrItemStatusText.Display(result.Status));
    }

    // ── H / I. OCR Pack が無い・壊れている ─────────────────

    [Fact]
    public void AScanPdf_WithoutTheOcrPack_IsBlockedWithAClearReason()
    {
        using var dir = new TempDir();
        var pdf = dir.File("スキャン.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var preview = Preview(pdf, OcrPack.Inspect(Path.Combine(dir.Root, "ocr-なし")));

        Assert.Equal(PdfReadStage.Blocked, preview.Stage);
        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue =>
            issue.Message.Contains("OCR が必要", StringComparison.Ordinal)
            && issue.Message.Contains("Offline OCR Pack", StringComparison.Ordinal));
    }

    [Fact]
    public void ABornDigitalPdf_WorksWithoutTheOcrPack()
    {
        using var dir = new TempDir();
        var pdf = dir.File("報告書.pdf");
        TestPdfFactory.CreateText(pdf, [["月次報告書(架空)", "金額: 1,200 円"]]);

        var preview = Preview(pdf, pack: null);

        Assert.Equal(PdfReadStage.Ready, preview.Stage);
        Assert.True(preview.CanExecute);
    }

    [Fact]
    public void AMissingPack_IsReportedWithoutThrowing()
    {
        using var dir = new TempDir();

        var status = OcrPack.Inspect(Path.Combine(dir.Root, "ocr"));

        Assert.False(status.IsPresent);
        Assert.False(status.IsUsable);
        Assert.Contains("Offline OCR Pack", status.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("manifest-missing")]
    [InlineData("manifest-broken")]
    [InlineData("wrong-version")]
    [InlineData("file-missing")]
    [InlineData("wrong-length")]
    [InlineData("wrong-hash")]
    public void ABrokenPack_IsBlockedBeforeAnythingNativeIsLoaded(string damage)
    {
        using var dir = new TempDir();
        var pack = Path.Combine(dir.Root, "ocr");
        BuildPack(pack, damage);

        var status = OcrPack.Inspect(pack);

        Assert.True(status.IsPresent);
        Assert.False(status.IsUsable);
        Assert.Contains("入れ直して", status.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnIntactPack_PassesTheCheck()
    {
        using var dir = new TempDir();
        var pack = Path.Combine(dir.Root, "ocr");
        BuildPack(pack, damage: null);

        var status = OcrPack.Inspect(pack);

        Assert.True(status.IsUsable);
    }

    // ── J. 中止 ─────────────────────────────────────

    [Fact]
    public void Cancelling_StopsReadingAndLeavesNothingBehind()
    {
        using var dir = new TempDir();
        var pdf = dir.File("長い.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 3);
        var before = Snapshot(pdf);

        using var cancellation = new CancellationTokenSource();
        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.Agreed("1 ページ目", 0.99))
            .Page(2, FakeOcrEngine.Agreed("2 ページ目", 0.99))
            .Page(3, FakeOcrEngine.Agreed("3 ページ目", 0.99));
        engine.OnRead = page =>
        {
            if (page == 2)
            {
                cancellation.Cancel();
            }
        };

        Assert.Throws<OperationCanceledException>(
            () => new PdfScanReader().Read(engine, pdf, [1, 2, 3], null, cancellation.Token));

        Assert.Equal(before, Snapshot(pdf));
        Assert.Empty(Directory.GetFiles(dir.Root, "*.xlsx"));
        Assert.Empty(Directory.GetFiles(dir.Root, "*.csv"));
        Assert.Empty(Directory.GetFiles(dir.Root, "*.audit.json"));
        Assert.Empty(Directory.GetFiles(dir.Root, "*.tmp"));
    }

    // ── K. 読み取り中にデータ元が変わった ────────────────

    [Fact]
    public void IfThePdfChangesAfterReading_TheRunIsRefused()
    {
        using var dir = new TempDir();
        var pdf = dir.File("差し替え.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var preview = Preview(pdf, UsablePack);
        var engine = new FakeOcrEngine().Page(1, FakeOcrEngine.Agreed("読み取った内容", 0.99));
        var reading = new PdfScanReader().Read(engine, pdf, preview.OcrPageNumbers);
        var done = new PdfReadPlanner().CompleteWithOcr(preview, reading);
        Assert.True(done.CanExecute);

        // 確認が終わったあとで、元の PDF が別のものに置き換わった。
        TestPdfFactory.CreateImageOnly(pdf, pages: 2);

        var result = new PdfReader().Execute(done);

        Assert.False(result.Success);
        Assert.Contains("変更されています", result.Message, StringComparison.Ordinal);
        Assert.Empty(Directory.GetFiles(dir.Root, "*.xlsx"));
    }

    // ── L / M. 同名衝突と後始末 ──────────────────────

    [Fact]
    public void AnExistingOutput_IsNeverOverwritten()
    {
        using var dir = new TempDir();
        var pdf = dir.File("控え.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var existing = dir.File("控え_PDF抽出.xlsx");
        File.WriteAllText(existing, "先にあったファイル");

        var done = ReadAndConfirm(pdf, engine => engine.Page(1, FakeOcrEngine.Agreed("内容", 0.99)));

        Assert.False(done.CanExecute);
        Assert.Contains(done.Blocks, issue => issue.Message.Contains("すでにあります", StringComparison.Ordinal));
        Assert.Equal("先にあったファイル", File.ReadAllText(existing));
    }

    [Fact]
    public void AfterASuccessfulRun_NoWorkFileIsLeftBehind()
    {
        using var dir = new TempDir();
        var pdf = dir.File("後始末.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var done = ReadAndConfirm(pdf, engine => engine.Page(1, FakeOcrEngine.Agreed("内容", 0.99)));
        var result = new PdfReader().Execute(done);

        Assert.True(result.Success, result.Message);
        Assert.Empty(Directory.GetFiles(dir.Root, "*.tmp"));
    }

    // ── N. データ元は読み取りのみ ────────────────────

    [Fact]
    public void TheSourcePdf_IsNeverChanged()
    {
        using var dir = new TempDir();
        var pdf = dir.File("元.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);
        var before = Snapshot(pdf);

        var done = ReadAndConfirm(pdf, engine => engine.Page(1, FakeOcrEngine.Agreed("内容", 0.99)));
        var result = new PdfReader().Execute(done);

        Assert.True(result.Success, result.Message);
        Assert.Equal(before, Snapshot(pdf));
    }

    // ── O / P. 修正する・元のままで正しいと確認する ─────────

    [Fact]
    public void AUserEdit_ReplacesTheOutputTextAndKeepsTheOriginal()
    {
        using var dir = new TempDir();
        var pdf = dir.File("修正.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine().Page(1, FakeOcrEngine.Split("架空商亊", 0.9, "架空商事", 0.9));
        var reading = Read(engine, pdf, [1]);
        var item = Assert.Single(reading.Items);

        item.Confirm("架空商事株式会社");

        Assert.Equal(OcrItemStatus.UserConfirmed, item.Status);
        Assert.True(item.IsUserEdited);
        Assert.Equal("架空商事株式会社", item.FinalText);

        // 元の読みは残す(あとから「元は何だったか」を見られるようにする)。
        Assert.Equal("架空商事", item.Text);
        Assert.Equal("架空商亊", item.OriginalEngineResults[0].Text);
        Assert.Equal("架空商事", item.OriginalEngineResults[1].Text);

        Assert.Equal("架空商事株式会社", LinesOf(reading)[0].Text);
    }

    [Fact]
    public void ConfirmingTheOriginal_ResolvesTheItemWithoutEditing()
    {
        var item = SingleItem(FakeOcrEngine.Split("1200", 0.9, "12OO", 0.5));

        item.Confirm();

        Assert.Equal(OcrItemStatus.UserConfirmed, item.Status);
        Assert.False(item.IsUserEdited);
        Assert.Equal("1200", item.FinalText);
    }

    [Fact]
    public void JustLookingAtAnItem_DoesNotResolveIt()
    {
        var item = SingleItem(FakeOcrEngine.Agreed("架空商事", 0.5));

        _ = item.Text;
        _ = item.OriginalEngineResults;

        Assert.Equal(OcrItemStatus.NeedsReview, item.Status);
        Assert.False(item.IsResolved);
    }

    [Fact]
    public void ConfirmationCanBeUndone()
    {
        var item = SingleItem(FakeOcrEngine.Split("1200", 0.9, "12OO", 0.5));
        item.Confirm("1,200");

        item.Unconfirm();

        Assert.Equal(OcrItemStatus.NeedsReview, item.Status);
        Assert.False(item.IsUserEdited);
        Assert.Equal("1200", item.FinalText);
    }

    // ── Q / R. 未確認が残っていると出力できない ───────────

    [Fact]
    public void UnresolvedItems_BlockTheOutput()
    {
        using var dir = new TempDir();
        var pdf = dir.File("未確認.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine().Page(1,
            FakeOcrEngine.Agreed("確かな行", 0.99, y: 100),
            FakeOcrEngine.Split("怪しい行", 0.6, "怪しい行?", 0.5, y: 200));

        var preview = Preview(pdf, UsablePack);
        var reading = new PdfScanReader().Read(engine, pdf, preview.OcrPageNumbers);
        var done = new PdfReadPlanner().CompleteWithOcr(preview, reading);

        Assert.False(done.CanExecute);
        Assert.Contains(done.Blocks, issue =>
            issue.Message.Contains("確認が済んでいない", StringComparison.Ordinal));

        var result = new PdfReader().Execute(done);
        Assert.False(result.Success);
        Assert.Empty(Directory.GetFiles(dir.Root, "*.xlsx"));
    }

    [Fact]
    public void OnceEverythingIsResolved_TheOutputIsAllowed()
    {
        using var dir = new TempDir();
        var pdf = dir.File("確認済み.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine().Page(1,
            FakeOcrEngine.Agreed("確かな行", 0.99, y: 100),
            FakeOcrEngine.Split("怪しい行", 0.6, "怪しい行?", 0.5, y: 200));

        var preview = Preview(pdf, UsablePack);
        var reading = new PdfScanReader().Read(engine, pdf, preview.OcrPageNumbers);

        foreach (var item in reading.Items.Where(item => !item.IsResolved))
        {
            item.Confirm();
        }

        var done = new PdfReadPlanner().CompleteWithOcr(preview, reading);

        Assert.True(done.CanExecute);
        Assert.True(new PdfReader().Execute(done).Success);
    }

    // ── S. NFKC ──────────────────────────────────────

    [Fact]
    public void ReadTextIsNormalizedWithNfkc()
    {
        // 康熙部首(⽉ U+2F49)は見た目が同じで別の文字。NFKC で揃える。
        var item = SingleItem(FakeOcrEngine.Agreed("⽉次報告書", 0.99));

        Assert.Equal("月次報告書", item.Text);
        Assert.Equal("月次報告書", item.OriginalEngineResults[0].Text);
    }

    [Fact]
    public void AUserEditIsNormalizedTheSameWay()
    {
        var item = SingleItem(FakeOcrEngine.Agreed("報告書", 0.99));

        item.Confirm("⽉次報告書");

        Assert.Equal("月次報告書", item.FinalText);
    }

    // ── T / U. CSV と XLSX の出力 ─────────────────────

    [Fact]
    public void TheOutputCanBeCsvWithSpecialCharactersQuoted()
    {
        using var dir = new TempDir();
        var pdf = dir.File("記号.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var done = ReadAndConfirm(
            pdf,
            engine => engine.Page(1, FakeOcrEngine.Agreed("金額: 1,200 円「特価」", 0.99)),
            format: PdfOutputFormat.Csv);

        Assert.True(new PdfReader().Execute(done).Success);

        var csv = File.ReadAllText(dir.File("記号_PDF抽出.csv"), Encoding.UTF8);

        Assert.Contains("ページ,行,内容\r\n", csv, StringComparison.Ordinal);
        Assert.Contains("\"金額: 1,200 円「特価」\"", csv, StringComparison.Ordinal);
    }

    [Fact]
    public void TheOutputCanBeXlsxWithPageAndLineColumns()
    {
        using var dir = new TempDir();
        var pdf = dir.File("帳票.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var done = ReadAndConfirm(
            pdf,
            engine => engine.Page(1,
                FakeOcrEngine.Agreed("架空商事株式会社", 0.99, y: 100),
                FakeOcrEngine.Agreed("金額: 1,200 円", 0.99, y: 200)));

        Assert.True(new PdfReader().Execute(done).Success);
        Assert.True(File.Exists(dir.File("帳票_PDF抽出.xlsx")));

        var rows = PdfReader.ToRows(done);

        Assert.Equal(["ページ", "行", "内容"], rows[0]);
        Assert.Equal(["1", "1", "架空商事株式会社"], rows[1]);
        Assert.Equal(["1", "2", "金額: 1,200 円"], rows[2]);
    }

    // ── V. 控え ─────────────────────────────────────

    [Fact]
    public void TheAuditRecordsTheCountsButNeverTheReadText()
    {
        using var dir = new TempDir();
        var pdf = dir.File("控え確認.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine().Page(1,
            FakeOcrEngine.Agreed("自動で確定する行", 0.99, y: 100),
            FakeOcrEngine.Split("要確認の行", 0.6, "要確認の行?", 0.5, y: 200),
            FakeOcrEngine.Split(string.Empty, 0.0, string.Empty, 0.0, y: 300));

        var preview = Preview(pdf, UsablePack);
        var reading = new PdfScanReader().Read(engine, pdf, preview.OcrPageNumbers);

        reading.Items[1].Confirm("要確認の行");
        reading.Items[2].Confirm("読めなかったので入れ直した");

        var done = new PdfReadPlanner().CompleteWithOcr(preview, reading);
        Assert.True(new PdfReader().Execute(done).Success);

        var audit = JsonDocument.Parse(
            File.ReadAllText(dir.File("控え確認_PDF抽出.xlsx.audit.json"))).RootElement;
        var ocr = audit.GetProperty("ocr");

        Assert.Equal(2, audit.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("pdf-extract", audit.GetProperty("operation").GetString());
        Assert.Equal("ocr-dual-read", audit.GetProperty("source").GetProperty("extractionMethod").GetString());
        Assert.Equal("NFKC", audit.GetProperty("source").GetProperty("normalization").GetString());

        Assert.Equal(1, ocr.GetProperty("ocrPageCount").GetInt32());
        Assert.Equal(3, ocr.GetProperty("itemCount").GetInt32());
        Assert.Equal(1, ocr.GetProperty("autoAcceptedCount").GetInt32());
        Assert.Equal(1, ocr.GetProperty("needsReviewCount").GetInt32());
        Assert.Equal(1, ocr.GetProperty("unreadableCount").GetInt32());
        Assert.Equal(2, ocr.GetProperty("userConfirmedCount").GetInt32());
        Assert.Equal(2, ocr.GetProperty("userEditedCount").GetInt32());
        Assert.Equal(0.98, ocr.GetProperty("autoAcceptThreshold").GetDouble());

        // 読み取った文字そのものは控えへ複製しない。
        var text = File.ReadAllText(dir.File("控え確認_PDF抽出.xlsx.audit.json"));
        Assert.DoesNotContain("自動で確定する行", text, StringComparison.Ordinal);
        Assert.DoesNotContain("要確認の行", text, StringComparison.Ordinal);
        Assert.DoesNotContain(dir.Root, text, StringComparison.Ordinal);
    }

    [Fact]
    public void ABornDigitalAuditKeepsSchemaVersionOne()
    {
        using var dir = new TempDir();
        var pdf = dir.File("文字だけ.pdf");
        TestPdfFactory.CreateText(pdf, [["月次報告書(架空)", "金額: 1,200 円"]]);

        var preview = Preview(pdf, UsablePack);
        Assert.True(new PdfReader().Execute(preview).Success);

        var audit = JsonDocument.Parse(
            File.ReadAllText(dir.File("文字だけ_PDF抽出.xlsx.audit.json"))).RootElement;

        Assert.Equal(1, audit.GetProperty("schemaVersion").GetInt32());
        Assert.False(audit.TryGetProperty("ocr", out _));
    }

    // ── 傾き・表らしいスキャン(次の段階へ送る) ────────────

    [Fact]
    public void ATiltedScan_IsBlockedBeforeAnyRecognitionRuns()
    {
        using var dir = new TempDir();
        var pdf = dir.File("傾き.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 2);

        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.Agreed("内容", 0.99))
            .Page(2, FakeOcrEngine.Agreed("内容", 0.99))
            .Probe(2, skew: 3.4, horizontal: 0, vertical: 0);

        var reading = Read(engine, pdf, [1, 2]);

        Assert.Equal([2], reading.NeedsDeskewPages);
        Assert.Contains(reading.Issues, issue => issue.Message.Contains("傾いています", StringComparison.Ordinal));

        // 何分もかかる認識を始める前に止める。
        Assert.Empty(engine.ReadPages);
        Assert.Empty(reading.Items);
    }

    [Fact]
    public void ATableLikeScan_IsBlockedAndKeptForTheNextStage()
    {
        using var dir = new TempDir();
        var pdf = dir.File("表スキャン.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.Agreed("商品コード", 0.99))
            .Probe(1, skew: 0, horizontal: 6, vertical: 5);

        var reading = Read(engine, pdf, [1]);

        Assert.Equal([1], reading.TableLikePages);
        Assert.Contains(reading.Issues, issue =>
            issue.Message.Contains("スキャンされた表の可能性", StringComparison.Ordinal)
            && issue.Message.Contains("次の段階", StringComparison.Ordinal));
        Assert.Empty(engine.ReadPages);
    }

    [Fact]
    public void UnderlinesAlone_AreNotTreatedAsATable()
    {
        using var dir = new TempDir();
        var pdf = dir.File("記入欄.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 1);

        // 記入欄の下線のように横線だけが並ぶページは表とみなさない。
        var engine = new FakeOcrEngine()
            .Page(1, FakeOcrEngine.Agreed("氏名", 0.99))
            .Probe(1, skew: 0, horizontal: 8, vertical: 0);

        var reading = Read(engine, pdf, [1]);

        Assert.Empty(reading.TableLikePages);
        Assert.Equal([1], engine.ReadPages);
    }

    // ── 行の組み立て ──────────────────────────────────

    [Fact]
    public void RegionsOnTheSameLine_AreJoinedLeftToRight()
    {
        var raw = new[]
        {
            FakeOcrEngine.Agreed("2 列目", 0.99, y: 100, x: 400),
            FakeOcrEngine.Agreed("1 列目", 0.99, y: 102, x: 60),
            FakeOcrEngine.Agreed("次の行", 0.99, y: 200, x: 60),
        };

        var items = PdfScanReader.BuildItems(1, raw);

        Assert.Equal([1, 1, 2], items.Select(item => item.LineNumber));
        Assert.Equal([0, 1, 0], items.Select(item => item.IndexInLine));
        Assert.Equal("1 列目", items[0].Text);
        Assert.Equal("2 列目", items[1].Text);
    }

    // ── ヘルパー ────────────────────────────────────

    private static OcrPackStatus UsablePack
        => new(IsPresent: true, IsUsable: true, "OCR Pack を使えます。", "テスト");

    private static PdfReadPreview Preview(string pdf, OcrPackStatus? pack)
        => new PdfReadPlanner().CreatePreview(new PdfReadRequest { SourceFilePath = pdf }, pack);

    private static OcrDocumentReading Read(FakeOcrEngine engine, string pdf, IReadOnlyList<int> pages)
        => new PdfScanReader().Read(engine, pdf, pages);

    private static IReadOnlyList<PdfTextLine> LinesOf(OcrDocumentReading reading)
        => PdfReadPlanner.ToLines(reading);

    private static OcrItem SingleItem(OcrRawLine line)
        => PdfScanReader.BuildItems(1, [line]).Single();

    private static PdfReadPreview ReadAndConfirm(
        string pdf,
        Func<FakeOcrEngine, FakeOcrEngine> setup,
        PdfOutputFormat format = PdfOutputFormat.Xlsx)
    {
        var planner = new PdfReadPlanner();
        var preview = planner.CreatePreview(
            new PdfReadRequest { SourceFilePath = pdf, OutputFormat = format }, UsablePack);

        var engine = setup(new FakeOcrEngine());
        var reading = new PdfScanReader().Read(engine, pdf, preview.OcrPageNumbers);

        foreach (var item in reading.Items.Where(item => !item.IsResolved))
        {
            item.Confirm();
        }

        return planner.CompleteWithOcr(preview, reading);
    }

    private static (string Sha, long Length, DateTime Modified) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }

    private static string[][] SampleRows(int count)
    {
        var rows = new List<string[]> { new[] { "商品コード", "商品名", "単価", "在庫" } };
        for (var index = 1; index <= count; index++)
        {
            rows.Add([$"A{index:D4}", "架空の商品", (1000 + (index * 7)).ToString(), (index * 3).ToString()]);
        }

        return [.. rows];
    }

    /// <summary>検査用の小さな Pack を作る(本物のモデルは要らない)。</summary>
    private static void BuildPack(string directory, string? damage)
    {
        Directory.CreateDirectory(directory);
        var enginePath = Path.Combine(directory, "ExcelBatchTool.Ocr.dll");
        File.WriteAllText(enginePath, "架空の中身");

        var bytes = File.ReadAllBytes(enginePath);
        var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes));

        var manifest = $$"""
            {
              "schemaVersion": {{(damage == "wrong-version" ? 99 : 1)}},
              "engineAssembly": "ExcelBatchTool.Ocr.dll",
              "engineType": "ExcelBatchTool.Ocr.PaddleOcrEngine",
              "files": [
                {
                  "path": "{{(damage == "file-missing" ? "ない.dll" : "ExcelBatchTool.Ocr.dll")}}",
                  "length": {{(damage == "wrong-length" ? bytes.Length + 1 : bytes.Length)}},
                  "sha256": "{{(damage == "wrong-hash" ? new string('0', 64) : hash)}}"
                }
              ]
            }
            """;

        var manifestPath = Path.Combine(directory, "pack.json");
        if (damage == "manifest-missing")
        {
            return;
        }

        File.WriteAllText(manifestPath, damage == "manifest-broken" ? "{ これは JSON ではない" : manifest);
    }
}
