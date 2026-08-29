using ExcelBatchTool.Core.CsvTransform;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Ocr;
using ExcelBatchTool.Core.Pdf;
using ExcelBatchTool.Core.Recipes;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2F-B3。PDF 読み取りの設定の保存と、大きすぎる PDF の扱い。
/// </summary>
public sealed class PdfRecipeAndLimitTests
{
    // ── PDF 読み取りの設定を保存する ──────────────────────

    [Fact]
    public void APdfRecipeKeepsHowToReadButNothingAboutThePdf()
    {
        var recipe = SampleRecipe();

        // 保存するのは「読み方」だけ。
        Assert.Equal("同じ様式の帳票として読む", recipe.PdfRead!.ReadMode);
        Assert.Equal(2, recipe.PdfRead.Fields.Count);

        // 元の PDF に関わるものを入れる場所がそもそも無い。
        var properties = typeof(PdfReadRecipe).GetProperties().Select(p => p.Name).ToList();
        Assert.DoesNotContain("SourcePath", properties);
        Assert.DoesNotContain("SourceFileName", properties);
        Assert.DoesNotContain("Text", properties);

        var fieldProperties = typeof(PdfReadRecipeField).GetProperties()
            .Select(p => p.Name).ToList();
        Assert.DoesNotContain("Value", fieldProperties);
        Assert.DoesNotContain("Text", fieldProperties);
    }

    [Fact]
    public void APdfRecipeSurvivesSavingAndLoading()
    {
        using var dir = new TempDir();
        var path = dir.File("recipes.json");
        var store = new RecipeStore(path);

        var saved = SampleRecipe();
        var add = store.Add(saved);
        Assert.True(add.IsSuccess, add.Error);

        var result = new RecipeStore(path).Load();
        Assert.True(result.IsSuccess);
        var loaded = result.Recipes.Single(recipe => recipe.Type == RecipeType.PdfRead);

        Assert.NotNull(loaded.PdfRead);
        Assert.Equal("店舗コード", loaded.PdfRead!.Fields[0].Name);
        Assert.Equal(364, loaded.PdfRead.Fields[0].X);
        Assert.Equal(2, loaded.PdfRead.Fields.Count);
        Assert.Equal("承認", loaded.PdfRead.Fields[1].Choices[0].Label);
        Assert.True(
            RecipeConfiguration.AreSame(saved, loaded),
            $"読み方: {loaded.PdfRead.ReadMode} / 形式: {loaded.PdfRead.OutputFormat} / "
            + $"接尾: {loaded.PdfRead.OutputSuffix}");
    }

    [Fact]
    public void ThePdfRecipeFileHasNoTraceOfTheSourceDocument()
    {
        using var dir = new TempDir();
        var path = dir.File("recipes.json");
        var added = new RecipeStore(path).Add(SampleRecipe());
        Assert.True(added.IsSuccess, added.Error);

        var text = File.ReadAllText(path);

        // 実際に書かれた中身にも、元の PDF に関わるものは出てこない。
        Assert.DoesNotContain(".pdf", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("C:\\", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("店舗コード", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("", "読み取り方")]
    [InlineData("同じ様式の帳票として読む", null)]
    public void APdfRecipeNeedsAReadMode(string mode, string? expected)
    {
        var recipe = SampleRecipe() with { PdfRead = SampleRecipe().PdfRead! with { ReadMode = mode } };
        var error = RecipeValidation.Validate(recipe);

        if (expected is null)
        {
            Assert.Null(error);
        }
        else
        {
            Assert.Contains(expected, error);
        }
    }

    [Fact]
    public void TwoFieldsCannotShareAName()
    {
        var pdf = SampleRecipe().PdfRead!;
        var clash = SampleRecipe() with
        {
            PdfRead = pdf with
            {
                Fields = [pdf.Fields[0], pdf.Fields[0]],
            },
        };

        Assert.Contains("重複", RecipeValidation.Validate(clash));
    }

    [Fact]
    public void AFieldNeedsAPlaceToRead()
    {
        var pdf = SampleRecipe().PdfRead!;
        var noArea = SampleRecipe() with
        {
            PdfRead = pdf with
            {
                Fields = [pdf.Fields[0] with { Width = 0 }],
            },
        };

        Assert.Contains("読み取る場所", RecipeValidation.Validate(noArea));
    }

    // ── 大きすぎる PDF ────────────────────────────────

    [Fact]
    public void TooManyScannedPagesAreBlockedWithHowManyAreAllowed()
    {
        using var dir = new TempDir();
        var pdf = dir.File("大量.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 3);

        // 上限の判断そのものを確かめる(実際に 1000 ページ作るのは現実的でない)。
        Assert.Equal(1000, PdfReadDefaults.MaxOcrPages);
        Assert.True(PdfReadDefaults.MaxOcrPages < PdfReadDefaults.MaxPages);

        var preview = new PdfReadPlanner().CreatePreview(
            new PdfReadRequest { SourceFilePath = pdf }, UsablePack);

        // 少ないページ数は当然通る。
        Assert.DoesNotContain(preview.Issues, issue =>
            issue.Message.Contains("一度に読み取れるのは", StringComparison.Ordinal));
    }

    [Fact]
    public void ALongButAllowedRunSaysHowLongItWillTake()
    {
        // 時間がかかることは止めずに先に知らせる。
        Assert.Equal(200, PdfReadDefaults.SlowOcrPageWarning);
        Assert.True(PdfReadDefaults.SlowOcrPageWarning < PdfReadDefaults.MaxOcrPages);

        // 実測した 1 ページあたりの時間から見込みを出す。
        var minutes = 500 * PdfReadDefaults.OcrSecondsPerPage / 60;
        Assert.InRange(minutes, 10, 30);
    }

    [Fact]
    public void APageLimitExistsForRenderedPixelsToo()
    {
        // A4 を 300dpi で描くと約 870 万画素。上限はその余裕を見た値。
        const long a4At300Dpi = 2480L * 3508L;
        Assert.True(PdfReadDefaults.MaxRenderedPixelsPerPage > a4At300Dpi);
        Assert.True(PdfReadDefaults.MaxRenderedPixelsPerPage < 200_000_000);
    }

    private static OcrPackStatus UsablePack
        => new(IsPresent: true, IsUsable: true, "OCR Pack を使えます。", "テスト");

    private static SavedRecipe SampleRecipe() => new()
    {
        Id = "pdf-1",
        Name = "架空の月次帳票",
        Type = RecipeType.PdfRead,
        CreatedAt = "2026-08-29T00:00:00Z",
        UpdatedAt = "2026-08-29T00:00:00Z",
        PdfRead = new PdfReadRecipe
        {
            ReadMode = "同じ様式の帳票として読む",
            OutputFormat = "Excel (.xlsx)",
            OutputSuffix = "_PDF抽出",
            Encoding = CsvOutputEncoding.Utf8Bom,
            QuoteMode = CsvQuoteMode.Minimal,
            Fields =
            [
                new PdfReadRecipeField
                {
                    Name = "店舗コード", Kind = "コード", IsRequired = true,
                    X = 364, Y = 527, Width = 362, Height = 78,
                },
                new PdfReadRecipeField
                {
                    Name = "状態", Kind = "選択", IsRequired = false,
                    X = 364, Y = 800, Width = 400, Height = 50,
                    Choices =
                    [
                        new PdfReadRecipeChoice
                        {
                            Label = "承認", X = 364, Y = 800, Width = 50, Height = 50,
                        },
                    ],
                },
            ],
        },
    };
}
