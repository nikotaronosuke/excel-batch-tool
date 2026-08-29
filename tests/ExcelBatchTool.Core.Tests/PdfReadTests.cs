using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelBatchTool.Core.CsvTransform;
using ExcelBatchTool.Core.Pdf;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2F-A。文字情報を持つ PDF(文章・表)から、行 / 列のデータを取り出して
/// Excel / CSV にする。元の PDF は読み取りのみ。すべて架空データ。
/// </summary>
public sealed class PdfReadTests
{
    // ── 自動判定 ─────────────────────────────────────────

    [Fact]
    public void APlainTextPdf_IsDetectedAsText()
    {
        using var dir = new TempDir();
        var pdf = dir.File("報告書.pdf");
        TestPdfFactory.CreateText(pdf, [["月次報告書(架空)", "会社名: 架空商事株式会社", "金額: 1,200 円"]]);

        var preview = Preview(pdf);

        Assert.Equal(PdfDocumentKind.Text, preview.Kind);
        Assert.Equal("通常の文字 PDF", preview.KindDisplay);
        Assert.True(preview.CanExecute);
    }

    [Fact]
    public void ALinedTablePdf_IsDetectedAsTable()
    {
        using var dir = new TempDir();
        var pdf = dir.File("商品表.pdf");
        TestPdfFactory.CreateTable(pdf, [SampleRows(5)]);

        var preview = Preview(pdf);

        Assert.Equal(PdfDocumentKind.Table, preview.Kind);
        Assert.True(preview.TableFromRulings);
        Assert.True(preview.CanExecute);
    }

    // ── 通常の文字 PDF ───────────────────────────────────

    [Fact]
    public void TextIsExtractedWithItsPageAndLineNumbers()
    {
        using var dir = new TempDir();
        var pdf = dir.File("報告書.pdf");
        TestPdfFactory.CreateText(pdf,
        [
            ["月次報告書(架空)", "会社名: 架空商事株式会社", "金額: 1,200 円"],
            ["翌月の予定", "担当者: 架空 太郎"],
        ]);

        var preview = Preview(pdf);

        Assert.Equal(2, preview.PageCount);
        Assert.Equal(5, preview.Lines.Count);
        Assert.Equal(new PdfTextLine(1, 1, "月次報告書(架空)"), preview.Lines[0]);
        Assert.Equal(new PdfTextLine(1, 2, "会社名: 架空商事株式会社"), preview.Lines[1]);
        Assert.Equal(new PdfTextLine(1, 3, "金額: 1,200 円"), preview.Lines[2]);
        Assert.Equal(new PdfTextLine(2, 1, "翌月の予定"), preview.Lines[3]);
        Assert.Equal(new PdfTextLine(2, 2, "担当者: 架空 太郎"), preview.Lines[4]);
    }

    [Fact]
    public void ATwoColumnPageKeepsBothColumnsOnTheSameLine()
    {
        using var dir = new TempDir();
        var pdf = dir.File("二段組.pdf");
        TestPdfFactory.CreateTwoColumnText(pdf,
            ["左の段の 1 行目", "左の段の 2 行目"],
            ["右の段の 1 行目", "右の段の 2 行目"]);

        var preview = Preview(pdf);

        // 同じ高さにある文字は 1 行として、左から右の順に並ぶ。
        Assert.Equal(2, preview.Lines.Count);
        Assert.Contains("左の段の 1 行目", preview.Lines[0].Text, StringComparison.Ordinal);
        Assert.Contains("右の段の 1 行目", preview.Lines[0].Text, StringComparison.Ordinal);
    }

    [Fact]
    public void NfkcNormalizationIsApplied()
    {
        using var dir = new TempDir();
        var pdf = dir.File("正規化.pdf");

        // 康熙部首の ⽉(U+2F49)は、見た目が同じでも 月(U+6708)とは別の文字。
        TestPdfFactory.CreateText(pdf, [["⽉次報告書", "ＡＢＣ 123"]]);

        var preview = Preview(pdf);

        Assert.Equal("月次報告書", preview.Lines[0].Text);
        Assert.DoesNotContain('⽉', preview.Lines[0].Text);

        // NFKC の範囲(全角英数字 → 半角)は行うが、それ以上の変換はしない。
        Assert.Equal("ABC 123", preview.Lines[1].Text);
    }

    [Fact]
    public void NormalizationDoesNotChangeMeaningBeyondNfkc()
    {
        // 大文字小文字・かな・記号は変えない。前後の空白だけ落とす。
        Assert.Equal("Abc-001", PdfTextNormalization.Normalize("  Abc-001  "));
        Assert.Equal("ひらがな", PdfTextNormalization.Normalize("ひらがな"));
        Assert.Equal("カタカナ", PdfTextNormalization.Normalize("ｶﾀｶﾅ"));
        Assert.Equal("0123", PdfTextNormalization.Normalize("0123"));
    }

    // ── 表 PDF ───────────────────────────────────────────

    [Fact]
    public void ALinedTableIsExtractedExactly()
    {
        using var dir = new TempDir();
        var pdf = dir.File("商品表.pdf");
        var rows = SampleRows(20);
        TestPdfFactory.CreateTable(pdf, [rows]);

        var preview = Preview(pdf);

        Assert.Equal(PdfDocumentKind.Table, preview.Kind);
        AssertRowsEqual(rows, preview.TableRows);
    }

    [Fact]
    public void ABorderlessTableIsExtractedExactly()
    {
        using var dir = new TempDir();
        var pdf = dir.File("罫線なし.pdf");
        var rows = SampleRows(20);
        TestPdfFactory.CreateTable(pdf, [rows], lined: false);

        var preview = Preview(pdf);

        Assert.False(preview.TableFromRulings);
        AssertRowsEqual(rows, preview.TableRows);
    }

    [Fact]
    public void AMultiPageTableKeepsEveryRow()
    {
        using var dir = new TempDir();
        var pdf = dir.File("複数ページ.pdf");
        var first = SampleRows(15);
        var second = SampleRows(12, startIndex: 16, includeHeader: false);
        TestPdfFactory.CreateTable(pdf, [first, second]);

        var preview = Preview(pdf);

        Assert.Equal(2, preview.PageCount);
        Assert.Equal(first.Count + second.Count, preview.TableRows.Count);
        Assert.Equal("A0001", preview.TableRows[1][0]);
        Assert.Equal("A0027", preview.TableRows[^1][0]);
    }

    [Fact]
    public void ARepeatedHeaderOnLaterPagesIsJoinedOnce()
    {
        using var dir = new TempDir();
        var pdf = dir.File("見出し繰り返し.pdf");
        var first = SampleRows(10);
        var second = SampleRows(10, startIndex: 11); // 2 ページ目も同じ見出しから始まる
        TestPdfFactory.CreateTable(pdf, [first, second]);

        var preview = Preview(pdf);

        // 見出しは 1 つだけになり、データ行は全部残る。
        Assert.Equal(1, preview.TableRows.Count(row => row[0] == "商品コード"));
        Assert.Equal(21, preview.TableRows.Count);
        Assert.Contains(preview.Warnings, issue => issue.Message.Contains("見出しの繰り返し"));
    }

    [Fact]
    public void ADifferentSecondPageHeaderIsKeptAsData()
    {
        using var dir = new TempDir();
        var pdf = dir.File("別見出し.pdf");
        var first = SampleRows(5);
        var second = new List<string[]>
        {
            new[] { "商品コード", "商品名", "単価", "在庫数" }, // 4 列目だけ違う
            new[] { "A0100", "架空みかん", "500", "9" },
        };

        TestPdfFactory.CreateTable(pdf, [first, second]);

        var preview = Preview(pdf);

        // 完全に同じでなければ、勝手に消さずデータ行として残す。
        Assert.Equal(2, preview.TableRows.Count(row => row[0] == "商品コード"));
    }

    // ── 扱えない PDF ─────────────────────────────────────

    [Fact]
    public void AnImageOnlyPdf_IsBlockedAsNeedingOcr()
    {
        using var dir = new TempDir();
        var pdf = dir.File("スキャン.pdf");
        TestPdfFactory.CreateImageOnly(pdf, pages: 3);

        var preview = Preview(pdf);

        Assert.Equal(PdfDocumentKind.Scan, preview.Kind);
        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("OCR"));
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("スキャン画像"));
    }

    [Fact]
    public void AMixedPdf_IsBlockedInsteadOfPartlyExtracted()
    {
        using var dir = new TempDir();
        var pdf = dir.File("混在.pdf");
        TestPdfFactory.CreateMixed(pdf, textPages: 8, imagePages: 2);

        var preview = Preview(pdf);

        Assert.Equal(PdfDocumentKind.Mixed, preview.Kind);
        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue =>
            issue.Message.Contains("2 ページ") && issue.Message.Contains("OCR"));

        // Phase 2F-B1 で OCR に対応したので、止める理由は「OCR Pack が無い」になった。
        // 一部のページだけを取り出したファイルを作らないことは変えていない。
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("Offline OCR Pack"));

        // 実行しても何も作らない。
        Assert.False(new PdfReader().Execute(preview).Success);
        Assert.Empty(Directory.GetFiles(dir.Root, "混在_*"));
    }

    [Fact]
    public void ACorruptedPdf_IsBlockedWithoutShowingAnException()
    {
        using var dir = new TempDir();
        var pdf = dir.File("壊れた.pdf");
        TestPdfFactory.CreateCorrupted(pdf);

        var preview = Preview(pdf);

        Assert.False(preview.CanExecute);
        var block = Assert.Single(preview.Blocks);
        Assert.Contains("読み取れません", block.Message);

        // 例外の内容(型名・スタックトレース)は画面に出さない。
        Assert.DoesNotContain("Exception", block.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("   at ", block.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void APdfWithNoPages_IsBlocked()
    {
        using var dir = new TempDir();
        var pdf = dir.File("空.pdf");
        TestPdfFactory.CreateEmpty(pdf);

        var preview = Preview(pdf);

        Assert.False(preview.CanExecute);
        Assert.NotEmpty(preview.Blocks);
    }

    [Fact]
    public void AMissingOrNonPdfFile_IsBlocked()
    {
        using var dir = new TempDir();

        var missing = Preview(dir.File("ない.pdf"));
        Assert.Contains(missing.Blocks, issue => issue.Message.Contains("見つかりません"));

        var notPdf = dir.File("表.xlsx");
        File.WriteAllText(notPdf, "これは PDF ではありません");
        Assert.Contains(Preview(notPdf).Blocks, issue => issue.Message.Contains("PDF ファイル"));
    }

    // ── 出力 ─────────────────────────────────────────────

    [Fact]
    public void TextGoesToXlsxAsPageLineText()
    {
        using var dir = new TempDir();
        var pdf = dir.File("報告書.pdf");
        TestPdfFactory.CreateText(pdf, [["月次報告書(架空)", "金額: 1,200 円"]]);

        var (preview, result) = Execute(pdf);

        Assert.True(result.Success, result.Message);
        var rows = ReadXlsx(dir.File("報告書_PDF抽出.xlsx"));

        Assert.Equal(["ページ", "行", "内容"], rows[0]);
        Assert.Equal(["1", "1", "月次報告書(架空)"], rows[1]);
        Assert.Equal(["1", "2", "金額: 1,200 円"], rows[2]);
        Assert.Equal(PdfDocumentKind.Text, preview.Kind);
    }

    [Fact]
    public void ATableGoesToXlsxWithItsColumns()
    {
        using var dir = new TempDir();
        var pdf = dir.File("商品表.pdf");
        var expected = SampleRows(10);
        TestPdfFactory.CreateTable(pdf, [expected]);

        var (_, result) = Execute(pdf);

        Assert.True(result.Success, result.Message);
        var rows = ReadXlsx(dir.File("商品表_PDF抽出.xlsx"));
        AssertRowsEqual(expected, rows.Select(row => row.ToArray()).ToList());
    }

    [Fact]
    public void ATableGoesToCsvThroughTheExistingWriter()
    {
        using var dir = new TempDir();
        var pdf = dir.File("商品表.pdf");
        var expected = SampleRows(5);
        TestPdfFactory.CreateTable(pdf, [expected]);

        var (_, result) = Execute(pdf, format: PdfOutputFormat.Csv);

        Assert.True(result.Success, result.Message);
        var path = dir.File("商品表_PDF抽出.csv");

        // 既定は UTF-8(BOM あり)・CRLF。既存の CSV 出力と同じ決まり。
        Assert.Equal([0xEF, 0xBB, 0xBF], File.ReadAllBytes(path).Take(3));
        var text = File.ReadAllText(path, new UTF8Encoding(false)).TrimStart('﻿');
        Assert.StartsWith("商品コード,商品名,単価,在庫\r\n", text, StringComparison.Ordinal);
    }

    [Fact]
    public void CsvKeepsCommasQuotesAndNewlinesInsideValues()
    {
        using var dir = new TempDir();
        var pdf = dir.File("記号.pdf");
        var rows = new List<string[]>
        {
            new[] { "商品コード", "産地", "備考" },
            new[] { "A0001", "東京,大阪", "12\"の板" },
        };

        TestPdfFactory.CreateTable(pdf, [rows], columnX: [60, 200, 380]);

        var (preview, result) = Execute(pdf, format: PdfOutputFormat.Csv);

        Assert.True(result.Success, result.Message);
        Assert.Equal("東京,大阪", preview.TableRows[1][1]);

        var text = File.ReadAllText(dir.File("記号_PDF抽出.csv"), new UTF8Encoding(false));
        Assert.Contains("\"東京,大阪\"", text, StringComparison.Ordinal);
        Assert.Contains("\"12\"\"の板\"", text, StringComparison.Ordinal);
    }

    [Fact]
    public void CodesThatLookNumericAreKeptAsText()
    {
        using var dir = new TempDir();
        var pdf = dir.File("コード.pdf");
        var rows = new List<string[]>
        {
            new[] { "商品コード", "電話", "郵便番号", "数量" },
            new[] { "0123", "000-1234-5678", "0001234", "42" },
        };

        TestPdfFactory.CreateTable(pdf, [rows]);
        var (_, result) = Execute(pdf);
        Assert.True(result.Success, result.Message);

        using var document = SpreadsheetDocument.Open(dir.File("コード_PDF抽出.xlsx"), false);
        var worksheet = FirstWorksheet(document);
        var cells = worksheet
            .Descendants<Cell>()
            .Where(cell => (cell.CellReference?.Value ?? string.Empty).EndsWith('2'))
            .ToList();

        // 先頭 0・記号入りは文字のまま。純粋な数値だけ数値にする。
        Assert.Equal(CellValues.InlineString, cells[0].DataType?.Value);
        Assert.Equal(CellValues.InlineString, cells[1].DataType?.Value);
        Assert.Equal(CellValues.InlineString, cells[2].DataType?.Value);
        Assert.Null(cells[3].DataType?.Value);
        Assert.Equal("42", cells[3].CellValue?.Text);
    }

    [Theory]
    [InlineData("0123", false)]
    [InlineData("000-1234-5678", false)]
    [InlineData("1,200", false)]
    [InlineData("A001", false)]
    [InlineData("007", false)]
    [InlineData("１２３", false)]
    [InlineData("42", true)]
    [InlineData("-3", true)]
    [InlineData("1.5", true)]
    [InlineData("0", true)]
    public void OnlyValuesThatSurviveARoundTripBecomeNumbers(string value, bool expected)
        => Assert.Equal(expected, PdfWorkbookWriter.IsSafeNumber(value, out _));

    // ── 安全性 ───────────────────────────────────────────

    [Fact]
    public void TheSourcePdfIsNotChanged()
    {
        using var dir = new TempDir();
        var pdf = dir.File("報告書.pdf");
        TestPdfFactory.CreateText(pdf, [["月次報告書(架空)"]]);
        var before = Fingerprint(pdf);

        var (_, result) = Execute(pdf);

        Assert.True(result.Success, result.Message);
        Assert.Equal(before, Fingerprint(pdf));
    }

    [Fact]
    public void APdfChangedAfterThePreview_StopsTheRun()
    {
        using var dir = new TempDir();
        var pdf = dir.File("報告書.pdf");
        TestPdfFactory.CreateText(pdf, [["月次報告書(架空)"]]);
        var preview = Preview(pdf);

        TestPdfFactory.CreateText(pdf, [["差し替えた内容(架空)"]]);

        var result = new PdfReader().Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("変更されています", result.Message);
        Assert.False(File.Exists(dir.File("報告書_PDF抽出.xlsx")));
    }

    [Fact]
    public void AnExistingOutputFile_IsNotOverwritten()
    {
        using var dir = new TempDir();
        var pdf = dir.File("報告書.pdf");
        TestPdfFactory.CreateText(pdf, [["月次報告書(架空)"]]);
        File.WriteAllText(dir.File("報告書_PDF抽出.xlsx"), "先にあった内容");

        var preview = Preview(pdf);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("すでにあります"));
        Assert.Equal("先にあった内容", File.ReadAllText(dir.File("報告書_PDF抽出.xlsx")));
    }

    [Fact]
    public void FilesThatWereAlreadyThere_AreLeftAlone()
    {
        using var dir = new TempDir();
        var pdf = dir.File("報告書.pdf");
        TestPdfFactory.CreateText(pdf, [["月次報告書(架空)"]]);

        // 実行前から置いてある、この実行のものではないファイル。
        var decoys = new[]
        {
            ("報告書_PDF抽出.xlsx.tmp", "前の作業で残った出力"),
            ("報告書_PDF抽出.xlsx.audit.json.tmp", "前の作業で残った控え"),
            ("無関係.tmp", "まったく関係のないファイル"),
        };

        foreach (var (name, content) in decoys)
        {
            File.WriteAllText(dir.File(name), content);
        }

        // 控えの名前をフォルダーで塞ぎ、確定の途中で失敗させる。
        Directory.CreateDirectory(dir.File("報告書_PDF抽出.xlsx.audit.json"));

        var result = new PdfReader().Execute(Preview(pdf));

        Assert.False(result.Success);
        Assert.Contains("取り消しました", result.Message);
        Assert.False(File.Exists(dir.File("報告書_PDF抽出.xlsx")));

        foreach (var (name, content) in decoys)
        {
            Assert.True(File.Exists(dir.File(name)), $"{name} が消えています。");
            Assert.Equal(content, File.ReadAllText(dir.File(name)));
        }
    }

    [Fact]
    public void WhenTheCleanupFails_OnlyThisRunsFilesAreNamed()
    {
        using var dir = new TempDir();
        var pdf = dir.File("報告書.pdf");
        TestPdfFactory.CreateText(pdf, [["月次報告書(架空)"]]);
        File.WriteAllText(dir.File("無関係.tmp"), "まったく関係のないファイル");
        Directory.CreateDirectory(dir.File("報告書_PDF抽出.xlsx.audit.json"));

        var result = new PdfReader { FileDeleter = _ => false }.Execute(Preview(pdf));

        Assert.False(result.Success);
        Assert.Contains("取り消せなかったファイル", result.Message);
        Assert.Contains("報告書_PDF抽出.xlsx", result.Message);
        Assert.DoesNotContain("無関係.tmp", result.Message, StringComparison.Ordinal);
        Assert.Contains("元の PDF は変更していません", result.Message);
    }

    [Fact]
    public void ASuccessfulRunLeavesNoWorkFilesBehind()
    {
        using var dir = new TempDir();
        var pdf = dir.File("報告書.pdf");
        TestPdfFactory.CreateText(pdf, [["月次報告書(架空)"]]);

        var (_, result) = Execute(pdf);

        Assert.True(result.Success, result.Message);
        Assert.Empty(Directory.GetFiles(dir.Root, "*.tmp"));
    }

    // ── 控え ─────────────────────────────────────────────

    [Fact]
    public void TheAuditRecordsWhatWasExtracted()
    {
        using var dir = new TempDir();
        var pdf = dir.File("商品表.pdf");
        TestPdfFactory.CreateTable(pdf, [SampleRows(4)]);

        var (_, result) = Execute(pdf);
        Assert.True(result.Success, result.Message);

        var path = dir.File("商品表_PDF抽出.xlsx.audit.json");
        using var document = System.Text.Json.JsonDocument.Parse(File.ReadAllText(path));
        var root = document.RootElement;

        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("pdf-extract", root.GetProperty("operation").GetString());

        var source = root.GetProperty("source");
        Assert.Equal("商品表.pdf", source.GetProperty("fileName").GetString());
        Assert.NotEmpty(source.GetProperty("sha256").GetString()!);
        Assert.True(source.GetProperty("length").GetInt64() > 0);
        Assert.Equal(1, source.GetProperty("pageCount").GetInt32());
        Assert.Equal("table", source.GetProperty("detectedKind").GetString());
        Assert.Equal("NFKC", source.GetProperty("normalization").GetString());
        Assert.Contains("pdfpig", source.GetProperty("extractionMethod").GetString()!);

        var output = root.GetProperty("output");
        Assert.Equal("商品表_PDF抽出.xlsx", output.GetProperty("fileName").GetString());
        Assert.Equal("xlsx", output.GetProperty("type").GetString());
        Assert.Equal(5, output.GetProperty("rowCount").GetInt32());
        Assert.Equal(4, output.GetProperty("columnCount").GetInt32());

        // 絶対パスも、PDF の本文そのものも残さない。
        var json = File.ReadAllText(path);
        Assert.DoesNotContain(dir.Root, json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(":\\", json, StringComparison.Ordinal);
        Assert.DoesNotContain("架空りんご", json, StringComparison.Ordinal);
    }

    // ── 補助 ─────────────────────────────────────────────

    private static List<string[]> SampleRows(
        int dataRows, int startIndex = 1, bool includeHeader = true)
    {
        var rows = new List<string[]>();
        if (includeHeader)
        {
            rows.Add(new[] { "商品コード", "商品名", "単価", "在庫" });
        }

        string[] names = ["架空りんご", "架空みかん", "架空ぶどう", "架空の緑茶", "架空ノート"];
        for (var index = 0; index < dataRows; index++)
        {
            var number = startIndex + index;
            rows.Add(new[]
            {
                $"A{number:D4}",
                names[index % names.Length],
                (1000 + number * 7).ToString(System.Globalization.CultureInfo.InvariantCulture),
                (number * 3).ToString(System.Globalization.CultureInfo.InvariantCulture),
            });
        }

        return rows;
    }

    private static void AssertRowsEqual(
        IReadOnlyList<string[]> expected, IReadOnlyList<string[]> actual)
    {
        Assert.Equal(expected.Count, actual.Count);
        for (var row = 0; row < expected.Count; row++)
        {
            Assert.Equal(expected[row], actual[row]);
        }
    }

    private static PdfReadPreview Preview(
        string pdfPath, PdfOutputFormat format = PdfOutputFormat.Xlsx)
        => new PdfReadPlanner().CreatePreview(new PdfReadRequest
        {
            SourceFilePath = pdfPath,
            OutputFormat = format,
        });

    private static (PdfReadPreview Preview, PdfReadResult Result) Execute(
        string pdfPath, PdfOutputFormat format = PdfOutputFormat.Xlsx)
    {
        var preview = Preview(pdfPath, format);
        return (preview, new PdfReader().Execute(preview));
    }

    private static Worksheet FirstWorksheet(SpreadsheetDocument document)
        => document.WorkbookPart?.WorksheetParts.FirstOrDefault()?.Worksheet
            ?? throw new InvalidOperationException("Worksheet が見つかりません。");

    private static List<List<string>> ReadXlsx(string path)
    {
        using var document = SpreadsheetDocument.Open(path, isEditable: false);
        var worksheet = FirstWorksheet(document);

        var rows = new List<List<string>>();
        foreach (var row in worksheet.Descendants<Row>())
        {
            var values = new List<string>();
            foreach (var cell in row.Elements<Cell>())
            {
                values.Add(cell.DataType?.Value == CellValues.InlineString
                    ? cell.InlineString?.Text?.Text ?? string.Empty
                    : cell.CellValue?.Text ?? string.Empty);
            }

            rows.Add(values);
        }

        return rows;
    }

    private static (long Length, DateTime Written, string Hash) Fingerprint(string path)
    {
        var info = new FileInfo(path);
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
        return (info.Length, info.LastWriteTimeUtc, hash);
    }
}
