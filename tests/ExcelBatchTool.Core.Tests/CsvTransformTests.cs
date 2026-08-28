using System.Text;
using ExcelBatchTool.Core.CsvTransform;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2E。Excel / CSV の表を、指定した項目名・並び順の新しい CSV にする。
/// 元のファイルは読み取りのみ。すべて架空データ。
/// </summary>
public sealed class CsvTransformTests
{
    // ── データ元が CSV のとき ─────────────────────────────

    [Fact]
    public void Utf8WithBom_IsRead()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品名", "A001,商品A"], withBom: true);

        var lines = Run(dir, source, [Column("コード", "商品コード"), Column("名前", "商品名")]);

        Assert.Equal(["コード,名前", "A001,商品A"], lines);
    }

    [Fact]
    public void Utf8WithoutBom_IsRead()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品名", "A001,商品A"]);

        var lines = Run(dir, source, [Column("コード", "商品コード")]);

        Assert.Equal(["コード", "A001"], lines);
    }

    [Fact]
    public void ShiftJisSource_IsRead()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品名", "A001,架空の商品"], encodingName: "shift_jis");

        var lines = Run(dir, source, [Column("名前", "商品名")]);

        Assert.Equal(["名前", "架空の商品"], lines);
    }

    [Fact]
    public void FieldWithACommaKeepsItsValue()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,産地", "A001,\"東京,大阪\""]);

        var lines = Run(dir, source, [Column("産地", "産地")]);

        Assert.Equal(["産地", "\"東京,大阪\""], lines);
    }

    [Fact]
    public void FieldWithAQuoteKeepsItsValue()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,名前", "A001,\"12\"\"の板\""]);

        var lines = Run(dir, source, [Column("名前", "名前")]);

        Assert.Equal(["名前", "\"12\"\"の板\""], lines);
    }

    [Fact]
    public void FieldWithANewlineKeepsItsValue()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,備考", "A001,\"1 行目", "2 行目\""]);

        var (path, preview, result) = Execute(
            dir, source, [Column("備考", "備考")]);

        Assert.True(result.Success);
        Assert.Equal(1, preview.OutputRowCount);

        // 引用符の中の改行はそのまま残り、読み直すと 1 項目として戻る。
        var text = File.ReadAllText(path, new UTF8Encoding(false));
        Assert.Contains("\"1 行目\r\n2 行目\"", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowsWhereEveryFieldIsBlank_AreSkipped()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品名", "A001,商品A", ",", "A002,商品B"]);

        var (_, preview, _) = Execute(dir, source, [Column("コード", "商品コード")]);

        Assert.Equal(2, preview.OutputRowCount);
        Assert.Equal(1, preview.BlankRowCount);
        Assert.Contains(preview.Warnings, issue => issue.Message.Contains("読み飛ばし"));
    }

    [Fact]
    public void PartlyBlankRows_AreKept()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品名", "A001,", "A002,商品B"]);

        var lines = Run(dir, source, [Column("コード", "商品コード"), Column("名前", "商品名")]);

        Assert.Equal(["コード,名前", "A001,", "A002,商品B"], lines);
    }

    [Fact]
    public void JapaneseHeadersAndValues_SurviveTheRoundTrip()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品名,備考", "A001,架空の商品,確認済み"]);

        var lines = Run(dir, source,
            [Column("商品番号", "商品コード"), Column("名称", "商品名"), Column("メモ", "備考")]);

        Assert.Equal(["商品番号,名称,メモ", "A001,架空の商品,確認済み"], lines);
    }

    // ── データ元が .xlsx のとき ───────────────────────────

    [Fact]
    public void SharedStrings_AreRead()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "商品",
            [["商品コード", "商品名"], ["A001", "共有文字列"]]);

        var lines = Run(dir, source, [Column("コード", "商品コード"), Column("名前", "商品名")],
            sheetName: "商品");

        Assert.Equal(["コード,名前", "A001,共有文字列"], lines);
    }

    [Fact]
    public void InlineStrings_AreRead()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        SourceTestCell[][] rows =
        [
            ["商品コード", "商品名"],
            ["A001", new SourceTestCell(new InlineSourceText("直接書かれた文字列"))],
        ];

        TestSourceTableFactory.CreateXlsx(source, "商品", rows);

        var lines = Run(dir, source, [Column("名前", "商品名")], sheetName: "商品");

        Assert.Equal(["名前", "直接書かれた文字列"], lines);
    }

    [Fact]
    public void ABooleanCell_StopsTheWholeTransform()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        SourceTestCell[][] rows =
        [
            ["商品コード", "公開"],
            ["A001", new SourceTestCell(true)],
        ];

        TestSourceTableFactory.CreateXlsx(source, "商品", rows);

        // TRUE / FALSE を「1」と書くか「TRUE」と書くかは推測になるので、止めて選び直してもらう。
        var preview = Preview(dir, source, [Column("公開", "公開")], sheetName: "商品");

        Assert.True(preview.HasBlocks);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("TRUE / FALSE"));
    }

    [Fact]
    public void PlainNumbers_AreWrittenAsTheyAre()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "商品",
            [["商品コード", "価格", "率"], ["A001", 1200, 0.5]]);

        var lines = Run(dir, source, [Column("価格", "価格"), Column("率", "率")], sheetName: "商品");

        Assert.Equal(["価格,率", "1200,0.5"], lines);
    }

    [Fact]
    public void BlankCells_BecomeEmptyFields()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        SourceTestCell[][] rows =
        [
            ["商品コード", "商品名"],
            ["A001", new SourceTestCell(null)],
        ];

        TestSourceTableFactory.CreateXlsx(source, "商品", rows);

        var lines = Run(dir, source, [Column("コード", "商品コード"), Column("名前", "商品名")],
            sheetName: "商品");

        Assert.Equal(["コード,名前", "A001,"], lines);
    }

    [Fact]
    public void AFormulaCell_StopsTheWholeTransform()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        SourceTestCell[][] rows =
        [
            ["商品コード", "価格"],
            ["A001", new SourceTestCell(new SourceTestFormula("1+1", "2"))],
        ];

        TestSourceTableFactory.CreateXlsx(source, "商品", rows);

        var preview = Preview(dir, source, [Column("価格", "価格")], sheetName: "商品");

        Assert.True(preview.HasBlocks);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("数式"));
        Assert.False(preview.CanExecute);
    }

    [Fact]
    public void ADateCell_StopsTheWholeTransform()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "商品",
            [["商品コード", "日付"], ["A001", new SourceTestCell(45000.0, StyleId: 1)]],
            styles: [new MutationTestStyle(14)]); // 組み込みの日付書式。

        var preview = Preview(dir, source, [Column("日付", "日付")], sheetName: "商品");

        Assert.True(preview.HasBlocks);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("日付"));
    }

    [Fact]
    public void APercentageCell_StopsTheWholeTransform()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "商品",
            [["商品コード", "率"], ["A001", new SourceTestCell(0.25, StyleId: 1)]],
            styles: [new MutationTestStyle(10)]); // 組み込みのパーセント書式。

        var preview = Preview(dir, source, [Column("率", "率")], sheetName: "商品");

        Assert.True(preview.HasBlocks);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("表示形式"));
    }

    // ── 列の作り方 ───────────────────────────────────────

    [Fact]
    public void ColumnsComeOutInTheOrderThatWasSpecified()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品名,価格", "A001,商品A,1200"]);

        var lines = Run(dir, source,
            [Column("価格", "価格"), Column("コード", "商品コード"), Column("名前", "商品名")]);

        Assert.Equal(["価格,コード,名前", "1200,A001,商品A"], lines);
    }

    [Fact]
    public void ColumnsCanBeRenamed()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,価格", "A001,1200"]);

        var lines = Run(dir, source, [Column("商品番号", "商品コード"), Column("販売価格", "価格")]);

        Assert.Equal(["商品番号,販売価格", "A001,1200"], lines);
    }

    [Fact]
    public void UnusedSourceColumnsAreLeftOut()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品名,内部メモ", "A001,商品A,確認済"]);

        var lines = Run(dir, source, [Column("コード", "商品コード"), Column("名前", "商品名")]);

        Assert.Equal(["コード,名前", "A001,商品A"], lines);
        Assert.DoesNotContain("内部メモ", lines[0], StringComparison.Ordinal);
        Assert.DoesNotContain("確認済", lines[1], StringComparison.Ordinal);
    }

    [Fact]
    public void AFixedValueGoesIntoEveryRow()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード", "A001", "A002"]);

        var lines = Run(dir, source, [Column("コード", "商品コード"), Fixed("公開状態", "1")]);

        Assert.Equal(["コード,公開状態", "A001,1", "A002,1"], lines);
    }

    [Fact]
    public void ABlankColumnIsEmptyInEveryRow()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード", "A001", "A002"]);

        var lines = Run(dir, source, [Column("コード", "商品コード"), BlankColumn("予備")]);

        Assert.Equal(["コード,予備", "A001,", "A002,"], lines);
    }

    [Fact]
    public void TheSameSourceColumnCanFeedSeveralOutputColumns()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品名", "A001,商品A"]);

        var lines = Run(dir, source,
            [Column("商品名", "商品名"), Column("検索用商品名", "商品名")]);

        Assert.Equal(["商品名,検索用商品名", "商品A,商品A"], lines);
    }

    [Fact]
    public void AMissingSourceColumn_Stops()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード", "A001"]);

        var preview = Preview(dir, source, [Column("価格", "存在しない項目")]);

        Assert.True(preview.HasBlocks);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("存在しない項目"));
    }

    [Fact]
    public void DuplicateOutputNames_Stop()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品名", "A001,商品A"]);

        var preview = Preview(dir, source,
            [Column("Price", "商品コード"), Column("price", "商品名")]);

        Assert.True(preview.HasBlocks);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("重複"));
    }

    [Fact]
    public void ABlankOutputName_Stops()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード", "A001"]);

        Assert.Contains(
            Preview(dir, source, [Column("  ", "商品コード")]).Blocks,
            issue => issue.Message.Contains("項目名が空"));

        Assert.Contains(
            Preview(dir, source, [Column("改行\n入り", "商品コード")]).Blocks,
            issue => issue.Message.Contains("特殊な文字"));
    }

    [Fact]
    public void NoOutputColumns_Stops()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード", "A001"]);

        var preview = Preview(dir, source, []);

        Assert.True(preview.HasBlocks);
        Assert.False(preview.CanExecute);
    }

    // ── 書き出し方 ───────────────────────────────────────

    [Fact]
    public void MinimalQuoting_OnlyQuotesWhatItMust()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["名前,産地", "商品A,\"東京,大阪\""]);

        var lines = Run(dir, source, [Column("名前", "名前"), Column("産地", "産地")]);

        Assert.Equal(["名前,産地", "商品A,\"東京,大阪\""], lines);
    }

    [Fact]
    public void QuoteAll_QuotesEveryField()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["名前,産地", "商品A,\"東京,大阪\""]);

        var lines = Run(dir, source, [Column("名前", "名前"), Column("産地", "産地")],
            quoteMode: CsvQuoteMode.All);

        Assert.Equal(["\"名前\",\"産地\"", "\"商品A\",\"東京,大阪\""], lines);
    }

    [Fact]
    public void TheOutputEncodingIsWhatWasChosen()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["名前", "架空の商品"]);

        var bom = Execute(dir, source, [Column("名前", "名前")],
            encoding: CsvOutputEncoding.Utf8Bom, suffix: "_bom");
        Assert.True(bom.Result.Success, bom.Result.Message);
        Assert.Equal([0xEF, 0xBB, 0xBF], File.ReadAllBytes(bom.Path).Take(3));

        var plain = Execute(dir, source, [Column("名前", "名前")],
            encoding: CsvOutputEncoding.Utf8, suffix: "_utf8");
        Assert.True(plain.Result.Success, plain.Result.Message);
        Assert.NotEqual<byte>(0xEF, File.ReadAllBytes(plain.Path)[0]);
        Assert.Contains("架空の商品",
            File.ReadAllText(plain.Path, new UTF8Encoding(false)), StringComparison.Ordinal);

        var sjis = Execute(dir, source, [Column("名前", "名前")],
            encoding: CsvOutputEncoding.ShiftJis, suffix: "_sjis");
        Assert.True(sjis.Result.Success, sjis.Result.Message);
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Assert.Contains("架空の商品",
            File.ReadAllText(sjis.Path, Encoding.GetEncoding(932)), StringComparison.Ordinal);

        // Shift_JIS は UTF-8 として読めないバイトになる(取り違えていないことの裏付け)。
        Assert.Throws<DecoderFallbackException>(
            () => new UTF8Encoding(false, throwOnInvalidBytes: true)
                .GetString(File.ReadAllBytes(sjis.Path)));
    }

    [Fact]
    public void EveryLineEndsWithCrLf()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["名前", "商品A", "商品B"], newLine: "\n");

        var (path, _, _) = Execute(dir, source, [Column("名前", "名前")]);

        var text = File.ReadAllText(path, new UTF8Encoding(false));
        Assert.Equal("名前\r\n商品A\r\n商品B\r\n", text.TrimStart('﻿'));
    }

    // ── 出力の安全性 ─────────────────────────────────────

    [Fact]
    public void TheSourceFileIsNotChanged()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品名", "A001,商品A"]);
        var before = Fingerprint(source);

        var (_, _, result) = Execute(dir, source, [Column("コード", "商品コード")]);

        Assert.True(result.Success);
        Assert.Equal(before, Fingerprint(source));
    }

    [Fact]
    public void AChangedSourceAfterThePreview_StopsTheRun()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード", "A001"]);
        var preview = Preview(dir, source, [Column("コード", "商品コード")]);

        File.WriteAllText(source, "商品コード\r\nB999\r\n", new UTF8Encoding(false));

        var result = new CsvTransformer().Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("変更されています", result.Message);
        Assert.False(File.Exists(dir.File("元データ_変換済み.csv")));
    }

    [Fact]
    public void AnExistingOutputFile_IsNotOverwritten()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード", "A001"]);
        File.WriteAllText(dir.File("元データ_変換済み.csv"), "先にあった内容");

        var preview = Preview(dir, source, [Column("コード", "商品コード")]);

        Assert.True(preview.HasBlocks);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("すでにあります"));
        Assert.Equal("先にあった内容", File.ReadAllText(dir.File("元データ_変換済み.csv")));
    }

    [Fact]
    public void AFailedRun_LeavesNoFilesBehind()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード", "A001"]);
        var preview = Preview(dir, source, [Column("コード", "商品コード")]);

        // 控えファイルの置き場所をフォルダーで塞ぎ、確定の途中で失敗させる。
        Directory.CreateDirectory(dir.File("元データ_変換済み.csv.audit.json"));

        var result = new CsvTransformer().Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("取り消しました", result.Message);
        Assert.False(File.Exists(dir.File("元データ_変換済み.csv")));
        Assert.False(File.Exists(dir.File("元データ_変換済み.csv.tmp")));
    }

    [Fact]
    public void WhenTheCleanupFails_TheRemainingFileIsNamed()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード", "A001"]);
        var preview = Preview(dir, source, [Column("コード", "商品コード")]);

        Directory.CreateDirectory(dir.File("元データ_変換済み.csv.audit.json"));

        // 消せない状況を作る。ファイルが残ったことを断定せずに知らせる。
        var transformer = new CsvTransformer { FileDeleter = _ => false };
        var result = transformer.Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("取り消せなかったファイル", result.Message);
        Assert.Contains("元データ_変換済み.csv", result.Message);
        Assert.Contains("元のファイルは変更していません", result.Message);

        // 元のデータ元は読み取りだけなので、そこは断定してよい。
        Assert.DoesNotContain("作成していません", result.Message);
    }

    [Fact]
    public void AMalformedSourceRow_Stops()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品名", "A001,商品A,余分"]);

        var preview = Preview(dir, source, [Column("コード", "商品コード")]);

        Assert.True(preview.HasBlocks);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("列数"));
    }

    [Fact]
    public void ADuplicateSourceHeader_Stops()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,商品コード", "A001,A002"]);

        var preview = Preview(dir, source, [Column("コード", "商品コード")]);

        Assert.True(preview.HasBlocks);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("重複"));
    }

    [Fact]
    public void TheOutputIsCheckedBeforeItIsKept()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,産地,備考", "A001,\"東京,大阪\",\"改行", "あり\""]);

        var (path, preview, result) = Execute(dir, source,
            [Column("コード", "商品コード"), Column("産地", "産地"), Column("備考", "備考")]);

        Assert.True(result.Success, result.Message);
        Assert.Equal(1, result.RowCount);

        // 読み直したとき、行数・列数・各項目が指定どおりであることを確かめている。
        var rows = ReadBack(path);
        Assert.Equal(2, rows.Count);
        Assert.Equal(["コード", "産地", "備考"], rows[0]);
        Assert.Equal(["A001", "東京,大阪", "改行\r\nあり"], rows[1]);
        Assert.Equal(3, preview.Columns.Count);
    }

    // ── 控え ─────────────────────────────────────────────

    [Fact]
    public void TheAuditRecordsWhatWasConverted()
    {
        using var dir = new TempDir();
        var source = Csv(dir, ["商品コード,価格", "A001,1200"]);

        var (_, _, result) = Execute(dir, source,
            [Column("商品番号", "商品コード"), Fixed("公開状態", "1")],
            quoteMode: CsvQuoteMode.All);

        Assert.True(result.Success, result.Message);

        using var document = System.Text.Json.JsonDocument.Parse(
            File.ReadAllText(dir.File("元データ_変換済み.csv.audit.json")));
        var root = document.RootElement;

        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("csv-transform", root.GetProperty("operation").GetString());

        var sourceInfo = root.GetProperty("source");
        Assert.Equal("元データ.csv", sourceInfo.GetProperty("fileName").GetString());
        Assert.Equal("csv", sourceInfo.GetProperty("type").GetString());
        Assert.NotEmpty(sourceInfo.GetProperty("sha256").GetString()!);

        var output = root.GetProperty("output");
        Assert.Equal("元データ_変換済み.csv", output.GetProperty("fileName").GetString());
        Assert.Equal("utf-8-bom", output.GetProperty("encoding").GetString());
        Assert.Equal("all", output.GetProperty("quoteMode").GetString());
        Assert.Equal("crlf", output.GetProperty("lineEnding").GetString());
        Assert.Equal(1, output.GetProperty("rowCount").GetInt32());
        Assert.Equal(2, output.GetProperty("columnCount").GetInt32());

        var columns = root.GetProperty("columns");
        Assert.Equal("商品番号", columns[0].GetProperty("outputName").GetString());
        Assert.Equal("source-column", columns[0].GetProperty("sourceKind").GetString());
        Assert.Equal("商品コード", columns[0].GetProperty("sourceColumn").GetString());
        Assert.Equal("fixed-text", columns[1].GetProperty("sourceKind").GetString());
        Assert.Equal("1", columns[1].GetProperty("fixedValue").GetString());

        // 絶対パスは残さない。
        var json = File.ReadAllText(dir.File("元データ_変換済み.csv.audit.json"));
        Assert.DoesNotContain(dir.Root, json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(":\\", json, StringComparison.Ordinal);
    }

    [Fact]
    public void ThePreviewShowsTheFirstRowsOnly()
    {
        using var dir = new TempDir();
        var lines = new List<string> { "商品コード" };
        for (var index = 1; index <= 100; index++)
        {
            lines.Add($"A{index:D4}");
        }

        var source = Csv(dir, lines);
        var preview = Preview(dir, source, [Column("コード", "商品コード")]);

        Assert.Equal(100, preview.OutputRowCount);
        Assert.Equal(CsvTransformDefaults.SampleRowCount, preview.SampleRows.Count);
        Assert.Equal("A0001", preview.SampleRows[0].Values[0]);
    }

    // ── 補助 ─────────────────────────────────────────────

    private static CsvOutputColumnRequest Column(string outputName, string sourceColumn) => new()
    {
        OutputName = outputName,
        ValueSourceKind = CsvValueSourceKind.SourceColumn,
        SourceColumn = sourceColumn,
    };

    private static CsvOutputColumnRequest Fixed(string outputName, string value) => new()
    {
        OutputName = outputName,
        ValueSourceKind = CsvValueSourceKind.FixedText,
        FixedValue = value,
    };

    private static CsvOutputColumnRequest BlankColumn(string outputName) => new()
    {
        OutputName = outputName,
        ValueSourceKind = CsvValueSourceKind.Blank,
    };

    private static string Csv(
        TempDir dir,
        IReadOnlyList<string> lines,
        string encodingName = "utf-8",
        bool withBom = false,
        string newLine = "\r\n")
    {
        var path = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(path, lines, encodingName, withBom, newLine);
        return path;
    }

    private static CsvTransformPreview Preview(
        TempDir dir,
        string sourceFilePath,
        IReadOnlyList<CsvOutputColumnRequest> columns,
        string? sheetName = null,
        CsvOutputEncoding encoding = CsvOutputEncoding.Utf8Bom,
        CsvQuoteMode quoteMode = CsvQuoteMode.Minimal,
        string suffix = CsvTransformDefaults.OutputSuffix)
        => new CsvTransformPlanner().CreatePreview(new CsvTransformRequest
        {
            SourceFilePath = sourceFilePath,
            SourceSheetName = sheetName,
            Columns = columns,
            Encoding = encoding,
            QuoteMode = quoteMode,
            OutputSuffix = suffix,
        });

    private static (string Path, CsvTransformPreview Preview, CsvTransformResult Result) Execute(
        TempDir dir,
        string sourceFilePath,
        IReadOnlyList<CsvOutputColumnRequest> columns,
        string? sheetName = null,
        CsvOutputEncoding encoding = CsvOutputEncoding.Utf8Bom,
        CsvQuoteMode quoteMode = CsvQuoteMode.Minimal,
        string suffix = CsvTransformDefaults.OutputSuffix)
    {
        var preview = Preview(dir, sourceFilePath, columns, sheetName, encoding, quoteMode, suffix);
        var result = new CsvTransformer().Execute(preview);
        return (dir.File(preview.OutputFileName), preview, result);
    }

    /// <summary>作った CSV を行ごとの文字列として読む(BOM は落とす)。</summary>
    private static IReadOnlyList<string> Run(
        TempDir dir,
        string sourceFilePath,
        IReadOnlyList<CsvOutputColumnRequest> columns,
        string? sheetName = null,
        CsvOutputEncoding encoding = CsvOutputEncoding.Utf8Bom,
        CsvQuoteMode quoteMode = CsvQuoteMode.Minimal)
    {
        var (path, _, result) = Execute(
            dir, sourceFilePath, columns, sheetName, encoding, quoteMode);

        Assert.True(result.Success, result.Message);

        return File.ReadAllText(path, new UTF8Encoding(false))
            .TrimStart('﻿')
            .Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
    }

    /// <summary>作った CSV を項目ごとに読み直す。</summary>
    private static IReadOnlyList<string[]> ReadBack(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var reader = new StreamReader(stream, new UTF8Encoding(false), true);
        using var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(reader)
        {
            TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited,
            HasFieldsEnclosedInQuotes = true,
            TrimWhiteSpace = false,
        };

        parser.SetDelimiters(",");

        var rows = new List<string[]>();
        while (!parser.EndOfData)
        {
            if (parser.ReadFields() is { } fields)
            {
                rows.Add(fields);
            }
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
