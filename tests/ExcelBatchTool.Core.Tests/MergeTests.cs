using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Tests;

/// <summary>Phase 1A(表データの縦結合)のテスト。すべて架空データで生成した Workbook を使う。</summary>
public sealed class MergeTests
{
    private const string OutputSheetName = "統合結果";

    private static readonly MergeOptions DefaultOptions = new();

    private static readonly MergeOptions NoMetadataOptions = new()
    {
        IncludeSourceFileColumn = false,
        IncludeSourceSheetColumn = false,
    };

    [Fact]
    public void Merge_SameHeadersAcrossWorkbooks_ProducesVerticalUnion()
    {
        using var dir = new TempDir();
        var osaka = dir.File("大阪.xlsx");
        var kyoto = dir.File("京都.xlsx");
        TestTableWorkbookFactory.CreateTable(osaka, "売上", ["商品", "売上"],
            [["A", 100], ["B", 200]]);
        TestTableWorkbookFactory.CreateTable(kyoto, "売上", ["商品", "売上"],
            [["C", 300], ["D", 400]]);

        var preview = CreatePreview(DefaultOptions, (osaka, "売上"), (kyoto, "売上"));

        Assert.True(preview.CanExecute);
        Assert.Equal(new[] { "商品", "売上" }, preview.DataHeaders);
        Assert.Equal(new[] { "元ファイル", "元シート", "商品", "売上" }, preview.OutputHeaders);
        Assert.Equal(2, preview.WorkbookCount);
        Assert.Equal(2, preview.SheetCount);
        Assert.Equal(4, preview.InputDataRowCount);
        Assert.Equal(5, preview.OutputRowCount);

        var output = dir.File("統合結果.xlsx");
        var result = new TableMerger().Execute(preview, DefaultOptions, output);

        Assert.True(result.Success, result.Message);
        Assert.Equal(4, result.DataRowCount);

        Assert.Equal(new[] { "元ファイル", "元シート", "商品", "売上" }, ReadHeaders(output));
        Assert.Equal(
            new[]
            {
                new[] { "大阪.xlsx", "売上", "A", "100" },
                ["大阪.xlsx", "売上", "B", "200"],
                ["京都.xlsx", "売上", "C", "300"],
                ["京都.xlsx", "売上", "D", "400"],
            },
            ReadRowTexts(output));
    }

    [Fact]
    public void Merge_DifferentHeaderOrder_AlignsByHeaderName()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        TestTableWorkbookFactory.CreateTable(a, "表", ["氏名", "電話", "金額"],
            [["架空太郎", "000-0000", 1000]]);
        // 列順が違うだけ。ヘッダー名で対応付けできる。
        TestTableWorkbookFactory.CreateTable(b, "表", ["金額", "氏名", "電話"],
            [[2000, "架空花子", "111-1111"]]);

        var preview = CreatePreview(NoMetadataOptions, (a, "表"), (b, "表"));
        Assert.True(preview.CanExecute);
        Assert.Equal(new[] { "氏名", "電話", "金額" }, preview.OutputHeaders);

        var output = dir.File("out.xlsx");
        Assert.True(new TableMerger().Execute(preview, NoMetadataOptions, output).Success);

        Assert.Equal(
            new[]
            {
                new[] { "架空太郎", "000-0000", "1000" },
                ["架空花子", "111-1111", "2000"],
            },
            ReadRowTexts(output));
    }

    [Fact]
    public void Preview_MissingHeader_IsBlockedWithMissingNames()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        TestTableWorkbookFactory.CreateTable(a, "表", ["氏名", "電話", "金額"], [["架空太郎", "000", 1]]);
        TestTableWorkbookFactory.CreateTable(b, "表", ["氏名", "金額"], [["架空花子", 2]]);

        var preview = CreatePreview(DefaultOptions, (a, "表"), (b, "表"));

        Assert.False(preview.CanExecute);
        var block = Assert.Single(preview.Blocks, issue => issue.FileName == "B.xlsx");
        Assert.Contains("不足ヘッダー", block.Message);
        Assert.Contains("電話", block.Message);
    }

    [Fact]
    public void Preview_ExtraHeader_IsBlockedWithExtraNames()
    {
        using var dir = new TempDir();
        var a = dir.File("A.xlsx");
        var b = dir.File("B.xlsx");
        TestTableWorkbookFactory.CreateTable(a, "表", ["氏名", "金額"], [["架空太郎", 1]]);
        TestTableWorkbookFactory.CreateTable(b, "表", ["氏名", "金額", "備考"], [["架空花子", 2, "なし"]]);

        var preview = CreatePreview(DefaultOptions, (a, "表"), (b, "表"));

        Assert.False(preview.CanExecute);
        var block = Assert.Single(preview.Blocks, issue => issue.FileName == "B.xlsx");
        Assert.Contains("余分なヘッダー", block.Message);
        Assert.Contains("備考", block.Message);
    }

    [Fact]
    public void Preview_EmptyHeaderCell_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestTableWorkbookFactory.CreateTable(path, "表", ["商品", "", "数量"], [["A", null, 1]]);

        var preview = CreatePreview(DefaultOptions, (path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("空のヘッダー"));
    }

    [Fact]
    public void Preview_DuplicateHeader_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestTableWorkbookFactory.CreateTable(path, "表", ["商品", "商品"], [["A", "B"]]);

        var preview = CreatePreview(DefaultOptions, (path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("ヘッダー「商品」"));
    }

    [Fact]
    public void Merge_BlankRowInMiddle_IsSkippedAndLaterRowsKept()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestTableWorkbookFactory.CreateTable(path, "表", ["商品", "売上"],
            [["A", 100], [], ["B", 200], [], [], ["C", 300]]);

        var preview = CreatePreview(NoMetadataOptions, (path, "表"));
        Assert.True(preview.CanExecute);
        Assert.Equal(3, preview.InputDataRowCount);

        var output = dir.File("out.xlsx");
        Assert.True(new TableMerger().Execute(preview, NoMetadataOptions, output).Success);

        Assert.Equal(
            new[] { new[] { "A", "100" }, ["B", "200"], ["C", "300"] },
            ReadRowTexts(output));
    }

    [Fact]
    public void Merge_SharedAndInlineStrings_ArePreserved()
    {
        using var dir = new TempDir();
        var shared = dir.File("shared.xlsx");
        var inline = dir.File("inline.xlsx");
        TestTableWorkbookFactory.CreateTable(shared, "表", ["商品"], [["共有文字列"], ["共有文字列"]],
            useSharedStrings: true);
        TestTableWorkbookFactory.CreateTable(inline, "表", ["商品"], [["インライン文字列"]]);

        var preview = CreatePreview(NoMetadataOptions, (shared, "表"), (inline, "表"));
        var output = dir.File("out.xlsx");
        Assert.True(new TableMerger().Execute(preview, NoMetadataOptions, output).Success);

        var rows = ReadRows(output);
        Assert.Equal(3, rows.Count);
        Assert.All(rows, row => Assert.Equal(MergeValueKind.Text, row[0].Kind));
        Assert.Equal(new[] { "共有文字列", "共有文字列", "インライン文字列" }, rows.Select(row => row[0].Text));
    }

    [Fact]
    public void Merge_NumberAndBoolean_KeepTheirCellTypes()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestTableWorkbookFactory.CreateTable(path, "表", ["数量", "有効"],
            [[123.45, true], [0, false]]);

        var preview = CreatePreview(NoMetadataOptions, (path, "表"));
        var output = dir.File("out.xlsx");
        Assert.True(new TableMerger().Execute(preview, NoMetadataOptions, output).Success);

        var rows = ReadRows(output);
        Assert.Equal(2, rows.Count);

        Assert.Equal(MergeValueKind.Number, rows[0][0].Kind);
        Assert.Equal(123.45, rows[0][0].Number, 6);
        Assert.Equal(MergeValueKind.Boolean, rows[0][1].Kind);
        Assert.True(rows[0][1].Boolean);

        Assert.Equal(MergeValueKind.Number, rows[1][0].Kind);
        Assert.Equal(0, rows[1][0].Number);
        Assert.Equal(MergeValueKind.Boolean, rows[1][1].Kind);
        Assert.False(rows[1][1].Boolean);
    }

    [Fact]
    public void Merge_DateValues_StayDatesInsteadOfPlainNumbers()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        var date = new DateTime(2026, 3, 15);
        var timestamp = new DateTime(2026, 3, 15, 9, 30, 0);
        TestTableWorkbookFactory.CreateTable(path, "表", ["日付", "日時", "時刻"],
            [[date, timestamp, new TimeOfDayValue(new TimeSpan(13, 45, 0))]]);

        var preview = CreatePreview(NoMetadataOptions, (path, "表"));
        var output = dir.File("out.xlsx");
        Assert.True(new TableMerger().Execute(preview, NoMetadataOptions, output).Success);

        var row = Assert.Single(ReadRows(output));

        Assert.Equal(MergeValueKind.Date, row[0].Kind);
        Assert.Equal(date, MergeCellValue.SerialToDateTime(row[0].Number));

        Assert.Equal(MergeValueKind.DateTime, row[1].Kind);
        Assert.Equal(timestamp, MergeCellValue.SerialToDateTime(row[1].Number));

        Assert.Equal(MergeValueKind.Time, row[2].Kind);
        Assert.Equal(new TimeSpan(13, 45, 0).TotalDays, row[2].Number, 6);
    }

    [Fact]
    public void Merge_Date1904Workbook_IsConvertedToTheSameDate()
    {
        using var dir = new TempDir();
        var system1900 = dir.File("1900.xlsx");
        var system1904 = dir.File("1904.xlsx");
        var date = new DateTime(2026, 3, 15);
        TestTableWorkbookFactory.CreateTable(system1900, "表", ["日付"], [[date]]);
        TestTableWorkbookFactory.CreateTable(system1904, "表", ["日付"], [[date]], date1904: true);

        var preview = CreatePreview(NoMetadataOptions, (system1900, "表"), (system1904, "表"));
        var output = dir.File("out.xlsx");
        Assert.True(new TableMerger().Execute(preview, NoMetadataOptions, output).Success);

        var rows = ReadRows(output);
        Assert.Equal(2, rows.Count);
        Assert.All(rows, row => Assert.Equal(MergeValueKind.Date, row[0].Kind));
        Assert.All(rows, row => Assert.Equal(date, MergeCellValue.SerialToDateTime(row[0].Number)));
    }

    [Fact]
    public void Merge_AmbiguousDateFormat_WarnsAndKeepsTheValueAsNumber()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestTableWorkbookFactory.CreateTable(path, "表", ["値"], [[new AmbiguousFormatted(5)]]);

        var preview = CreatePreview(NoMetadataOptions, (path, "表"));

        Assert.True(preview.CanExecute);
        Assert.Contains(preview.Warnings, issue => issue.Message.Contains("日付か数値か判断できない"));

        var output = dir.File("out.xlsx");
        Assert.True(new TableMerger().Execute(preview, NoMetadataOptions, output).Success);

        var row = Assert.Single(ReadRows(output));
        Assert.Equal(MergeValueKind.Number, row[0].Kind);
        Assert.Equal(5, row[0].Number);
    }

    [Fact]
    public void Preview_SelectedSheetWithFormula_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestTableWorkbookFactory.Create(path,
            [new TestSheetSpec { Name = "表", Headers = ["商品", "売上"], Rows = [["A", 100]], AddFormulaCell = true }]);

        var preview = CreatePreview(DefaultOptions, (path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("数式を含むため"));
    }

    [Fact]
    public void Preview_SelectedSheetWithMergedCellInTable_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestTableWorkbookFactory.Create(path,
            [new TestSheetSpec { Name = "表", Headers = ["商品", "売上"], Rows = [["A", 100]], MergeReference = "A1:B1" }]);

        var preview = CreatePreview(DefaultOptions, (path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("結合セル"));
    }

    [Fact]
    public void Preview_UnselectedSheetWithFormulaOrMergedCell_DoesNotBlockTheSelectedSheet()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestTableWorkbookFactory.Create(path,
        [
            new TestSheetSpec { Name = "きれいな表", Headers = ["商品", "売上"], Rows = [["A", 100]] },
            new TestSheetSpec { Name = "数式あり", Headers = ["商品", "売上"], Rows = [["B", 200]], AddFormulaCell = true },
            new TestSheetSpec { Name = "結合あり", Headers = ["商品", "売上"], Rows = [["C", 300]], MergeReference = "A1:B2" },
        ]);

        var preview = CreatePreview(NoMetadataOptions, (path, "きれいな表"));

        Assert.True(preview.CanExecute);
        Assert.Equal(1, preview.InputDataRowCount);

        var output = dir.File("out.xlsx");
        Assert.True(new TableMerger().Execute(preview, NoMetadataOptions, output).Success);
        Assert.Equal(new[] { new[] { "A", "100" } }, ReadRowTexts(output));
    }

    [Fact]
    public void Merge_SheetWithChartAndImage_MergesCellDataOnly()
    {
        using var dir = new TempDir();
        var path = dir.File("装飾あり.xlsx");
        TestTableWorkbookFactory.Create(path,
            [new TestSheetSpec
            {
                Name = "表",
                Headers = ["商品", "売上"],
                Rows = [["A", 100], ["B", 200]],
                AddChart = true,
                AddImage = true,
            }]);

        var preview = CreatePreview(NoMetadataOptions, (path, "表"));
        Assert.True(preview.CanExecute);

        var output = dir.File("out.xlsx");
        Assert.True(new TableMerger().Execute(preview, NoMetadataOptions, output).Success);

        // データだけが統合され、Chart / Image は出力へコピーされない。
        Assert.Equal(new[] { new[] { "A", "100" }, ["B", "200"] }, ReadRowTexts(output));

        var analysis = WorkbookAnalyzer.Analyze(output);
        Assert.Equal(AnalysisStatus.Succeeded, analysis.Status);
        Assert.Equal(SafetyLevel.Normal, analysis.Level);
        Assert.DoesNotContain(analysis.Findings, finding => finding.Type == FindingType.Chart);
        Assert.DoesNotContain(analysis.Findings, finding => finding.Type == FindingType.Image);
    }

    [Fact]
    public void Merge_WithoutMetadataColumns_OmitsThem()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestTableWorkbookFactory.CreateTable(path, "表", ["商品"], [["A"]]);

        var preview = CreatePreview(NoMetadataOptions, (path, "表"));

        Assert.Equal(0, preview.MetadataColumnCount);
        Assert.Equal(new[] { "商品" }, preview.OutputHeaders);
    }

    [Fact]
    public void Preview_MetadataColumnNameCollidesWithDataHeader_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestTableWorkbookFactory.CreateTable(path, "表", ["元ファイル", "売上"], [["A", 100]]);

        var preview = CreatePreview(DefaultOptions, (path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("重複しています"));
    }

    [Fact]
    public void Preview_NoSelection_IsBlocked()
    {
        var preview = new MergePlanner().CreatePreview([], null, DefaultOptions);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("選択されていません"));
    }

    [Fact]
    public void Preview_UnreadableFile_IsBlockedWithoutThrowing()
    {
        using var dir = new TempDir();
        var path = dir.File("corrupt.xlsx");
        TestWorkbookFactory.CreateCorrupt(path);

        var preview = CreatePreview(DefaultOptions, (path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains("読み取れません"));
    }

    // --- helpers -------------------------------------------------------

    /// <summary>最初の選択シートを基準にしてプレビューを作る(既定の使い方)。</summary>
    private static MergePreview CreatePreview(
        MergeOptions options,
        params (string Path, string Sheet)[] selections)
        => CreatePreviewWithBase(options, baseIndex: 0, selections);

    private static MergePreview CreatePreviewWithBase(
        MergeOptions options,
        int baseIndex,
        params (string Path, string Sheet)[] selections)
    {
        var list = selections.Select(s => new MergeSourceSelection(s.Path, s.Sheet)).ToList();
        return new MergePlanner().CreatePreview(list, list.ElementAtOrDefault(baseIndex), options);
    }

    private static IReadOnlyList<string> ReadHeaders(string path)
        => WorksheetTableScanner.Scan(path, OutputSheetName, CancellationToken.None).Headers;

    private static List<MergeCellValue[]> ReadRows(string path)
    {
        var headerCount = ReadHeaders(path).Count;
        return [.. WorksheetTableScanner.ReadDataRows(path, OutputSheetName, headerCount)];
    }

    private static string[][] ReadRowTexts(string path)
        => [.. ReadRows(path).Select(row => row.Select(value => value.ToDisplayString()).ToArray())];
}
