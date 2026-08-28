using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Mapping;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2C2: データ元の表と転記先の表をキーで突合し、既存行の指定列を更新する。
/// 両側に存在するキーだけを更新し、行の追加・削除はしない。
/// 書き込みは Phase 2B / 2C1 と同じ仕組みを使う。
/// </summary>
public sealed class TableUpdateTests
{
    private const string OutputSuffix = "_更新済み";

    // ── A. 基本の突合 ────────────────────────────────────

    [Fact]
    public void Execute_RowsInTheSameOrder_AreUpdated()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["商品コード,単価,在庫", "A001,1200,10", "A002,1500,20"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価", "在庫"],
            [["A001", 1100, 5], ["A002", 1400, 18]]));

        var result = Execute(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number), Map("在庫", "在庫", CellWriteKind.Number)));

        Assert.True(result.Success, result.Message);
        Assert.Equal(4, result.ChangedCellCount);

        var output = Output(dir, "マスタ");
        Assert.Equal("1200", Number(output, "商品一覧", "B2"));
        Assert.Equal("10", Number(output, "商品一覧", "C2"));
        Assert.Equal("1500", Number(output, "商品一覧", "B3"));
        Assert.Equal("20", Number(output, "商品一覧", "C3"));
    }

    [Fact]
    public void Execute_RowsInReversedOrder_MatchByKeyNotByPosition()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["商品コード,単価", "A001,1200", "A002,1500"]);

        // 転記先は逆順。行番号ではなくキーで対応付く。
        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価"],
            [["A002", 1400], ["A001", 1100]]));

        Assert.True(Execute(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number))).Success);

        var output = Output(dir, "マスタ");
        Assert.Equal("1500", Number(output, "商品一覧", "B2")); // A002 の行
        Assert.Equal("1200", Number(output, "商品一覧", "B3")); // A001 の行
    }

    [Fact]
    public void Execute_DifferentKeyColumnNames_AreAllowed()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["SKU,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価"],
            [["A001", 1100]]));

        var request = Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)) with
        {
            SourceKeyColumn = "SKU",
            TargetKeyColumn = "商品コード",
        };

        Assert.True(Execute(request).Success);
        Assert.Equal("1200", Number(Output(dir, "マスタ"), "商品一覧", "B2"));
    }

    [Fact]
    public void Execute_MixedTextAndNumberMappings_AreWrittenPerColumn()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["商品コード,商品名,単価,在庫,担当,区分", "A001,商品A改,1200,10,佐藤,新"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "商品名", "単価", "在庫", "担当", "区分"],
            [["A001", "商品A", 1100, 5, "田中", "旧"]]));

        var result = Execute(Request(source, [(target, "商品一覧")],
            Map("商品名", "商品名"),
            Map("単価", "単価", CellWriteKind.Number),
            Map("在庫", "在庫", CellWriteKind.Number),
            Map("担当", "担当"),
            Map("区分", "区分")));

        Assert.True(result.Success, result.Message);
        Assert.Equal(5, result.ChangedCellCount);

        var output = Output(dir, "マスタ");
        Assert.Equal("商品A改", Text(output, "商品一覧", "B2"));
        Assert.Equal("1200", Number(output, "商品一覧", "C2"));
        Assert.Equal("佐藤", Text(output, "商品一覧", "E2"));
    }

    [Fact]
    public void Execute_MultipleWorkbooksAndSheets_AreUpdatedIndependently()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["商品コード,在庫", "A001,10", "B001,20"]);

        var first = dir.File("東京.xlsx");
        CreateTarget(first,
            Table("1月", ["商品コード", "在庫"], [["A001", 1]]),
            Table("2月", ["商品コード", "在庫"], [["A001", 2]]));

        var second = dir.File("大阪.xlsx");
        CreateTarget(second, Table("1月", ["商品コード", "在庫"], [["B001", 3]]));

        var result = Execute(Request(source,
            [(first, "1月"), (first, "2月"), (second, "1月")],
            Map("在庫", "在庫", CellWriteKind.Number)));

        Assert.True(result.Success, result.Message);
        Assert.Equal(3, result.ChangedCellCount);

        // 別シートに同じキーがあるのは正常で、両方が同じデータ元の行で更新される。
        Assert.Equal("10", Number(Output(dir, "東京"), "1月", "B2"));
        Assert.Equal("10", Number(Output(dir, "東京"), "2月", "B2"));
        Assert.Equal("20", Number(Output(dir, "大阪"), "1月", "B2"));
    }

    [Fact]
    public void Execute_TargetHeaderOnRowThree_IsUsed()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価"],
            [["A001", 1100]],
            HeaderRow: 3));

        var request = Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)) with
        {
            TargetHeaderRow = 3,
        };

        Assert.True(Execute(request).Success);
        Assert.Equal("1200", Number(Output(dir, "マスタ"), "商品一覧", "B4"));
    }

    // ── B. 両側に存在するキーだけを更新する ──────────────

    [Fact]
    public void Execute_SourceOnlyAndTargetOnlyRows_AreLeftAloneWithWarnings()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["商品コード,単価", "A001,1200", "A002,1500", "A003,900"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価"],
            [["A001", 1100], ["A002", 1400], ["A999", 500]]));

        var preview = Preview(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)));

        Assert.True(preview.Mutation.CanExecute,
            string.Join(" / ", preview.Mutation.Blocks.Select(issue => issue.Message)));

        // 集計: 一致 2 / データ元のみ 1 / 転記先のみ 1。
        Assert.Equal(2, preview.Summary.MatchedKeyCount);
        Assert.Equal(1, preview.Summary.SourceOnlyKeyCount);
        Assert.Equal(1, preview.Summary.TargetOnlyKeyCount);
        Assert.Contains(preview.Mutation.Warnings, issue => issue.Message.Contains("行の追加はしません"));
        Assert.Contains(preview.Mutation.Warnings, issue => issue.Message.Contains("そのまま残します"));

        Assert.True(new CellMutator().Execute(preview.Mutation).Success);

        var output = Output(dir, "マスタ");

        // A999 の行はそのまま。行の追加も無い(データ行は 3 行のまま)。
        Assert.Equal("500", Number(output, "商品一覧", "B4"));
        Assert.Equal(4, RowCount(output, "商品一覧"));
    }

    [Fact]
    public void Preview_NoMatchingKeys_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["B999", 1100]]));

        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)), "一致する行がありません");
        Assert.Empty(Directory.GetFiles(dir.Root, "*更新済み*"));
    }

    [Fact]
    public void Execute_OneMatchAmongMany_UpdatesOnlyThatRow()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["商品コード,単価", "X100,50", "A002,1500", "X200,60"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価"],
            [["A001", 1100], ["A002", 1400], ["A003", 1300]]));

        var result = Execute(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)));

        Assert.True(result.Success, result.Message);
        Assert.Equal(1, result.ChangedCellCount);

        var output = Output(dir, "マスタ");
        Assert.Equal("1100", Number(output, "商品一覧", "B2"));
        Assert.Equal("1500", Number(output, "商品一覧", "B3"));
        Assert.Equal("1300", Number(output, "商品一覧", "B4"));
    }

    // ── C. 重複キー ──────────────────────────────────────

    [Fact]
    public void Preview_UsedSourceDuplicate_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["商品コード,単価", "A001,1200", "A001,1250"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)), "データ元に 2 件以上");
    }

    [Fact]
    public void Preview_UnusedSourceDuplicate_IsOnlyAWarning()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["商品コード,単価", "A001,1200", "Z900,10", "Z900,20"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        var preview = Preview(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)));

        Assert.True(preview.Mutation.CanExecute);
        Assert.Contains(preview.Mutation.Warnings, issue => issue.Message.Contains("重複"));
    }

    [Fact]
    public void Preview_UsedTargetDuplicate_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価"],
            [["A001", 1100], ["A001", 1150]]));

        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)), "このシートに 2 行以上");
    }

    [Fact]
    public void Preview_UnusedTargetDuplicate_IsOnlyAWarning()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価"],
            [["A001", 1100], ["Z900", 10], ["Z900", 20]]));

        var preview = Preview(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)));

        Assert.True(preview.Mutation.CanExecute,
            string.Join(" / ", preview.Mutation.Blocks.Select(issue => issue.Message)));
        Assert.Contains(preview.Mutation.Warnings, issue => issue.Message.Contains("重複"));
    }

    [Fact]
    public void Execute_SameKeyInDifferentSheets_IsNotADuplicate()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,在庫", "A001,10"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target,
            Table("1月", ["商品コード", "在庫"], [["A001", 1]]),
            Table("2月", ["商品コード", "在庫"], [["A001", 2]]));

        var result = Execute(Request(source, [(target, "1月"), (target, "2月")],
            Map("在庫", "在庫", CellWriteKind.Number)));

        Assert.True(result.Success, result.Message);
        Assert.Equal(2, result.ChangedCellCount);
    }

    // ── D. キーの扱い ────────────────────────────────────

    [Theory]
    [InlineData("a001")]
    [InlineData(" A001 ")]
    public void Preview_KeyDifferingByCaseOrSpace_DoesNotMatch(string targetKey)
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [[targetKey, 1100]]));

        // 表記のゆれを勝手に吸収しない。
        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)), "一致する行がありません");
    }

    [Fact]
    public void Execute_LeadingZeroKeys_ArePreserved()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "00123,700"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["00123", 650]]));

        Assert.True(Execute(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number))).Success);
        Assert.Equal("700", Number(Output(dir, "マスタ"), "商品一覧", "B2"));
    }

    [Fact]
    public void Execute_BlankTargetKeyRows_AreSkippedWithCounts()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価"],
            [
                ["A001", 1100],
                [null, null],       // 完全な空行
                [null, 999],        // キーだけ空欄で値あり → Warning
                ["A900", 500],
            ]));

        var preview = Preview(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)));

        Assert.True(preview.Mutation.CanExecute,
            string.Join(" / ", preview.Mutation.Blocks.Select(issue => issue.Message)));
        Assert.Contains(preview.Mutation.Warnings, issue => issue.Message.Contains("キーが空欄"));
        Assert.True(new CellMutator().Execute(preview.Mutation).Success);

        var output = Output(dir, "マスタ");
        Assert.Equal("1200", Number(output, "商品一覧", "B2"));
        Assert.Equal("999", Number(output, "商品一覧", "B4")); // 空欄キーの行はそのまま
    }

    [Theory]
    [InlineData("number")]
    [InlineData("formula")]
    [InlineData("richtext")]
    public void Preview_NonTextTargetKey_BlocksTheSheet(string kind)
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        var cells = new List<MutationTestCell>
        {
            new("A1", "商品コード"),
            new("B1", "単価"),
            new("A2", "A001"),
            new("B2", 1100),
            new("B3", 900),
        };

        if (kind == "number")
        {
            cells.Add(new MutationTestCell("A3", 123));
        }

        CreateTarget(target, new MutationTestSheet
        {
            Name = "商品一覧",
            Cells = [.. cells],
            FormulaCell = kind == "formula" ? "A3" : null,
            RichTextCell = kind == "richtext" ? "A3" : null,
        });

        // 「A001」と数値 1 を推測で比べない。キー列は文字列だけの表として扱う。
        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)),
            kind == "formula" ? "計算結果を保証できません" : "キー列");
    }

    // ── E. ヘッダー ──────────────────────────────────────

    [Fact]
    public void Preview_BlankTargetHeader_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, new MutationTestSheet
        {
            Name = "商品一覧",
            Cells =
            [
                new MutationTestCell("A1", "商品コード"),
                new MutationTestCell("C1", "単価"), // B1 が抜けている
                new MutationTestCell("A2", "A001"),
                new MutationTestCell("C2", 1100),
            ],
        });

        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)), "項目名が空");
    }

    [Fact]
    public void Preview_DuplicateTargetHeader_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価", "単価"],
            [["A001", 1100, 1150]]));

        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)), "重複");
    }

    [Fact]
    public void Preview_RequiredTargetColumnMissing_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "在庫"], [["A001", 5]]));

        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)), "「単価」がありません");
    }

    [Fact]
    public void Execute_ExtraTargetColumns_AreAllowedAndUntouched()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価", "備考", "担当"],
            [["A001", 1100, "残す", "田中"]]));

        Assert.True(Execute(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number))).Success);

        var output = Output(dir, "マスタ");
        Assert.Equal("1200", Number(output, "商品一覧", "B2"));
        Assert.Equal("残す", Text(output, "商品一覧", "C2"));
        Assert.Equal("田中", Text(output, "商品一覧", "D2"));
    }

    [Fact]
    public void Preview_RequiredColumnMissingOnOneOfTwoSheets_BlocksTheWholeBatch()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        // 1月には単価列があるが、2月には無い。列の位置が違うのは問題ない。
        var target = dir.File("マスタ.xlsx");
        CreateTarget(target,
            Table("1月", ["商品コード", "備考", "単価"], [["A001", "x", 1100]]),
            Table("2月", ["商品コード", "備考"], [["A001", "y"]]));

        var preview = Preview(Request(source, [(target, "1月"), (target, "2月")],
            Map("単価", "単価", CellWriteKind.Number)));

        Assert.False(preview.Mutation.CanExecute);
        Assert.Contains(preview.Mutation.Blocks, issue => issue.Message.Contains("「単価」がありません"));
        Assert.False(new CellMutator().Execute(preview.Mutation).Success);
        Assert.Empty(Directory.GetFiles(dir.Root, "*更新済み*"));
    }

    [Fact]
    public void Execute_ColumnsInDifferentPositionsPerSheet_AreResolvedByName()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target,
            Table("1月", ["商品コード", "単価"], [["A001", 1100]]),
            Table("2月", ["単価", "商品コード"], [[1150, "A001"]]));

        Assert.True(Execute(Request(source, [(target, "1月"), (target, "2月")],
            Map("単価", "単価", CellWriteKind.Number))).Success);

        var output = Output(dir, "マスタ");
        Assert.Equal("1200", Number(output, "1月", "B2"));
        Assert.Equal("1200", Number(output, "2月", "A2")); // 2月では単価が A 列
    }

    // ── F. 対応付け ──────────────────────────────────────

    [Fact]
    public void Preview_DuplicateTargetColumnInMappings_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,定価,売価", "A001,1200,1000"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "価格"], [["A001", 900]]));

        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("定価", "価格", CellWriteKind.Number),
            Map("売価", "価格", CellWriteKind.Number)), "重複");
    }

    [Fact]
    public void Execute_SameSourceColumnToTwoTargetColumns_IsAllowed()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,名称", "A001,商品A改"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "商品名", "表示名"],
            [["A001", "旧", "旧"]]));

        Assert.True(Execute(Request(source, [(target, "商品一覧")],
            Map("名称", "商品名"), Map("名称", "表示名"))).Success);

        var output = Output(dir, "マスタ");
        Assert.Equal("商品A改", Text(output, "商品一覧", "B2"));
        Assert.Equal("商品A改", Text(output, "商品一覧", "C2"));
    }

    [Fact]
    public void Preview_TargetKeyColumnAsMappingTarget_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,新コード", "A001,B001"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        // 照合に使っている列は更新できない。
        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("新コード", "商品コード")), "キーの列");
    }

    [Theory]
    [InlineData("source")]
    [InlineData("target")]
    public void Preview_UnknownColumnInAMapping_IsBlocked(string side)
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        var mapping = side == "source"
            ? Map("存在しない項目", "単価", CellWriteKind.Number)
            : Map("単価", "存在しない項目", CellWriteKind.Number);

        AssertBlocked(Request(source, [(target, "商品一覧")], mapping), "ありません");
    }

    // ── G. 値 ────────────────────────────────────────────

    [Fact]
    public void Preview_BlankMatchedSourceValue_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)), "空欄");
    }

    [Theory]
    [InlineData("formula")]
    [InlineData("percent")]
    public void Preview_UnsupportedXlsxSourceValue_IsBlocked(string kind)
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "データ",
            [
                [(SourceTestCell)"商品コード", (SourceTestCell)"単価"],
                [
                    (SourceTestCell)"A001",
                    kind == "formula"
                        ? new SourceTestCell(new SourceTestFormula("1+1", "2"))
                        : new SourceTestCell(0.15, StyleId: 1),
                ],
            ],
            styles: [new MutationTestStyle(NumberFormatId: 9)]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        var request = Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)) with
        {
            SourceSheetName = "データ",
        };

        AssertBlocked(request, kind == "formula" ? "数式" : "表示形式");
    }

    [Fact]
    public void Execute_CsvQuotedThousandsNumber_IsRead()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,\"1,500\""]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        Assert.True(Execute(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number))).Success);
        Assert.Equal("1500", Number(Output(dir, "マスタ"), "商品一覧", "B2"));
    }

    // ── H. 転記先の安全確認(既存 guard の再利用)────────

    [Theory]
    [InlineData("missing")]
    [InlineData("merged")]
    [InlineData("validation")]
    [InlineData("hyperlink")]
    [InlineData("richtext")]
    [InlineData("protected")]
    public void Preview_OneUnsafeUpdateCell_BlocksTheWholeBatch(string kind)
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["商品コード,単価,備考", "A001,1200,新", "A002,1500,新"]);

        var target = dir.File("マスタ.xlsx");
        var sheet = new MutationTestSheet
        {
            Name = "商品一覧",
            Cells =
            [
                new MutationTestCell("A1", "商品コード"),
                new MutationTestCell("B1", "単価"),
                new MutationTestCell("C1", "備考"),
                new MutationTestCell("A2", "A001"),
                new MutationTestCell("B2", 1100),
                new MutationTestCell("C2", "旧"),
                new MutationTestCell("A3", "A002"),
                new MutationTestCell("B3", 1400),
                // kind == "missing" では C3 を作らない。
                .. kind == "missing"
                    ? Array.Empty<MutationTestCell>()
                    : [new MutationTestCell("C3", "旧")],
            ],
            Merges = kind == "merged" ? ["C3:D3"] : [],
            DataValidationSqref = kind == "validation" ? "C3:C10" : null,
            HyperlinkReference = kind == "hyperlink" ? "C3" : null,
            RichTextCell = kind == "richtext" ? "C3" : null,
            AddProtection = kind == "protected",
        };

        CreateTarget(target, sheet);

        var preview = Preview(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number), Map("備考", "備考")));

        // 安全な行(A001)だけを更新する経路は作らない。
        Assert.False(preview.Mutation.CanExecute);
        Assert.False(new CellMutator().Execute(preview.Mutation).Success);
        Assert.Empty(Directory.GetFiles(dir.Root, "*更新済み*"));
    }

    [Fact]
    public void Preview_TargetWorkbookWithAFormula_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, new MutationTestSheet
        {
            Name = "商品一覧",
            Cells =
            [
                new MutationTestCell("A1", "商品コード"),
                new MutationTestCell("B1", "単価"),
                new MutationTestCell("A2", "A001"),
                new MutationTestCell("B2", 1100),
            ],
            FormulaCell = "Z99",
        });

        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)), "計算結果を保証できません");
    }

    [Fact]
    public void Preview_DateFormattedUpdateCell_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        TestMutationWorkbookFactory.Create(target,
            [
                new MutationTestSheet
                {
                    Name = "商品一覧",
                    Cells =
                    [
                        new MutationTestCell("A1", "商品コード"),
                        new MutationTestCell("B1", "単価"),
                        new MutationTestCell("A2", "A001"),
                        new MutationTestCell("B2", 45000, StyleId: 1), // 日付書式
                    ],
                },
            ],
            [new MutationTestStyle(NumberFormatId: 14)]);

        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)), "表示形式");
    }

    [Fact]
    public void Preview_TableBeyondTheTestedRowLimit_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価"],
            [["A001", 1100], ["A002", 1150], ["A003", 1300], ["A004", 1400]]));

        // 上限は実測済みの範囲(既定 100,000 行)。テストでは小さくして動作だけ確かめる。
        var preview = new TableUpdatePlanner(maxKeyedRowsPerSheet: 3).CreatePreview(
            Request(source, [(target, "商品一覧")], Map("単価", "単価", CellWriteKind.Number)));

        Assert.False(preview.Mutation.CanExecute);
        Assert.Contains(preview.Mutation.Blocks, issue => issue.Message.Contains("動作を確認した範囲"));
    }

    // ── I. No-op ─────────────────────────────────────────

    [Fact]
    public void Execute_PartialNoOp_WritesOnlyWhatDiffersAndAuditsOnlyThat()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["商品コード,単価", "A001,1100", "A002,1500"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "単価"],
            [["A001", 1100], ["A002", 1400]])); // A001 は既に同じ値

        var preview = Preview(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)));

        Assert.True(preview.Mutation.CanExecute);
        Assert.Equal(1, preview.Mutation.NoOpCount);
        Assert.Equal(1, preview.Mutation.ChangeCount);
        Assert.True(new CellMutator().Execute(preview.Mutation).Success);

        using var json = JsonDocument.Parse(
            File.ReadAllText(dir.File("マスタ" + OutputSuffix + ".xlsx.audit.json")));
        var change = Assert.Single(json.RootElement.GetProperty("changes").EnumerateArray());
        Assert.Equal("A002", change.GetProperty("key").GetString());
    }

    [Fact]
    public void Preview_AllMatchesAreNoOps_CannotExecute()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1100"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        var preview = Preview(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)));

        Assert.False(preview.Mutation.CanExecute);
        Assert.Equal(1, preview.Summary.MatchedKeyCount); // 一致はしている
        Assert.Empty(Directory.GetFiles(dir.Root, "*更新済み*"));
    }

    // ── J. 共通 pipeline の再利用 ─────────────────────────

    [Fact]
    public void Execute_ChangesOnlyTheTargetWorksheetPartAndKeepsStyles()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        TestMutationWorkbookFactory.Create(target,
            [
                new MutationTestSheet
                {
                    Name = "商品一覧",
                    Cells =
                    [
                        new MutationTestCell("A1", "商品コード"),
                        new MutationTestCell("B1", "単価"),
                        new MutationTestCell("A2", "A001"),
                        new MutationTestCell("B2", 1100, StyleId: 1),
                    ],
                    AddChart = true,
                    AddImage = true,
                    AddTable = true,
                    AddConditionalFormatting = true,
                },
                new MutationTestSheet
                {
                    Name = "参考",
                    Cells = [new MutationTestCell("A1", "そのまま")],
                },
            ],
            [new MutationTestStyle(NumberFormatId: 4)]);

        Assert.True(Execute(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number))).Success);

        var before = Entries(target);
        var after = Entries(Output(dir, "マスタ"));
        var changed = before.Keys.Union(after.Keys)
            .Where(name => !before.TryGetValue(name, out var left)
                || !after.TryGetValue(name, out var right)
                || left != right)
            .ToList();

        // 変わるのは対象 WorksheetPart のみ。Excel Table や図・グラフはそのまま残る。
        Assert.Equal(["xl/worksheets/sheet1.xml"], changed);
        Assert.Equal(1U, ReadCell(Output(dir, "マスタ"), "商品一覧", "B2").StyleIndex?.Value);
    }

    [Fact]
    public void Execute_OutputPassesTheOpenXmlValidator()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        Assert.True(Execute(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number))).Success);

        using var stream = new FileStream(
            Output(dir, "マスタ"), FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
    }

    // ── K. 安全な実行 ────────────────────────────────────

    [Fact]
    public void Execute_LeavesSourceAndTargetUnchanged()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        var before = new[] { Snapshot(source), Snapshot(target) };

        Assert.True(Execute(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number))).Success);

        Assert.Equal(before, new[] { Snapshot(source), Snapshot(target) });
    }

    [Fact]
    public void Execute_SourceChangedAfterPreview_AbortsTheWholeBatch()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        var preview = Preview(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)));

        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,9999"]);

        var result = new CellMutator().Execute(preview.Mutation);

        Assert.False(result.Success);
        Assert.Contains("データ元", result.Message);
        Assert.Empty(Directory.GetFiles(dir.Root, "*更新済み*"));
    }

    [Fact]
    public void Execute_TargetChangedAfterPreview_AbortsTheWholeBatch()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        var preview = Preview(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)));

        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 999]]));

        var result = new CellMutator().Execute(preview.Mutation);

        Assert.False(result.Success);
        Assert.Contains("プレビュー後に変更されました", result.Message);
        Assert.Empty(Directory.GetFiles(dir.Root, "*更新済み*"));
    }

    [Fact]
    public void Preview_OutputOrAuditCollision_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        File.WriteAllText(dir.File("マスタ" + OutputSuffix + ".xlsx"), "架空の既存ファイル");
        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)), "既にあります");

        File.Delete(dir.File("マスタ" + OutputSuffix + ".xlsx"));
        File.WriteAllText(dir.File("マスタ" + OutputSuffix + ".xlsx.audit.json"), "{}");
        AssertBlocked(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)), "既にあります");
    }

    [Fact]
    public void Execute_FailureRollsBackAndReportsAccurately()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,単価", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧", ["商品コード", "単価"], [["A001", 1100]]));

        var preview = Preview(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number)));

        Directory.CreateDirectory(dir.File("マスタ" + OutputSuffix + ".xlsx.audit.json"));

        var result = new CellMutator().Execute(preview.Mutation);

        Assert.False(result.Success);
        Assert.Contains("取り消しました", result.Message);
        Assert.DoesNotContain("作成していません", result.Message);
        Assert.False(File.Exists(Output(dir, "マスタ")));
        Assert.Empty(Directory.GetFiles(dir.Root, "~ebt-*"));
    }

    // ── L. 控えファイル(schemaVersion 3)──────────────────

    [Fact]
    public void Execute_AuditRecordsBothSidesOfTheMatch()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["商品コード,単価", "Z900,1", "A001,1200"]);

        var target = dir.File("マスタ.xlsx");
        CreateTarget(target, Table("商品一覧",
            ["商品コード", "備考", "単価"],
            [["B777", "x", 9], ["A001", "y", 1100]]));

        Assert.True(Execute(Request(source, [(target, "商品一覧")],
            Map("単価", "単価", CellWriteKind.Number))).Success);

        var auditPath = dir.File("マスタ" + OutputSuffix + ".xlsx.audit.json");
        using var json = JsonDocument.Parse(File.ReadAllText(auditPath));
        var root = json.RootElement;

        Assert.Equal(3, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("map-source-table-to-target-table", root.GetProperty("operation").GetString());

        var dataSource = root.GetProperty("dataSource");
        Assert.Equal("元データ.csv", dataSource.GetProperty("fileName").GetString());
        Assert.Equal("商品コード", dataSource.GetProperty("keyColumn").GetString());

        var table = root.GetProperty("targetTable");
        Assert.Equal(1, table.GetProperty("headerRow").GetInt32());
        Assert.Equal("商品コード", table.GetProperty("keyColumn").GetString());

        var change = Assert.Single(root.GetProperty("changes").EnumerateArray());
        Assert.Equal("A001", change.GetProperty("key").GetString());
        Assert.Equal("単価", change.GetProperty("sourceColumn").GetString());
        Assert.Equal("単価", change.GetProperty("targetColumn").GetString());
        Assert.Equal(3, change.GetProperty("sourceRowNumber").GetInt32()); // CSV レコード番号
        Assert.Equal(3, change.GetProperty("targetRowNumber").GetInt32()); // ワークシート行番号
        Assert.Equal("C3", change.GetProperty("cell").GetString());

        var text = File.ReadAllText(auditPath);
        Assert.DoesNotContain(dir.Root, text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(":\\", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Execute_FixedCellMappingAuditStaysAtSchemaVersionTwo()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", "OSAKA,佐藤"]);

        var target = dir.File("大阪.xlsx");
        TestMutationWorkbookFactory.Create(target,
        [
            new MutationTestSheet
            {
                Name = "月報",
                Cells = [new MutationTestCell("A1", "OSAKA"), new MutationTestCell("D5", "旧")],
            },
        ]);

        var preview = new SourceMappingPlanner().CreatePreview(new SourceMappingBatchRequest
        {
            SourceFilePath = source,
            HeaderRow = 1,
            KeyColumn = "店舗コード",
            TargetKeyCell = "A1",
            Targets = [new CellMutationTarget(target, "月報")],
            Mappings = [new SourceMappingRequest { SourceColumn = "担当者", TargetCell = "D5" }],
        });

        Assert.True(new CellMutator().Execute(preview).Success);

        using var json = JsonDocument.Parse(
            File.ReadAllText(dir.File("大阪_転記済み.xlsx.audit.json")));

        // 2C1 の控えの意味は変えない(targetTable も targetRowNumber も付かない)。
        Assert.Equal(2, json.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.False(json.RootElement.TryGetProperty("targetTable", out _));
        var change = Assert.Single(json.RootElement.GetProperty("changes").EnumerateArray());
        Assert.False(change.TryGetProperty("targetRowNumber", out _));
    }

    // ── ヘルパー ──────────────────────────────────────────

    private static TableColumnMappingRequest Map(
        string source, string target, CellWriteKind kind = CellWriteKind.Text)
        => new() { SourceColumn = source, TargetColumn = target, WriteKind = kind };

    private static TableUpdateBatchRequest Request(
        string sourcePath,
        (string Path, string Sheet)[] targets,
        params TableColumnMappingRequest[] mappings)
        => new()
        {
            SourceFilePath = sourcePath,
            SourceSheetName = sourcePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? "データ"
                : null,
            SourceHeaderRow = 1,
            SourceKeyColumn = "商品コード",
            TargetHeaderRow = 1,
            TargetKeyColumn = "商品コード",
            Targets = [.. targets.Select(t => new CellMutationTarget(t.Path, t.Sheet))],
            Mappings = mappings,
        };

    /// <summary>ヘッダー + データ行から、表の形をしたテスト用シート定義を作る。</summary>
    private static MutationTestSheet Table(
        string name, string[] header, object?[][] rows, int HeaderRow = 1)
    {
        var cells = new List<MutationTestCell>();

        for (var column = 0; column < header.Length; column++)
        {
            cells.Add(new MutationTestCell(
                $"{CellRangeParser.ColumnIndexToLetters(column + 1)}{HeaderRow}", header[column]));
        }

        for (var row = 0; row < rows.Length; row++)
        {
            for (var column = 0; column < rows[row].Length; column++)
            {
                if (rows[row][column] is { } value)
                {
                    cells.Add(new MutationTestCell(
                        $"{CellRangeParser.ColumnIndexToLetters(column + 1)}{HeaderRow + 1 + row}",
                        value));
                }
            }
        }

        return new MutationTestSheet { Name = name, Cells = [.. cells] };
    }

    private static void CreateTarget(string path, params MutationTestSheet[] sheets)
        => TestMutationWorkbookFactory.Create(path, sheets);

    private static TableUpdatePreview Preview(TableUpdateBatchRequest request)
        => new TableUpdatePlanner().CreatePreview(request);

    private static CellMutationResult Execute(TableUpdateBatchRequest request)
    {
        var preview = Preview(request);
        Assert.True(preview.Mutation.CanExecute,
            string.Join(" / ", preview.Mutation.Blocks.Select(issue => $"{issue.Location}: {issue.Message}")));
        return new CellMutator().Execute(preview.Mutation);
    }

    private static void AssertBlocked(TableUpdateBatchRequest request, string expectedFragment)
    {
        var preview = Preview(request);

        Assert.False(preview.Mutation.CanExecute);
        Assert.Contains(preview.Mutation.Blocks, issue => issue.Message.Contains(expectedFragment));
    }

    private static string Output(TempDir dir, string sourceName)
        => dir.File(sourceName + OutputSuffix + ".xlsx");

    private static string? Text(string path, string sheetName, string reference)
    {
        var cell = ReadCell(path, sheetName, reference);
        return cell.InlineString?.Text?.Text ?? SharedText(path, cell);
    }

    private static string? SharedText(string path, Cell cell)
    {
        if (cell.DataType?.Value != CellValues.SharedString)
        {
            return cell.CellValue?.InnerText;
        }

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var items = document.WorkbookPart!.SharedStringTablePart!.SharedStringTable!
            .Elements<SharedStringItem>().ToList();
        return items[int.Parse(cell.CellValue!.InnerText!)].InnerText;
    }

    private static string? Number(string path, string sheetName, string reference)
        => ReadCell(path, sheetName, reference).CellValue?.InnerText;

    private static int RowCount(string path, string sheetName)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name?.Value == sheetName);
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return worksheetPart.Worksheet!.Descendants<Row>().Count();
    }

    private static Cell ReadCell(string path, string sheetName, string reference)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name?.Value == sheetName);
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return (Cell)worksheetPart.Worksheet!.Descendants<Cell>()
            .Single(cell => string.Equals(
                cell.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase))
            .CloneNode(true);
    }

    private static IReadOnlyDictionary<string, string> Entries(string path)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        using var zip = ZipFile.OpenRead(path);
        foreach (var entry in zip.Entries)
        {
            using var stream = entry.Open();
            using var memory = new MemoryStream();
            stream.CopyTo(memory);
            map[entry.FullName] = Convert.ToHexString(SHA256.HashData(memory.ToArray()));
        }

        return map;
    }

    private static (string Hash, long Length, DateTime LastWriteUtc) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }
}
