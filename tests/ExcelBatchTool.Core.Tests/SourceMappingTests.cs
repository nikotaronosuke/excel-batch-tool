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
/// Phase 2C1: データ元の表からキーで 1 行を特定し、転記先の決まったセルへ入れる。
/// 転記先の安全確認・書き込みは Phase 2B と同じ仕組みを使う。
/// </summary>
public sealed class SourceMappingTests
{
    private const string OutputSuffix = "_転記済み";

    // ── A. データ元が .xlsx ──────────────────────────────

    [Fact]
    public void Execute_TextAndNumberFromXlsx_AreWrittenToTheFixedCells()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        CreateSource(source,
            ["店舗コード", "担当者", "売上"],
            ["OSAKA", "佐藤", 1500],
            ["KYOTO", "田中", 1800]);

        var osaka = dir.File("大阪.xlsx");
        var kyoto = dir.File("京都.xlsx");
        CreateTarget(osaka, "OSAKA");
        CreateTarget(kyoto, "KYOTO");

        var result = Execute(Request(source, [(osaka, "月報"), (kyoto, "月報")],
            Map("担当者", "D5"), Map("売上", "F8", CellWriteKind.Number)));

        Assert.True(result.Success, result.Message);
        Assert.Equal(4, result.ChangedCellCount);

        Assert.Equal("佐藤", Text(Output(dir, "大阪"), "月報", "D5"));
        Assert.Equal("1500", ReadCell(Output(dir, "大阪"), "月報", "F8").CellValue?.InnerText);
        Assert.Equal("田中", Text(Output(dir, "京都"), "月報", "D5"));
        Assert.Equal("1800", ReadCell(Output(dir, "京都"), "月報", "F8").CellValue?.InnerText);
    }

    [Fact]
    public void Execute_HeaderRowOtherThanTheFirst_IsUsed()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");

        // 1〜2 行目はタイトル行。項目名は 3 行目。
        TestSourceTableFactory.CreateXlsx(source, "データ",
        [
            ["店舗コード", "担当者"],
            ["OSAKA", "佐藤"],
        ],
        headerRow: 3);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        var request = Request(source, [(target, "月報")], Map("担当者", "D5")) with { HeaderRow = 3 };
        Assert.True(Execute(request).Success);

        Assert.Equal("佐藤", Text(Output(dir, "大阪"), "月報", "D5"));
    }

    [Fact]
    public void Execute_InlineStringKeyAndValue_AreRead()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "データ",
        [
            ["店舗コード", "担当者"],
            [new SourceTestCell(new InlineSourceText("OSAKA")), new SourceTestCell(new InlineSourceText("佐藤"))],
        ]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        Assert.True(Execute(Request(source, [(target, "月報")], Map("担当者", "D5"))).Success);
        Assert.Equal("佐藤", Text(Output(dir, "大阪"), "月報", "D5"));
    }

    [Fact]
    public void Preview_DuplicateSourceHeader_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        CreateSource(source, ["店舗コード", "担当者", "担当者"], ["OSAKA", "佐藤", "鈴木"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "重複");
    }

    [Fact]
    public void Preview_BlankSourceHeader_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "データ",
        [
            ["店舗コード", new SourceTestCell(null), "売上"],
            ["OSAKA", "佐藤", 1500],
        ]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        AssertBlocked(Request(source, [(target, "月報")], Map("売上", "F8", CellWriteKind.Number)), "項目名が空");
    }

    [Theory]
    [InlineData("formula")]
    [InlineData("number")]
    public void Preview_UnsupportedSourceKeyType_IsBlocked(string kind)
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");

        SourceTestCell key = kind == "formula"
            ? new SourceTestCell(new SourceTestFormula("\"OSAKA\"", "OSAKA"))
            : new SourceTestCell(123);

        TestSourceTableFactory.CreateXlsx(source, "データ",
        [
            ["店舗コード", "担当者"],
            [key, "佐藤"],
        ]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        // キーが文字列でない行は読み取れない(00123 と 123 を取り違えないため)。
        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "文字列");
    }

    [Fact]
    public void Preview_MappedSourceCellIsAFormula_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "データ",
        [
            ["店舗コード", "担当者"],
            ["OSAKA", new SourceTestCell(new SourceTestFormula("\"佐\"&\"藤\"", "佐藤"))],
        ]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "数式");
    }

    [Fact]
    public void Execute_FormulaInAnUnusedColumn_DoesNotBlock()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "データ",
        [
            ["店舗コード", "担当者", "計算"],
            ["OSAKA", "佐藤", new SourceTestCell(new SourceTestFormula("1+1", "2"))],
        ]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        // 今回読まない列の数式は転記結果に影響しない。
        Assert.True(Execute(Request(source, [(target, "月報")], Map("担当者", "D5"))).Success);
        Assert.Equal("佐藤", Text(Output(dir, "大阪"), "月報", "D5"));
    }

    [Fact]
    public void Preview_BlankMappedSourceValue_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "データ",
        [
            ["店舗コード", "担当者"],
            ["OSAKA", new SourceTestCell(null)],
        ]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        // 空欄を転記してセルを消すことはしない。
        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "空欄");
    }

    [Fact]
    public void Preview_PercentageFormattedSourceNumber_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "データ",
        [
            ["店舗コード", "達成率"],
            ["OSAKA", new SourceTestCell(0.15, StyleId: 1)],
        ],
        styles: [new MutationTestStyle(NumberFormatId: 9)]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        // 画面上 15% に見える 0.15 をそのまま転記しない。
        AssertBlocked(
            Request(source, [(target, "月報")], Map("達成率", "F8", CellWriteKind.Number)),
            "表示形式");
    }

    [Fact]
    public void Preview_RichTextSourceValue_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "データ",
        [
            ["店舗コード", "担当者"],
            ["OSAKA", new SourceTestCell(new RichSourceText("太字", "通常"))],
        ]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "文字ごとに書式");
    }

    [Fact]
    public void Preview_SourceWithAStructuralProblem_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        TestSourceTableFactory.CreateXlsx(source, "データ",
            [
                [(SourceTestCell)"店舗コード", (SourceTestCell)"担当者"],
                [(SourceTestCell)"OSAKA", (SourceTestCell)"佐藤"],
            ],
            addStyleSchemaError: true);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        // 値の意味を決めるパート(書式・共有文字列・シート一覧)が壊れているものは転記元にしない。
        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "Excel の形式として問題");
    }

    // ── B. データ元が .csv ──────────────────────────────

    [Theory]
    [InlineData("utf-8", false, "\r\n")]
    [InlineData("utf-8", true, "\r\n")]
    [InlineData("utf-8", false, "\n")]
    [InlineData("cp932", false, "\r\n")]
    public void Execute_CsvEncodingsAndLineEndings_AreRead(string encoding, bool bom, string newLine)
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["店舗コード,担当者,売上", "OSAKA,佐藤,1500"],
            encoding, bom, newLine);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        Assert.True(Execute(Request(source, [(target, "月報")],
            Map("担当者", "D5"), Map("売上", "F8", CellWriteKind.Number))).Success);

        Assert.Equal("佐藤", Text(Output(dir, "大阪"), "月報", "D5"));
        Assert.Equal("1500", ReadCell(Output(dir, "大阪"), "月報", "F8").CellValue?.InnerText);
    }

    [Theory]
    [InlineData("\"佐藤, 花子\"", "佐藤, 花子")]
    [InlineData("\"引用\"\"符\"", "引用\"符")]
    [InlineData("\"1 行目\n2 行目\"", "1 行目\n2 行目")]
    public void Execute_QuotedCsvFields_AreReadVerbatim(string field, string expected)
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", $"OSAKA,{field}"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        Assert.True(Execute(Request(source, [(target, "月報")], Map("担当者", "D5"))).Success);
        Assert.Equal(expected.Replace("\n", "\r\n", StringComparison.Ordinal),
            Text(Output(dir, "大阪"), "月報", "D5")?.Replace("\n", "\r\n", StringComparison.Ordinal));
    }

    [Fact]
    public void Execute_CsvKeyKeepsLeadingZeros()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["商品コード,担当者", "00123,佐藤"]);

        var target = dir.File("在庫.xlsx");
        CreateTarget(target, "00123");

        // 00123 を数値にしない。文字列として一致させる。
        Assert.True(Execute(Request(source, "商品コード", [(target, "月報")], Map("担当者", "D5"))).Success);
        Assert.Equal("佐藤", Text(Output(dir, "在庫"), "月報", "D5"));
    }

    [Fact]
    public void Execute_CsvThousandsSeparatedNumber_IsRead()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,売上", "OSAKA,\"1,500\""]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        Assert.True(Execute(Request(source, [(target, "月報")],
            Map("売上", "F8", CellWriteKind.Number))).Success);

        Assert.Equal("1500", ReadCell(Output(dir, "大阪"), "月報", "F8").CellValue?.InnerText);
    }

    [Fact]
    public void Preview_CsvWithADifferentColumnCount_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["店舗コード,担当者", "OSAKA,佐藤,余分"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "列数");
    }

    [Fact]
    public void Preview_MalformedCsv_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");

        // 閉じていない引用符。
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", "OSAKA,\"佐藤"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "読み取れません");
    }

    [Fact]
    public void Preview_DuplicateCsvHeader_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者,担当者", "OSAKA,佐藤,鈴木"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "重複");
    }

    [Fact]
    public void Preview_CsvNumberThatCannotBeParsed_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,売上", "OSAKA,たくさん"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        AssertBlocked(
            Request(source, [(target, "月報")], Map("売上", "F8", CellWriteKind.Number)),
            "数値として読み取れません");
    }

    // ── C. キーの照合 ────────────────────────────────────

    [Fact]
    public void Preview_KeyMissingFromTheSource_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", "KYOTO,田中"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "一致するデータ");
    }

    [Fact]
    public void Preview_RequiredKeyAppearsTwiceInTheSource_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["店舗コード,担当者", "OSAKA,佐藤", "OSAKA,鈴木"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "2 件以上");
    }

    [Fact]
    public void Preview_UnusedDuplicateKey_IsOnlyAWarning()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["店舗コード,担当者", "OSAKA,佐藤", "KYOTO,田中", "KYOTO,鈴木"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        var preview = Preview(Request(source, [(target, "月報")], Map("担当者", "D5")));

        // 今回使わないキーの重複では止めない。
        Assert.True(preview.CanExecute, string.Join(" / ", preview.Blocks.Select(i => i.Message)));
        Assert.Contains(preview.Warnings, issue => issue.Message.Contains("使いません"));
    }

    [Fact]
    public void Execute_SameKeyOnMultipleTargetSheets_UsesTheSameSourceRow()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", "OSAKA,佐藤"]);

        var target = dir.File("大阪.xlsx");
        TestMutationWorkbookFactory.Create(target,
        [
            KeySheet("1月", "OSAKA"),
            KeySheet("2月", "OSAKA"),
        ]);

        Assert.True(Execute(Request(source, [(target, "1月"), (target, "2月")], Map("担当者", "D5"))).Success);

        Assert.Equal("佐藤", Text(Output(dir, "大阪"), "1月", "D5"));
        Assert.Equal("佐藤", Text(Output(dir, "大阪"), "2月", "D5"));
    }

    [Theory]
    [InlineData("osaka")]
    [InlineData(" OSAKA ")]
    public void Preview_KeyDifferingOnlyByCaseOrSpace_DoesNotMatch(string targetKey)
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", "OSAKA,佐藤"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, targetKey);

        // 表記ゆれを勝手に吸収しない。利用者に不一致として見せる。
        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "一致するデータ");
    }

    [Theory]
    [InlineData("blank")]
    [InlineData("number")]
    [InlineData("merged")]
    public void Preview_UnusableTargetKeyCell_IsBlocked(string kind)
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", "OSAKA,佐藤"]);

        var target = dir.File("大阪.xlsx");
        var sheet = new MutationTestSheet
        {
            Name = "月報",
            Cells = kind switch
            {
                "blank" => [new MutationTestCell("A1", null), new MutationTestCell("D5", "旧")],
                "number" => [new MutationTestCell("A1", 123), new MutationTestCell("D5", "旧")],
                _ => [new MutationTestCell("A1", "OSAKA"), new MutationTestCell("D5", "旧")],
            },
            Merges = kind == "merged" ? ["A1:B1"] : [],
        };

        TestMutationWorkbookFactory.Create(target, [sheet]);

        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "照合キーのセル");
    }

    [Fact]
    public void Preview_MappingWritesToTheKeyCell_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,店舗名", "OSAKA,大阪店"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        // 照合に使っているセルを実行中に書き換えない。
        AssertBlocked(Request(source, [(target, "月報")], Map("店舗名", "A1")), "キーのセル");
    }

    // ── D. 対応付け ──────────────────────────────────────

    [Fact]
    public void Execute_TwentyMappings_AreAllApplied()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");

        var headers = new List<string> { "店舗コード" };
        var values = new List<string> { "OSAKA" };
        for (var i = 1; i <= 20; i++)
        {
            headers.Add($"項目{i}");
            values.Add($"値{i}");
        }

        TestSourceTableFactory.CreateCsv(source,
            [string.Join(",", headers), string.Join(",", values)]);

        var target = dir.File("大阪.xlsx");
        TestMutationWorkbookFactory.Create(target,
        [
            new MutationTestSheet
            {
                Name = "月報",
                Cells = [new MutationTestCell("A1", "OSAKA"),
                    .. Enumerable.Range(1, 20).Select(i => new MutationTestCell($"C{i}", $"旧{i}"))],
            },
        ]);

        var mappings = Enumerable.Range(1, 20).Select(i => Map($"項目{i}", $"C{i}")).ToArray();
        var result = Execute(Request(source, [(target, "月報")], mappings));

        Assert.True(result.Success, result.Message);
        Assert.Equal(20, result.ChangedCellCount);

        for (var i = 1; i <= 20; i++)
        {
            Assert.Equal($"値{i}", Text(Output(dir, "大阪"), "月報", $"C{i}"));
        }
    }

    [Fact]
    public void Preview_DuplicateTargetCellInTheMapping_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者,部署", "OSAKA,佐藤,営業"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        // D5 と $D$5 は同じセル。
        AssertBlocked(
            Request(source, [(target, "月報")], Map("担当者", "D5"), Map("部署", "$D$5")),
            "重複");
    }

    [Fact]
    public void Execute_SameSourceColumnToTwoCells_IsAllowed()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,店舗名", "OSAKA,大阪店"]);

        var target = dir.File("大阪.xlsx");
        TestMutationWorkbookFactory.Create(target,
        [
            new MutationTestSheet
            {
                Name = "月報",
                Cells =
                [
                    new MutationTestCell("A1", "OSAKA"),
                    new MutationTestCell("B2", "旧"),
                    new MutationTestCell("H10", "旧"),
                ],
            },
        ]);

        Assert.True(Execute(Request(source, [(target, "月報")],
            Map("店舗名", "B2"), Map("店舗名", "H10"))).Success);

        Assert.Equal("大阪店", Text(Output(dir, "大阪"), "月報", "B2"));
        Assert.Equal("大阪店", Text(Output(dir, "大阪"), "月報", "H10"));
    }

    [Fact]
    public void Preview_UnknownSourceColumn_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", "OSAKA,佐藤"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        AssertBlocked(Request(source, [(target, "月報")], Map("存在しない項目", "D5")), "ありません");
    }

    [Fact]
    public void Preview_BlankWriteKindInAMapping_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", "OSAKA,佐藤"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        AssertBlocked(
            Request(source, [(target, "月報")], Map("担当者", "D5", CellWriteKind.Blank)),
            "空欄");
    }

    [Fact]
    public void Preview_SourceUsedAsItsOwnTarget_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        CreateSource(source, ["店舗コード", "担当者"], ["OSAKA", "佐藤"]);

        AssertBlocked(Request(source, [(source, "データ")], Map("担当者", "D5")), "データ元のファイル");
    }

    // ── E. 転記先の安全確認(Phase 2B と同じ仕組み)────────

    [Theory]
    [InlineData("merged")]
    [InlineData("validation")]
    [InlineData("hyperlink")]
    [InlineData("richtext")]
    [InlineData("protected")]
    [InlineData("missing")]
    public void Preview_OneUnsafeTargetCell_BlocksTheWholeBatch(string kind)
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者,部署", "OSAKA,佐藤,営業"]);

        var target = dir.File("大阪.xlsx");
        TestMutationWorkbookFactory.Create(target,
        [
            new MutationTestSheet
            {
                Name = "月報",
                Cells =
                [
                    new MutationTestCell("A1", "OSAKA"),
                    new MutationTestCell("D5", "旧"),
                    .. kind == "missing"
                        ? Array.Empty<MutationTestCell>()
                        : [new MutationTestCell("H10", "旧")],
                ],
                Merges = kind == "merged" ? ["H10:I10"] : [],
                DataValidationSqref = kind == "validation" ? "H1:H20" : null,
                HyperlinkReference = kind == "hyperlink" ? "H10" : null,
                RichTextCell = kind == "richtext" ? "H10" : null,
                AddProtection = kind == "protected",
            },
        ]);

        var preview = Preview(Request(source, [(target, "月報")],
            Map("担当者", "D5"), Map("部署", "H10")));

        // 安全な D5 だけを転記する経路は作らない。
        Assert.False(preview.CanExecute);
        Assert.False(new CellMutator().Execute(preview).Success);
        Assert.Empty(Directory.GetFiles(dir.Root, "*転記済み*"));
    }

    [Fact]
    public void Preview_TargetWorkbookWithAFormula_IsBlocked()
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
                FormulaCell = "Z9",
            },
        ]);

        // データ元と違い、転記先はブック内に数式が 1 件でもあれば対象外。
        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "計算結果を保証できません");
    }

    [Fact]
    public void Execute_KeepsTheTargetStyleIndex()
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
                    Cells = [new MutationTestCell("A1", "OSAKA"), new MutationTestCell("D5", "旧", StyleId: 1)],
                },
            ],
            [new MutationTestStyle(NumberFormatId: 49)]);

        Assert.True(Execute(Request(source, [(target, "月報")], Map("担当者", "D5"))).Success);
        Assert.Equal(1U, ReadCell(Output(dir, "大阪"), "月報", "D5").StyleIndex?.Value);
    }

    [Fact]
    public void Execute_ChangesOnlyTheTargetWorksheetPart()
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
                AddChart = true,
                AddImage = true,
                AddTable = true,
                AddConditionalFormatting = true,
            },
            new MutationTestSheet { Name = "参考", Cells = [new MutationTestCell("A1", "そのまま")] },
        ]);

        Assert.True(Execute(Request(source, [(target, "月報")], Map("担当者", "D5"))).Success);

        var before = Entries(target);
        var after = Entries(Output(dir, "大阪"));
        var changed = before.Keys.Union(after.Keys)
            .Where(name => !before.TryGetValue(name, out var left)
                || !after.TryGetValue(name, out var right)
                || left != right)
            .ToList();

        Assert.Equal(["xl/worksheets/sheet1.xml"], changed);
    }

    [Fact]
    public void Execute_PartialNoOp_WritesOnlyWhatDiffers()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者,部署", "OSAKA,佐藤,営業"]);

        var target = dir.File("大阪.xlsx");
        TestMutationWorkbookFactory.Create(target,
        [
            new MutationTestSheet
            {
                Name = "月報",
                Cells =
                [
                    new MutationTestCell("A1", "OSAKA"),
                    new MutationTestCell("D5", "佐藤"), // 既に同じ値
                    new MutationTestCell("H10", "旧"),
                ],
            },
        ]);

        var preview = Preview(Request(source, [(target, "月報")],
            Map("担当者", "D5"), Map("部署", "H10")));

        Assert.True(preview.CanExecute);
        Assert.Equal(1, preview.NoOpCount);
        Assert.Equal(1, preview.ChangeCount);
        Assert.True(new CellMutator().Execute(preview).Success);

        using var json = JsonDocument.Parse(
            File.ReadAllText(dir.File("大阪" + OutputSuffix + ".xlsx.audit.json")));
        var change = Assert.Single(json.RootElement.GetProperty("changes").EnumerateArray());
        Assert.Equal("H10", change.GetProperty("cell").GetString());
    }

    [Fact]
    public void Preview_EverythingIsNoOp_CannotExecute()
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
                Cells = [new MutationTestCell("A1", "OSAKA"), new MutationTestCell("D5", "佐藤")],
            },
        ]);

        var preview = Preview(Request(source, [(target, "月報")], Map("担当者", "D5")));

        Assert.False(preview.CanExecute);
        Assert.Empty(Directory.GetFiles(dir.Root, "*転記済み*"));
    }

    // ── F. 安全な実行 ────────────────────────────────────

    [Fact]
    public void Execute_LeavesTheSourceAndTargetsUnchanged()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", "OSAKA,佐藤", "KYOTO,田中"]);

        var osaka = dir.File("大阪.xlsx");
        var kyoto = dir.File("京都.xlsx");
        CreateTarget(osaka, "OSAKA");
        CreateTarget(kyoto, "KYOTO");

        var before = new[] { Snapshot(source), Snapshot(osaka), Snapshot(kyoto) };

        Assert.True(Execute(Request(source, [(osaka, "月報"), (kyoto, "月報")], Map("担当者", "D5"))).Success);

        Assert.Equal(before, new[] { Snapshot(source), Snapshot(osaka), Snapshot(kyoto) });
    }

    [Theory]
    [InlineData("元データ.csv")]
    [InlineData("元データ.xlsx")]
    public void Execute_SourceChangedAfterPreview_AbortsTheWholeBatch(string sourceName)
    {
        using var dir = new TempDir();
        var source = dir.File(sourceName);
        WriteSource(source, "佐藤");

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        var preview = Preview(Request(source, [(target, "月報")], Map("担当者", "D5")));
        Assert.True(preview.CanExecute);

        // プレビュー後にデータ元を差し替える。
        WriteSource(source, "鈴木");

        var result = new CellMutator().Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("データ元", result.Message);
        Assert.Contains("プレビュー後に変更されました", result.Message);
        Assert.Empty(Directory.GetFiles(dir.Root, "*転記済み*"));

        void WriteSource(string path, string person)
        {
            if (path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                TestSourceTableFactory.CreateCsv(path, ["店舗コード,担当者", $"OSAKA,{person}"]);
            }
            else
            {
                CreateSource(path, ["店舗コード", "担当者"], ["OSAKA", person]);
            }
        }
    }

    [Fact]
    public void Execute_TargetChangedAfterPreview_AbortsTheWholeBatch()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", "OSAKA,佐藤"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        var preview = Preview(Request(source, [(target, "月報")], Map("担当者", "D5")));
        CreateTarget(target, "OSAKA", oldValue: "別の値");

        var result = new CellMutator().Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("プレビュー後に変更されました", result.Message);
        Assert.Empty(Directory.GetFiles(dir.Root, "*転記済み*"));
    }

    [Fact]
    public void Preview_OutputOrAuditAlreadyExists_IsBlocked()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", "OSAKA,佐藤"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        File.WriteAllText(dir.File("大阪" + OutputSuffix + ".xlsx"), "架空の既存ファイル");
        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "既にあります");

        File.Delete(dir.File("大阪" + OutputSuffix + ".xlsx"));
        File.WriteAllText(dir.File("大阪" + OutputSuffix + ".xlsx.audit.json"), "{}");
        AssertBlocked(Request(source, [(target, "月報")], Map("担当者", "D5")), "既にあります");
    }

    [Fact]
    public void Execute_FailureRollsBackAndReportsAccurately()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source, ["店舗コード,担当者", "OSAKA,佐藤"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        var preview = Preview(Request(source, [(target, "月報")], Map("担当者", "D5")));

        // 控えファイルの置き場所をフォルダーで塞ぎ、確定の途中で失敗させる。
        Directory.CreateDirectory(dir.File("大阪" + OutputSuffix + ".xlsx.audit.json"));

        var result = new CellMutator().Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("取り消しました", result.Message);
        Assert.DoesNotContain("作成していません", result.Message);
        Assert.False(File.Exists(Output(dir, "大阪")));
        Assert.Empty(Directory.GetFiles(dir.Root, "~ebt-*"));
    }

    [Fact]
    public void Execute_OutputPassesTheOpenXmlValidator()
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
                AddChart = true,
                AddTable = true,
                AddConditionalFormatting = true,
            },
        ]);

        Assert.True(Execute(Request(source, [(target, "月報")], Map("担当者", "D5"))).Success);

        using var stream = new FileStream(Output(dir, "大阪"), FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
    }

    // ── G. 控えファイル(schemaVersion 2)──────────────────

    [Fact]
    public void Execute_AuditRecordsWhereEachValueCameFrom()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(source,
            ["店舗コード,担当者,売上", "KYOTO,田中,1800", "OSAKA,佐藤,1500"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        Assert.True(Execute(Request(source, [(target, "月報")],
            Map("担当者", "D5"), Map("売上", "F8", CellWriteKind.Number))).Success);

        using var json = JsonDocument.Parse(
            File.ReadAllText(dir.File("大阪" + OutputSuffix + ".xlsx.audit.json")));
        var root = json.RootElement;

        Assert.Equal(2, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("map-source-to-cells", root.GetProperty("operation").GetString());

        var dataSource = root.GetProperty("dataSource");
        Assert.Equal("元データ.csv", dataSource.GetProperty("fileName").GetString());
        Assert.Equal("csv", dataSource.GetProperty("type").GetString());
        Assert.Equal("店舗コード", dataSource.GetProperty("keyColumn").GetString());
        Assert.Equal(1, dataSource.GetProperty("headerRow").GetInt32());
        Assert.Equal(Sha256(source), dataSource.GetProperty("sha256").GetString());

        var changes = root.GetProperty("changes").EnumerateArray().ToList();
        Assert.Equal(2, changes.Count);
        Assert.Equal("担当者", changes[0].GetProperty("sourceColumn").GetString());
        Assert.Equal("OSAKA", changes[0].GetProperty("key").GetString());

        // OSAKA は CSV の 3 レコード目(ヘッダーが 1)。
        Assert.Equal(3, changes[0].GetProperty("sourceRowNumber").GetInt32());
        Assert.Equal("佐藤", changes[0].GetProperty("newValue").GetString());

        var text = File.ReadAllText(dir.File("大阪" + OutputSuffix + ".xlsx.audit.json"));
        Assert.DoesNotContain(dir.Root, text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(":\\", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Execute_AuditForXlsxSourceRecordsSheetAndRow()
    {
        using var dir = new TempDir();
        var source = dir.File("元データ.xlsx");
        CreateSource(source, ["店舗コード", "担当者"], ["KYOTO", "田中"], ["OSAKA", "佐藤"]);

        var target = dir.File("大阪.xlsx");
        CreateTarget(target, "OSAKA");

        Assert.True(Execute(Request(source, [(target, "月報")], Map("担当者", "D5"))).Success);

        using var json = JsonDocument.Parse(
            File.ReadAllText(dir.File("大阪" + OutputSuffix + ".xlsx.audit.json")));

        var dataSource = json.RootElement.GetProperty("dataSource");
        Assert.Equal("xlsx", dataSource.GetProperty("type").GetString());
        Assert.Equal("データ", dataSource.GetProperty("sheetName").GetString());

        var change = Assert.Single(json.RootElement.GetProperty("changes").EnumerateArray());
        Assert.Equal(3, change.GetProperty("sourceRowNumber").GetInt32()); // 実際のワークシート行番号
    }

    [Fact]
    public void Execute_ManualInputAuditStaysAtSchemaVersionOne()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        TestMutationWorkbookFactory.Create(path,
            [new MutationTestSheet { Name = "月報", Cells = [new MutationTestCell("B2", "旧")] }]);

        var preview = new CellMutationPlanner().CreatePreview(new CellMutationRequest
        {
            Targets = [new CellMutationTarget(path, "月報")],
            Operations =
            [
                new CellMutationOperationRequest
                {
                    CellReference = "B2",
                    WriteKind = CellWriteKind.Text,
                    TextValue = "新",
                },
            ],
        });

        Assert.True(new CellMutator().Execute(preview).Success);

        using var json = JsonDocument.Parse(
            File.ReadAllText(dir.File("大阪_変更済み.xlsx.audit.json")));

        // Phase 2A / 2B の控えの意味は変えない。
        Assert.Equal(1, json.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("set-cell-value", json.RootElement.GetProperty("operation").GetString());
        Assert.False(json.RootElement.TryGetProperty("dataSource", out _));
    }

    // ── ヘルパー ──────────────────────────────────────────

    private static SourceMappingRequest Map(
        string column, string cell, CellWriteKind kind = CellWriteKind.Text)
        => new() { SourceColumn = column, TargetCell = cell, WriteKind = kind };

    private static SourceMappingBatchRequest Request(
        string sourcePath,
        (string Path, string Sheet)[] targets,
        params SourceMappingRequest[] mappings)
        => Request(sourcePath, "店舗コード", targets, mappings);

    private static SourceMappingBatchRequest Request(
        string sourcePath,
        string keyColumn,
        (string Path, string Sheet)[] targets,
        params SourceMappingRequest[] mappings)
        => new()
        {
            SourceFilePath = sourcePath,
            SourceSheetName = sourcePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? "データ"
                : null,
            HeaderRow = 1,
            KeyColumn = keyColumn,
            TargetKeyCell = "A1",
            Targets = [.. targets.Select(t => new CellMutationTarget(t.Path, t.Sheet))],
            Mappings = mappings,
        };

    private static void CreateSource(
        string path, string[] header, params SourceTestCell[][] rows)
        => TestSourceTableFactory.CreateXlsx(path, "データ",
            [[.. header.Select(name => (SourceTestCell)name)], .. rows]);

    private static MutationTestSheet KeySheet(string name, string key, string oldValue = "旧")
        => new()
        {
            Name = name,
            Cells =
            [
                new MutationTestCell("A1", key),
                new MutationTestCell("D5", oldValue),
                new MutationTestCell("F8", 0),
                new MutationTestCell("H10", oldValue),
                new MutationTestCell("B2", oldValue),
            ],
        };

    private static void CreateTarget(string path, string key, string oldValue = "旧")
        => TestMutationWorkbookFactory.Create(path, [KeySheet("月報", key, oldValue)]);

    private static CellMutationPreview Preview(SourceMappingBatchRequest request)
        => new SourceMappingPlanner().CreatePreview(request);

    private static CellMutationResult Execute(SourceMappingBatchRequest request)
    {
        var preview = Preview(request);
        Assert.True(preview.CanExecute,
            string.Join(" / ", preview.Blocks.Select(issue => $"{issue.Location}: {issue.Message}")));
        return new CellMutator().Execute(preview);
    }

    private static void AssertBlocked(SourceMappingBatchRequest request, string expectedFragment)
    {
        var preview = Preview(request);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(expectedFragment));
    }

    private static string Output(TempDir dir, string sourceName)
        => dir.File(sourceName + OutputSuffix + ".xlsx");

    private static string? Text(string path, string sheetName, string reference)
        => ReadCell(path, sheetName, reference).InlineString?.Text?.Text;

    private static Cell ReadCell(string path, string sheetName, string reference)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name?.Value == sheetName);
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return worksheetPart.Worksheet!.Descendants<Cell>()
            .Single(cell => string.Equals(
                cell.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
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

    private static string Sha256(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static (string Hash, long Length, DateTime LastWriteUtc) Snapshot(string path)
    {
        var info = new FileInfo(path);
        return (Sha256(path), info.Length, info.LastWriteTimeUtc);
    }
}
