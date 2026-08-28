using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2A: 複数 Workbook の同じセルを一括変更する。
/// 元ファイルは読み取り専用で、コピーの対象 WorksheetPart だけを書き換える。
/// </summary>
public sealed class CellMutationTests
{
    private const string OutputSuffix = "_変更済み";

    /// <summary>
    /// このテスト群のフィクスチャ自体が Excel の形式として正しいこと。
    /// Phase 2A は「元ファイルが検証を通ること」を前提にしているので、
    /// フィクスチャが汚れていると、狙いと違う理由で Block されてしまう。
    /// </summary>
    [Fact]
    public void Fixtures_AreValidatorClean()
    {
        using var dir = new TempDir();

        var rich = dir.File("全部入り.xlsx");
        CreateWorkbook(rich,
            new MutationTestSheet
            {
                Name = "月報",
                Cells = [Cell("A1", "項目"), Cell("A2", "架空A"), Cell("B1", 1), Cell("B2", "未確認")],
                AddChart = true,
                AddImage = true,
                AddTable = true,
                AddConditionalFormatting = true,
                AddProtection = true,
                Merges = ["F1:G1"],
                DataValidationSqref = "D1:D5",
                X14ValidationSqref = "E1:E5",
                HyperlinkReference = "A1",
                FormulaCell = "C1",
                RichTextCell = "C2",
                MetadataCell = "C3",
            },
            Sheet("参考", Cell("A1", "参考データ")));

        var pivot = dir.File("ピボット.xlsx");
        CreateWorkbook(pivot, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("A1", "項目"), Cell("A2", "架空"), Cell("B1", 1), Cell("B2", 2)],
            AddPivotTable = true,
        });

        var external = dir.File("外部参照.xlsx");
        TestMutationWorkbookFactory.Create(
            external, [Sheet("月報", Cell("B2", "未確認"))], addExternalLink: true);

        foreach (var path in new[] { rich, pivot, external })
        {
            var errors = DescribeValidationErrors(path);
            Assert.True(errors.Length == 0, errors);
        }
    }

    [Fact]
    public void Fixture_WithASchemaError_IsNotValidatorClean()
    {
        using var dir = new TempDir();
        var path = dir.File("壊れている.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("B2", "未確認")],
            AddSchemaError = true,
        });

        Assert.NotEqual(string.Empty, DescribeValidationErrors(path));
    }

    // ── A. 基本動作 ──────────────────────────────────────

    [Fact]
    public void Execute_TextValue_IsWrittenAsInlineString()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        var result = Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"));
        Assert.True(result.Success, result.Message);

        var output = Output(dir, "大阪");
        var cell = ReadCell(output, "月報", "B2");
        Assert.Equal(CellValues.InlineString, cell.DataType?.Value);
        Assert.Equal("確認済み", cell.InlineString?.Text?.Text);
    }

    [Fact]
    public void Execute_NumberValue_IsWrittenAsNumber()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", 100)));

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Number, "-1.5")).Success);

        var cell = ReadCell(Output(dir, "大阪"), "月報", "B2");
        Assert.Null(cell.DataType?.Value);
        Assert.Equal("-1.5", cell.CellValue?.InnerText);
    }

    [Fact]
    public void Execute_Blank_ClearsTheValueButKeepsTheCell()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認", StyleId: 1)),
            [new MutationTestStyle(NumberFormatId: 14)]);

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Blank)).Success);

        // 空欄にするだけなら表示形式によらず安全なので、日付書式でも実行できる。
        var cell = ReadCell(Output(dir, "大阪"), "月報", "B2");
        Assert.Null(cell.CellValue);
        Assert.Null(cell.InlineString);
        Assert.Null(cell.DataType?.Value);
        Assert.Equal(1U, cell.StyleIndex?.Value);
    }

    [Fact]
    public void Execute_MultipleWorkbooks_CreatesOneOutputEach()
    {
        using var dir = new TempDir();
        var first = dir.File("大阪.xlsx");
        var second = dir.File("京都.xlsx");
        CreateWorkbook(first, Sheet("月報", Cell("B2", "未確認")));
        CreateWorkbook(second, Sheet("月報", Cell("B2", "未確認")));

        var result = Execute(new CellMutationRequest
        {
            Targets = [new CellMutationTarget(first, "月報"), new CellMutationTarget(second, "月報")],
            CellReference = "B2",
            WriteKind = CellWriteKind.Text,
            TextValue = "確認済み",
        });

        Assert.True(result.Success, result.Message);
        Assert.Equal(2, result.ChangedCellCount);
        Assert.Equal("確認済み", ReadCell(Output(dir, "大阪"), "月報", "B2").InlineString?.Text?.Text);
        Assert.Equal("確認済み", ReadCell(Output(dir, "京都"), "月報", "B2").InlineString?.Text?.Text);
    }

    [Fact]
    public void Execute_MultipleSheetsInOneWorkbook_CreatesASingleOutput()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path,
            Sheet("1月", Cell("B2", "未確認")),
            Sheet("2月", Cell("B2", "未確認")),
            Sheet("3月", Cell("B2", "未確認")));

        var result = Execute(new CellMutationRequest
        {
            Targets =
            [
                new CellMutationTarget(path, "1月"),
                new CellMutationTarget(path, "2月"),
                new CellMutationTarget(path, "3月"),
            ],
            CellReference = "B2",
            WriteKind = CellWriteKind.Text,
            TextValue = "完了",
        });

        Assert.True(result.Success, result.Message);
        Assert.Equal(3, result.ChangedCellCount);
        Assert.Single(Directory.GetFiles(dir.Root, "*変更済み.xlsx"));

        var output = Output(dir, "大阪");
        foreach (var sheetName in new[] { "1月", "2月", "3月" })
        {
            Assert.Equal("完了", ReadCell(output, sheetName, "B2").InlineString?.Text?.Text);
        }
    }

    [Fact]
    public void Execute_KeepsTheExistingStyleIndex()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認", StyleId: 1)),
            [new MutationTestStyle(NumberFormatId: 49)]);

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み")).Success);

        Assert.Equal(1U, ReadCell(Output(dir, "大阪"), "月報", "B2").StyleIndex?.Value);
    }

    [Fact]
    public void Execute_SharedStringSource_DoesNotGrowTheSharedStringTable()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み")).Success);

        Assert.Equal(
            Entries(path)["xl/sharedStrings.xml"],
            Entries(Output(dir, "大阪"))["xl/sharedStrings.xml"]);
    }

    [Fact]
    public void Execute_UsesTheOutputSuffixNextToTheSource()
    {
        using var dir = new TempDir();
        var path = dir.File("売上.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        var result = Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み") with
        {
            OutputSuffix = "_確認",
        });

        Assert.True(result.Success, result.Message);
        Assert.Equal("売上_確認.xlsx", Assert.Single(result.OutputFileNames));
        Assert.True(File.Exists(dir.File("売上_確認.xlsx")));
    }

    [Fact]
    public void Preview_OutputAlreadyExists_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("売上.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));
        File.WriteAllText(dir.File("売上" + OutputSuffix + ".xlsx"), "架空の既存ファイル");

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "既にあります");
    }

    [Fact]
    public void Preview_AuditFileAlreadyExists_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("売上.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));
        File.WriteAllText(dir.File("売上" + OutputSuffix + ".xlsx.audit.json"), "{}");

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "既にあります");
    }

    // ── B. 対象セルの guard ──────────────────────────────

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("2B")]
    [InlineData("B")]
    [InlineData("12")]
    [InlineData("B2 C3")]
    [InlineData("#REF!")]
    public void Preview_InvalidCellAddress_IsBlocked(string reference)
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        AssertBlocked(Request(path, "月報", reference, CellWriteKind.Text, "確認済み"), "セル");
    }

    [Fact]
    public void Preview_RangeAddress_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        AssertBlocked(Request(path, "月報", "A1:B5", CellWriteKind.Text, "確認済み"), "範囲");
    }

    [Fact]
    public void Preview_SheetQualifiedAddress_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        AssertBlocked(Request(path, "月報", "月報!B2", CellWriteKind.Text, "確認済み"), "シート名");
    }

    [Fact]
    public void Execute_AddressAtTheExcelUpperBound_IsAccepted()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("XFD1048576", "未確認")));

        Assert.True(Execute(Request(path, "月報", "XFD1048576", CellWriteKind.Text, "確認済み")).Success);
        Assert.Equal("確認済み", ReadCell(Output(dir, "大阪"), "月報", "XFD1048576").InlineString?.Text?.Text);
    }

    [Theory]
    [InlineData("XFE1")]
    [InlineData("A1048577")]
    public void Preview_AddressBeyondExcelLimits_IsBlocked(string reference)
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        AssertBlocked(Request(path, "月報", reference, CellWriteKind.Text, "確認済み"), "範囲を超えています");
    }

    [Fact]
    public void Preview_MissingPhysicalCell_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        AssertBlocked(Request(path, "月報", "Z500", CellWriteKind.Text, "確認済み"), "存在しない");
    }

    [Fact]
    public void Preview_FormulaTargetCell_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("A1", "項目")],
            FormulaCell = "B2",
        });

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "数式");
    }

    [Fact]
    public void Preview_FormulaAnywhereInTheWorkbook_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path,
            Sheet("月報", Cell("B2", "未確認")),
            new MutationTestSheet { Name = "集計", Cells = [Cell("A1", "計")], FormulaCell = "B1" });

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "計算結果を保証できません");
    }

    [Theory]
    [InlineData("B2:C3")]
    [InlineData("A1:B2")]
    public void Preview_MergedTargetCell_IsBlocked(string merge)
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("A1", "項目"), Cell("B2", "未確認")],
            Merges = [merge],
        });

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "結合セル");
    }

    [Fact]
    public void Preview_ProtectedSheet_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("B2", "未確認")],
            AddProtection = true,
        });

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "保護");
    }

    [Fact]
    public void Preview_DataValidationOnTheTargetCell_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("B2", "未確認")],
            DataValidationSqref = "B1:B10",
        });

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "入力規則");
    }

    [Fact]
    public void Preview_X14DataValidationOnTheTargetCell_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("B2", "未確認")],
            X14ValidationSqref = "A1:C5",
        });

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "入力規則");
    }

    [Fact]
    public void Preview_DataValidationOnAnotherCell_DoesNotBlock()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("B2", "未確認")],
            DataValidationSqref = "D1:D10",
        });

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み")).Success);
    }

    [Fact]
    public void Preview_HyperlinkOnTheTargetCell_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("B2", "未確認")],
            HyperlinkReference = "B2",
        });

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "ハイパーリンク");
    }

    [Fact]
    public void Preview_RichTextTargetCell_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("A1", "項目")],
            RichTextCell = "B2",
        });

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "文字ごとに書式");
    }

    [Fact]
    public void Preview_CellWithValueMetadata_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("A1", "項目")],
            MetadataCell = "B2",
        });

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "特別なデータ");
    }

    [Fact]
    public void Preview_PivotTableOnTheSheet_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("A1", "項目"), Cell("A2", "架空"), Cell("B1", 1), Cell("B2", 2)],
            AddPivotTable = true,
        });

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Number, "5"), "ピボットテーブル");
    }

    [Fact]
    public void Preview_SourceWithSchemaErrors_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("B2", "未確認")],
            AddSchemaError = true,
        });

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "Excel の形式として問題");
    }

    [Fact]
    public void Preview_ExternalLink_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        TestMutationWorkbookFactory.Create(
            path, [Sheet("月報", Cell("B2", "未確認"))], addExternalLink: true);

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"), "外部参照");
    }

    [Theory]
    [InlineData(14U, null, CellWriteKind.Text)]
    [InlineData(14U, null, CellWriteKind.Number)]
    [InlineData(9U, null, CellWriteKind.Number)]
    [InlineData(9U, null, CellWriteKind.Text)]
    [InlineData(8U, null, CellWriteKind.Number)]
    [InlineData(164U, "yyyy\"年\"m\"月\"", CellWriteKind.Text)]
    [InlineData(165U, "#,##0\"円\"", CellWriteKind.Number)]
    [InlineData(49U, null, CellWriteKind.Number)]
    public void Preview_UnsafeNumberFormatForTheValueKind_IsBlocked(
        uint numberFormatId, string? formatCode, CellWriteKind kind)
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", 1, StyleId: 1)),
            [new MutationTestStyle(numberFormatId, formatCode)]);

        AssertBlocked(Request(path, "月報", "B2", kind, "5"), "表示形式");
    }

    [Theory]
    [InlineData(0U, CellWriteKind.Text)]
    [InlineData(49U, CellWriteKind.Text)]
    [InlineData(0U, CellWriteKind.Number)]
    [InlineData(1U, CellWriteKind.Number)]
    [InlineData(2U, CellWriteKind.Number)]
    [InlineData(3U, CellWriteKind.Number)]
    [InlineData(4U, CellWriteKind.Number)]
    public void Execute_SafeNumberFormatForTheValueKind_IsAccepted(uint numberFormatId, CellWriteKind kind)
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", 1, StyleId: 1)),
            [new MutationTestStyle(numberFormatId)]);

        Assert.True(Execute(Request(path, "月報", "B2", kind, "5")).Success);
    }

    [Theory]
    [InlineData("")]
    [InlineData("abc")]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("1,000")]
    public void Preview_InvalidNumberInput_IsBlocked(string input)
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", 1)));

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Number, input), "数値");
    }

    [Fact]
    public void Preview_EmptyTextValue_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        AssertBlocked(Request(path, "月報", "B2", CellWriteKind.Text, string.Empty), "新しい値");
    }

    [Theory]
    [InlineData("")]
    [InlineData("a/b")]
    public void Preview_InvalidOutputSuffix_IsBlocked(string suffix)
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        AssertBlocked(
            Request(path, "月報", "B2", CellWriteKind.Text, "確認済み") with { OutputSuffix = suffix },
            "出力ファイル名");
    }

    // ── C. 変更しない部分の保持 ──────────────────────────

    [Fact]
    public void Execute_ChangesOnlyTheTargetWorksheetPart()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path,
            new MutationTestSheet
            {
                Name = "月報",
                Cells = [Cell("A1", "項目"), Cell("B2", "未確認")],
                AddChart = true,
                AddImage = true,
                AddTable = true,
                AddConditionalFormatting = true,
                HyperlinkReference = "A1",
                DataValidationSqref = "D1:D5",
            },
            Sheet("参考", Cell("A1", "参考データ")));

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み")).Success);

        var before = Entries(path);
        var after = Entries(Output(dir, "大阪"));

        // 対象 WorksheetPart 以外は、展開後の内容が完全に一致すること。
        var changed = before.Keys.Union(after.Keys)
            .Where(name => !before.TryGetValue(name, out var left)
                || !after.TryGetValue(name, out var right)
                || left != right)
            .ToList();

        Assert.Equal(["xl/worksheets/sheet1.xml"], changed);
    }

    [Theory]
    [InlineData("xl/styles.xml")]
    [InlineData("xl/sharedStrings.xml")]
    [InlineData("xl/workbook.xml")]
    [InlineData("xl/theme/theme1.xml")]
    [InlineData("[Content_Types].xml")]
    [InlineData("_rels/.rels")]
    [InlineData("xl/_rels/workbook.xml.rels")]
    [InlineData("xl/worksheets/_rels/sheet1.xml.rels")]
    [InlineData("xl/worksheets/sheet2.xml")]
    public void Execute_LeavesTheseEntriesUnchanged(string entryName)
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path,
            new MutationTestSheet
            {
                Name = "月報",
                Cells = [Cell("A1", "項目"), Cell("B2", "未確認")],
                AddChart = true,
                AddImage = true,
                AddTable = true,
                HyperlinkReference = "A1",
            },
            Sheet("参考", Cell("A1", "参考データ")));

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み")).Success);

        var before = Entries(path);
        var after = Entries(Output(dir, "大阪"));

        Assert.True(before.ContainsKey(entryName), $"元ファイルに {entryName} がありません。");
        Assert.Equal(before[entryName], after[entryName]);
    }

    [Fact]
    public void Execute_LeavesChartDrawingImageAndTablePartsUnchanged()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("A1", "項目"), Cell("B2", "未確認")],
            AddChart = true,
            AddImage = true,
            AddTable = true,
        });

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み")).Success);

        var before = Entries(path);
        var after = Entries(Output(dir, "大阪"));

        var parts = before.Keys
            .Where(name => name.Contains("chart", StringComparison.OrdinalIgnoreCase)
                || name.Contains("drawing", StringComparison.OrdinalIgnoreCase)
                || name.Contains("media", StringComparison.OrdinalIgnoreCase)
                || name.Contains("table", StringComparison.OrdinalIgnoreCase))
            .ToList();

        Assert.NotEmpty(parts);
        Assert.All(parts, name => Assert.Equal(before[name], after[name]));
    }

    [Fact]
    public void Execute_LeavesOtherCellsInTheTargetSheetUnchanged()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報",
            Cell("A1", "項目"), Cell("A2", "架空A"), Cell("B1", "値"), Cell("B2", "未確認"), Cell("C3", 42)));

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み")).Success);

        var output = Output(dir, "大阪");
        Assert.Equal("0", ReadCell(output, "月報", "A1").CellValue?.InnerText);
        Assert.Equal(CellValues.SharedString, ReadCell(output, "月報", "A2").DataType?.Value);
        Assert.Equal("42", ReadCell(output, "月報", "C3").CellValue?.InnerText);
    }

    // ── D. 安全な実行の流れ ──────────────────────────────

    [Fact]
    public void Execute_LeavesEverySourceFileUnchanged()
    {
        using var dir = new TempDir();
        var first = dir.File("大阪.xlsx");
        var second = dir.File("京都.xlsx");
        CreateWorkbook(first, Sheet("月報", Cell("B2", "未確認")));
        CreateWorkbook(second, Sheet("月報", Cell("B2", "未確認")));

        var before = new[] { Snapshot(first), Snapshot(second) };

        Assert.True(Execute(new CellMutationRequest
        {
            Targets = [new CellMutationTarget(first, "月報"), new CellMutationTarget(second, "月報")],
            CellReference = "B2",
            WriteKind = CellWriteKind.Text,
            TextValue = "確認済み",
        }).Success);

        Assert.Equal(before, new[] { Snapshot(first), Snapshot(second) });
    }

    [Fact]
    public void Execute_SourceChangedAfterPreview_AbortsTheWholeBatch()
    {
        using var dir = new TempDir();
        var first = dir.File("大阪.xlsx");
        var second = dir.File("京都.xlsx");
        CreateWorkbook(first, Sheet("月報", Cell("B2", "未確認")));
        CreateWorkbook(second, Sheet("月報", Cell("B2", "未確認")));

        var preview = Preview(new CellMutationRequest
        {
            Targets = [new CellMutationTarget(first, "月報"), new CellMutationTarget(second, "月報")],
            CellReference = "B2",
            WriteKind = CellWriteKind.Text,
            TextValue = "確認済み",
        });

        Assert.True(preview.CanExecute);

        // プレビュー後に 2 つ目だけ差し替える。
        CreateWorkbook(second, Sheet("月報", Cell("B2", "別の値")));

        var result = new CellMutator().Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("プレビュー後に変更されました", result.Message);
        Assert.Empty(Directory.GetFiles(dir.Root, "*変更済み*"));
    }

    [Fact]
    public void Execute_OutputCreatedAfterPreview_AbortsTheWholeBatch()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        var preview = Preview(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"));
        Assert.True(preview.CanExecute);

        File.WriteAllText(dir.File("大阪" + OutputSuffix + ".xlsx"), "架空の既存ファイル");

        var result = new CellMutator().Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("上書きしません", result.Message);
        Assert.Equal("架空の既存ファイル", File.ReadAllText(dir.File("大阪" + OutputSuffix + ".xlsx")));
    }

    [Fact]
    public void Execute_OneTargetFails_LeavesNoOutputAtAll()
    {
        using var dir = new TempDir();
        var first = dir.File("大阪.xlsx");
        var second = dir.File("京都.xlsx");
        CreateWorkbook(first, Sheet("月報", Cell("B2", "未確認")));
        CreateWorkbook(second, Sheet("月報", Cell("B2", "未確認")));

        var preview = Preview(new CellMutationRequest
        {
            Targets = [new CellMutationTarget(first, "月報"), new CellMutationTarget(second, "月報")],
            CellReference = "B2",
            WriteKind = CellWriteKind.Text,
            TextValue = "確認済み",
        });

        // 2 つ目を壊してから実行する(1 つ目だけ作られないこと)。
        File.WriteAllText(second, "壊れたファイル");

        var result = new CellMutator().Execute(preview);

        Assert.False(result.Success);
        Assert.Empty(Directory.GetFiles(dir.Root, "*変更済み*"));
    }

    [Fact]
    public void Execute_LeavesNoTemporaryFilesBehind()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み")).Success);

        Assert.Empty(Directory.GetFiles(dir.Root, "~ebt-*"));
    }

    [Fact]
    public void Execute_FailedBatch_LeavesNoTemporaryFilesBehind()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        var preview = Preview(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"));
        File.WriteAllText(path, "壊れたファイル");

        Assert.False(new CellMutator().Execute(preview).Success);
        Assert.Empty(Directory.GetFiles(dir.Root, "~ebt-*"));
    }

    // ── D2. 取り消し(rollback)の結果を正確に伝える ────────

    [Fact]
    public void Execute_FailedMidBatch_SaysItRolledBackRatherThanNeverCreating()
    {
        using var dir = new TempDir();
        var first = dir.File("大阪.xlsx");
        var second = dir.File("京都.xlsx");
        CreateWorkbook(first, Sheet("月報", Cell("B2", "未確認")));
        CreateWorkbook(second, Sheet("月報", Cell("B2", "未確認")));

        var before = new[] { Snapshot(first), Snapshot(second) };

        var preview = Preview(new CellMutationRequest
        {
            Targets = [new CellMutationTarget(first, "月報"), new CellMutationTarget(second, "月報")],
            CellReference = "B2",
            WriteKind = CellWriteKind.Text,
            TextValue = "確認済み",
        });

        // 2 ファイル目の控えファイルの置き場所をフォルダーで塞ぐ。1 ファイル目を確定した
        // あとで失敗するので、確定済みのものまで取り消す必要がある。
        Directory.CreateDirectory(dir.File("京都" + OutputSuffix + ".xlsx.audit.json"));

        var result = new CellMutator().Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("取り消しました", result.Message);
        Assert.DoesNotContain("作成していません", result.Message);
        Assert.Contains("元のファイルは変更していません", result.Message);

        Assert.Empty(Directory.GetFiles(dir.Root, "*変更済み.xlsx"));
        Assert.Empty(Directory.GetFiles(dir.Root, "*.audit.json"));
        Assert.Empty(Directory.GetFiles(dir.Root, "~ebt-*"));
        Assert.Equal(before, new[] { Snapshot(first), Snapshot(second) });
    }

    [Fact]
    public void Execute_Cancelled_UsesTheSameRollbackWording()
    {
        using var dir = new TempDir();
        var first = dir.File("大阪.xlsx");
        var second = dir.File("京都.xlsx");
        CreateWorkbook(first, Sheet("月報", Cell("B2", "未確認")));
        CreateWorkbook(second, Sheet("月報", Cell("B2", "未確認")));

        var before = new[] { Snapshot(first), Snapshot(second) };

        var preview = Preview(new CellMutationRequest
        {
            Targets = [new CellMutationTarget(first, "月報"), new CellMutationTarget(second, "月報")],
            CellReference = "B2",
            WriteKind = CellWriteKind.Text,
            TextValue = "確認済み",
        });

        // 1 ファイル目を作り終えた時点で中止する。
        using var cancellation = new CancellationTokenSource();
        var progress = new Progress<CellMutationProgress>(_ => cancellation.Cancel());

        var result = new CellMutator().Execute(preview, progress, cancellation.Token);

        Assert.False(result.Success);
        Assert.Contains("中止しました", result.Message);
        Assert.Contains("取り消しました", result.Message);
        Assert.DoesNotContain("作成していません", result.Message);
        Assert.Empty(Directory.GetFiles(dir.Root, "*変更済み*"));
        Assert.Empty(Directory.GetFiles(dir.Root, "~ebt-*"));
        Assert.Equal(before, new[] { Snapshot(first), Snapshot(second) });
    }

    [Fact]
    public void Execute_FailureAfterAWorkbookWasCommitted_TakesTheCommittedFileBack()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        var before = Snapshot(path);
        var preview = Preview(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"));

        // 控えファイルの置き場所をフォルダーで塞ぐ。File.Exists は false のままなので
        // 事前確認は通り、Workbook を確定したあとの控えファイル確定で失敗する。
        Directory.CreateDirectory(dir.File("大阪" + OutputSuffix + ".xlsx.audit.json"));

        var result = new CellMutator().Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("取り消しました", result.Message);
        Assert.DoesNotContain("作成していません", result.Message);

        // いったん確定した Workbook も取り消されていること。
        Assert.False(File.Exists(Output(dir, "大阪")));
        Assert.Empty(Directory.GetFiles(dir.Root, "~ebt-*"));
        Assert.Equal(before, Snapshot(path));
    }

    [Fact]
    public void Execute_WhenTheCommittedWorkbookCannotBeDeleted_ReportsWhatRemains()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        var before = Snapshot(path);
        var preview = Preview(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"));

        Directory.CreateDirectory(dir.File("大阪" + OutputSuffix + ".xlsx.audit.json"));

        // 確定済み Workbook だけ消せない状況(ロックされている等)を再現する。
        var mutator = new CellMutator
        {
            FileDeleter = candidate => candidate.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? false
                : CellMutator.TryDeleteFile(candidate),
        };

        var result = mutator.Execute(preview);

        Assert.False(result.Success);
        Assert.True(File.Exists(Output(dir, "大阪")));

        // 消せていないのに「作成していません」と言い切らないこと。
        Assert.DoesNotContain("作成していません", result.Message);
        Assert.Contains("残っている可能性があります", result.Message);
        Assert.Contains("大阪" + OutputSuffix + ".xlsx", result.Message);

        // 表示するのはファイル名だけ。絶対パスは出さない。
        Assert.DoesNotContain(dir.Root, result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(":\\", result.Message, StringComparison.Ordinal);

        // 元ファイルについては引き続き断定してよい。
        Assert.Contains("元のファイルは変更していません", result.Message);
        Assert.Equal(before, Snapshot(path));
    }

    [Fact]
    public void Execute_WhenOnlyAnAuditFileCannotBeDeleted_ReportsThatFile()
    {
        using var dir = new TempDir();
        var first = dir.File("大阪.xlsx");
        var second = dir.File("京都.xlsx");
        CreateWorkbook(first, Sheet("月報", Cell("B2", "未確認")));
        CreateWorkbook(second, Sheet("月報", Cell("B2", "未確認")));

        var preview = Preview(new CellMutationRequest
        {
            Targets = [new CellMutationTarget(first, "月報"), new CellMutationTarget(second, "月報")],
            CellReference = "B2",
            WriteKind = CellWriteKind.Text,
            TextValue = "確認済み",
        });

        // 2 ファイル目の控えファイルの確定で失敗させる。
        // このとき 1 ファイル目は Workbook も控えも確定済みになっている。
        Directory.CreateDirectory(dir.File("京都" + OutputSuffix + ".xlsx.audit.json"));

        var mutator = new CellMutator
        {
            FileDeleter = candidate => candidate.EndsWith(".json", StringComparison.OrdinalIgnoreCase)
                ? false
                : CellMutator.TryDeleteFile(candidate),
        };

        var result = mutator.Execute(preview);

        Assert.False(result.Success);
        Assert.DoesNotContain("作成していません", result.Message);
        Assert.Contains("残っている可能性があります", result.Message);
        Assert.Contains("大阪" + OutputSuffix + ".xlsx.audit.json", result.Message);
        Assert.DoesNotContain(dir.Root, result.Message, StringComparison.OrdinalIgnoreCase);

        // Workbook 側は消せているので、残存として挙げない。
        Assert.False(File.Exists(Output(dir, "大阪")));
        Assert.False(File.Exists(Output(dir, "京都")));
    }

    [Fact]
    public void Execute_SuccessMessage_IsUnchangedByTheRollbackWording()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        var result = Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"));

        Assert.True(result.Success);
        Assert.Contains("元のファイルは変更していません", result.Message);
        Assert.DoesNotContain("取り消しました", result.Message);
        Assert.DoesNotContain("残っている可能性があります", result.Message);
    }

    [Fact]
    public void Execute_OutputPassesTheOpenXmlValidatorAndReopens()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("A1", "項目"), Cell("B2", "未確認")],
            AddChart = true,
            AddImage = true,
            AddTable = true,
            AddConditionalFormatting = true,
        });

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み")).Success);

        using var stream = new FileStream(Output(dir, "大阪"), FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
    }

    // ── 控えファイル(audit) ────────────────────────────

    [Fact]
    public void Execute_WritesAnAuditFileNextToTheOutput()
    {
        using var dir = new TempDir();
        var path = dir.File("売上.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み")).Success);

        var auditPath = dir.File("売上" + OutputSuffix + ".xlsx.audit.json");
        Assert.True(File.Exists(auditPath));

        using var json = JsonDocument.Parse(File.ReadAllText(auditPath));
        var root = json.RootElement;

        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("売上.xlsx", root.GetProperty("sourceFileName").GetString());
        Assert.Equal("売上" + OutputSuffix + ".xlsx", root.GetProperty("outputFileName").GetString());
        Assert.Equal("set-cell-value", root.GetProperty("operation").GetString());
        Assert.Equal(Sha256(path), root.GetProperty("sourceSha256").GetString());
        Assert.Equal(Sha256(Output(dir, "売上")), root.GetProperty("outputSha256").GetString());
        Assert.True(DateTimeOffset.TryParse(root.GetProperty("createdAt").GetString(), out _));

        var change = Assert.Single(root.GetProperty("changes").EnumerateArray());
        Assert.Equal("月報", change.GetProperty("sheetName").GetString());
        Assert.Equal("B2", change.GetProperty("cell").GetString());
        Assert.Equal("未確認", change.GetProperty("oldValue").GetString());
        Assert.Equal("text", change.GetProperty("oldType").GetString());
        Assert.Equal("確認済み", change.GetProperty("newValue").GetString());
        Assert.Equal("text", change.GetProperty("newType").GetString());
    }

    [Fact]
    public void Execute_AuditFile_ContainsNoAbsolutePath()
    {
        using var dir = new TempDir();
        var path = dir.File("売上.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み")).Success);

        var text = File.ReadAllText(dir.File("売上" + OutputSuffix + ".xlsx.audit.json"));
        Assert.DoesNotContain(dir.Root, text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(":\\", text, StringComparison.Ordinal);
        Assert.DoesNotContain("/", text.Replace("\\/", string.Empty, StringComparison.Ordinal));
    }

    [Fact]
    public void Execute_AuditFile_RecordsBlankAsTheNewType()
    {
        using var dir = new TempDir();
        var path = dir.File("売上.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", 12.5)));

        Assert.True(Execute(Request(path, "月報", "B2", CellWriteKind.Blank)).Success);

        using var json = JsonDocument.Parse(
            File.ReadAllText(dir.File("売上" + OutputSuffix + ".xlsx.audit.json")));
        var change = Assert.Single(json.RootElement.GetProperty("changes").EnumerateArray());

        Assert.Equal("12.5", change.GetProperty("oldValue").GetString());
        Assert.Equal("number", change.GetProperty("oldType").GetString());
        Assert.Equal("blank", change.GetProperty("newType").GetString());
    }

    [Fact]
    public void Execute_TwoFiles_ProducesBothOutputsAndBothAuditFiles()
    {
        using var dir = new TempDir();
        var first = dir.File("大阪.xlsx");
        var second = dir.File("京都.xlsx");
        CreateWorkbook(first, Sheet("月報", Cell("B2", "未確認")));
        CreateWorkbook(second, Sheet("月報", Cell("B2", "未確認")));

        Assert.True(Execute(new CellMutationRequest
        {
            Targets = [new CellMutationTarget(first, "月報"), new CellMutationTarget(second, "月報")],
            CellReference = "B2",
            WriteKind = CellWriteKind.Text,
            TextValue = "確認済み",
        }).Success);

        Assert.Equal(2, Directory.GetFiles(dir.Root, "*変更済み.xlsx").Length);
        Assert.Equal(2, Directory.GetFiles(dir.Root, "*.audit.json").Length);
    }

    // ── No-op ────────────────────────────────────────────

    [Fact]
    public void Preview_EveryTargetIsNoOp_CannotExecute()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "確認済み")));

        var preview = Preview(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み"));

        Assert.False(preview.CanExecute);
        Assert.Equal(1, preview.NoOpCount);
        Assert.All(preview.Targets, target => Assert.True(target.IsNoOp));
        Assert.Empty(Directory.GetFiles(dir.Root, "*変更済み*"));
    }

    [Fact]
    public void Execute_PartiallyNoOp_ChangesOnlyWhatDiffers()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path,
            Sheet("1月", Cell("B2", "確認済み")),
            Sheet("2月", Cell("B2", "未確認")));

        var request = new CellMutationRequest
        {
            Targets = [new CellMutationTarget(path, "1月"), new CellMutationTarget(path, "2月")],
            CellReference = "B2",
            WriteKind = CellWriteKind.Text,
            TextValue = "確認済み",
        };

        var preview = Preview(request);
        Assert.True(preview.CanExecute);
        Assert.Equal(1, preview.NoOpCount);
        Assert.Equal(1, preview.ChangeCount);

        Assert.True(new CellMutator().Execute(preview).Success);

        var output = Output(dir, "大阪");

        // 変更不要だったシートは共有文字列のまま、変更したシートだけ InlineString になる。
        Assert.Equal(CellValues.SharedString, ReadCell(output, "1月", "B2").DataType?.Value);
        Assert.Equal(CellValues.InlineString, ReadCell(output, "2月", "B2").DataType?.Value);
    }

    [Fact]
    public void Preview_NoOpForNumberAndBlank_IsDetected()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", 42), Cell("C2", null, StyleId: null)));

        Assert.True(Assert.Single(
            Preview(Request(path, "月報", "B2", CellWriteKind.Number, "42")).Targets).IsNoOp);
        Assert.True(Assert.Single(
            Preview(Request(path, "月報", "C2", CellWriteKind.Blank)).Targets).IsNoOp);
    }

    [Fact]
    public void Preview_DuplicateSheetSelection_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("大阪.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        AssertBlocked(
            new CellMutationRequest
            {
                Targets = [new CellMutationTarget(path, "月報"), new CellMutationTarget(path, "月報")],
                CellReference = "B2",
                WriteKind = CellWriteKind.Text,
                TextValue = "確認済み",
            },
            "複数回選択");
    }

    [Fact]
    public void Preview_ShowsCurrentAndNewValueAndOutputName()
    {
        using var dir = new TempDir();
        var path = dir.File("売上.xlsx");
        CreateWorkbook(path, Sheet("月報", Cell("B2", "未確認")));

        var target = Assert.Single(
            Preview(Request(path, "月報", "B2", CellWriteKind.Text, "確認済み")).Targets);

        Assert.Equal("売上.xlsx", target.FileName);
        Assert.Equal("月報", target.SheetName);
        Assert.Equal("B2", target.CellReference);
        Assert.Equal("未確認", target.CurrentValueDisplay);
        Assert.Equal("確認済み", target.NewValueDisplay);
        Assert.Equal("売上" + OutputSuffix + ".xlsx", target.OutputFileName);
        Assert.False(target.IsBlocked);
    }

    [Fact]
    public void Preview_OneBlockedSheet_BlocksTheWholeBatch()
    {
        using var dir = new TempDir();
        var first = dir.File("大阪.xlsx");
        var second = dir.File("京都.xlsx");
        CreateWorkbook(first, Sheet("月報", Cell("B2", "未確認")));
        CreateWorkbook(second, new MutationTestSheet
        {
            Name = "月報",
            Cells = [Cell("A1", "項目")],
            FormulaCell = "B2",
        });

        var preview = Preview(new CellMutationRequest
        {
            Targets = [new CellMutationTarget(first, "月報"), new CellMutationTarget(second, "月報")],
            CellReference = "B2",
            WriteKind = CellWriteKind.Text,
            TextValue = "確認済み",
        });

        Assert.False(preview.CanExecute);
        Assert.Empty(Directory.GetFiles(dir.Root, "*変更済み*"));
    }

    // ── ヘルパー ──────────────────────────────────────────

    private static MutationTestSheet Sheet(string name, params MutationTestCell[] cells)
        => new() { Name = name, Cells = cells };

    private static MutationTestCell Cell(string reference, object? value, int? StyleId = null)
        => new(reference, value, StyleId);

    private static void CreateWorkbook(
        string path,
        params MutationTestSheet[] sheets)
        => TestMutationWorkbookFactory.Create(path, sheets);

    private static void CreateWorkbook(
        string path,
        MutationTestSheet sheet,
        IReadOnlyList<MutationTestStyle> styles)
        => TestMutationWorkbookFactory.Create(path, [sheet], styles);

    private static CellMutationRequest Request(
        string path, string sheet, string reference, CellWriteKind kind, string? value = null) => new()
        {
            Targets = [new CellMutationTarget(path, sheet)],
            CellReference = reference,
            WriteKind = kind,
            TextValue = kind == CellWriteKind.Text ? value : null,
            NumberText = kind == CellWriteKind.Number ? value : null,
        };

    private static CellMutationPreview Preview(CellMutationRequest request)
        => new CellMutationPlanner().CreatePreview(request);

    private static CellMutationResult Execute(CellMutationRequest request)
    {
        var preview = Preview(request);
        Assert.True(preview.CanExecute,
            string.Join(" / ", preview.Blocks.Select(issue => $"{issue.Location}: {issue.Message}")));
        return new CellMutator().Execute(preview);
    }

    private static void AssertBlocked(CellMutationRequest request, string expectedFragment)
    {
        var preview = Preview(request);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(expectedFragment));
    }

    private static string Output(TempDir dir, string sourceName)
        => dir.File(sourceName + OutputSuffix + ".xlsx");

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

    /// <summary>ZIP entry 名 → 展開後の内容のハッシュ。圧縮結果ではなく中身を比べる。</summary>
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

    /// <summary>検証エラーを読める形にまとめる(問題なしなら空文字)。</summary>
    private static string DescribeValidationErrors(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);

        var errors = new OpenXmlValidator().Validate(document).ToList();
        return errors.Count == 0
            ? string.Empty
            : $"{Path.GetFileName(path)}: " + string.Join(
                " / ", errors.Take(5).Select(error => $"{error.Path?.XPath} {error.Description}"));
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
