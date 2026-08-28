using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Aggregation;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 1B.2B3A: 数式を使わない標準の条件付き書式
/// (duplicateValues / uniqueValues / top10 / aboveAverage)だけを移植する。
/// それ以外は黙って落とさず、シート全体を Block する。
/// </summary>
public sealed class SheetAggregationConditionalFormattingTests
{
    private static readonly TestDifferentialFormat RedFill = new() { FillArgb = "FFFFC7CE" };

    // ── 対応するルールの移植 ──────────────────────────────

    [Theory]
    [InlineData("duplicateValues")]
    [InlineData("uniqueValues")]
    public void Aggregate_DuplicateOrUniqueRule_IsKept(string type)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule { Type = type });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var rule = SingleRule(output, "表");
        Assert.Equal(type, rule.Type?.InnerText);
        Assert.Equal(1, rule.Priority?.Value);
        Assert.Equal(0U, rule.FormatId?.Value);
    }

    [Theory]
    [InlineData(10U, null, null)]
    [InlineData(5U, null, true)]
    [InlineData(20U, true, null)]
    [InlineData(0U, true, null)]
    [InlineData(100U, true, true)]
    [InlineData(1U, null, null)]
    [InlineData(1000U, null, null)]
    public void Aggregate_Top10Rule_KeepsRankPercentAndBottom(uint rank, bool? percent, bool? bottom)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule
        {
            Type = "top10",
            Rank = rank,
            Percent = percent,
            Bottom = bottom,
        });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var rule = SingleRule(output, "表");
        Assert.Equal("top10", rule.Type?.InnerText);
        Assert.Equal(rank, rule.Rank?.Value);
        Assert.Equal(percent, rule.Percent?.Value);
        Assert.Equal(bottom, rule.Bottom?.Value);
    }

    [Theory]
    [InlineData(null, null, null)]
    [InlineData(false, null, null)]
    [InlineData(true, null, null)]
    [InlineData(null, true, null)]
    [InlineData(null, null, 1)]
    [InlineData(false, null, 2)]
    public void Aggregate_AboveAverageRule_KeepsItsAttributes(
        bool? aboveAverage, bool? equalAverage, int? stdDev)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule
        {
            Type = "aboveAverage",
            AboveAverage = aboveAverage,
            EqualAverage = equalAverage,
            StdDev = stdDev,
        });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var rule = SingleRule(output, "表");
        Assert.Equal("aboveAverage", rule.Type?.InnerText);
        Assert.Equal(aboveAverage, rule.AboveAverage?.Value);
        Assert.Equal(equalAverage, rule.EqualAverage?.Value);
        Assert.Equal(stdDev, rule.StdDev?.Value);
    }

    [Fact]
    public void Aggregate_RuleWithoutOptionalAttributes_DoesNotAddDefaults()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule { Type = "aboveAverage" });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        // 元ファイルに無い属性を出力側で足さない(足すと Excel での見え方が変わりうる)。
        var rule = SingleRule(output, "表");
        Assert.Null(rule.AboveAverage);
        Assert.Null(rule.EqualAverage);
        Assert.Null(rule.StdDev);
        Assert.Null(rule.StopIfTrue);
        Assert.Null(rule.Rank);
        Assert.Null(rule.Percent);
        Assert.Null(rule.Bottom);
        Assert.Null(rule.Operator);
        Assert.Null(rule.Text);
        Assert.Null(rule.TimePeriod);
    }

    // ── 適用範囲(sqref)──────────────────────────────────

    [Theory]
    [InlineData("A1:A10")]
    [InlineData("B5")]
    [InlineData("A1:A10 C1:C10 E5")]
    [InlineData("$A$1:$B$20")]
    public void Aggregate_Sqref_IsKeptExactly(string sqref)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormatting
        {
            Sqref = sqref,
            Rules = [new TestConditionalFormattingRule { Type = "duplicateValues" }],
        });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        Assert.Equal(sqref, SingleFormatting(output, "表").SequenceOfReferences?.InnerText);
    }

    [Theory]
    [InlineData("")]
    [InlineData("#REF!")]
    [InlineData("A1:#REF!")]
    [InlineData("ぜんぶ")]
    [InlineData("A1:XFE10")]
    [InlineData("A1:A1048577")]
    public void Preview_BrokenSqref_IsBlocked(string sqref)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormatting
        {
            Sqref = sqref,
            Rules = [new TestConditionalFormattingRule { Type = "duplicateValues" }],
        });

        AssertBlocked(path, "適用範囲");
    }

    // ── 優先順位(priority)──────────────────────────────

    [Fact]
    public void Aggregate_Priorities_AreKeptAsIsNotRenumbered()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path,
            new TestConditionalFormatting
            {
                Sqref = "A1:A10",
                Rules =
                [
                    new TestConditionalFormattingRule { Type = "duplicateValues", Priority = 7 },
                    new TestConditionalFormattingRule { Type = "uniqueValues", Priority = 3 },
                ],
            },
            new TestConditionalFormatting
            {
                Sqref = "C1:C10",
                Rules = [new TestConditionalFormattingRule { Type = "top10", Priority = 12, Rank = 5U }],
            });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        // 1,2,3 へ振り直すと stopIfTrue の評価順が変わるので、元の値をそのまま使う。
        Assert.Equal(
            new[] { 7, 3, 12 },
            Formattings(output, "表")
                .SelectMany(item => item.Elements<ConditionalFormattingRule>())
                .Select(rule => rule.Priority!.Value)
                .ToArray());
    }

    [Fact]
    public void Preview_DuplicatePriorityAcrossRanges_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path,
            new TestConditionalFormatting
            {
                Sqref = "A1:A10",
                Rules = [new TestConditionalFormattingRule { Type = "duplicateValues", Priority = 2 }],
            },
            new TestConditionalFormatting
            {
                Sqref = "C1:C10",
                Rules = [new TestConditionalFormattingRule { Type = "uniqueValues", Priority = 2 }],
            });

        AssertBlocked(path, "優先順位");
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(null)]
    public void Preview_InvalidPriority_IsBlocked(int? priority)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule
        {
            Type = "duplicateValues",
            Priority = priority,
        });

        AssertBlocked(path, "優先順位");
    }

    // ── stopIfTrue ──────────────────────────────────────

    [Fact]
    public void Aggregate_StopIfTrue_IsKept()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path,
            new TestConditionalFormatting
            {
                Sqref = "A1:A10",
                Rules =
                [
                    new TestConditionalFormattingRule
                    {
                        Type = "duplicateValues", Priority = 1, StopIfTrue = true,
                    },
                    new TestConditionalFormattingRule { Type = "uniqueValues", Priority = 2 },
                ],
            });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var rules = SingleFormatting(output, "表").Elements<ConditionalFormattingRule>().ToList();
        Assert.True(rules[0].StopIfTrue!.Value);
        Assert.Null(rules[1].StopIfTrue);
    }

    // ── 対応していないルールの種類 ──────────────────────

    [Theory]
    [InlineData("expression", "数式")]
    [InlineData("cellIs", "数式")]
    [InlineData("colorScale", "カラースケール")]
    [InlineData("dataBar", "データバー")]
    [InlineData("iconSet", "アイコンセット")]
    [InlineData("containsText", "条件付き書式")]
    [InlineData("notContainsText", "条件付き書式")]
    [InlineData("beginsWith", "条件付き書式")]
    [InlineData("endsWith", "条件付き書式")]
    [InlineData("containsBlanks", "条件付き書式")]
    [InlineData("notContainsBlanks", "条件付き書式")]
    [InlineData("containsErrors", "条件付き書式")]
    [InlineData("notContainsErrors", "条件付き書式")]
    [InlineData("timePeriod", "条件付き書式")]
    [InlineData("架空の種類", "条件付き書式")]
    public void Preview_UnsupportedRuleType_IsBlocked(string type, string expectedFragment)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule { Type = type });

        AssertBlocked(path, expectedFragment);
    }

    [Fact]
    public void Preview_SupportedTypeWithAFormula_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");

        // 対応 type でも数式を持つなら、数式を無視してコピーはしない。
        CreateWorkbook(path, new TestConditionalFormattingRule
        {
            Type = "duplicateValues",
            Formula = "A1>0",
        });

        AssertBlocked(path, "数式");
    }

    [Theory]
    [InlineData("greaterThan", null, null)]
    [InlineData(null, "架空", null)]
    [InlineData(null, null, "today")]
    public void Preview_SupportedTypeWithForeignAttributes_IsBlocked(
        string? op, string? text, string? timePeriod)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule
        {
            Type = "duplicateValues",
            Operator = op,
            Text = text,
            TimePeriod = timePeriod,
        });

        AssertBlocked(path, "構造");
    }

    [Fact]
    public void Preview_RuleWithUnknownAttribute_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule
        {
            Type = "duplicateValues",
            AddUnknownAttribute = true,
        });

        AssertBlocked(path, "対応していない設定");
    }

    [Fact]
    public void Preview_ConditionalFormattingWithUnknownAttribute_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormatting
        {
            AddUnknownAttribute = true,
            Rules = [new TestConditionalFormattingRule { Type = "duplicateValues" }],
        });

        AssertBlocked(path, "対応していない設定");
    }

    [Fact]
    public void Preview_ConditionalFormattingWithUnknownChild_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormatting
        {
            AddUnknownChild = true,
            Rules = [new TestConditionalFormattingRule { Type = "duplicateValues" }],
        });

        AssertBlocked(path, "対応していない内容");
    }

    [Fact]
    public void Preview_PivotConditionalFormatting_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormatting
        {
            Pivot = true,
            Rules = [new TestConditionalFormattingRule { Type = "duplicateValues" }],
        });

        AssertBlocked(path, "ピボットテーブル");
    }

    [Fact]
    public void Preview_ConditionalFormattingWithoutRules_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormatting { Rules = [] });

        AssertBlocked(path, "ルールがありません");
    }

    // ── 種類ごとの属性の整合 ────────────────────────────

    [Fact]
    public void Preview_Top10WithoutRank_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule { Type = "top10" });

        AssertBlocked(path, "件数が指定されていません");
    }

    [Theory]
    [InlineData(0U, null)]
    [InlineData(1001U, null)]
    [InlineData(101U, true)]
    public void Preview_Top10WithOutOfRangeRank_IsBlocked(uint rank, bool? percent)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule
        {
            Type = "top10",
            Rank = rank,
            Percent = percent,
        });

        AssertBlocked(path, "有効な範囲");
    }

    [Theory]
    [InlineData("duplicateValues")]
    [InlineData("uniqueValues")]
    [InlineData("aboveAverage")]
    public void Preview_RankAttributesOnANonTop10Rule_AreBlocked(string type)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule { Type = type, Rank = 5U });

        AssertBlocked(path, "構造");
    }

    [Theory]
    [InlineData("duplicateValues")]
    [InlineData("top10")]
    public void Preview_AverageAttributesOnANonAboveAverageRule_AreBlocked(string type)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule
        {
            Type = type,
            Rank = type == "top10" ? 5U : null,
            AboveAverage = true,
        });

        AssertBlocked(path, "構造");
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Preview_AboveAverageWithANonPositiveStdDev_IsBlocked(int stdDev)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule
        {
            Type = "aboveAverage",
            StdDev = stdDev,
        });

        AssertBlocked(path, "平均条件");
    }

    [Fact]
    public void Preview_EqualAverageCombinedWithStdDev_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule
        {
            Type = "aboveAverage",
            EqualAverage = true,
            StdDev = 1,
        });

        AssertBlocked(path, "平均条件");
    }

    // ── 書式(dxf)────────────────────────────────────────

    [Fact]
    public void Aggregate_DxfContents_AreCopied()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path,
            [new TestDifferentialFormat
            {
                Bold = true,
                FontArgb = "FF9C0006",
                NumberFormatCode = "0.0%",
                FillArgb = "FFFFC7CE",
                ThinBorder = true,
            }],
            new TestConditionalFormattingRule { Type = "duplicateValues" });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var dxf = Assert.Single(Dxfs(output));
        Assert.True(dxf.Font?.GetFirstChild<Bold>() is not null);
        Assert.Equal("FF9C0006", dxf.Font?.GetFirstChild<Color>()?.Rgb?.Value);
        Assert.Equal("0.0%", dxf.NumberingFormat?.FormatCode?.Value);
        Assert.Equal("FFFFC7CE", dxf.Fill?.PatternFill?.BackgroundColor?.Rgb?.Value);
        Assert.NotNull(dxf.Border);
    }

    [Fact]
    public void Aggregate_DxfAlignmentAndProtection_AreCopied()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path,
            [new TestDifferentialFormat { CenterAlignment = true, Unlocked = true }],
            new TestConditionalFormattingRule { Type = "duplicateValues" });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var dxf = Assert.Single(Dxfs(output));
        Assert.Equal(HorizontalAlignmentValues.Center, dxf.Alignment?.Horizontal?.Value);
        Assert.False(dxf.Protection?.Locked?.Value);
    }

    [Fact]
    public void Preview_DxfWithAThemeColor_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path,
            [new TestDifferentialFormat { FontThemeColor = 4U }],
            new TestConditionalFormattingRule { Type = "duplicateValues" });

        // 出力ブックのテーマ次第で色が変わるため、変換もコピーもしない。
        AssertBlocked(path, "テーマの色");
    }

    [Theory]
    [InlineData(14U)]
    [InlineData(180U)]
    public void Preview_DxfNumberFormatWithoutAFormatCode_IsBlocked(uint numberFormatId)
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path,
            [new TestDifferentialFormat { NumberFormatIdWithoutCode = numberFormatId }],
            new TestConditionalFormattingRule { Type = "duplicateValues" });

        AssertBlocked(path, "表示形式");
    }

    [Fact]
    public void Preview_DxfWithUnknownAttribute_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path,
            [new TestDifferentialFormat { Bold = true, AddUnknownAttribute = true }],
            new TestConditionalFormattingRule { Type = "duplicateValues" });

        AssertBlocked(path, "対応していない書式設定");
    }

    [Fact]
    public void Preview_MissingDxfId_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule
        {
            Type = "duplicateValues",
            DxfId = null,
        });

        AssertBlocked(path, "書式の指定がありません");
    }

    [Fact]
    public void Preview_DxfIdOutOfRange_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule
        {
            Type = "duplicateValues",
            DxfId = 5U,
        });

        AssertBlocked(path, "書式情報");
    }

    // ── dxf の対応付け(CellStyle とは別管理)────────────

    [Fact]
    public void Aggregate_WithoutConditionalFormatting_DoesNotCreateDxfs()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        TestSheetWorkbookFactory.Create(path,
            [new TestAggregationSheetSpec { Name = "表", Rows = [["A", 1]] }]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        Assert.Null(Stylesheet(output).DifferentialFormats);
    }

    [Fact]
    public void Aggregate_UnreferencedDxfs_AreNotCopied()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path,
            [
                new TestDifferentialFormat { Bold = true },
                new TestDifferentialFormat { FillArgb = "FF00FF00" },
                new TestDifferentialFormat { ThinBorder = true },
            ],
            new TestConditionalFormattingRule { Type = "duplicateValues", DxfId = 1U });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        var dxf = Assert.Single(Dxfs(output));
        Assert.Equal("FF00FF00", dxf.Fill?.PatternFill?.BackgroundColor?.Rgb?.Value);
        Assert.Equal(0U, SingleRule(output, "表").FormatId?.Value);
    }

    [Fact]
    public void Aggregate_SameDxfUsedTwiceInAWorkbook_IsStoredOnce()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path,
            [RedFill],
            new TestConditionalFormatting
            {
                Sqref = "A1:A10",
                Rules =
                [
                    new TestConditionalFormattingRule { Type = "duplicateValues", Priority = 1 },
                    new TestConditionalFormattingRule { Type = "uniqueValues", Priority = 2 },
                ],
            });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        Assert.Single(Dxfs(output));
        Assert.All(
            SingleFormatting(output, "表").Elements<ConditionalFormattingRule>(),
            rule => Assert.Equal(0U, rule.FormatId?.Value));
    }

    [Fact]
    public void Aggregate_SameDxfIdInDifferentWorkbooks_DoesNotCollide()
    {
        using var dir = new TempDir();
        var first = dir.File("A.xlsx");
        var second = dir.File("B.xlsx");

        // どちらも dxfId=0 を参照するが、書式の中身は別物。
        CreateWorkbook(first,
            [new TestDifferentialFormat { FillArgb = "FFFF0000" }],
            new TestConditionalFormattingRule { Type = "duplicateValues" });
        CreateWorkbook(second,
            [new TestDifferentialFormat { Bold = true }],
            new TestConditionalFormattingRule { Type = "uniqueValues" });

        var preview = CreatePreview((first, "表"), (second, "表"));
        Assert.True(preview.CanExecute);

        var output = dir.File("out.xlsx");
        Assert.True(new SheetAggregator().Execute(preview, output).Success);

        var dxfs = Dxfs(output);
        Assert.Equal(2, dxfs.Count);

        var firstDxfId = SingleRule(output, preview.Sheets[0].OutputSheetName).FormatId!.Value;
        var secondDxfId = SingleRule(output, preview.Sheets[1].OutputSheetName).FormatId!.Value;
        Assert.NotEqual(firstDxfId, secondDxfId);

        Assert.Equal("FFFF0000", dxfs[(int)firstDxfId].Fill?.PatternFill?.BackgroundColor?.Rgb?.Value);
        Assert.NotNull(dxfs[(int)secondDxfId].Font?.GetFirstChild<Bold>());
    }

    [Fact]
    public void Aggregate_ConditionalFormatting_DoesNotShiftCellStyleIndexes()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");

        TestSheetWorkbookFactory.Create(path,
            [
                new TestAggregationSheetSpec
                {
                    Name = "表",
                    Rows = [[new Styled("太字", 0), new Styled("塗り", 1)]],
                    ConditionalFormattings =
                    [
                        new TestConditionalFormatting
                        {
                            Sqref = "A1:B1",
                            Rules = [new TestConditionalFormattingRule { Type = "duplicateValues" }],
                        },
                    ],
                },
            ],
            styles: [new TestStyle { Bold = true }, new TestStyle { FillArgb = "FF00B0F0" }],
            differentialFormats: [RedFill]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        // セル書式と条件付き書式の書式は別の索引体系。片方を足しても他方はずれない。
        var stylesheet = Stylesheet(output);
        var cellFormats = stylesheet.CellFormats!.Elements<CellFormat>().ToList();
        var worksheet = Worksheet(output, "表");
        var cells = worksheet.Descendants<Cell>().ToList();

        var boldFormat = cellFormats[(int)cells[0].StyleIndex!.Value];
        var fillFormat = cellFormats[(int)cells[1].StyleIndex!.Value];

        Assert.NotNull(stylesheet.Fonts!.Elements<Font>()
            .ElementAt((int)boldFormat.FontId!.Value).GetFirstChild<Bold>());
        Assert.Equal(
            "FF00B0F0",
            stylesheet.Fills!.Elements<Fill>()
                .ElementAt((int)fillFormat.FillId!.Value)
                .PatternFill?.ForegroundColor?.Rgb?.Value);
    }

    [Fact]
    public void Aggregate_CustomNumberFormats_DoNotCollideBetweenCellsAndDxfs()
    {
        using var dir = new TempDir();
        var first = dir.File("A.xlsx");
        var second = dir.File("B.xlsx");

        // どちらも元ブックでは numFmtId=164。片方はセル書式、片方は条件付き書式の書式。
        TestSheetWorkbookFactory.Create(first,
            [new TestAggregationSheetSpec { Name = "表", Rows = [[new Styled(1.5, 0)]] }],
            styles: [new TestStyle { NumberFormatCode = "0.00\"円\"" }]);

        CreateWorkbook(second,
            [new TestDifferentialFormat { NumberFormatCode = "0.0%" }],
            new TestConditionalFormattingRule { Type = "duplicateValues" });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (first, "表"), (second, "表")).Success);

        var stylesheet = Stylesheet(output);
        var numberFormats = stylesheet.NumberingFormats!.Elements<NumberingFormat>().ToList();

        Assert.Equal(2, numberFormats.Count);
        Assert.Equal(2, numberFormats.Select(format => format.NumberFormatId!.Value).Distinct().Count());

        var dxfNumberFormat = Assert.Single(Dxfs(output)).NumberingFormat!;
        Assert.Equal("0.0%", dxfNumberFormat.FormatCode?.Value);

        // ID からブックの定義を引いても、dxf が持つ formatCode と一致すること。
        var registered = numberFormats
            .Single(format => format.NumberFormatId!.Value == dxfNumberFormat.NumberFormatId!.Value);
        Assert.Equal("0.0%", registered.FormatCode?.Value);
    }

    // ── 新しい形式(x14)の条件付き書式 ────────────────

    [Fact]
    public void Preview_X14ConditionalFormatting_IsBlocked()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");

        TestSheetWorkbookFactory.Create(path,
            [
                new TestAggregationSheetSpec
                {
                    Name = "表",
                    Rows = [["A", 1]],
                    ConditionalFormattings =
                    [
                        new TestConditionalFormatting
                        {
                            Rules = [new TestConditionalFormattingRule { Type = "duplicateValues" }],
                        },
                    ],
                    AddX14ConditionalFormatting = true,
                },
            ],
            differentialFormats: [RedFill]);

        // 標準形式だけ写して x14 側を黙って落とすことはしない。
        AssertBlocked(path, "新しい形式の条件付き書式");
    }

    [Fact]
    public void Analyze_X14ConditionalFormatting_IsReportedAsConditionalFormatting()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");

        TestSheetWorkbookFactory.Create(path,
            [
                new TestAggregationSheetSpec
                {
                    Name = "表",
                    Rows = [["A", 1]],
                    AddX14ConditionalFormatting = true,
                },
            ]);

        var result = WorkbookAnalyzer.Analyze(path);

        Assert.Single(result.Findings, finding => finding.Type == FindingType.ConditionalFormatting);
    }

    // ── シート単位の Block と出力の安全性 ────────────────

    [Fact]
    public void Preview_OneUnsupportedRule_BlocksTheWholeSheet()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path,
            new TestConditionalFormatting
            {
                Sqref = "A1:A10",
                Rules = [new TestConditionalFormattingRule { Type = "duplicateValues", Priority = 1 }],
            },
            new TestConditionalFormatting
            {
                Sqref = "C1:C10",
                Rules = [new TestConditionalFormattingRule { Type = "cellIs", Priority = 2, Formula = "1" }],
            });

        // 対応分だけ部分的にコピーせず、シートごと Block する。
        var preview = CreatePreview((path, "表"));
        Assert.False(preview.CanExecute);
        Assert.All(preview.Sheets, sheet => Assert.True(sheet.IsBlocked));
    }

    [Fact]
    public void Preview_UnsupportedRuleOnAnUnselectedSheet_DoesNotBlockTheSelectedSheet()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");

        TestSheetWorkbookFactory.Create(path,
            [
                new TestAggregationSheetSpec
                {
                    Name = "きれいな表",
                    Rows = [["A", 1]],
                    ConditionalFormattings =
                    [
                        new TestConditionalFormatting
                        {
                            Rules = [new TestConditionalFormattingRule { Type = "duplicateValues" }],
                        },
                    ],
                },
                new TestAggregationSheetSpec
                {
                    Name = "問題あり",
                    Rows = [["B", 2]],
                    ConditionalFormattings =
                    [
                        new TestConditionalFormatting
                        {
                            Rules =
                            [
                                new TestConditionalFormattingRule { Type = "dataBar" },
                            ],
                        },
                    ],
                },
            ],
            differentialFormats: [RedFill]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "きれいな表")).Success);

        Assert.Single(Formattings(output, "きれいな表"));
    }

    [Fact]
    public void Aggregate_ConditionalFormatting_ComesBeforeDataValidationsAndHyperlinks()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");

        TestSheetWorkbookFactory.Create(path,
            [
                new TestAggregationSheetSpec
                {
                    Name = "表",
                    Rows = [["A", 1]],
                    Merges = ["D1:E1"],
                    ConditionalFormattings =
                    [
                        new TestConditionalFormatting
                        {
                            Rules = [new TestConditionalFormattingRule { Type = "duplicateValues" }],
                        },
                    ],
                    DataValidations = [new TestDataValidation("B1:B5", "list", Formula1: "\"はい,いいえ\"")],
                    Hyperlinks = [new TestHyperlink("C1", ExternalTarget: "https://example.invalid/")],
                    AddPrintOptions = true,
                    AddPageMargins = true,
                },
            ],
            differentialFormats: [RedFill]);

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        // CT_Worksheet の順序: mergeCells → conditionalFormatting → dataValidations
        //                     → hyperlinks → printOptions
        var order = Worksheet(output, "表").ChildElements.Select(child => child.LocalName).ToList();
        Assert.True(order.IndexOf("mergeCells") < order.IndexOf("conditionalFormatting"));
        Assert.True(order.IndexOf("conditionalFormatting") < order.IndexOf("dataValidations"));
        Assert.True(order.IndexOf("dataValidations") < order.IndexOf("hyperlinks"));
        Assert.True(order.IndexOf("hyperlinks") < order.IndexOf("printOptions"));
    }

    [Fact]
    public void Aggregate_ConditionalFormatting_PassesTheOpenXmlValidator()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path,
            [
                new TestDifferentialFormat
                {
                    Bold = true,
                    NumberFormatCode = "0.0%",
                    FillArgb = "FFFFC7CE",
                    CenterAlignment = true,
                    ThinBorder = true,
                    Unlocked = true,
                },
            ],
            new TestConditionalFormatting
            {
                Sqref = "A1:A10 C1:C10",
                Rules =
                [
                    new TestConditionalFormattingRule { Type = "duplicateValues", Priority = 1 },
                    new TestConditionalFormattingRule
                    {
                        Type = "top10", Priority = 2, Rank = 10U, Bottom = true,
                    },
                    new TestConditionalFormattingRule
                    {
                        Type = "aboveAverage", Priority = 3, StdDev = 2, StopIfTrue = true,
                    },
                ],
            });

        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        using var stream = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
    }

    [Fact]
    public void Aggregate_ConditionalFormatting_LeavesTheInputFileUnchanged()
    {
        using var dir = new TempDir();
        var path = dir.File("A.xlsx");
        CreateWorkbook(path, new TestConditionalFormattingRule { Type = "duplicateValues" });

        var before = Snapshot(path);
        var output = dir.File("out.xlsx");
        Assert.True(Aggregate(output, (path, "表")).Success);

        Assert.Equal(before, Snapshot(path));
    }

    // ── ヘルパー ──────────────────────────────────────────

    private static void CreateWorkbook(string path, params TestConditionalFormatting[] formattings)
        => CreateWorkbook(path, [RedFill], formattings);

    private static void CreateWorkbook(string path, TestConditionalFormattingRule rule)
        => CreateWorkbook(path, [RedFill], rule);

    private static void CreateWorkbook(
        string path,
        IReadOnlyList<TestDifferentialFormat> differentialFormats,
        TestConditionalFormattingRule rule)
        => CreateWorkbook(
            path,
            differentialFormats,
            [new TestConditionalFormatting { Sqref = "A1:A10", Rules = [rule] }]);

    private static void CreateWorkbook(
        string path,
        IReadOnlyList<TestDifferentialFormat> differentialFormats,
        params TestConditionalFormatting[] formattings)
        => TestSheetWorkbookFactory.Create(
            path,
            [
                new TestAggregationSheetSpec
                {
                    Name = "表",
                    Rows = [["A", 1], ["B", 2]],
                    ConditionalFormattings = formattings,
                },
            ],
            differentialFormats: differentialFormats);

    private static void AssertBlocked(string path, string expectedFragment)
    {
        var preview = CreatePreview((path, "表"));

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.Blocks, issue => issue.Message.Contains(expectedFragment));
    }

    private static SheetAggregationPreview CreatePreview(params (string Path, string Sheet)[] selections)
        => new SheetAggregationPlanner().CreatePreview(
            [.. selections.Select(s => new SheetSelection(s.Path, s.Sheet))]);

    private static SheetAggregationResult Aggregate(
        string output, params (string Path, string Sheet)[] selections)
    {
        var preview = CreatePreview(selections);
        Assert.True(preview.CanExecute,
            string.Join(" / ", preview.Blocks.Select(issue => $"{issue.Location}: {issue.Message}")));
        return new SheetAggregator().Execute(preview, output);
    }

    private static Worksheet Worksheet(string path, string sheetName)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(s => s.Name?.Value == sheetName);
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet!;
    }

    private static Stylesheet Stylesheet(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        return document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
    }

    private static IReadOnlyList<DifferentialFormat> Dxfs(string path)
        => [.. Stylesheet(path).DifferentialFormats?.Elements<DifferentialFormat>() ?? []];

    private static IReadOnlyList<ConditionalFormatting> Formattings(string path, string sheetName)
        => [.. Worksheet(path, sheetName).Elements<ConditionalFormatting>()];

    private static ConditionalFormatting SingleFormatting(string path, string sheetName)
        => Assert.Single(Formattings(path, sheetName));

    private static ConditionalFormattingRule SingleRule(string path, string sheetName)
        => Assert.Single(SingleFormatting(path, sheetName).Elements<ConditionalFormattingRule>());

    private static (string Hash, long Length, DateTime LastWriteUtc) Snapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var info = new FileInfo(path);
        return (hash, info.Length, info.LastWriteTimeUtc);
    }
}
