namespace ExcelBatchTool.Core.Tests;

public sealed class WorkbookAnalyzerTests
{
    [Fact]
    public void Analyze_NormalWorkbook_ReturnsNormalLevelAndSheetInfo()
    {
        using var dir = new TempDir();
        var path = dir.File("normal.xlsx");
        TestWorkbookFactory.CreateNormal(path);

        var result = WorkbookAnalyzer.Analyze(path);

        Assert.Equal(AnalysisStatus.Succeeded, result.Status);
        Assert.Equal(SafetyLevel.Normal, result.Level);
        Assert.Empty(result.Findings);
        var sheet = Assert.Single(result.Sheets);
        Assert.Equal("データ", sheet.Name);
        Assert.Equal("A1:B3", sheet.UsedRange);
        Assert.Equal(3, sheet.EstimatedRowCount);
        Assert.Equal(2, sheet.EstimatedColumnCount);
        Assert.Equal(path, result.FilePath);
        Assert.True(result.FileSizeBytes > 0);
    }

    [Fact]
    public void Analyze_MultiSheetWorkbook_ReturnsAllSheetNames()
    {
        using var dir = new TempDir();
        var path = dir.File("multi.xlsx");
        TestWorkbookFactory.CreateMultiSheet(path, "売上", "在庫", "月次まとめ");

        var result = WorkbookAnalyzer.Analyze(path);

        Assert.Equal(AnalysisStatus.Succeeded, result.Status);
        Assert.Equal(new[] { "売上", "在庫", "月次まとめ" }, result.Sheets.Select(sheet => sheet.Name));
    }

    [Fact]
    public void Analyze_WorkbookWithFormulas_ReportsFormulaFinding()
    {
        using var dir = new TempDir();
        var path = dir.File("formulas.xlsx");
        TestWorkbookFactory.CreateWithFormulas(path);

        var result = WorkbookAnalyzer.Analyze(path);

        Assert.Equal(AnalysisStatus.Succeeded, result.Status);
        Assert.Equal(SafetyLevel.NeedsAttention, result.Level);
        var finding = Assert.Single(result.Findings, finding => finding.Type == FindingType.Formula);
        Assert.Equal(SafetyLevel.NeedsAttention, finding.Level);
        Assert.Equal(1, finding.Count);
        Assert.Contains("計算", finding.SheetNames);
    }

    [Fact]
    public void Analyze_WorkbookWithMergedCells_ReportsMergedCellFinding()
    {
        using var dir = new TempDir();
        var path = dir.File("merged.xlsx");
        TestWorkbookFactory.CreateWithMergedCells(path);

        var result = WorkbookAnalyzer.Analyze(path);

        Assert.Equal(SafetyLevel.NeedsAttention, result.Level);
        var finding = Assert.Single(result.Findings, finding => finding.Type == FindingType.MergedCell);
        Assert.Equal(1, finding.Count);
    }

    [Fact]
    public void Analyze_WorkbookWithChart_ReportsChartFinding()
    {
        using var dir = new TempDir();
        var path = dir.File("chart.xlsx");
        TestWorkbookFactory.CreateWithChart(path);

        var result = WorkbookAnalyzer.Analyze(path);

        Assert.Equal(AnalysisStatus.Succeeded, result.Status);
        Assert.Equal(SafetyLevel.NeedsAttention, result.Level);
        var finding = Assert.Single(result.Findings, finding => finding.Type == FindingType.Chart);
        Assert.Contains("グラフ元", finding.SheetNames);
    }

    [Fact]
    public void Analyze_WorkbookWithImage_ReportsImageFinding()
    {
        using var dir = new TempDir();
        var path = dir.File("image.xlsx");
        TestWorkbookFactory.CreateWithImage(path);

        var result = WorkbookAnalyzer.Analyze(path);

        Assert.Equal(SafetyLevel.NeedsAttention, result.Level);
        Assert.Single(result.Findings, finding => finding.Type == FindingType.Image);
    }

    [Fact]
    public void Analyze_ProtectedSheet_ReportsSheetProtectionFinding()
    {
        using var dir = new TempDir();
        var path = dir.File("protected.xlsx");
        TestWorkbookFactory.CreateWithSheetProtection(path);

        var result = WorkbookAnalyzer.Analyze(path);

        Assert.Equal(SafetyLevel.NeedsAttention, result.Level);
        var finding = Assert.Single(result.Findings, finding => finding.Type == FindingType.SheetProtection);
        Assert.Contains("保護", finding.SheetNames);
    }

    [Fact]
    public void Analyze_WorkbookWithExternalLink_ReportsExternalLinkFinding()
    {
        using var dir = new TempDir();
        var path = dir.File("external.xlsx");
        TestWorkbookFactory.CreateWithExternalLink(path);

        var result = WorkbookAnalyzer.Analyze(path);

        Assert.Equal(SafetyLevel.NeedsAttention, result.Level);
        var finding = Assert.Single(result.Findings, finding => finding.Type == FindingType.ExternalLink);
        Assert.True(finding.Count >= 1);
    }

    [Fact]
    public void Analyze_CorruptFile_FailsAsUnsupportedWithoutThrowing()
    {
        using var dir = new TempDir();
        var path = dir.File("corrupt.xlsx");
        TestWorkbookFactory.CreateCorrupt(path);

        var result = WorkbookAnalyzer.Analyze(path);

        Assert.Equal(AnalysisStatus.Failed, result.Status);
        Assert.Equal(SafetyLevel.UnsupportedForNow, result.Level);
        Assert.NotNull(result.ErrorMessage);
        Assert.Single(result.Findings, finding => finding.Type == FindingType.OpenFailed);
    }

    [Fact]
    public void Analyze_NonXlsxExtension_FailsAsUnsupportedFileType()
    {
        using var dir = new TempDir();
        var path = dir.File("legacy.xls");
        File.WriteAllText(path, "架空の旧形式ファイル");

        var result = WorkbookAnalyzer.Analyze(path);

        Assert.Equal(AnalysisStatus.Failed, result.Status);
        Assert.Equal(SafetyLevel.UnsupportedForNow, result.Level);
        Assert.Single(result.Findings, finding => finding.Type == FindingType.UnsupportedFileType);
    }

    [Fact]
    public void Analyze_MissingFile_FailsWithoutThrowing()
    {
        using var dir = new TempDir();

        var result = WorkbookAnalyzer.Analyze(dir.File("not-found.xlsx"));

        Assert.Equal(AnalysisStatus.Failed, result.Status);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public async Task BatchAnalyzer_OneCorruptFile_DoesNotStopOtherFiles()
    {
        using var dir = new TempDir();
        var corrupt = dir.File("corrupt.xlsx");
        var normal = dir.File("normal.xlsx");
        TestWorkbookFactory.CreateCorrupt(corrupt);
        TestWorkbookFactory.CreateNormal(normal);

        var reported = 0;
        var progress = new Progress<WorkbookAnalysisResult>(_ => Interlocked.Increment(ref reported));

        var results = await new BatchAnalyzer().AnalyzeAsync([corrupt, normal], progress);

        Assert.Equal(2, results.Count);
        Assert.Equal(AnalysisStatus.Failed, results[0].Status);
        Assert.Equal(AnalysisStatus.Succeeded, results[1].Status);
        Assert.Equal(SafetyLevel.Normal, results[1].Level);
    }
}
