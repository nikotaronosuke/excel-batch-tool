using System.Security.Cryptography;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 0 の解析が対象ファイルを一切変更しないことを、
/// SHA-256 ハッシュ・ファイルサイズ・最終更新日時の比較で証明する。
/// </summary>
public sealed class FileImmutabilityTests
{
    private sealed record FileSnapshot(string Sha256, long Length, DateTime LastWriteTimeUtc);

    [Fact]
    public void Analyze_AllGeneratedWorkbooks_DoesNotModifySourceFiles()
    {
        using var dir = new TempDir();
        var paths = CreateAllTestWorkbooks(dir);

        var before = paths.ToDictionary(path => path, TakeSnapshot);

        foreach (var path in paths)
        {
            WorkbookAnalyzer.Analyze(path);
        }

        foreach (var path in paths)
        {
            Assert.Equal(before[path], TakeSnapshot(path));
        }
    }

    [Fact]
    public async Task BatchAnalyze_AllGeneratedWorkbooks_DoesNotModifySourceFiles()
    {
        using var dir = new TempDir();
        var paths = CreateAllTestWorkbooks(dir);

        var before = paths.ToDictionary(path => path, TakeSnapshot);

        var results = await new BatchAnalyzer().AnalyzeAsync(paths);

        Assert.Equal(paths.Count, results.Count);
        foreach (var path in paths)
        {
            Assert.Equal(before[path], TakeSnapshot(path));
        }
    }

    private static List<string> CreateAllTestWorkbooks(TempDir dir)
    {
        TestWorkbookFactory.CreateNormal(dir.File("normal.xlsx"));
        TestWorkbookFactory.CreateMultiSheet(dir.File("multi.xlsx"), "一", "二", "三");
        TestWorkbookFactory.CreateWithFormulas(dir.File("formulas.xlsx"));
        TestWorkbookFactory.CreateWithMergedCells(dir.File("merged.xlsx"));
        TestWorkbookFactory.CreateWithChart(dir.File("chart.xlsx"));
        TestWorkbookFactory.CreateWithImage(dir.File("image.xlsx"));
        TestWorkbookFactory.CreateWithSheetProtection(dir.File("protected.xlsx"));
        TestWorkbookFactory.CreateWithExternalLink(dir.File("external.xlsx"));
        TestWorkbookFactory.CreateCorrupt(dir.File("corrupt.xlsx"));

        return Directory.EnumerateFiles(dir.Root, "*.xlsx").OrderBy(path => path).ToList();
    }

    private static FileSnapshot TakeSnapshot(string path)
    {
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var info = new FileInfo(path);
        return new FileSnapshot(hash, info.Length, info.LastWriteTimeUtc);
    }
}
