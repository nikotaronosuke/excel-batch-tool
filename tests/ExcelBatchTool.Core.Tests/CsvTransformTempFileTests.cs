using System.Text;
using ExcelBatchTool.Core.CsvTransform;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// Phase 2E.1。作業用ファイルは、この実行が新しく作ったものだけを扱う。
/// 実行前からあったファイルは上書きも削除もしない。すべて架空データ。
/// </summary>
public sealed class CsvTransformTempFileTests
{
    /// <summary>実行前から置いておく、この実行のものではないファイル。</summary>
    private static readonly (string Name, string Content)[] Decoys =
    [
        // Phase 2E で使っていた固定名。実行前から残っていることがあり得る。
        ("元データ_変換済み.csv.tmp", "前の作業で残った CSV"),
        ("元データ_変換済み.csv.audit.json.tmp", "前の作業で残った控え"),
        ("無関係.tmp", "まったく関係のないファイル"),
        ("元データ_変換済み.txt", "名前が似ているだけのファイル"),
    ];

    [Fact]
    public void FilesThatWereAlreadyThere_AreLeftAlone()
    {
        using var dir = new TempDir();
        var source = Source(dir);
        PlaceDecoys(dir);

        var preview = Preview(dir, source);
        var result = new CsvTransformer().Execute(preview);

        Assert.True(result.Success, result.Message);
        AssertDecoysUntouched(dir);
    }

    [Fact]
    public void AnExistingTempFile_IsNotUsedAndNotOverwritten()
    {
        using var dir = new TempDir();
        var source = Source(dir);
        PlaceDecoys(dir);

        var preview = Preview(dir, source);
        Assert.True(new CsvTransformer().Execute(preview).Success);

        // 作業用ファイルは固定名を使わないので、前の残りには手を付けない。
        Assert.Equal(
            "前の作業で残った CSV",
            File.ReadAllText(dir.File("元データ_変換済み.csv.tmp")));
        Assert.Equal(
            "前の作業で残った控え",
            File.ReadAllText(dir.File("元データ_変換済み.csv.audit.json.tmp")));
    }

    [Fact]
    public void WhenTheOutputCannotBeFinished_OnlyThisRunsFilesAreRemoved()
    {
        using var dir = new TempDir();
        var source = Source(dir);
        PlaceDecoys(dir);

        // 出力先の名前をフォルダーで塞ぐ。作業用ファイルを書いたあとで確定に失敗する。
        Directory.CreateDirectory(dir.File("元データ_変換済み.csv"));

        var preview = Preview(dir, source);
        var result = new CsvTransformer().Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("取り消しました", result.Message);
        AssertDecoysUntouched(dir);
        AssertNoLeftoverWorkFiles(dir);
    }

    [Fact]
    public void WhenTheAuditCannotBeFinished_OnlyThisRunsFilesAreRemoved()
    {
        using var dir = new TempDir();
        var source = Source(dir);
        PlaceDecoys(dir);

        // 控えの名前をフォルダーで塞ぐ。CSV を書き終えたあとで失敗する。
        Directory.CreateDirectory(dir.File("元データ_変換済み.csv.audit.json"));

        var preview = Preview(dir, source);
        var result = new CsvTransformer().Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("取り消しました", result.Message);

        // 先に移した CSV も取り消す。
        Assert.False(File.Exists(dir.File("元データ_変換済み.csv")));
        AssertDecoysUntouched(dir);
        AssertNoLeftoverWorkFiles(dir);
    }

    [Fact]
    public void WhenTheCleanupFails_OnlyThisRunsFilesAreNamed()
    {
        using var dir = new TempDir();
        var source = Source(dir);
        PlaceDecoys(dir);

        Directory.CreateDirectory(dir.File("元データ_変換済み.csv.audit.json"));

        var preview = Preview(dir, source);
        var result = new CsvTransformer { FileDeleter = _ => false }.Execute(preview);

        Assert.False(result.Success);
        Assert.Contains("取り消せなかったファイル", result.Message);
        Assert.Contains("元データ_変換済み.csv", result.Message);

        // 実行前からあったファイルは、名前にも挙げない。
        Assert.DoesNotContain("無関係.tmp", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("元データ_変換済み.txt", result.Message, StringComparison.Ordinal);
        AssertDecoysUntouched(dir);
    }

    [Fact]
    public void ASuccessfulRunLeavesNoWorkFilesBehind()
    {
        using var dir = new TempDir();
        var source = Source(dir);

        var result = new CsvTransformer().Execute(Preview(dir, source));

        Assert.True(result.Success, result.Message);
        Assert.True(File.Exists(dir.File("元データ_変換済み.csv")));
        Assert.True(File.Exists(dir.File("元データ_変換済み.csv.audit.json")));
        AssertNoLeftoverWorkFiles(dir);
    }

    [Fact]
    public void TwoRunsInTheSameFolderDoNotShareWorkFiles()
    {
        using var dir = new TempDir();
        var source = Source(dir);

        var first = new CsvTransformer().Execute(Preview(dir, source));
        Assert.True(first.Success, first.Message);

        // 同じデータ元から別名で作る。作業用ファイルの名前がぶつからない。
        var second = new CsvTransformer().Execute(Preview(dir, source, suffix: "_変換済み2"));

        Assert.True(second.Success, second.Message);
        Assert.True(File.Exists(dir.File("元データ_変換済み.csv")));
        Assert.True(File.Exists(dir.File("元データ_変換済み2.csv")));
        AssertNoLeftoverWorkFiles(dir);
    }

    [Fact]
    public void TheWorkFileNameIsNotKeptAnywhere()
    {
        using var dir = new TempDir();
        var source = Source(dir);

        Assert.True(new CsvTransformer().Execute(Preview(dir, source)).Success);

        // 出来上がるファイル名にも控えにも、作業用の名前は残らない。
        Assert.Equal(
            ["元データ.csv", "元データ_変換済み.csv", "元データ_変換済み.csv.audit.json"],
            Directory.GetFiles(dir.Root).Select(Path.GetFileName).Order(StringComparer.Ordinal));

        Assert.DoesNotContain(
            ".tmp",
            File.ReadAllText(dir.File("元データ_変換済み.csv.audit.json")),
            StringComparison.Ordinal);
    }

    [Fact]
    public void TheSourceFileIsStillUnchanged()
    {
        using var dir = new TempDir();
        var source = Source(dir);
        PlaceDecoys(dir);
        var before = Fingerprint(source);

        Directory.CreateDirectory(dir.File("元データ_変換済み.csv.audit.json"));
        Assert.False(new CsvTransformer().Execute(Preview(dir, source)).Success);

        Assert.Equal(before, Fingerprint(source));
    }

    // ── 補助 ─────────────────────────────────────────────

    private static string Source(TempDir dir)
    {
        var path = dir.File("元データ.csv");
        TestSourceTableFactory.CreateCsv(path, ["商品コード,商品名", "A001,商品A", "A002,商品B"]);
        return path;
    }

    private static void PlaceDecoys(TempDir dir)
    {
        foreach (var (name, content) in Decoys)
        {
            File.WriteAllText(dir.File(name), content, new UTF8Encoding(false));
        }
    }

    private static void AssertDecoysUntouched(TempDir dir)
    {
        foreach (var (name, content) in Decoys)
        {
            Assert.True(File.Exists(dir.File(name)), $"{name} が消えています。");
            Assert.Equal(content, File.ReadAllText(dir.File(name), new UTF8Encoding(false)));
        }
    }

    /// <summary>この実行が作った作業用ファイルが残っていないこと。</summary>
    private static void AssertNoLeftoverWorkFiles(TempDir dir)
    {
        var known = new HashSet<string>(Decoys.Select(decoy => decoy.Name), StringComparer.Ordinal);

        Assert.Empty(Directory.GetFiles(dir.Root, "*.tmp")
            .Select(Path.GetFileName)
            .Where(name => !known.Contains(name!)));
    }

    private static CsvTransformPreview Preview(
        TempDir dir, string sourceFilePath, string suffix = CsvTransformDefaults.OutputSuffix)
        => new CsvTransformPlanner().CreatePreview(new CsvTransformRequest
        {
            SourceFilePath = sourceFilePath,
            Columns =
            [
                new CsvOutputColumnRequest
                {
                    OutputName = "コード",
                    ValueSourceKind = CsvValueSourceKind.SourceColumn,
                    SourceColumn = "商品コード",
                },
            ],
            OutputSuffix = suffix,
        });

    private static (long Length, DateTime Written, string Hash) Fingerprint(string path)
    {
        var info = new FileInfo(path);
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
        return (info.Length, info.LastWriteTimeUtc, hash);
    }
}
