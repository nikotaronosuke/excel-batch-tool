namespace ExcelBatchTool.Core.Tests;

/// <summary>テストごとに使い捨てる一時ディレクトリ。</summary>
public sealed class TempDir : IDisposable
{
    public TempDir()
    {
        Root = Path.Combine(Path.GetTempPath(), "ExcelBatchTool.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Root);
    }

    public string Root { get; }

    public string File(string name) => Path.Combine(Root, name);

    public void Dispose()
    {
        try
        {
            Directory.Delete(Root, recursive: true);
        }
        catch (IOException)
        {
            // テスト一時ファイルの削除失敗は無視する。
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
