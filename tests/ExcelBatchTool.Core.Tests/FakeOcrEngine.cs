using ExcelBatchTool.Core.Ocr;

namespace ExcelBatchTool.Core.Tests;

/// <summary>
/// テスト用の OCR。Offline OCR Pack を置かずに、統合・確認・出力までの筋道を確かめる。
///
/// 本物の認識器と同じ形(領域ごとに 2 つのモデルの読みと自信)を返すので、
/// 統合の規則と確認の決まりはこのまま製品と同じコードが走る。
/// </summary>
internal sealed class FakeOcrEngine : IOcrEngine
{
    private readonly Dictionary<int, List<OcrRawLine>> _pages = [];
    private readonly Dictionary<int, OcrPageProbe> _probes = [];

    public OcrEngineInfo Info { get; init; }
        = new("テスト多言語", "テスト日本語", "テスト", "テスト", 300);

    /// <summary>読み取りのたびに数える(ページ単位で処理していることの確認に使う)。</summary>
    public List<int> ReadPages { get; } = [];

    public List<int> ProbedPages { get; } = [];

    /// <summary>確認用に描いたページ(手元に置く枚数の確認に使う)。</summary>
    public List<int> RenderedPages { get; } = [];

    public int OpenCount { get; private set; }

    public bool IsDisposed { get; private set; }

    /// <summary>ページを読むたびに呼ばれる。中止やファイル差し替えの再現に使う。</summary>
    public Action<int>? OnRead { get; set; }

    public FakeOcrEngine Page(int page, params OcrRawLine[] lines)
    {
        _pages[page] = [.. lines];
        return this;
    }

    public FakeOcrEngine Probe(int page, double skew, int horizontal, int vertical)
    {
        _probes[page] = new OcrPageProbe(page, skew, horizontal, vertical);
        return this;
    }

    /// <summary>2 つのモデルが同じ文字を、同じ自信で読んだ領域。</summary>
    public static OcrRawLine Agreed(string text, double score, double y = 0, double x = 0)
        => new(new OcrBox(x, y, Math.Max(text.Length, 1) * 20, 30), text, score, text, score);

    /// <summary>2 つのモデルで読みが割れた領域。</summary>
    public static OcrRawLine Split(
        string multi, double multiScore, string japan, double japanScore, double y = 0, double x = 0)
        => new(
            new OcrBox(x, y, Math.Max(Math.Max(multi.Length, japan.Length), 1) * 20, 30),
            multi, multiScore, japan, japanScore);

    public IOcrPageSource Open(string pdfFilePath)
    {
        OpenCount++;
        return new Source(this);
    }

    public void Dispose() => IsDisposed = true;

    private sealed class Source(FakeOcrEngine engine) : IOcrPageSource
    {
        public int PageCount => engine._pages.Count;

        public OcrPageProbe Probe(int pageNumber, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            engine.ProbedPages.Add(pageNumber);
            return engine._probes.TryGetValue(pageNumber, out var probe)
                ? probe
                : new OcrPageProbe(pageNumber, 0, 0, 0);
        }

        public IReadOnlyList<OcrRawLine> Read(int pageNumber, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            engine.OnRead?.Invoke(pageNumber);
            engine.ReadPages.Add(pageNumber);
            return engine._pages.TryGetValue(pageNumber, out var lines) ? lines : [];
        }

        public OcrPageImage RenderPage(int pageNumber, int dpi, CancellationToken cancellationToken)
        {
            engine.RenderedPages.Add(pageNumber);

            // 実際の画像は要らない。大きさと倍率が正しく伝わることだけを確かめる。
            return new OcrPageImage(
                pageNumber, [0x89, 0x50, 0x4E, 0x47], 1240, 1754, (double)dpi / 300);
        }

        public void Dispose()
        {
        }
    }
}
