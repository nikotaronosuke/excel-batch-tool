namespace ExcelBatchTool.Core.Ocr;

/// <summary>
/// 確認中のページ画像を、決まった枚数だけ手元に置く。
///
/// 120 ページを 300dpi の画像のまま抱えるとメモリが際限なく増えるので、
/// **今見ているページと、その前後だけ**を残して古いものから捨てる。
/// 長く確認していてもメモリが増え続けないようにするのがここの目的。
/// </summary>
public sealed class OcrPageImageCache(Func<int, OcrPageImage> render, int capacity = 3)
{
    /// <summary>手元に置くページ数の既定値(今のページ + 前後 1 ページ)。</summary>
    public const int DefaultCapacity = 3;

    private readonly Dictionary<int, OcrPageImage> _pages = [];

    /// <summary>使った順。末尾がいちばん新しい。</summary>
    private readonly List<int> _order = [];

    private readonly int _capacity = Math.Max(capacity, 1);

    /// <summary>いま手元にあるページ数。</summary>
    public int Count => _pages.Count;

    /// <summary>実際に描いた回数(測定とテスト用)。</summary>
    public int RenderCount { get; private set; }

    /// <summary>手元にあるページ番号(古い順)。</summary>
    public IReadOnlyList<int> Pages => _order;

    public OcrPageImage Get(int page)
    {
        if (_pages.TryGetValue(page, out var cached))
        {
            Touch(page);
            return cached;
        }

        var image = render(page);
        RenderCount++;

        _pages[page] = image;
        Touch(page);
        Evict();

        return image;
    }

    /// <summary>前後のページを先に用意しておく(ページ送りを待たせないため)。</summary>
    public void Preload(int page, int pageCount)
    {
        foreach (var neighbour in new[] { page, page - 1, page + 1 })
        {
            if (neighbour >= 1 && neighbour <= pageCount && !_pages.ContainsKey(neighbour))
            {
                // 今のページを追い出さないよう、余裕があるときだけ。
                if (_pages.Count >= _capacity)
                {
                    return;
                }

                Get(neighbour);
            }
        }
    }

    public void Clear()
    {
        _pages.Clear();
        _order.Clear();
    }

    private void Touch(int page)
    {
        _order.Remove(page);
        _order.Add(page);
    }

    private void Evict()
    {
        while (_order.Count > _capacity)
        {
            var oldest = _order[0];
            _order.RemoveAt(0);
            _pages.Remove(oldest);
        }
    }
}
