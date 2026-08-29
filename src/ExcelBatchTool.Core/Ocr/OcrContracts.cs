namespace ExcelBatchTool.Core.Ocr;

/// <summary>ページ画像上の位置(左上が原点、単位は画素)。</summary>
public readonly record struct OcrBox(double X, double Y, double Width, double Height)
{
    public double Bottom => Y + Height;

    public double Right => X + Width;

    public double CenterY => Y + (Height / 2);
}

/// <summary>
/// 1 つの検出領域を 2 つの認識モデルで読んだ、統合前の結果。
///
/// 検出は 1 回だけ行い、同じ切り出し画像を両モデルへ通す。こうすると領域が 1 対 1 で
/// 対応するので、統合のときに位置合わせをしなくてよい。
/// </summary>
public sealed record OcrRawLine(
    OcrBox Box,
    string MultiText,
    double MultiScore,
    string JapanText,
    double JapanScore);

/// <summary>
/// OCR にかける前に、ページ画像だけを見て分かること。
/// 傾きと罫線の格子は「この段階で扱えるページか」の判断に使うので、
/// 何分もかかる認識より先に測る。
/// </summary>
public sealed record OcrPageProbe(
    int Page,
    double SkewDegrees,
    int HorizontalRulings,
    int VerticalRulings);

/// <summary>使っているモデルと版。控えと画面に出す。</summary>
public sealed record OcrEngineInfo(
    string MultiModel,
    string JapanModel,
    string Runtime,
    string Backend,
    int Dpi);

/// <summary>読み取りの進み具合(画面の「23 / 120 ページ」表示に使う)。</summary>
public sealed record OcrProgress(int DonePages, int TotalPages, bool IsProbe = false)
{
    public string Text => IsProbe
        ? $"{DonePages:N0} / {TotalPages:N0} ページを確認中"
        : $"{DonePages:N0} / {TotalPages:N0} ページを読み取り中";
}

/// <summary>1 つの PDF を開いた状態。ページ単位で読み、ページごとに解放する。</summary>
public interface IOcrPageSource : IDisposable
{
    int PageCount { get; }

    /// <summary>認識にかけず、ページ画像の性質だけを測る。</summary>
    OcrPageProbe Probe(int pageNumber, CancellationToken cancellationToken);

    /// <summary>ページを 2 つのモデルで読む。</summary>
    IReadOnlyList<OcrRawLine> Read(int pageNumber, CancellationToken cancellationToken);

    /// <summary>
    /// 確認画面に出すためのページ画像。OCR より粗い解像度でよいので、
    /// 「OCR の座標 → この画像の座標」の倍率も一緒に返す。
    /// </summary>
    OcrPageImage RenderPage(int pageNumber, int dpi, CancellationToken cancellationToken);
}

/// <summary>
/// OCR の実体。製品本体はこのインターフェースだけを知っていて、
/// 実装は Offline OCR Pack 側にある(本体に OCR ランタイムを混ぜ込まない)。
/// </summary>
public interface IOcrEngine : IDisposable
{
    OcrEngineInfo Info { get; }

    IOcrPageSource Open(string pdfFilePath);
}
