namespace ExcelBatchTool.Core.Ocr;

/// <summary>ページ画像上の位置(左上が原点、単位は画素)。</summary>
public readonly record struct OcrBox(double X, double Y, double Width, double Height)
{
    public double Bottom => Y + Height;

    public double Right => X + Width;

    public double CenterY => Y + (Height / 2);

    public double CenterX => X + (Width / 2);
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
/// 傾き・罫線・下線は「このページをどう扱うか」の判断に使うので、
/// 何分もかかる認識より先に測る。
/// </summary>
public sealed record OcrPageProbe(
    int Page,
    double SkewDegrees,
    int HorizontalRulings,
    int VerticalRulings)
{
    /// <summary>傾きの推定が信頼できるか(行らしい塊が十分に見つかったか)。</summary>
    public bool SkewReliable { get; init; } = true;

    /// <summary>記入欄の下線のように、縦線と組にならない横線の本数。</summary>
    public int UnderlineCount { get; init; }
}

/// <summary>
/// 1 ページを読んだ結果。
///
/// 座標は**傾きを直したあとの画像**のもの。元のページへ戻すには
/// <see cref="Transform"/> を通す。戻す責任は呼び出し側(Core)にあり、
/// Pack 側は「直した画像で読んだ結果」と「戻し方」だけを返す。
/// </summary>
public sealed record OcrPageRead(
    int Page,
    IReadOnlyList<OcrRawLine> Lines,
    DeskewTransform Transform)
{
    /// <summary>横罫線の位置(直した画像の座標)。</summary>
    public IReadOnlyList<double> RowRulings { get; init; } = [];

    /// <summary>縦罫線の位置(直した画像の座標)。</summary>
    public IReadOnlyList<double> ColumnRulings { get; init; } = [];
}

/// <summary>確認用に描いた 1 ページの画像。</summary>
public sealed record OcrPageImage(int Page, byte[] Png, int Width, int Height, double ScaleFromOcr);

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

    /// <summary>
    /// ページを 2 つのモデルで読む。
    /// <paramref name="deskewDegrees"/> が 0 でなければ、その角度だけ直してから読む。
    /// </summary>
    OcrPageRead Read(int pageNumber, double deskewDegrees, CancellationToken cancellationToken);

    /// <summary>
    /// 指定した場所の黒い画素の割合を測る(チェックや丸の判定に使う)。
    /// 座標は「直した画像」のもの。1 回の描画でまとめて測る。
    /// </summary>
    IReadOnlyList<double> InkRatios(
        int pageNumber,
        IReadOnlyList<OcrBox> areas,
        double deskewDegrees,
        CancellationToken cancellationToken);

    /// <summary>確認画面に出すためのページ画像(元のページのまま。傾きは直さない)。</summary>
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
