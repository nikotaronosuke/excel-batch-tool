using ExcelBatchTool.Core.Ocr;
using OpenCvSharp;
using Sdcb.PaddleInference;
using Sdcb.PaddleOCR;
using Sdcb.PaddleOCR.Models;
using Sdcb.PaddleOCR.Models.Details;

namespace ExcelBatchTool.Ocr;

/// <summary>
/// Offline OCR Pack の本体。モデルはすべて Pack の中のファイルから読む
/// (実行時のダウンロードはしない。Python / Java / PATH 設定も要らない)。
///
/// 検出は 1 回だけ行い、同じ切り出し画像を 2 つの認識モデルへ通す。
/// 統合の規則は製品側(ExcelBatchTool.Core の OcrFusion)が持つ。
/// </summary>
public sealed class PaddleOcrEngine : IOcrEngine
{
    /// <summary>OCR にかける解像度。Phase 2F-R で 300dpi を実測して採用した。</summary>
    public const int OcrDpi = 300;

    /// <summary>ページの性質を測るだけの解像度(認識より先に安く回す)。</summary>
    public const int ProbeDpi = 100;

    private readonly string _root;
    private readonly PaddleOcrDetector _detector;
    private readonly PaddleOcrClassifier _classifier;
    private readonly PaddleOcrRecognizer _multi;
    private readonly PaddleOcrRecognizer _japan;

    public PaddleOcrEngine()
        : this(Path.GetDirectoryName(typeof(PaddleOcrEngine).Assembly.Location)!)
    {
    }

    public PaddleOcrEngine(string packDirectory)
    {
        _root = packDirectory;

        var models = Path.Combine(_root, "models");

        // 推論の実行方式は実測で選んだ(同じモデル・同じページでの 1 ページあたり):
        //   OpenBLAS 22.8 秒 / oneDNN 23.2 秒 / **ONNX 3.3 秒**
        // Intel MKL は再配布条件に曖昧さが残るため採らない。ONNX Runtime(MIT)と
        // paddle2onnx(Apache-2.0)はどちらも条件が明確で、いちばん速い。
        var threads = Math.Max(Environment.ProcessorCount, 1);
        void Device(PaddleConfig config) => PaddleDevice.Onnx(cpuMathThreadCount: threads)(config);

        _detector = new PaddleOcrDetector(
            new FileDetectionModel(Path.Combine(models, "det"), ModelVersion.V4), Device)
        {
            // 検出した枠の「確からしさ」で切り落とす処理は使わない。
            //
            // この絞り込み(Sdcb の GetScore)は、枠が画像の外へわずかにはみ出したときに
            // native 側で保護されていないメモリを読み、プロセスごと落ちることを実測した
            // (120 ページの帳票で再現。落ちるページは実行ごとに変わる)。
            // .NET からは捕まえられない種類の落ち方なので、呼ばない形にする。
            //
            // 外して困らないことも実測で確かめた: 120 ページで枠は 2,169 → 2,191 件
            // (+1%)しか増えず、増えた枠は認識の自信が低いので「要確認」へ回る。
            // 確からしさの判断は、認識モデル 2 つの一致と自信で行っている。
            BoxScoreThreahold = null,
        };
        _classifier = new PaddleOcrClassifier(
            new FileClassificationModel(Path.Combine(models, "cls"), ModelVersion.V2), Device);

        // 向きの判定は既定のまま使う。効きを弱めると読み取りが落ちることを実測した
        // (罫線表のセル一致 61.3% → 38.7%、罫線なし 87.9% → 64.3%)。
        // 検出した枠は上下が逆のまま返ることが多く、この判定がそれを直している。
        _multi = new PaddleOcrRecognizer(
            new FileRecognizationModel(
                Path.Combine(models, "rec-multi"),
                Path.Combine(models, "rec-multi", "dict.txt"),
                ModelVersion.V4),
            Device);
        _japan = new PaddleOcrRecognizer(
            new FileRecognizationModel(
                Path.Combine(models, "rec-japan"),
                Path.Combine(models, "rec-japan", "dict.txt"),
                ModelVersion.V4),
            Device);
    }

    public OcrEngineInfo Info { get; } = new(
        "ch_PP-OCRv4_rec",
        "japan_PP-OCRv4_rec",
        "Paddle Inference 2.6.1",
        "ONNX Runtime",
        OcrDpi);

    public IOcrPageSource Open(string pdfFilePath)
        => new PageSource(File.ReadAllBytes(pdfFilePath), this);

    public void Dispose()
    {
        _japan.Dispose();
        _multi.Dispose();
        _classifier.Dispose();
        _detector.Dispose();
    }

    private sealed class PageSource(byte[] pdf, PaddleOcrEngine engine) : IOcrPageSource
    {
        public int PageCount { get; } = PDFtoImage.Conversion.GetPageCount(pdf);

        public OcrPageProbe Probe(int pageNumber, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            using var mat = Render(pdf, pageNumber, ProbeDpi);
            using var binary = PageImageOps.Binarize(mat);
            using var forRulings = PageImageOps.BinarizeForRulings(mat);

            var (degrees, reliable) = PageImageOps.Skew(binary);
            var (rows, columns) = PageImageOps.Rulings(forRulings);

            return new OcrPageProbe(pageNumber, degrees, rows.Count, columns.Count)
            {
                SkewReliable = reliable,
                UnderlineCount = PageImageOps.Underlines(forRulings),
            };
        }

        public OcrPageRead Read(
            int pageNumber, double deskewDegrees, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            using var page = Render(pdf, pageNumber, OcrDpi);

            // 傾きを直してから読む。直した画像の座標と、元へ戻す変換を一緒に返す。
            var transform = DeskewTransform.None;
            using var mat = deskewDegrees == 0
                ? page.Clone()
                : PageImageOps.Rotate(page, deskewDegrees, out transform);

            using var binary = PageImageOps.BinarizeForRulings(mat);
            var (rowRulings, columnRulings) = PageImageOps.Rulings(binary);

            var rects = engine._detector.Run(mat);
            if (rects.Length == 0)
            {
                return new OcrPageRead(pageNumber, [], transform)
                {
                    RowRulings = rowRulings,
                    ColumnRulings = columnRulings,
                };
            }

            var crops = new List<Mat>(rects.Length);
            try
            {
                foreach (var rect in rects)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // 切り出しただけでは 180 度回ったままのことがある(実測)。
                    // 向きの判定を通してから認識へ渡す。回転不要のときは
                    // 同じ Mat がそのまま返るので、その場合は解放しない。
                    var raw = PaddleOcrAll.GetRotateCropImage(mat, rect);
                    var oriented = engine._classifier.Run(raw);
                    if (!ReferenceEquals(oriented, raw))
                    {
                        raw.Dispose();
                    }

                    crops.Add(oriented);
                }

                var array = crops.ToArray();
                var multi = engine._multi.Run(array, 0);
                var japan = engine._japan.Run(array, 0);

                var lines = new List<OcrRawLine>(rects.Length);
                for (var index = 0; index < rects.Length; index++)
                {
                    var box = rects[index].BoundingRect();
                    lines.Add(new OcrRawLine(
                        new OcrBox(box.X, box.Y, box.Width, box.Height),
                        multi[index].Text,
                        OcrFusion.Finite(multi[index].Score),
                        japan[index].Text,
                        OcrFusion.Finite(japan[index].Score)));
                }

                return new OcrPageRead(pageNumber, lines, transform)
                {
                    RowRulings = rowRulings,
                    ColumnRulings = columnRulings,
                };
            }
            finally
            {
                foreach (var crop in crops)
                {
                    crop.Dispose();
                }
            }
        }

        /// <summary>指定した場所の黒い画素の割合。1 回の描画でまとめて測る。</summary>
        public IReadOnlyList<double> InkRatios(
            int pageNumber,
            IReadOnlyList<OcrBox> areas,
            double deskewDegrees,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (areas.Count == 0)
            {
                return [];
            }

            using var page = Render(pdf, pageNumber, OcrDpi);
            using var mat = deskewDegrees == 0
                ? page.Clone()
                : PageImageOps.Rotate(page, -deskewDegrees, out _);
            using var gray = mat.CvtColor(ColorConversionCodes.BGR2GRAY);

            return [.. areas.Select(area => PageImageOps.InkRatio(gray, area))];
        }

        public OcrPageImage RenderPage(int pageNumber, int dpi, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            using var bitmap = PDFtoImage.Conversion.ToImage(
                pdf, page: pageNumber - 1, options: new PDFtoImage.RenderOptions(Dpi: dpi));
            using var image = SkiaSharp.SKImage.FromBitmap(bitmap);
            using var encoded = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 80);

            return new OcrPageImage(
                pageNumber, encoded.ToArray(), bitmap.Width, bitmap.Height, (double)dpi / OcrDpi);
        }

        public void Dispose()
        {
            // ページ画像はページごとに解放しているので、ここで抱えているものは無い。
        }

        private static Mat Render(byte[] bytes, int pageNumber, int dpi)
        {
            using var bitmap = PDFtoImage.Conversion.ToImage(
                bytes, page: pageNumber - 1, options: new PDFtoImage.RenderOptions(Dpi: dpi));
            using var image = SkiaSharp.SKImage.FromBitmap(bitmap);
            using var encoded = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
            return Cv2.ImDecode(encoded.ToArray(), ImreadModes.Color);
        }

    }
}
