using System.Diagnostics;
using OpenCvSharp;
using Sdcb.PaddleInference;
using Sdcb.PaddleOCR;
using Sdcb.PaddleOCR.Models;
using Sdcb.PaddleOCR.Models.Online;

namespace PdfBench;

/// <summary>1 領域を 2 つの認識モデルで読んだ結果。統合方式の比較はこれを使って後から行う。</summary>
public sealed record DualRegion(
    int Page,
    double X,
    double Y,
    double Width,
    double Height,
    string MultiText,
    double MultiScore,
    string JapanText,
    double JapanScore);

public sealed record DualCapture(string Fixture, int Dpi, int Pages, double Seconds, List<DualRegion> Regions);

/// <summary>
/// 検出は 1 回だけ行い、同じ切り出し画像を 2 つの認識モデルへ通す。
/// こうすると領域が 1 対 1 で対応するので、統合方式を bbox 照合なしで比較できる。
/// </summary>
public static class OcrCapture
{
    public static DualCapture Run(string pdfPath, int dpi, int maxPages)
    {
        var bytes = File.ReadAllBytes(pdfPath);
        var pageCount = Math.Min(PDFtoImage.Conversion.GetPageCount(bytes), maxPages);

        var detectionModel = OnlineDetectionModel.ChineseV4.DownloadAsync().GetAwaiter().GetResult();
        var classificationModel = OnlineClassificationModel.ChineseMobileV2
            .DownloadAsync().GetAwaiter().GetResult();
        var multiModel = LocalDictOnlineRecognizationModel.ChineseV4.DownloadAsync().GetAwaiter().GetResult();
        var japanModel = LocalDictOnlineRecognizationModel.JapanV4.DownloadAsync().GetAwaiter().GetResult();

        using var detector = new PaddleOcrDetector(detectionModel, PaddleDevice.Mkldnn());
        using var classifier = new PaddleOcrClassifier(classificationModel, PaddleDevice.Mkldnn());
        using var multi = new PaddleOcrRecognizer(multiModel, PaddleDevice.Mkldnn());
        using var japan = new PaddleOcrRecognizer(japanModel, PaddleDevice.Mkldnn());

        var regions = new List<DualRegion>();
        var timer = Stopwatch.StartNew();

        for (var page = 0; page < pageCount; page++)
        {
            using var mat = Render(bytes, page, dpi);
            var rects = detector.Run(mat);

            var crops = new List<Mat>(rects.Length);
            try
            {
                foreach (var rect in rects)
                {
                    // 切り出しただけでは 180 度回ったままのことがある(実測: 電話番号が
                    // 逆さに読まれた)。向きの判定を通してから認識へ渡す。
                    // 回転不要のときは同じ Mat がそのまま返るので、その場合は解放しない。
                    var raw = PaddleOcrAll.GetRotateCropImage(mat, rect);
                    var oriented = classifier.Run(raw);
                    if (!ReferenceEquals(oriented, raw))
                    {
                        raw.Dispose();
                    }

                    crops.Add(oriented);
                }

                var multiResults = multi.Run(crops.ToArray(), 0);
                var japanResults = japan.Run(crops.ToArray(), 0);

                for (var index = 0; index < rects.Length; index++)
                {
                    var box = rects[index].BoundingRect();
                    regions.Add(new DualRegion(
                        page + 1,
                        box.X,
                        box.Y,
                        box.Width,
                        box.Height,
                        multiResults[index].Text,
                        Finite(multiResults[index].Score),
                        japanResults[index].Text,
                        Finite(japanResults[index].Score)));
                }
            }
            finally
            {
                foreach (var crop in crops)
                {
                    crop.Dispose();
                }
            }

            if ((page + 1) % 10 == 0)
            {
                Console.WriteLine($"  {page + 1}/{pageCount} pages");
            }
        }

        timer.Stop();
        return new DualCapture(
            Path.GetFileName(pdfPath), dpi, pageCount, timer.Elapsed.TotalSeconds, regions);
    }

    /// <summary>
    /// 認識器は NaN / Infinity を返すことがある(空に近い切り出しで実測)。
    /// そのまま比べると「自信 = 無限大」で閾値を通ってしまうので、必ず 0 に倒す。
    /// </summary>
    private static double Finite(float score)
        => double.IsFinite(score) ? score : 0;

    private static Mat Render(byte[] pdfBytes, int page, int dpi)
    {
        using var bitmap = PDFtoImage.Conversion.ToImage(
            pdfBytes, page: page, options: new PDFtoImage.RenderOptions(Dpi: dpi));
        using var image = SkiaSharp.SKImage.FromBitmap(bitmap);
        using var encoded = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
        return Cv2.ImDecode(encoded.ToArray(), ImreadModes.Color);
    }
}
