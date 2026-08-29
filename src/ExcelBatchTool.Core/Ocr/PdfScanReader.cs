using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Pdf;

namespace ExcelBatchTool.Core.Ocr;

/// <summary>
/// スキャンされたページを OCR して、確認・修正できる形にする。
///
/// ここでは出力ファイルを一切作らない。作れるのは「人が確認し終えたあと」だけ、
/// という順序を型で守るために、この段階の結果は <see cref="OcrDocumentReading"/> に留める。
///
/// 認識は 1 ページ数秒かかるので、まず画像だけを見る安い確認を全ページに通し、
/// この段階で扱えないページ(傾き・表)が見つかったら、認識を始める前に止める。
/// </summary>
public sealed class PdfScanReader
{
    /// <summary>これ以上傾いていたら、この段階では確定させない(傾き補正は次の段階)。</summary>
    public const double MaxSkewDegrees = 1.5;

    /// <summary>画像から罫線の格子が見つかったら「表らしいスキャン」とみなす。</summary>
    public const int TableRulingThreshold = 3;

    public OcrDocumentReading Read(
        IOcrEngine engine,
        string pdfFilePath,
        IReadOnlyList<int> pages,
        IProgress<OcrProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        using var source = engine.Open(pdfFilePath);

        var probes = new List<OcrPageProbe>();
        progress?.Report(new OcrProgress(0, pages.Count, IsProbe: true));

        foreach (var page in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            probes.Add(source.Probe(page, cancellationToken));
            progress?.Report(new OcrProgress(probes.Count, pages.Count, IsProbe: true));
        }

        var needsDeskew = probes
            .Where(probe => Math.Abs(probe.SkewDegrees) > MaxSkewDegrees)
            .Select(probe => probe.Page)
            .ToList();

        var tableLike = probes
            .Where(probe => probe.HorizontalRulings >= TableRulingThreshold
                && probe.VerticalRulings >= TableRulingThreshold)
            .Select(probe => probe.Page)
            .ToList();

        var issues = new List<MergeIssue>();

        if (tableLike.Count > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"スキャンされた表の可能性があります({Describe(tableLike)})。"
                    + "表としての読み取りは次の段階で対応します。"
                    + "文章として無理に取り出すことはしません。"));
        }

        if (needsDeskew.Count > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"ページが傾いています({Describe(needsDeskew)})。"
                    + "傾きの補正は次の段階で対応します。"
                    + "傾いたまま読み取った結果を確定させることはしません。"));
        }

        // 扱えないページがあるなら、何分もかかる認識を始めない。
        if (issues.Count > 0)
        {
            return new OcrDocumentReading
            {
                Items = [],
                OcrPages = pages,
                EngineInfo = engine.Info,
                NeedsDeskewPages = needsDeskew,
                TableLikePages = tableLike,
                Issues = issues,
            };
        }

        var items = new List<OcrItem>();
        var done = 0;
        progress?.Report(new OcrProgress(0, pages.Count));

        foreach (var page in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var lines = source.Read(page, cancellationToken);
            items.AddRange(BuildItems(page, lines));

            done++;
            progress?.Report(new OcrProgress(done, pages.Count));
        }

        if (items.Count == 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                "スキャンされたページから文字を読み取れませんでした。"
                    + "白紙のページか、文字として認識できない内容の可能性があります。"));
        }

        return new OcrDocumentReading
        {
            Items = items,
            OcrPages = pages,
            EngineInfo = engine.Info,
            NeedsDeskewPages = needsDeskew,
            TableLikePages = tableLike,
            Issues = issues,
        };
    }

    /// <summary>1 ページ分の領域を、行に組み立てながら確認対象の項目にする。</summary>
    internal static IReadOnlyList<OcrItem> BuildItems(int page, IReadOnlyList<OcrRawLine> raw)
    {
        var fused = raw.Select(OcrFusion.Fuse).ToList();
        var lines = OcrLineLayout.BuildLines(raw);
        var items = new List<OcrItem>();

        foreach (var line in lines)
        {
            var position = 0;
            foreach (var index in line.RegionIndexes)
            {
                var result = fused[index];
                var region = raw[index];

                items.Add(new OcrItem
                {
                    PageNumber = page,
                    LineNumber = line.LineNumber,
                    IndexInLine = position++,
                    Text = result.Text,
                    BoundingBox = region.Box,
                    Confidence = result.Confidence,
                    Reason = result.Reason,
                    OriginalEngineResults =
                    [
                        new OcrEngineReading(
                            OcrFusion.MultiEngineName,
                            PdfTextNormalization.Normalize(region.MultiText),
                            OcrFusion.Finite(region.MultiScore)),
                        new OcrEngineReading(
                            OcrFusion.JapanEngineName,
                            PdfTextNormalization.Normalize(region.JapanText),
                            OcrFusion.Finite(region.JapanScore)),
                    ],
                    InitialStatus = result.Status,
                    Status = result.Status,
                });
            }
        }

        return items;
    }

    private static string Describe(IReadOnlyList<int> pages)
        => pages.Count <= 5
            ? string.Join(" / ", pages.Select(page => $"{page} ページ目"))
            : string.Join(" / ", pages.Take(5).Select(page => $"{page} ページ目")) + " ほか";
}
