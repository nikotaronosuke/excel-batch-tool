using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Pdf;

namespace ExcelBatchTool.Core.Ocr;

/// <summary>読み取ったページをどう組み立てるか。</summary>
public enum OcrReadMode
{
    /// <summary>ページごとに種類を見て決める。</summary>
    Auto = 0,

    /// <summary>文章として、ページ / 行 / 内容で取り出す。</summary>
    Lines,

    /// <summary>表として、行・列で取り出す。</summary>
    Table,

    /// <summary>同じ様式の帳票として、1 ページ 1 件で取り出す。</summary>
    FixedForm,
}

/// <summary>読み取りの指定。</summary>
public sealed record OcrReadOptions
{
    public OcrReadMode Mode { get; init; } = OcrReadMode.Auto;

    /// <summary>帳票として読むときの項目指定。</summary>
    public FormTemplate? Template { get; init; }

    /// <summary>傾きを直す(既定で行う)。</summary>
    public bool Deskew { get; init; } = true;
}

/// <summary>
/// スキャンされたページを OCR して、確認・修正できる形にする。
///
/// ここでは出力ファイルを一切作らない。作れるのは「人が確認し終えたあと」だけ、
/// という順序を型で守るために、この段階の結果は <see cref="OcrDocumentReading"/> に留める。
///
/// 認識は 1 ページ数秒かかるので、まず画像だけを見る安い確認を全ページに通し、
/// 傾きの有無とページの種類をそこで決めてから認識へ入る。
/// </summary>
public sealed class PdfScanReader
{
    /// <summary>これ以上傾いていたら、直さずに人へ回す。</summary>
    public const double MaxSkewDegrees = DeskewPolicy.MaximumAngle;

    /// <summary>PDF を開いて読み、終わったら閉じる(テストと測定用の入口)。</summary>
    public OcrDocumentReading Read(
        IOcrEngine engine,
        string pdfFilePath,
        IReadOnlyList<int> pages,
        IProgress<OcrProgress>? progress = null,
        CancellationToken cancellationToken = default)
        => Read(engine, pdfFilePath, pages, new OcrReadOptions(), progress, cancellationToken);

    public OcrDocumentReading Read(
        IOcrEngine engine,
        string pdfFilePath,
        IReadOnlyList<int> pages,
        OcrReadOptions options,
        IProgress<OcrProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        using var source = engine.Open(pdfFilePath);
        return Read(source, engine.Info, pages, options, progress, cancellationToken);
    }

    public OcrDocumentReading Read(
        IOcrPageSource source,
        OcrEngineInfo engineInfo,
        IReadOnlyList<int> pages,
        IProgress<OcrProgress>? progress = null,
        CancellationToken cancellationToken = default)
        => Read(source, engineInfo, pages, new OcrReadOptions(), progress, cancellationToken);

    /// <summary>
    /// 開いたままの PDF を読む。
    ///
    /// 確認画面はページ画像を出すために読み取りのあとも PDF を開いたままにするので、
    /// 開閉の責任は呼び出し側に持たせる。
    /// </summary>
    public OcrDocumentReading Read(
        IOcrPageSource source,
        OcrEngineInfo engineInfo,
        IReadOnlyList<int> pages,
        OcrReadOptions options,
        IProgress<OcrProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        // ── 1. 画像だけを見る安い確認 ──────────────────────
        var probes = new List<OcrPageProbe>();
        progress?.Report(new OcrProgress(0, pages.Count, IsProbe: true));

        foreach (var page in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            probes.Add(source.Probe(page, cancellationToken));
            progress?.Report(new OcrProgress(probes.Count, pages.Count, IsProbe: true));
        }

        var issues = new List<MergeIssue>();

        // 大きすぎる傾きは直さずに止める(端が切れるくらい傾いた紙は撮り直しが速い)。
        var tooTilted = probes
            .Where(probe => DeskewPolicy.IsTooTilted(probe.SkewDegrees, probe.SkewReliable))
            .Select(probe => probe.Page)
            .ToList();

        if (tooTilted.Count > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"傾きが大きすぎるページがあります({Describe(tooTilted)})。"
                    + $"{DeskewPolicy.MaximumAngle:0.#} 度を超える傾きは直しません。"
                    + "まっすぐに取り込み直してください。"));

            return Blocked(pages, engineInfo, probes, issues);
        }

        // ── 2. ページごとに読む ────────────────────────
        var items = new List<OcrItem>();
        var deskewed = new List<int>();
        var tables = new List<int>();
        var forms = new List<int>();
        var tiltedTables = new List<int>();
        var done = 0;

        progress?.Report(new OcrProgress(0, pages.Count));

        foreach (var probe in probes)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var angle = options.Deskew
                && DeskewPolicy.ShouldDeskew(probe.SkewDegrees, probe.SkewReliable)
                    ? -probe.SkewDegrees
                    : 0;

            if (angle != 0)
            {
                deskewed.Add(probe.Page);
            }

            var read = source.Read(probe.Page, angle, cancellationToken);
            var kind = Decide(options, probe, read);

            // 傾いていた表は、まだ実用にならない。罫線の位置がわずかな残り傾き
            // (推定の誤差は平均 0.33 度)でずれ、行と列が崩れる。実測では
            // まっすぐな罫線表のセル一致 61.3% に対し、2 度傾いた表は 4.8% で、
            // 誤確定も 8 件出た。崩れた表をそれらしく出すより、理由を示して止める。
            //
            // 罫線の有無では分けない。傾いた罫線表は罫線そのものも拾えなくなり、
            // 「罫線なしの表」として同じように崩れるため。
            //
            // 「傾きを測れなかった表」も同じ扱いにする。表には文字が十分にあるので、
            // 本来は角度を測れる。測れないのは行が斜めに走って塊が繋がってしまうとき、
            // つまり傾いているときだった(実測: まっすぐな表は測れて、
            // 2 度傾いた表は測れず、罫線も 1 本も拾えなかった)。
            if (kind is ScanPageKind.RuledTable or ScanPageKind.BorderlessTable
                && (angle != 0 || !probe.SkewReliable))
            {
                tiltedTables.Add(probe.Page);
                done++;
                progress?.Report(new OcrProgress(done, pages.Count));
                continue;
            }

            switch (kind)
            {
                case ScanPageKind.RuledTable or ScanPageKind.BorderlessTable:
                    tables.Add(probe.Page);
                    items.AddRange(BuildTableItems(read, kind));
                    break;

                // 項目がまだ決まっていないときは、ふつうに文章として読む。
                // その読み取り結果から項目の候補を作れるようにするため。
                case ScanPageKind.FixedForm when options.Template is { Fields.Count: > 0 } template:
                    forms.Add(probe.Page);
                    items.AddRange(BuildFormItems(source, read, template, cancellationToken));
                    break;

                default:
                    items.AddRange(BuildItems(read.Page, Original(read)));
                    break;
            }

            done++;
            progress?.Report(new OcrProgress(done, pages.Count));
        }

        if (tiltedTables.Count > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"傾いた表があります({Describe(tiltedTables)})。"
                    + "傾いた表を行と列へ戻すのは次の段階で対応します。"
                    + "崩れた表を出すことはしません。"));
        }

        if (items.Count == 0 && tiltedTables.Count == 0)
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
            EngineInfo = engineInfo,
            NeedsDeskewPages = deskewed,
            TableLikePages = tables,
            FormPages = forms,
            ResolvedMode = forms.Count > 0
                ? OcrReadMode.FixedForm
                : tables.Count == pages.Count && tables.Count > 0
                    ? OcrReadMode.Table
                    : OcrReadMode.Lines,
            FieldNames = options.Template is { } fields
                ? [.. fields.Fields.Select(field => field.Name)]
                : [],
            Issues = issues,
        };
    }

    /// <summary>このページをどう組み立てるか。</summary>
    private static ScanPageKind Decide(
        OcrReadOptions options, OcrPageProbe probe, OcrPageRead read)
    {
        switch (options.Mode)
        {
            case OcrReadMode.Lines:
                return ScanPageKind.Prose;
            case OcrReadMode.Table:
                return read.ColumnRulings.Count > ScanTableBuilder.MinimumColumns
                    ? ScanPageKind.RuledTable
                    : ScanPageKind.BorderlessTable;
            case OcrReadMode.FixedForm:
                return ScanPageKind.FixedForm;
        }

        var rows = ScanTableBuilder.GroupRows(read.Lines);
        var columns = ScanTableBuilder.ColumnStarts(rows, read.Lines);
        var aligned = rows.Count(row => row.Count >= columns.Count && columns.Count >= 2);

        return ScanPageClassifier.Classify(new ScanPageMetrics(
            probe.HorizontalRulings,
            probe.VerticalRulings,
            rows.Count,
            aligned,
            columns.Count,
            probe.UnderlineCount));
    }

    /// <summary>表として組み立てる。座標は元のページへ戻してから項目にする。</summary>
    private static IReadOnlyList<OcrItem> BuildTableItems(OcrPageRead read, ScanPageKind kind)
    {
        var table = kind == ScanPageKind.RuledTable
            ? ScanTableBuilder.FromRulings(
                read.Lines, read.RowRulings, read.ColumnRulings, Fuse)
            : ScanTableBuilder.FromAlignment(read.Lines, Fuse);

        // 表として組み立てられなければ、文章として出す(捨てない)。
        if (table is null)
        {
            return BuildItems(read.Page, Original(read));
        }

        var items = new List<OcrItem>();

        // 列ごとの多数派の形。上下逆に読まれたセルは、両モデルが一致していても
        // 同じ列の他の行と文字の種類が食い違うので、そこで気づける。
        var majority = new Dictionary<int, ColumnShapeGuard.Shape>();
        foreach (var column in table.Cells.GroupBy(cell => cell.Column))
        {
            majority[column.Key] = ColumnShapeGuard.MajorityShape(
                column.Where(cell => !cell.IsEmpty).Select(cell => cell.Text));
        }

        foreach (var cell in table.Cells.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column))
        {
            // このセルを組み立てた読み取り(確認画面でモデルごとの読みを見せる)。
            var sources = read.Lines
                .Where(line => FormFieldExtractor.Overlap(line.Box, cell.Box) > 0.5)
                .ToList();

            var shapeOk = ColumnShapeGuard.CanAutoAccept(
                majority.GetValueOrDefault(cell.Column), cell.Text);

            var status = cell.IsEmpty
                ? OcrItemStatus.Unreadable
                : cell.Confidence >= OcrFusion.AutoAcceptThreshold
                    && sources.All(line => Agree(line))
                    && shapeOk
                        ? OcrItemStatus.AutoAccepted
                        : OcrItemStatus.NeedsReview;

            items.Add(new OcrItem
            {
                PageNumber = read.Page,
                LineNumber = cell.Row + 1,
                IndexInLine = cell.Column,
                Row = cell.Row,
                Column = cell.Column,
                Text = cell.Text,
                BoundingBox = read.Transform.ToOriginal(cell.Box),
                Confidence = cell.Confidence,
                Reason = cell.IsEmpty
                    ? "このセルからは何も読み取れませんでした"
                    : !shapeOk
                        ? ColumnShapeGuard.Reason
                        : $"{table.RowCount} 行 × {table.ColumnCount} 列の表の "
                            + $"{cell.Row + 1} 行 {cell.Column + 1} 列目",
                OriginalEngineResults = sources.Count == 0
                    ? [new OcrEngineReading(OcrFusion.MultiEngineName, string.Empty, 0)]
                    : [.. sources.SelectMany(EngineReadings)],
                InitialStatus = status,
                Status = status,
            });
        }

        return items;
    }

    /// <summary>
    /// 帳票として組み立てる。
    /// **指定した項目の数と、作る件数は必ず同じ。** 読み取りが見つからなくても
    /// 「見つからない」として 1 件残す(項目ごと消えるのを防ぐ)。
    /// </summary>
    private static IReadOnlyList<OcrItem> BuildFormItems(
        IOcrPageSource source,
        OcrPageRead read,
        FormTemplate template,
        CancellationToken cancellationToken)
    {
        // 指定した領域は「直した画像」の座標へ移してから使う。
        var deskewedTemplate = template with
        {
            Fields = [.. template.Fields.Select(field => field with
            {
                Area = read.Transform.ToDeskewed(field.Area),
                Choices = [.. field.Choices.Select(choice => choice with
                {
                    Area = read.Transform.ToDeskewed(choice.Area),
                })],
            })],
            Anchors = [.. template.Anchors.Select(anchor => anchor with
            {
                Area = read.Transform.ToDeskewed(anchor.Area),
            })],
        };

        var offset = FormFieldExtractor.FindOffset(
            deskewedTemplate, read.Lines, line => Fuse(line).Text);

        var readings = FormFieldExtractor.Read(deskewedTemplate, read.Lines, offset, Fuse);
        var marks = ReadMarks(source, read, deskewedTemplate, offset, cancellationToken);

        var items = new List<OcrItem>(readings.Count);
        var index = 0;

        foreach (var reading in readings)
        {
            var field = template.Fields[index];
            var place = read.Transform.ToOriginal(reading.Area);

            if (field.Kind == FormFieldKind.Choice)
            {
                var mark = marks[field.Name];
                items.Add(new OcrItem
                {
                    PageNumber = read.Page,
                    LineNumber = index + 1,
                    IndexInLine = 0,
                    FieldName = field.Name,
                    Text = mark.Text,
                    BoundingBox = place,
                    Confidence = mark.Confidence,
                    Reason = mark.Reason,
                    OriginalEngineResults =
                        [new OcrEngineReading("印の判定", mark.Text, mark.Confidence)],
                    InitialStatus = MarkClassifier.ToStatus(mark),
                    Status = MarkClassifier.ToStatus(mark),
                });

                index++;
                continue;
            }

            // 自信が足りていても、項目の種類として形が怪しければ自動確定しない。
            // 2 つのモデルは同じ字形の取り違えを共有するので、一致は根拠にならない。
            var shapeOk = FieldAutoAcceptPolicy.CanAutoAccept(field.Kind, reading.Text);

            var status = !reading.WasFound
                ? OcrItemStatus.Missing
                : reading.Confidence >= OcrFusion.AutoAcceptThreshold && shapeOk
                    ? OcrItemStatus.AutoAccepted
                    : OcrItemStatus.NeedsReview;

            var reason = reading.WasFound
                && !shapeOk
                && reading.Confidence >= OcrFusion.AutoAcceptThreshold
                ? FieldAutoAcceptPolicy.ReasonFor(field.Kind)
                : reading.Reason;

            items.Add(new OcrItem
            {
                PageNumber = read.Page,
                LineNumber = index + 1,
                IndexInLine = 0,
                FieldName = field.Name,
                IsMissing = !reading.WasFound,
                Text = reading.Text,
                BoundingBox = place,
                Confidence = reading.Confidence,
                Reason = reason,
                OriginalEngineResults =
                    [new OcrEngineReading(OcrFusion.MultiEngineName, reading.Text, reading.Confidence)],
                InitialStatus = status,
                Status = status,
            });

            index++;
        }

        return items;
    }

    private static Dictionary<string, MarkResult> ReadMarks(
        IOcrPageSource source,
        OcrPageRead read,
        FormTemplate template,
        FormOffset offset,
        CancellationToken cancellationToken)
    {
        var results = new Dictionary<string, MarkResult>(StringComparer.Ordinal);
        var choiceFields = template.Fields
            .Where(field => field.Kind == FormFieldKind.Choice && field.Choices.Count > 0)
            .ToList();

        if (choiceFields.Count == 0)
        {
            return results;
        }

        // 1 回の描画でまとめて測る。箱の中と、丸囲みの線が通る上下の帯。
        var areas = new List<OcrBox>();
        foreach (var field in choiceFields)
        {
            foreach (var choice in field.Choices)
            {
                var area = offset.Apply(choice.Area);
                areas.Add(Inset(area, 0.2));
                areas.Add(RingBand(area, above: true));
                areas.Add(RingBand(area, above: false));
            }
        }

        var ink = source.InkRatios(
            read.Page, areas, read.Transform.AngleDegrees, cancellationToken);

        var cursor = 0;
        foreach (var field in choiceFields)
        {
            var samples = new List<MarkSample>();
            foreach (var choice in field.Choices)
            {
                samples.Add(new MarkSample(
                    choice.Label, ink[cursor], Math.Max(ink[cursor + 1], ink[cursor + 2])));
                cursor += 3;
            }

            results[field.Name] = MarkClassifier.Classify(samples);
        }

        return results;
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
                    OriginalEngineResults = [.. EngineReadings(region)],
                    InitialStatus = result.Status,
                    Status = result.Status,
                });
            }
        }

        return items;
    }

    /// <summary>読み取り位置を元のページの座標へ戻した行。</summary>
    private static IReadOnlyList<OcrRawLine> Original(OcrPageRead read)
        => read.Transform.IsIdentity
            ? read.Lines
            : [.. read.Lines.Select(line => line with { Box = read.Transform.ToOriginal(line.Box) })];

    private static (string Text, double Confidence) Fuse(OcrRawLine line)
    {
        var result = OcrFusion.Fuse(line);
        return (result.Text, result.Confidence);
    }

    private static bool Agree(OcrRawLine line)
        => OcrFusion.Fuse(line).Status == OcrItemStatus.AutoAccepted;

    private static IEnumerable<OcrEngineReading> EngineReadings(OcrRawLine region)
    {
        yield return new OcrEngineReading(
            OcrFusion.MultiEngineName,
            PdfTextNormalization.Normalize(region.MultiText),
            OcrFusion.Finite(region.MultiScore));
        yield return new OcrEngineReading(
            OcrFusion.JapanEngineName,
            PdfTextNormalization.Normalize(region.JapanText),
            OcrFusion.Finite(region.JapanScore));
    }

    /// <summary>枠線を避けて内側を測る。</summary>
    private static OcrBox Inset(OcrBox box, double ratio)
    {
        var dx = box.Width * ratio;
        var dy = box.Height * ratio;
        return new OcrBox(box.X + dx, box.Y + dy, box.Width - (dx * 2), box.Height - (dy * 2));
    }

    /// <summary>
    /// 丸囲みの線が通る、ラベルの上下の細い帯。
    ///
    /// 丸は箱ではなく**ラベルの側**に描かれる。広い四角で測ると、細い線が
    /// 面積に埋もれて濃さがほとんど出ない(実測: 丸の回は濃さ約 0.03 にしかならず、
    /// 6 回とも「薄い」と判断されて 0/6 だった)。
    /// 線が通るところだけを細い帯で測れば、同じ線でも濃さが十分に出る。
    /// ラベルの文字そのものは帯に入らないので、印が無いときは薄いままになる。
    /// </summary>
    private static OcrBox RingBand(OcrBox box, bool above)
    {
        var height = Math.Max(box.Height * 0.25, 2);
        var top = above
            ? box.Y - (box.Height * 0.45)
            : box.Y + box.Height + (box.Height * 0.2);

        return new OcrBox(
            box.X + (box.Width * 0.9), top, box.Width * 4.5, height);
    }

    private static OcrDocumentReading Blocked(
        IReadOnlyList<int> pages,
        OcrEngineInfo engineInfo,
        List<OcrPageProbe> probes,
        List<MergeIssue> issues)
        => new()
        {
            Items = [],
            OcrPages = pages,
            EngineInfo = engineInfo,
            NeedsDeskewPages = [.. probes
                .Where(probe => DeskewPolicy.ShouldDeskew(probe.SkewDegrees, probe.SkewReliable))
                .Select(probe => probe.Page)],
            Issues = issues,
        };

    private static string Describe(IReadOnlyList<int> pages)
        => pages.Count <= 5
            ? string.Join(" / ", pages.Select(page => $"{page} ページ目"))
            : string.Join(" / ", pages.Take(5).Select(page => $"{page} ページ目")) + " ほか";
}
