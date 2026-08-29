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
        var formPages = new List<FormPageReading>();
        var done = 0;

        progress?.Report(new OcrProgress(0, pages.Count));

        foreach (var probe in probes)
        {
            cancellationToken.ThrowIfCancellationRequested();

            // 回す向きは実測で決めた。傾き 1.98 度と測ったページを -1.98 度回すと
            // 罫線が 1 本も拾えず(rows=0)、+1.98 度回すと 22 行 3 列の格子が
            // そのまま出てきた。**符号を逆にしていたので、傾きを直すつもりで
            // 倍にしていた**(2F-B2 の「傾いた表 4.8%」「帳票 傾き 50.8%」は
            // これが原因)。
            var angle = options.Deskew
                && DeskewPolicy.ShouldDeskew(probe.SkewDegrees, probe.SkewReliable)
                    ? probe.SkewDegrees
                    : 0;

            if (angle != 0)
            {
                deskewed.Add(probe.Page);
            }

            var read = source.Read(probe.Page, angle, cancellationToken);
            var kind = Decide(options, probe, read);

            switch (kind)
            {
                // 傾いた表も、傾きを直したうえで行と列へ戻すところまで通す
                // (Phase 2F-B2 では止めていた。当時はまっすぐな表でもセル一致
                //  61.3% しか無く、傾けると 4.8% まで落ちたため。
                //  Paddle Inference を 3.3.1 へ上げてまっすぐが 94.3% になったので
                //  測り直した)。
                //
                // 戻せたかどうかは、出来上がった表の形で判断する。
                // 崩れているとき(セルがほとんど空、行や列が少なすぎる)は、
                // それらしい表を出さずに理由を示して止める。
                case ScanPageKind.RuledTable or ScanPageKind.BorderlessTable:
                {
                    var built = BuildTableItems(source, read, kind, angle, cancellationToken);
                    if (built is null)
                    {
                        tiltedTables.Add(probe.Page);
                        break;
                    }

                    tables.Add(probe.Page);
                    items.AddRange(built);
                    break;
                }

                // 項目がまだ決まっていないときは、ふつうに文章として読む。
                // その読み取り結果から項目の候補を作れるようにするため。
                case ScanPageKind.FixedForm when options.Template is { Fields.Count: > 0 } template:
                    // 項目ごとの「その項目らしい形」を全ページから学んでから状態を決めるので、
                    // ここでは読み取りだけ貯めて、items はページを読み終えてから作る。
                    forms.Add(probe.Page);
                    formPages.Add(ReadFormPage(source, read, template, cancellationToken));
                    break;

                default:
                    items.AddRange(BuildItems(read.Page, Original(read)));
                    break;
            }

            done++;
            progress?.Report(new OcrProgress(done, pages.Count));
        }

        if (formPages.Count > 0 && options.Template is { Fields.Count: > 0 } formTemplate)
        {
            items.AddRange(BuildFormItems(formPages, formTemplate));
            items.Sort((left, right) =>
            {
                var byPage = left.PageNumber.CompareTo(right.PageNumber);
                if (byPage != 0)
                {
                    return byPage;
                }

                var byLine = left.LineNumber.CompareTo(right.LineNumber);
                return byLine != 0 ? byLine : left.IndexInLine.CompareTo(right.IndexInLine);
            });
        }

        if (tiltedTables.Count > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"表の行と列を戻せないページがあります({Describe(tiltedTables)})。"
                    + "傾きが大きい、罫線が読み取れない、"
                    + "などの理由で表の形が定まりませんでした。"
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
    /// <summary>
    /// 戻した表が使いものになっていないか。
    ///
    /// 傾きを直しきれないと、罫線の位置が少しずつずれて区画が噛み合わなくなり、
    /// 「行と列はあるのに、ほとんどのセルが空」という形になる。
    /// この形のまま出すと、表として正しそうに見えて中身が抜け落ちる。
    /// 中身が入っているセルが半分に満たなければ、表として扱わない。
    /// </summary>
    /// <summary>この割合を超える黒い画素があれば、その区画には文字があったとみなす。</summary>
    internal const double BlankCellInkRatio = 0.01;

    /// <summary>
    /// 罫線から作った格子のうち、読み取りが 1 つも割り当たらなかった区画。
    /// 位置は格子から作るので、読めていなくても確認画面で場所を示せる。
    /// </summary>
    internal static IReadOnlyList<(int Row, int Column, OcrBox Box)> BlankCells(ScanTable table)
    {
        var taken = table.Cells
            .Where(cell => !cell.IsEmpty)
            .Select(cell => (cell.Row, cell.Column))
            .ToHashSet();

        var byRow = table.Cells.GroupBy(cell => cell.Row)
            .ToDictionary(group => group.Key, group => ScanTableBuilder.Union([.. group.Select(c => c.Box)]));
        var byColumn = table.Cells.GroupBy(cell => cell.Column)
            .ToDictionary(group => group.Key, group => ScanTableBuilder.Union([.. group.Select(c => c.Box)]));

        var blanks = new List<(int, int, OcrBox)>();
        for (var row = 0; row < table.RowCount; row++)
        {
            if (!byRow.TryGetValue(row, out var rowBox))
            {
                continue;
            }

            for (var column = 0; column < table.ColumnCount; column++)
            {
                if (taken.Contains((row, column))
                    || !byColumn.TryGetValue(column, out var columnBox))
                {
                    continue;
                }

                blanks.Add((row, column,
                    new OcrBox(columnBox.X, rowBox.Y, columnBox.Width, rowBox.Height)));
            }
        }

        return blanks;
    }

    internal static bool IsBrokenTable(ScanTable table)
    {
        var grid = table.RowCount * table.ColumnCount;
        if (grid == 0)
        {
            return true;
        }

        var filled = table.Cells.Count(cell => !cell.IsEmpty);
        return (double)filled / grid < BrokenTableFilledRatio;
    }

    /// <summary>中身のあるセルがこの割合に満たなければ、表として扱わない。</summary>
    internal const double BrokenTableFilledRatio = 0.5;

    private static IReadOnlyList<OcrItem>? BuildTableItems(
        IOcrPageSource source,
        OcrPageRead read,
        ScanPageKind kind,
        double angle,
        CancellationToken cancellationToken)
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

        if (IsBrokenTable(table))
        {
            return null;
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

        // 文字が 1 つも割り当たらなかった区画を拾う。
        //
        // ここを飛ばすと、**元の表には文字があるのに検出されなかったセル**が
        // 黙って空欄として出てしまう(Phase 2F の最重要不変条件に反する)。
        // かといって空欄をすべて人へ回すと、もともと空のセルが多い表で
        // 確認の手間が現実的でなくなる。
        //
        // そこで、その区画に**黒い画素があるか**を測って分ける。
        // 文字があるのに読めなかった区画だけを「読取不能」として人へ回し、
        // もともと空の区画はそのまま空欄として出す。
        var blanks = BlankCells(table);
        var inked = blanks.Count == 0
            ? []
            : source.InkRatios(read.Page, [.. blanks.Select(b => b.Box)], angle, cancellationToken);

        for (var index = 0; index < blanks.Count; index++)
        {
            var ratio = index < inked.Count ? inked[index] : 0;
            if (ratio < BlankCellInkRatio)
            {
                continue;
            }

            var blank = blanks[index];
            items.Add(new OcrItem
            {
                PageNumber = read.Page,
                LineNumber = blank.Row + 1,
                IndexInLine = blank.Column,
                Row = blank.Row,
                Column = blank.Column,
                Text = string.Empty,
                BoundingBox = read.Transform.ToOriginal(blank.Box),
                Confidence = 0,
                Reason = "このセルには文字があるようですが、読み取れませんでした。"
                    + "元のページと見比べて入力してください。",
                OriginalEngineResults =
                    [new OcrEngineReading(OcrFusion.MultiEngineName, string.Empty, 0)],
                InitialStatus = OcrItemStatus.Unreadable,
                Status = OcrItemStatus.Unreadable,
            });
        }

        foreach (var cell in table.Cells.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column))
        {
            // このセルを組み立てた読み取り(確認画面でモデルごとの読みを見せる)。
            var sources = read.Lines
                .Where(line => FormFieldExtractor.Overlap(line.Box, cell.Box) > 0.5)
                .ToList();

            var shapeOk = ColumnShapeGuard.CanAutoAccept(
                majority.GetValueOrDefault(cell.Column), cell.Text)
                && !FieldAutoAcceptPolicy.IsUpsideDownAmbiguous(cell.Text);

            var status = cell.IsEmpty
                ? OcrItemStatus.Unreadable
                : cell.Confidence >= OcrFusion.AutoAcceptThreshold
                    && sources.All(line => Agree(line))
                    && shapeOk
                    && !cell.IsMerged
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
                    : cell.IsMerged
                        ? "縦に離れた 2 か所を 1 つのセルにまとめています。"
                            + "行の区切りを取り違えている可能性があります。"
                            + "元のページと見比べてください。"
                        : FieldAutoAcceptPolicy.IsUpsideDownAmbiguous(cell.Text)
                            ? FieldAutoAcceptPolicy.UpsideDownReason
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
    /// <summary>1 ページぶんの帳票の読み取り(状態はまだ決めない)。</summary>
    private sealed record FormPageReading(
        int Page,
        DeskewTransform Transform,
        IReadOnlyList<FormFieldReading> Readings,
        IReadOnlyDictionary<string, MarkResult> Marks);

    private static FormPageReading ReadFormPage(
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

        return new FormPageReading(
            read.Page,
            read.Transform,
            FormFieldExtractor.Read(deskewedTemplate, read.Lines, offset, Fuse),
            ReadMarks(source, read, deskewedTemplate, offset, cancellationToken));
    }

    private static IReadOnlyList<OcrItem> BuildFormItems(
        IReadOnlyList<FormPageReading> pages, FormTemplate template)
    {
        // 項目ごとに「その項目らしい形」を全ページの読みから学ぶ。
        // 学んだ形は**自動確定してよいかの判断にだけ**使い、値は書き換えない。
        var patterns = new Dictionary<string, string?>();
        for (var field = 0; field < template.Fields.Count; field++)
        {
            var name = template.Fields[field].Name;
            patterns[name] = FieldShapePattern.Learn(
                pages.Where(page => field < page.Readings.Count)
                    .Where(page => page.Readings[field].WasFound)
                    .Select(page => page.Readings[field].Text));
        }

        var items = new List<OcrItem>(pages.Count * template.Fields.Count);

        foreach (var page in pages)
        {
            for (var index = 0; index < page.Readings.Count; index++)
            {
                var reading = page.Readings[index];
                var field = template.Fields[index];
                var place = page.Transform.ToOriginal(reading.Area);

                if (field.Kind == FormFieldKind.Choice)
                {
                    var mark = page.Marks[field.Name];
                    items.Add(new OcrItem
                    {
                        PageNumber = page.Page,
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

                    continue;
                }

                // 自信が足りていても、形が怪しければ自動確定しない。
                // 2 つのモデルは同じ字形の取り違えを共有するので、一致は根拠にならない。
                //
                // 判断は 2 段構え。全ページから形を学べた項目は、その形どおりに
                // 読めていれば自動確定してよい(取り違えやすい字を含んでいても、
                // 同じ形で 120 ページ読めているならその読みは信用できる)。
                // 形を学べなかった項目だけ、B2 の粗い決まりへ落とす。
                var pattern = patterns.GetValueOrDefault(field.Name);
                var byPattern = pattern is not null;
                var shapeOk = (byPattern
                        ? FieldShapePattern.Matches(pattern, reading.Text)
                        : FieldAutoAcceptPolicy.CanAutoAccept(field.Kind, reading.Text))
                    && !FieldAutoAcceptPolicy.IsUpsideDownAmbiguous(reading.Text);

                var status = !reading.WasFound
                    ? OcrItemStatus.Missing
                    : reading.Confidence >= OcrFusion.AutoAcceptThreshold && shapeOk
                        ? OcrItemStatus.AutoAccepted
                        : OcrItemStatus.NeedsReview;

                var reason = reading.WasFound
                    && !shapeOk
                    && reading.Confidence >= OcrFusion.AutoAcceptThreshold
                    ? byPattern
                        ? FieldShapePattern.Reason
                        : FieldAutoAcceptPolicy.ReasonFor(field.Kind)
                    : reading.Reason;

                items.Add(new OcrItem
                {
                    PageNumber = page.Page,
                    LineNumber = index + 1,
                    IndexInLine = 0,
                    FieldName = field.Name,
                    IsMissing = !reading.WasFound,
                    Text = reading.Text,
                    BoundingBox = place,
                    Confidence = reading.Confidence,
                    Reason = reason,
                    OriginalEngineResults =
                        [new OcrEngineReading(
                            OcrFusion.MultiEngineName, reading.Text, reading.Confidence)],
                    InitialStatus = status,
                    Status = status,
                });
            }
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
