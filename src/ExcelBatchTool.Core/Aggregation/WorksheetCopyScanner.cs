using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>
/// 1 つの Worksheet を集約できるか調べた結果と、書き出しに必要な構造メタデータ。
/// セルの中身は保持しない(書き出し時にストリーミングで読み直す)。
/// </summary>
internal sealed class SheetCopyScan
{
    /// <summary>このシートを集約できない理由。1 件でもあれば Block。</summary>
    public IReadOnlyList<string> BlockReasons { get; init; } = Array.Empty<string>();

    /// <summary>集約はできるが、利用者に伝えるべきこと。</summary>
    public IReadOnlyList<string> WarningReasons { get; init; } = Array.Empty<string>();

    public bool IsBlocked => BlockReasons.Count > 0;

    /// <summary>元シートの表示状態(visible / hidden / veryHidden)。そのまま出力へ引き継ぐ。</summary>
    public SheetVisibility Visibility { get; init; } = SheetVisibility.Visible;

    public string? DimensionReference { get; init; }

    /// <summary>行の既定の高さなど。</summary>
    public SheetFormatProperties? SheetFormat { get; init; }

    /// <summary>列の幅・非表示などの定義。</summary>
    public IReadOnlyList<Column> Columns { get; init; } = Array.Empty<Column>();

    /// <summary>ウィンドウ枠の固定。</summary>
    public Pane? FreezePane { get; init; }

    public IReadOnlyList<Selection> Selections { get; init; } = Array.Empty<Selection>();

    public bool ShowGridLines { get; init; } = true;

    public bool ShowRowColHeaders { get; init; } = true;

    public bool RightToLeft { get; init; }

    public uint? ZoomScale { get; init; }

    /// <summary>シートの保護設定。</summary>
    public SheetProtection? Protection { get; init; }

    /// <summary>結合セルの範囲(A1 形式)。</summary>
    public IReadOnlyList<string> MergeReferences { get; init; } = Array.Empty<string>();

    /// <summary>印刷の拡大縮小設定(sheetPr/pageSetUpPr)。</summary>
    public PageSetupProperties? PageSetupProperties { get; init; }

    public PrintOptions? PrintOptions { get; init; }

    public PageMargins? PageMargins { get; init; }

    /// <summary>プリンター固有の設定(r:id)を持たない pageSetup のみ保持する。</summary>
    public PageSetup? PageSetup { get; init; }

    /// <summary>文字列だけのヘッダー・フッター。</summary>
    public HeaderFooter? HeaderFooter { get; init; }

    public RowBreaks? RowBreaks { get; init; }

    public ColumnBreaks? ColumnBreaks { get; init; }

    /// <summary>印刷範囲・印刷タイトル(シート名を除いた範囲部分)。</summary>
    public IReadOnlyList<PrintDefinedNameInfo> PrintDefinedNames { get; init; }
        = Array.Empty<PrintDefinedNameInfo>();

    /// <summary>ハイパーリンク。別シート宛の解決は Planner が行う。</summary>
    public IReadOnlyList<HyperlinkInfo> Hyperlinks { get; init; } = Array.Empty<HyperlinkInfo>();
}

/// <summary>解析済みの印刷範囲・印刷タイトル。範囲はシート名を含まない。</summary>
internal sealed record PrintDefinedNameInfo(PrintDefinedNameKind Kind, IReadOnlyList<string> Ranges);

/// <summary>
/// Worksheet を集約できるか検証する。対象ファイルは読み取り専用でしか開かない。
/// Phase 1B.1 で保持できない要素を見つけたら、黙って落とさず Block 理由として返す。
/// </summary>
internal static class WorksheetCopyScanner
{
    /// <summary>Workbook 単位の検証結果。</summary>
    public sealed record WorkbookScan(IReadOnlyList<string> BlockReasons, string? ThemeFingerprint);

    /// <summary>Workbook 全体を集約対象にできない場合の理由(シートを切り離すと意味が壊れるもの)。</summary>
    public static WorkbookScan ScanWorkbook(string filePath)
    {
        if (!File.Exists(filePath))
        {
            return new WorkbookScan(["ファイルが見つかりません。"], null);
        }

        if (!string.Equals(Path.GetExtension(filePath), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return new WorkbookScan(["現在のバージョンで扱えるのは .xlsx のみです。"], null);
        }

        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var document = SpreadsheetDocument.Open(stream, isEditable: false);

            var workbookPart = document.WorkbookPart;
            if (workbookPart is null)
            {
                return new WorkbookScan(["Workbook 情報が見つかりません。"], null);
            }

            var reasons = new List<string>();

            if (workbookPart.VbaProjectPart is not null
                || workbookPart.ContentType.Contains("macroEnabled", StringComparison.OrdinalIgnoreCase))
            {
                reasons.Add("マクロ (VBA) を含むため、Phase 1B.1 では集約できません。");
            }

            var hasExternalLink = workbookPart.ExternalWorkbookParts.Any()
                || (workbookPart.Workbook?.GetFirstChild<ExternalReferences>()?
                    .Elements<ExternalReference>().Any() ?? false)
                || workbookPart.ConnectionsPart is not null;

            if (hasExternalLink)
            {
                reasons.Add("他のブックへの外部参照(外部リンク)を含むため、Phase 1B.1 では集約できません。");
            }

            return new WorkbookScan(reasons, ComputeThemeFingerprint(workbookPart));
        }
        catch (Exception ex) when (ex is InvalidDataException or FileFormatException or OpenXmlPackageException)
        {
            return new WorkbookScan(
                ["ファイルを読み取れません。パスワード保護(暗号化)されているか、破損している可能性があります。"], null);
        }
        catch (Exception ex)
        {
            return new WorkbookScan([$"読み取りエラー: {ex.Message}"], null);
        }
    }

    /// <summary>テーマ(配色定義)の同一性を比べるための指紋。テーマが無い場合は null。</summary>
    private static string? ComputeThemeFingerprint(WorkbookPart workbookPart)
    {
        var themePart = workbookPart.ThemePart;
        if (themePart is null)
        {
            return null;
        }

        try
        {
            using var stream = themePart.GetStream(FileMode.Open, FileAccess.Read);
            return Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>選択された 1 シートを走査する。Workbook 単位の問題は <see cref="ScanWorkbook"/> で扱う。</summary>
    public static SheetCopyScan ScanSheet(string filePath, string sheetName, CancellationToken cancellationToken)
    {
        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var document = SpreadsheetDocument.Open(stream, isEditable: false);

            var workbookPart = document.WorkbookPart;
            if (workbookPart is null)
            {
                return Blocked("Workbook 情報が見つかりません。");
            }

            var sheet = workbookPart.Workbook?.Sheets?.Elements<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal));

            if (sheet?.Id?.Value is not { } relationshipId)
            {
                return Blocked($"ワークシート「{sheetName}」が見つかりません。");
            }

            OpenXmlPart? part;
            try
            {
                part = workbookPart.GetPartById(relationshipId);
            }
            catch (ArgumentOutOfRangeException)
            {
                part = null;
            }

            if (part is not WorksheetPart worksheetPart)
            {
                return Blocked("グラフシート・マクロシート等は集約対象にできません(通常のワークシートのみ)。");
            }

            var visibility = WorkbookAnalyzer.ResolveVisibility(sheet.State?.Value);
            var sheetIndex = workbookPart.Workbook!.Sheets!.Elements<Sheet>().ToList().IndexOf(sheet);

            return ScanWorksheetPart(
                worksheetPart, workbookPart, sheetName, sheetIndex, visibility, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex) when (ex is InvalidDataException or FileFormatException or OpenXmlPackageException)
        {
            return Blocked("ファイルを読み取れません。パスワード保護(暗号化)されているか、破損している可能性があります。");
        }
        catch (Exception ex)
        {
            return Blocked($"読み取りエラー: {ex.Message}");
        }

        static SheetCopyScan Blocked(string reason) => new() { BlockReasons = [reason] };
    }

    private static SheetCopyScan ScanWorksheetPart(
        WorksheetPart worksheetPart,
        WorkbookPart workbookPart,
        string sheetName,
        int sheetIndex,
        SheetVisibility visibility,
        CancellationToken cancellationToken)
    {
        var blocks = new List<string>();
        var warnings = new List<string>();

        AddPartLevelBlocks(worksheetPart, blocks);

        var context = WorkbookReadContext.Create(workbookPart);

        string? dimension = null;
        SheetFormatProperties? sheetFormat = null;
        var columns = new List<Column>();
        Pane? pane = null;
        var selections = new List<Selection>();
        var showGridLines = true;
        var showRowColHeaders = true;
        var rightToLeft = false;
        uint? zoomScale = null;
        SheetProtection? protection = null;
        var mergeReferences = new List<string>();

        var hyperlinks = new List<HyperlinkInfo>();
        PageSetupProperties? pageSetupProperties = null;
        PrintOptions? printOptions = null;
        PageMargins? pageMargins = null;
        PageSetup? pageSetup = null;
        HeaderFooter? headerFooter = null;
        RowBreaks? rowBreaks = null;
        ColumnBreaks? columnBreaks = null;

        var hasFormula = false;
        var hasRichText = false;
        var hasPhoneticText = false;
        var sheetViewSeen = false;

        using (var reader = OpenXmlReader.Create(worksheetPart))
        {
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (!reader.IsStartElement)
                {
                    continue;
                }

                var type = reader.ElementType;

                if (type == typeof(Cell))
                {
                    var cell = (Cell)reader.LoadCurrentElement()!;
                    hasFormula |= cell.CellFormula is not null;
                    hasRichText |= context.ReferencesRichText(cell);
                    hasPhoneticText |= context.ReferencesPhoneticText(cell);
                }
                else if (type == typeof(SheetDimension))
                {
                    dimension = ((SheetDimension)reader.LoadCurrentElement()!).Reference?.Value;
                }
                else if (type == typeof(SheetFormatProperties))
                {
                    sheetFormat = (SheetFormatProperties)reader.LoadCurrentElement()!.CloneNode(true);
                }
                else if (type == typeof(Columns))
                {
                    var element = (Columns)reader.LoadCurrentElement()!;
                    columns.AddRange(element.Elements<Column>().Select(column => (Column)column.CloneNode(true)));
                }
                else if (type == typeof(SheetView))
                {
                    var view = (SheetView)reader.LoadCurrentElement()!;
                    if (!sheetViewSeen)
                    {
                        sheetViewSeen = true;
                        pane = view.GetFirstChild<Pane>() is { } sourcePane
                            ? (Pane)sourcePane.CloneNode(true)
                            : null;
                        selections.AddRange(view.Elements<Selection>()
                            .Select(selection => (Selection)selection.CloneNode(true)));
                        showGridLines = view.ShowGridLines?.Value ?? true;
                        showRowColHeaders = view.ShowRowColHeaders?.Value ?? true;
                        rightToLeft = view.RightToLeft?.Value ?? false;
                        zoomScale = view.ZoomScale?.Value;
                    }
                }
                else if (type == typeof(SheetProtection))
                {
                    protection = (SheetProtection)reader.LoadCurrentElement()!.CloneNode(true);
                }
                else if (type == typeof(MergeCell))
                {
                    var reference = ((MergeCell)reader.LoadCurrentElement()!).Reference?.Value;
                    if (string.IsNullOrWhiteSpace(reference) || !CellRangeParser.TryParseRange(reference, out _))
                    {
                        blocks.Add("結合セルの定義を解釈できません(壊れている可能性があります)。");
                    }
                    else
                    {
                        mergeReferences.Add(reference);
                    }
                }
                else if (type == typeof(ConditionalFormatting))
                {
                    AddOnce(blocks, "条件付き書式を含むため、Phase 1B.1 では集約できません。");
                }
                else if (type == typeof(DataValidation) || type == typeof(DataValidations))
                {
                    AddOnce(blocks, "データの入力規則を含むため、Phase 1B.1 では集約できません。");
                }
                else if (type == typeof(Hyperlink))
                {
                    var hyperlink = (Hyperlink)reader.LoadCurrentElement()!;
                    hyperlinks.Add(HyperlinkScanner.Scan(hyperlink, sheetName, worksheetPart));
                }
                else if (type == typeof(AutoFilter))
                {
                    AddOnce(blocks, "オートフィルターを含むため、Phase 1B.1 では集約できません。");
                }
                else if (type == typeof(PrintOptions))
                {
                    printOptions = (PrintOptions)reader.LoadCurrentElement()!.CloneNode(true);
                }
                else if (type == typeof(PageMargins))
                {
                    pageMargins = (PageMargins)reader.LoadCurrentElement()!.CloneNode(true);
                }
                else if (type == typeof(PageSetup))
                {
                    var setup = (PageSetup)reader.LoadCurrentElement()!;
                    if (setup.Id?.Value is not null)
                    {
                        // r:id はプリンター設定パート (devMode) への参照。
                        // そのまま持ち込むことも黙って外すこともできないので Block する。
                        AddOnce(blocks,
                            "プリンター固有の設定を含むため、現在のバージョンでは安全に集約できません。");
                    }
                    else
                    {
                        pageSetup = (PageSetup)setup.CloneNode(true);
                    }
                }
                else if (type == typeof(HeaderFooter))
                {
                    var element = (HeaderFooter)reader.LoadCurrentElement()!;
                    if (ContainsHeaderFooterImage(element))
                    {
                        AddOnce(blocks,
                            "ヘッダー・フッターに画像を含むため、現在のバージョンでは集約できません。");
                    }
                    else
                    {
                        headerFooter = (HeaderFooter)element.CloneNode(true);
                    }
                }
                else if (type == typeof(RowBreaks))
                {
                    rowBreaks = (RowBreaks)reader.LoadCurrentElement()!.CloneNode(true);
                }
                else if (type == typeof(ColumnBreaks))
                {
                    columnBreaks = (ColumnBreaks)reader.LoadCurrentElement()!.CloneNode(true);
                }
                else if (type == typeof(DrawingHeaderFooter) || type == typeof(LegacyDrawingHeaderFooter))
                {
                    AddOnce(blocks, "ヘッダー・フッターに画像を含むため、現在のバージョンでは集約できません。");
                }
                else if (type == typeof(SheetProperties))
                {
                    var properties = (SheetProperties)reader.LoadCurrentElement()!;
                    if (properties.PageSetupProperties is { } setupProperties)
                    {
                        pageSetupProperties = (PageSetupProperties)setupProperties.CloneNode(true);
                    }

                    if (properties.TabColor is not null)
                    {
                        AddOnce(warnings, "シート見出しの色は引き継がれません。");
                    }
                }
                else if (type == typeof(PhoneticProperties))
                {
                    AddOnce(warnings, "ふりがな(phonetic)情報は引き継がれません。");
                }
            }
        }

        if (hasFormula)
        {
            blocks.Add("数式を含むため、Phase 1B.1 では集約できません(参照先が壊れる可能性があるため)。");
        }

        if (hasRichText)
        {
            blocks.Add("文字単位で書式が設定されたセル(リッチテキスト)を含むため、Phase 1B.1 では集約できません。");
        }

        if (hasPhoneticText)
        {
            AddOnce(warnings, "ふりがな(phonetic)情報は引き継がれません。");
        }

        if (HasDuplicateOrOverlappingMerges(mergeReferences))
        {
            blocks.Add("結合セルの範囲が重複しています(定義が壊れている可能性があります)。");
        }

        if (FindBrokenBreak(rowBreaks, maxId: 1_048_575) is { } rowBreakError)
        {
            blocks.Add($"行の改ページ位置の定義が壊れています({rowBreakError})。");
        }

        if (FindBrokenBreak(columnBreaks, maxId: 16_383) is { } columnBreakError)
        {
            blocks.Add($"列の改ページ位置の定義が壊れています({columnBreakError})。");
        }

        var printDefinedNames = ReadPrintDefinedNames(workbookPart, sheetName, sheetIndex, blocks);

        return new SheetCopyScan
        {
            BlockReasons = blocks,
            WarningReasons = warnings,
            Visibility = visibility,
            PageSetupProperties = pageSetupProperties,
            PrintOptions = printOptions,
            PageMargins = pageMargins,
            PageSetup = pageSetup,
            HeaderFooter = headerFooter,
            RowBreaks = rowBreaks,
            ColumnBreaks = columnBreaks,
            PrintDefinedNames = printDefinedNames,
            Hyperlinks = hyperlinks,
            DimensionReference = dimension,
            SheetFormat = sheetFormat,
            Columns = columns,
            FreezePane = pane,
            Selections = selections,
            ShowGridLines = showGridLines,
            ShowRowColHeaders = showRowColHeaders,
            RightToLeft = rightToLeft,
            ZoomScale = zoomScale,
            Protection = protection,
            MergeReferences = mergeReferences,
        };
    }

    private static void AddPartLevelBlocks(WorksheetPart worksheetPart, List<string> blocks)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart is not null)
        {
            if (drawingsPart.ChartParts.Any())
            {
                blocks.Add("グラフを含むため、Phase 1B.1 では集約できません。");
            }

            if (drawingsPart.ImageParts.Any())
            {
                blocks.Add("画像を含むため、Phase 1B.1 では集約できません。");
            }

            if (!drawingsPart.ChartParts.Any() && !drawingsPart.ImageParts.Any())
            {
                blocks.Add("図形を含むため、Phase 1B.1 では集約できません。");
            }
        }

        if (worksheetPart.ImageParts.Any())
        {
            blocks.Add("シートの背景画像を含むため、Phase 1B.1 では集約できません。");
        }

        if (worksheetPart.VmlDrawingParts.Any())
        {
            blocks.Add("旧形式の図形・コメント枠を含むため、Phase 1B.1 では集約できません。");
        }

        if (worksheetPart.TableDefinitionParts.Any())
        {
            blocks.Add("テーブル(ListObject)を含むため、Phase 1B.1 では集約できません。");
        }

        if (worksheetPart.PivotTableParts.Any())
        {
            blocks.Add("ピボットテーブルを含むため、Phase 1B.1 では集約できません。");
        }

        if (worksheetPart.WorksheetCommentsPart is not null
            || worksheetPart.GetPartsOfType<WorksheetThreadedCommentsPart>().Any())
        {
            blocks.Add("コメントを含むため、Phase 1B.1 では集約できません。");
        }

        if (worksheetPart.EmbeddedObjectParts.Any() || worksheetPart.EmbeddedPackageParts.Any())
        {
            blocks.Add("埋め込みオブジェクト (OLE) を含むため、Phase 1B.1 では集約できません。");
        }

        if (worksheetPart.EmbeddedControlPersistenceParts.Any() || worksheetPart.ControlPropertiesParts.Any())
        {
            blocks.Add("ActiveX コントロールを含むため、Phase 1B.1 では集約できません。");
        }
    }

    /// <summary>
    /// ヘッダー・フッターの書式コード「&amp;G」は画像の差し込みを表す。
    /// 画像は移植しないので、文字だけコピーして黙って落とさないよう検出する。
    /// </summary>
    private static bool ContainsHeaderFooterImage(HeaderFooter headerFooter)
        => headerFooter.Descendants<OpenXmlLeafTextElement>()
            .Any(element => element.Text?.Contains("&G", StringComparison.Ordinal) == true);

    /// <summary>改ページ定義の壊れを探す。問題が無ければ null。</summary>
    private static string? FindBrokenBreak(OpenXmlCompositeElement? breaks, uint maxId)
    {
        if (breaks is null)
        {
            return null;
        }

        foreach (var item in breaks.Elements<Break>())
        {
            if (item.Id?.Value is not { } id)
            {
                return "位置が指定されていません";
            }

            if (id == 0 || id > maxId)
            {
                return $"位置 {id} は範囲外です";
            }

            if (item.Min?.Value is { } min && item.Max?.Value is { } max && min > max)
            {
                return $"範囲が逆転しています({min} > {max})";
            }
        }

        return null;
    }

    /// <summary>
    /// このシートに紐づく Defined Name を読む。印刷範囲・印刷タイトルだけを対象にし、
    /// それ以外のシート固有 Defined Name や解釈できない参照は Block 理由にする。
    /// </summary>
    private static List<PrintDefinedNameInfo> ReadPrintDefinedNames(
        WorkbookPart workbookPart,
        string sheetName,
        int sheetIndex,
        List<string> blocks)
    {
        var results = new List<PrintDefinedNameInfo>();
        var definedNames = workbookPart.Workbook?.DefinedNames?.Elements<DefinedName>().ToList() ?? [];
        if (definedNames.Count == 0)
        {
            return results;
        }

        if (sheetIndex < 0)
        {
            blocks.Add("シートの位置を特定できないため、印刷範囲・印刷タイトルを安全に扱えません。");
            return results;
        }

        var localNames = definedNames
            .Where(name => name.LocalSheetId?.Value == (uint)sheetIndex)
            .ToList();

        foreach (var group in localNames.GroupBy(name => name.Name?.Value ?? string.Empty, StringComparer.Ordinal))
        {
            var kind = group.Key switch
            {
                PrintDefinedNameParser.PrintAreaName => (PrintDefinedNameKind?)PrintDefinedNameKind.PrintArea,
                PrintDefinedNameParser.PrintTitlesName => PrintDefinedNameKind.PrintTitles,
                _ => null,
            };

            if (kind is null)
            {
                blocks.Add($"このシートに固有の名前定義「{group.Key}」は現在のバージョンでは引き継げません。");
                continue;
            }

            if (group.Count() > 1)
            {
                blocks.Add(
                    $"{PrintDefinedNameParser.DisplayNameOf(kind.Value)}の定義がこのシートに複数あります。");
                continue;
            }

            var definedName = group.Single();
            if (PrintDefinedNameParser.TryParse(definedName.Text, sheetName, kind.Value, out var ranges, out var error))
            {
                results.Add(new PrintDefinedNameInfo(kind.Value, ranges));
            }
            else if (error is not null)
            {
                blocks.Add(error);
            }
        }

        return results;
    }

    private static bool HasDuplicateOrOverlappingMerges(IReadOnlyList<string> references)
    {
        var ranges = new List<CellRangeParser.CellRange>(references.Count);
        foreach (var reference in references)
        {
            if (!CellRangeParser.TryParseRange(reference, out var range))
            {
                continue;
            }

            foreach (var existing in ranges)
            {
                if (range.FirstColumn <= existing.LastColumn
                    && range.LastColumn >= existing.FirstColumn
                    && range.FirstRow <= existing.LastRow
                    && range.LastRow >= existing.FirstRow)
                {
                    return true;
                }
            }

            ranges.Add(range);
        }

        return false;
    }

    private static void AddOnce(List<string> blocks, string reason)
    {
        if (!blocks.Contains(reason))
        {
            blocks.Add(reason);
        }
    }
}
