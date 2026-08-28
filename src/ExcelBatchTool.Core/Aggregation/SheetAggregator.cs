using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>
/// プレビュー済みの計画にしたがって、選択された Worksheet を新規 Workbook へ集約する。
/// 入力ファイルは読み取り専用でしか開かず、既存ファイルを上書きしない。
/// 出力は「一時ファイルへ生成 → 再度開いて検証 → 最終パスへ移動」で確定する。
/// </summary>
public sealed class SheetAggregator
{
    /// <summary>検証エラーを表示する最大件数。</summary>
    private const int MaxReportedValidationErrors = 5;

    public SheetAggregationResult Execute(
        SheetAggregationPreview preview,
        string outputPath,
        IProgress<SheetAggregationProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (!preview.CanExecute)
        {
            return SheetAggregationResult.Failed("解決していない問題があるため集約を実行できません。");
        }

        string fullOutputPath;
        try
        {
            fullOutputPath = Path.GetFullPath(outputPath);
        }
        catch (Exception ex)
        {
            return SheetAggregationResult.Failed($"出力先のパスを解釈できません: {ex.Message}");
        }

        foreach (var sheet in preview.Sheets)
        {
            if (string.Equals(Path.GetFullPath(sheet.FilePath), fullOutputPath, StringComparison.OrdinalIgnoreCase))
            {
                return SheetAggregationResult.Failed(
                    $"出力先が入力ファイル「{sheet.FileName}」と同じです。入力ファイルは変更しません。別の保存先を指定してください。");
            }
        }

        if (File.Exists(fullOutputPath))
        {
            return SheetAggregationResult.Failed(
                "同じ名前のファイルが既にあります。既存ファイルは上書きしません。別の名前を指定してください。");
        }

        var directory = Path.GetDirectoryName(fullOutputPath);
        if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
        {
            return SheetAggregationResult.Failed("保存先のフォルダーが見つかりません。");
        }

        var tempPath = Path.Combine(directory, $"~ebt-sheets-{Guid.NewGuid():N}.xlsx");

        try
        {
            WriteWorkbook(tempPath, preview, progress, cancellationToken);

            if (Validate(tempPath, preview) is { } validationError)
            {
                DeleteQuietly(tempPath);
                return SheetAggregationResult.Failed($"出力ファイルの検証に失敗しました: {validationError}");
            }

            File.Move(tempPath, fullOutputPath);
        }
        catch (OperationCanceledException)
        {
            DeleteQuietly(tempPath);
            return SheetAggregationResult.Failed("集約を中止しました。出力ファイルは作成していません。");
        }
        catch (Exception ex)
        {
            DeleteQuietly(tempPath);
            return SheetAggregationResult.Failed($"集約に失敗しました: {ex.Message}(出力ファイルは作成していません)");
        }

        return new SheetAggregationResult
        {
            Success = true,
            OutputPath = fullOutputPath,
            WorkbookCount = preview.WorkbookCount,
            SheetCount = preview.SheetCount,
            Message = $"{preview.WorkbookCount:N0} ファイルから {preview.SheetCount:N0} シートを"
                + "1 つのブックにまとめました。入力ファイルは変更していません。",
        };
    }

    private static void WriteWorkbook(
        string tempPath,
        SheetAggregationPreview preview,
        IProgress<SheetAggregationProgress>? progress,
        CancellationToken cancellationToken)
    {
        using var document = SpreadsheetDocument.Create(tempPath, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();

        var styles = new OutputStylesheetBuilder();
        var sheets = new Sheets();
        var definedNames = new DefinedNames();
        var themeCopied = false;
        var completed = 0;
        var sheetIndex = 0;
        uint sheetId = 1;

        // 同じ Source を続けて処理するので、Workbook は必要なときだけ開き直す。
        SourceWorkbook? current = null;

        try
        {
            foreach (var plan in preview.Sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (current is null || !current.Matches(plan.FilePath))
                {
                    current?.Dispose();
                    current = SourceWorkbook.Open(plan.FilePath);
                }

                if (!themeCopied && current.WorkbookPart.ThemePart is { } sourceTheme)
                {
                    CopyThemePart(workbookPart, sourceTheme);
                    themeCopied = true;
                }

                var styleMap = styles.AddSource(current.Key, current.WorkbookPart);
                var scan = WorksheetCopyScanner.ScanSheet(plan.FilePath, plan.SheetName, cancellationToken);
                if (scan.IsBlocked)
                {
                    throw new InvalidOperationException(
                        $"{plan.SourceDisplay}: {string.Join(" / ", scan.BlockReasons)}");
                }

                var sourceWorksheetPart = current.GetWorksheetPart(plan.SheetName)
                    ?? throw new InvalidOperationException($"{plan.SourceDisplay}: ワークシートが見つかりません。");

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                // 外部リンクは出力側で relationship を張り直す(Source の r:id は使わない)。
                var hyperlinks = BuildHyperlinks(worksheetPart, plan.Hyperlinks);

                WriteWorksheet(
                    worksheetPart, sourceWorksheetPart, current.Context, scan, styleMap, hyperlinks, cancellationToken);

                var sheet = new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = sheetId++,
                    Name = plan.OutputSheetName,
                };

                if (ToSheetState(plan.Visibility) is { } state)
                {
                    sheet.State = state;
                }

                sheets.Append(sheet);

                // 印刷範囲・印刷タイトルは Workbook 側の Defined Name。
                // シート名は出力名で組み立て直し、localSheetId は出力での並び順に合わせる。
                foreach (var printName in scan.PrintDefinedNames)
                {
                    definedNames.Append(new DefinedName(
                        PrintDefinedNameParser.Format(plan.OutputSheetName, printName.Ranges))
                    {
                        Name = PrintDefinedNameParser.NameOf(printName.Kind),
                        LocalSheetId = (uint)sheetIndex,
                    });
                }

                sheetIndex++;
                completed++;
                progress?.Report(new SheetAggregationProgress(completed, preview.Sheets.Count));
            }
        }
        finally
        {
            current?.Dispose();
        }

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = styles.Build();
        stylesPart.Stylesheet.Save();

        // CT_Workbook では definedNames は sheets の後ろに置く。
        workbookPart.Workbook = new Workbook(sheets);
        if (definedNames.Any())
        {
            workbookPart.Workbook.Append(definedNames);
        }

        workbookPart.Workbook.Save();
    }

    /// <summary>表示状態を sheet/@state へ変換する(Visible は属性を書かない)。</summary>
    private static SheetStateValues? ToSheetState(SheetVisibility visibility) => visibility switch
    {
        SheetVisibility.Hidden => SheetStateValues.Hidden,
        SheetVisibility.VeryHidden => SheetStateValues.VeryHidden,
        _ => null,
    };

    private static void CopyThemePart(WorkbookPart workbookPart, ThemePart sourceTheme)
    {
        var themePart = workbookPart.AddNewPart<ThemePart>();
        using var source = sourceTheme.GetStream(FileMode.Open, FileAccess.Read);
        themePart.FeedData(source);
    }

    /// <summary>
    /// 出力 WorksheetPart 用のハイパーリンク要素を作る。外部リンクは
    /// 出力側に新しい HyperlinkRelationship を作り、その r:id を設定する。
    /// </summary>
    private static Hyperlinks? BuildHyperlinks(
        WorksheetPart worksheetPart,
        IReadOnlyList<ResolvedHyperlink> resolved)
    {
        if (resolved.Count == 0)
        {
            return null;
        }

        var hyperlinks = new Hyperlinks();
        foreach (var link in resolved)
        {
            var element = new Hyperlink { Reference = link.Reference };

            if (link.ExternalTarget is { } target)
            {
                var relationship = worksheetPart.AddHyperlinkRelationship(
                    new Uri(target, UriKind.Absolute), isExternal: true);
                element.Id = relationship.Id;
            }

            if (!string.IsNullOrEmpty(link.Location))
            {
                element.Location = link.Location;
            }

            if (!string.IsNullOrEmpty(link.Tooltip))
            {
                element.Tooltip = link.Tooltip;
            }

            if (!string.IsNullOrEmpty(link.Display))
            {
                element.Display = link.Display;
            }

            hyperlinks.Append(element);
        }

        return hyperlinks;
    }

    private static void WriteWorksheet(
        WorksheetPart worksheetPart,
        WorksheetPart sourceWorksheetPart,
        WorkbookReadContext context,
        SheetCopyScan scan,
        uint[] styleMap,
        Hyperlinks? hyperlinks,
        CancellationToken cancellationToken)
    {
        // CT_Worksheet の子要素順にしたがって書く(順序を崩すと Open XML の検証に落ちる)。
        using var writer = OpenXmlWriter.Create(worksheetPart);
        writer.WriteStartElement(new Worksheet());

        if (scan.PageSetupProperties is { } pageSetupProperties)
        {
            writer.WriteElement(new SheetProperties
            {
                PageSetupProperties = (PageSetupProperties)pageSetupProperties.CloneNode(true),
            });
        }

        if (scan.DimensionReference is { } dimension)
        {
            writer.WriteElement(new SheetDimension { Reference = dimension });
        }

        writer.WriteElement(BuildSheetViews(scan));

        if (scan.SheetFormat is { } sheetFormat)
        {
            writer.WriteElement((SheetFormatProperties)sheetFormat.CloneNode(true));
        }

        if (scan.Columns.Count > 0)
        {
            writer.WriteStartElement(new Columns());
            foreach (var column in scan.Columns)
            {
                var copy = (Column)column.CloneNode(true);
                copy.Style = column.Style is { } style
                    ? OutputStylesheetBuilder.MapStyleIndex(styleMap, style.Value)
                    : null;
                writer.WriteElement(copy);
            }

            writer.WriteEndElement();
        }

        writer.WriteStartElement(new SheetData());
        WriteRows(writer, sourceWorksheetPart, context, styleMap, cancellationToken);
        writer.WriteEndElement();

        if (scan.Protection is { } protection)
        {
            writer.WriteElement((SheetProtection)protection.CloneNode(true));
        }

        if (scan.MergeReferences.Count > 0)
        {
            writer.WriteElement(new MergeCells(
                scan.MergeReferences.Select(reference => new MergeCell { Reference = reference }))
            {
                Count = (uint)scan.MergeReferences.Count,
            });
        }

        // hyperlinks は mergeCells より後、printOptions より前。
        if (hyperlinks is not null)
        {
            writer.WriteElement(hyperlinks);
        }

        // printOptions → pageMargins → pageSetup → headerFooter → rowBreaks → colBreaks の順。
        if (scan.PrintOptions is { } printOptions)
        {
            writer.WriteElement((PrintOptions)printOptions.CloneNode(true));
        }

        if (scan.PageMargins is { } pageMargins)
        {
            writer.WriteElement((PageMargins)pageMargins.CloneNode(true));
        }

        if (scan.PageSetup is { } pageSetup)
        {
            writer.WriteElement((PageSetup)pageSetup.CloneNode(true));
        }

        if (scan.HeaderFooter is { } headerFooter)
        {
            writer.WriteElement((HeaderFooter)headerFooter.CloneNode(true));
        }

        if (CloneBreaks<RowBreaks>(scan.RowBreaks) is { } rowBreaks)
        {
            writer.WriteElement(rowBreaks);
        }

        if (CloneBreaks<ColumnBreaks>(scan.ColumnBreaks) is { } columnBreaks)
        {
            writer.WriteElement(columnBreaks);
        }

        writer.WriteEndElement();
    }

    /// <summary>改ページ定義を写し、count / manualBreakCount を実際の内容に合わせ直す。</summary>
    private static T? CloneBreaks<T>(OpenXmlCompositeElement? source)
        where T : OpenXmlCompositeElement
    {
        if (source is null)
        {
            return null;
        }

        var clone = (T)source.CloneNode(true);
        var items = clone.Elements<Break>().ToList();
        if (items.Count == 0)
        {
            return null;
        }

        var manual = (uint)items.Count(item => item.ManualPageBreak?.Value == true);
        switch (clone)
        {
            case RowBreaks rows:
                rows.Count = (uint)items.Count;
                rows.ManualBreakCount = manual;
                break;

            case ColumnBreaks columns:
                columns.Count = (uint)items.Count;
                columns.ManualBreakCount = manual;
                break;
        }

        return clone;
    }

    private static SheetViews BuildSheetViews(SheetCopyScan scan)
    {
        var view = new SheetView { WorkbookViewId = 0U };

        if (!scan.ShowGridLines)
        {
            view.ShowGridLines = false;
        }

        if (!scan.ShowRowColHeaders)
        {
            view.ShowRowColHeaders = false;
        }

        if (scan.RightToLeft)
        {
            view.RightToLeft = true;
        }

        if (scan.ZoomScale is { } zoom)
        {
            view.ZoomScale = zoom;
        }

        if (scan.FreezePane is { } pane)
        {
            view.Append((Pane)pane.CloneNode(true));
        }

        foreach (var selection in scan.Selections)
        {
            view.Append((Selection)selection.CloneNode(true));
        }

        return new SheetViews(view);
    }

    /// <summary>行を 1 行ずつ読み書きする(シート全体を一度に展開しない)。</summary>
    private static void WriteRows(
        OpenXmlWriter writer,
        WorksheetPart sourceWorksheetPart,
        WorkbookReadContext context,
        uint[] styleMap,
        CancellationToken cancellationToken)
    {
        using var reader = OpenXmlReader.Create(sourceWorksheetPart);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!reader.IsStartElement || reader.ElementType != typeof(Row))
            {
                continue;
            }

            var sourceRow = (Row)reader.LoadCurrentElement()!;
            var row = new Row { RowIndex = sourceRow.RowIndex };

            if (sourceRow.Height is { } height)
            {
                row.Height = height.Value;
            }

            if (sourceRow.CustomHeight?.Value == true)
            {
                row.CustomHeight = true;
            }

            if (sourceRow.Hidden?.Value == true)
            {
                row.Hidden = true;
            }

            if (sourceRow.OutlineLevel is { } outlineLevel)
            {
                row.OutlineLevel = outlineLevel.Value;
            }

            if (sourceRow.Collapsed?.Value == true)
            {
                row.Collapsed = true;
            }

            if (sourceRow.CustomFormat?.Value == true)
            {
                row.CustomFormat = true;
                row.StyleIndex = OutputStylesheetBuilder.MapStyleIndex(styleMap, sourceRow.StyleIndex?.Value);
            }

            foreach (var sourceCell in sourceRow.Elements<Cell>())
            {
                if (ConvertCell(sourceCell, context, styleMap) is { } cell)
                {
                    row.Append(cell);
                }
            }

            writer.WriteElement(row);
        }
    }

    /// <summary>
    /// セルを出力用に作り直す。書式は remap し、共有文字列は InlineString へ、
    /// 日付は 1900 date system の serial へ正規化する(元の表示形式はそのまま使う)。
    /// </summary>
    private static Cell? ConvertCell(Cell sourceCell, WorkbookReadContext context, uint[] styleMap)
    {
        var styleIndex = OutputStylesheetBuilder.MapStyleIndex(styleMap, sourceCell.StyleIndex?.Value);
        var cell = new Cell { CellReference = sourceCell.CellReference?.Value };

        if (styleIndex != 0)
        {
            cell.StyleIndex = styleIndex;
        }

        // エラー値は型を保ったまま持ち越す。
        if (sourceCell.DataType?.Value == CellValues.Error)
        {
            cell.DataType = CellValues.Error;
            cell.CellValue = new CellValue(sourceCell.CellValue?.InnerText ?? "#N/A");
            return cell;
        }

        var value = context.ReadCell(sourceCell, out _);
        switch (value.Kind)
        {
            case MergeValueKind.Blank:
                // 値が無く書式も既定なら、セル自体を書かない。
                return styleIndex == 0 ? null : cell;

            case MergeValueKind.Text:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(
                    new Text(value.Text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
                return cell;

            case MergeValueKind.Boolean:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(value.Boolean ? "1" : "0");
                return cell;

            default:
                cell.CellValue = new CellValue(value.Number.ToString(CultureInfo.InvariantCulture));
                return cell;
        }
    }

    /// <summary>生成した .xlsx を開き直し、構造と Open XML の妥当性を検証する。</summary>
    private static string? Validate(string path, SheetAggregationPreview preview)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var document = SpreadsheetDocument.Open(stream, isEditable: false);

            var workbookPart = document.WorkbookPart;
            if (workbookPart is null)
            {
                return "Workbook 情報がありません。";
            }

            var sheets = workbookPart.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
            if (sheets.Count != preview.Sheets.Count)
            {
                return $"シート数が想定と異なります(想定 {preview.Sheets.Count} / 実際 {sheets.Count})。";
            }

            var cellFormatCount = workbookPart.WorkbookStylesPart?.Stylesheet?.CellFormats?.Count() ?? 0;

            for (var index = 0; index < sheets.Count; index++)
            {
                var expected = preview.Sheets[index];
                var actual = sheets[index];

                if (!string.Equals(actual.Name?.Value, expected.OutputSheetName, StringComparison.Ordinal))
                {
                    return $"{index + 1} 番目のシート名が想定と異なります"
                        + $"(想定「{expected.OutputSheetName}」/ 実際「{actual.Name?.Value}」)。";
                }

                var actualVisibility = WorkbookAnalyzer.ResolveVisibility(actual.State?.Value);
                if (actualVisibility != expected.Visibility)
                {
                    return $"シート「{expected.OutputSheetName}」の表示状態が想定と異なります"
                        + $"(想定 {expected.Visibility} / 実際 {actualVisibility})。";
                }

                if (actual.Id?.Value is not { } relationshipId
                    || workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
                {
                    return $"シート「{expected.OutputSheetName}」のワークシートがありません。";
                }

                if (CheckWorksheet(worksheetPart, expected, cellFormatCount) is { } sheetError)
                {
                    return sheetError;
                }

                if (CheckPrintLayout(worksheetPart, expected) is { } printError)
                {
                    return printError;
                }

                if (CheckPrintDefinedNames(workbookPart, expected, index) is { } definedNameError)
                {
                    return definedNameError;
                }

                if (CheckHyperlinks(worksheetPart, expected) is { } hyperlinkError)
                {
                    return hyperlinkError;
                }
            }

            var validationErrors = new OpenXmlValidator().Validate(document).ToList();
            if (validationErrors.Count > 0)
            {
                var details = validationErrors
                    .Take(MaxReportedValidationErrors)
                    .Select(error => $"{error.Path?.XPath}: {error.Description}");
                return $"Open XML の検証エラーが {validationErrors.Count} 件あります。{string.Join(" / ", details)}";
            }

            return null;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }

    /// <summary>ハイパーリンクが想定どおり出力され、外部リンクの参照先が一致するか確かめる。</summary>
    private static string? CheckHyperlinks(WorksheetPart worksheetPart, SheetAggregationPlan expected)
    {
        var written = worksheetPart.Worksheet?.GetFirstChild<Hyperlinks>()?.Elements<Hyperlink>().ToList() ?? [];

        if (written.Count != expected.Hyperlinks.Count)
        {
            return $"シート「{expected.OutputSheetName}」のリンク数が想定と異なります"
                + $"(想定 {expected.Hyperlinks.Count} / 実際 {written.Count})。";
        }

        for (var index = 0; index < written.Count; index++)
        {
            var actual = written[index];
            var wanted = expected.Hyperlinks[index];

            if (!string.Equals(actual.Reference?.Value, wanted.Reference, StringComparison.Ordinal))
            {
                return $"シート「{expected.OutputSheetName}」のリンク位置が想定と異なります"
                    + $"(想定 {wanted.Reference} / 実際 {actual.Reference?.Value})。";
            }

            if (!string.Equals(actual.Location?.Value, NullIfEmpty(wanted.Location), StringComparison.Ordinal))
            {
                return $"シート「{expected.OutputSheetName}」のセル {wanted.Reference} のリンク先が想定と異なります"
                    + $"(想定「{wanted.Location}」/ 実際「{actual.Location?.Value}」)。";
            }

            if (!string.Equals(actual.Tooltip?.Value, NullIfEmpty(wanted.Tooltip), StringComparison.Ordinal)
                || !string.Equals(actual.Display?.Value, NullIfEmpty(wanted.Display), StringComparison.Ordinal))
            {
                return $"シート「{expected.OutputSheetName}」のセル {wanted.Reference} のリンク表示設定が"
                    + "想定と異なります。";
            }

            if (!wanted.IsExternal)
            {
                if (actual.Id?.Value is not null)
                {
                    return $"シート「{expected.OutputSheetName}」のセル {wanted.Reference} に"
                        + "不要な外部リンク参照が出力されています。";
                }

                continue;
            }

            if (actual.Id?.Value is not { } relationshipId)
            {
                return $"シート「{expected.OutputSheetName}」のセル {wanted.Reference} の"
                    + "外部リンク参照が出力されていません。";
            }

            var relationship = worksheetPart.HyperlinkRelationships
                .FirstOrDefault(item => string.Equals(item.Id, relationshipId, StringComparison.Ordinal));

            if (relationship is null || !relationship.IsExternal)
            {
                return $"シート「{expected.OutputSheetName}」のセル {wanted.Reference} の"
                    + "外部リンク参照が解決できません。";
            }

            if (!string.Equals(relationship.Uri?.OriginalString, wanted.ExternalTarget, StringComparison.Ordinal))
            {
                return $"シート「{expected.OutputSheetName}」のセル {wanted.Reference} のリンク先が想定と異なります"
                    + $"(想定「{wanted.ExternalTarget}」/ 実際「{relationship.Uri?.OriginalString}」)。";
            }
        }

        return null;
    }

    private static string? NullIfEmpty(string? value) => string.IsNullOrEmpty(value) ? null : value;

    /// <summary>印刷・ページレイアウト情報が想定どおり出力されているか確かめる。</summary>
    private static string? CheckPrintLayout(WorksheetPart worksheetPart, SheetAggregationPlan expected)
    {
        var layout = expected.PrintLayout;
        if (layout.IsEmpty)
        {
            return null;
        }

        var worksheet = worksheetPart.Worksheet;
        if (worksheet is null)
        {
            return $"シート「{expected.OutputSheetName}」の内容を読み取れません。";
        }

        string? Mismatch(string what) =>
            $"シート「{expected.OutputSheetName}」の{what}が出力されていません。";

        if (layout.HasPageSetupProperties
            && worksheet.GetFirstChild<SheetProperties>()?.PageSetupProperties is null)
        {
            return Mismatch("印刷の拡大縮小設定");
        }

        if (layout.HasPrintOptions && worksheet.GetFirstChild<PrintOptions>() is null)
        {
            return Mismatch("印刷オプション");
        }

        if (layout.HasPageMargins && worksheet.GetFirstChild<PageMargins>() is null)
        {
            return Mismatch("余白設定");
        }

        if (layout.HasPageSetup && worksheet.GetFirstChild<PageSetup>() is null)
        {
            return Mismatch("ページ設定");
        }

        if (layout.HasHeaderFooter && worksheet.GetFirstChild<HeaderFooter>() is null)
        {
            return Mismatch("ヘッダー・フッター");
        }

        var rowBreaks = worksheet.GetFirstChild<RowBreaks>()?.Elements<Break>().Count() ?? 0;
        if (rowBreaks != layout.RowBreakCount)
        {
            return $"シート「{expected.OutputSheetName}」の行の改ページ数が想定と異なります"
                + $"(想定 {layout.RowBreakCount} / 実際 {rowBreaks})。";
        }

        var columnBreaks = worksheet.GetFirstChild<ColumnBreaks>()?.Elements<Break>().Count() ?? 0;
        if (columnBreaks != layout.ColumnBreakCount)
        {
            return $"シート「{expected.OutputSheetName}」の列の改ページ数が想定と異なります"
                + $"(想定 {layout.ColumnBreakCount} / 実際 {columnBreaks})。";
        }

        return null;
    }

    /// <summary>印刷範囲・印刷タイトルが、出力シート名と出力の並び順で作られているか確かめる。</summary>
    private static string? CheckPrintDefinedNames(
        WorkbookPart workbookPart,
        SheetAggregationPlan expected,
        int outputSheetIndex)
    {
        var definedNames = workbookPart.Workbook?.DefinedNames?.Elements<DefinedName>().ToList() ?? [];

        foreach (var (kind, ranges) in new[]
        {
            (PrintDefinedNameKind.PrintArea, expected.PrintLayout.PrintAreaRanges),
            (PrintDefinedNameKind.PrintTitles, expected.PrintLayout.PrintTitleRanges),
        })
        {
            var name = PrintDefinedNameParser.NameOf(kind);
            var matches = definedNames
                .Where(definedName => definedName.Name?.Value == name
                    && definedName.LocalSheetId?.Value == (uint)outputSheetIndex)
                .ToList();

            if (ranges.Count == 0)
            {
                if (matches.Count > 0)
                {
                    return $"シート「{expected.OutputSheetName}」に想定していない"
                        + $"{PrintDefinedNameParser.DisplayNameOf(kind)}が出力されています。";
                }

                continue;
            }

            if (matches.Count != 1)
            {
                return $"シート「{expected.OutputSheetName}」の"
                    + $"{PrintDefinedNameParser.DisplayNameOf(kind)}が正しく出力されていません。";
            }

            var wanted = PrintDefinedNameParser.Format(expected.OutputSheetName, ranges);
            if (!string.Equals(matches[0].Text, wanted, StringComparison.Ordinal))
            {
                return $"シート「{expected.OutputSheetName}」の"
                    + $"{PrintDefinedNameParser.DisplayNameOf(kind)}が想定と異なります"
                    + $"(想定「{wanted}」/ 実際「{matches[0].Text}」)。";
            }
        }

        return null;
    }

    private static string? CheckWorksheet(WorksheetPart worksheetPart, SheetAggregationPlan expected, int cellFormatCount)
    {
        var mergeCount = 0;
        using var reader = OpenXmlReader.Create(worksheetPart);

        while (reader.Read())
        {
            if (!reader.IsStartElement)
            {
                continue;
            }

            if (reader.ElementType == typeof(Cell))
            {
                var cell = (Cell)reader.LoadCurrentElement()!;
                if (cell.StyleIndex?.Value is { } styleIndex && styleIndex >= (uint)cellFormatCount)
                {
                    return $"シート「{expected.OutputSheetName}」に無効な書式参照があります(StyleIndex {styleIndex})。";
                }
            }
            else if (reader.ElementType == typeof(MergeCell))
            {
                mergeCount++;
                reader.LoadCurrentElement();
            }
        }

        return null;
    }

    private static void DeleteQuietly(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    /// <summary>読み取り専用で開いた Source Workbook。</summary>
    private sealed class SourceWorkbook : IDisposable
    {
        private readonly FileStream _stream;
        private readonly SpreadsheetDocument _document;

        private SourceWorkbook(string key, FileStream stream, SpreadsheetDocument document, WorkbookPart workbookPart)
        {
            Key = key;
            _stream = stream;
            _document = document;
            WorkbookPart = workbookPart;
            Context = WorkbookReadContext.Create(workbookPart);
        }

        public string Key { get; }

        public WorkbookPart WorkbookPart { get; }

        public WorkbookReadContext Context { get; }

        public static SourceWorkbook Open(string filePath)
        {
            var key = Path.GetFullPath(filePath);
            var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            try
            {
                var document = SpreadsheetDocument.Open(stream, isEditable: false);
                var workbookPart = document.WorkbookPart
                    ?? throw new InvalidOperationException($"{Path.GetFileName(filePath)}: Workbook 情報が見つかりません。");
                return new SourceWorkbook(key, stream, document, workbookPart);
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }

        public bool Matches(string filePath)
            => string.Equals(Key, Path.GetFullPath(filePath), StringComparison.OrdinalIgnoreCase);

        public WorksheetPart? GetWorksheetPart(string sheetName)
        {
            var sheet = WorkbookPart.Workbook?.Sheets?.Elements<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal));

            if (sheet?.Id?.Value is not { } relationshipId)
            {
                return null;
            }

            try
            {
                return WorkbookPart.GetPartById(relationshipId) as WorksheetPart;
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        public void Dispose()
        {
            _document.Dispose();
            _stream.Dispose();
        }
    }
}
