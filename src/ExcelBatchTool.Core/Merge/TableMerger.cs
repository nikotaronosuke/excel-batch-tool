using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBatchTool.Core.Merge;

/// <summary>
/// プレビュー済みの計画にしたがって、新規 Workbook へ表を縦結合する。
/// 入力ファイルは読み取り専用でしか開かず、既存ファイルを上書きしない。
/// 出力は「一時ファイルへ生成 → 再度開いて検証 → 最終パスへ移動」で確定する。
/// </summary>
public sealed class TableMerger
{
    private const uint StyleDefault = 0;
    private const uint StyleHeader = 1;
    private const uint StyleDate = 2;
    private const uint StyleDateTime = 3;
    private const uint StyleTime = 4;

    public MergeExecutionResult Execute(
        MergePreview preview,
        MergeOptions options,
        string outputPath,
        IProgress<MergeProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (!preview.CanExecute)
        {
            return MergeExecutionResult.Failed("解決していない問題があるため統合を実行できません。");
        }

        string fullOutputPath;
        try
        {
            fullOutputPath = Path.GetFullPath(outputPath);
        }
        catch (Exception ex)
        {
            return MergeExecutionResult.Failed($"出力先のパスを解釈できません: {ex.Message}");
        }

        foreach (var source in preview.Sources)
        {
            if (string.Equals(Path.GetFullPath(source.FilePath), fullOutputPath, StringComparison.OrdinalIgnoreCase))
            {
                return MergeExecutionResult.Failed(
                    $"出力先が入力ファイル「{source.FileName}」と同じです。入力ファイルは変更しません。別の保存先を指定してください。");
            }
        }

        if (File.Exists(fullOutputPath))
        {
            return MergeExecutionResult.Failed(
                "同じ名前のファイルが既にあります。既存ファイルは上書きしません。別の名前を指定してください。");
        }

        var directory = Path.GetDirectoryName(fullOutputPath);
        if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
        {
            return MergeExecutionResult.Failed("保存先のフォルダーが見つかりません。");
        }

        // 一時ファイルは最終パスと同じフォルダーに作る(移動を同一ボリューム内で完結させるため)。
        var tempPath = Path.Combine(directory, $"~ebt-merge-{Guid.NewGuid():N}.xlsx");
        var writtenDataRows = 0;

        try
        {
            writtenDataRows = WriteWorkbook(tempPath, preview, options, progress, cancellationToken);

            if (Validate(tempPath, preview.OutputHeaders, writtenDataRows, options.OutputSheetName) is { } validationError)
            {
                DeleteQuietly(tempPath);
                return MergeExecutionResult.Failed($"出力ファイルの検証に失敗しました: {validationError}");
            }

            // 最終パスへ移動。既存ファイルがあれば移動は失敗し、上書きされない。
            File.Move(tempPath, fullOutputPath);
        }
        catch (OperationCanceledException)
        {
            DeleteQuietly(tempPath);
            return MergeExecutionResult.Failed("統合を中止しました。出力ファイルは作成していません。");
        }
        catch (Exception ex)
        {
            DeleteQuietly(tempPath);
            return MergeExecutionResult.Failed($"統合に失敗しました: {ex.Message}(出力ファイルは作成していません)");
        }

        return new MergeExecutionResult
        {
            Success = true,
            OutputPath = fullOutputPath,
            WorkbookCount = preview.WorkbookCount,
            SheetCount = preview.SheetCount,
            DataRowCount = writtenDataRows,
            Message = $"{preview.WorkbookCount:N0} ファイル / {preview.SheetCount:N0} シート / "
                + $"{writtenDataRows:N0} 行を統合しました。入力ファイルは変更していません。",
        };
    }

    private static int WriteWorkbook(
        string tempPath,
        MergePreview preview,
        MergeOptions options,
        IProgress<MergeProgress>? progress,
        CancellationToken cancellationToken)
    {
        var writtenDataRows = 0;

        using (var document = SpreadsheetDocument.Create(tempPath, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = BuildStylesheet();
            stylesPart.Stylesheet.Save();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            using (var writer = OpenXmlWriter.Create(worksheetPart))
            {
                writer.WriteStartElement(new Worksheet());

                // ヘッダー行を固定表示にする(sheetData より前に書く必要がある)。
                writer.WriteElement(new SheetViews(new SheetView
                {
                    WorkbookViewId = 0U,
                    Pane = new Pane
                    {
                        VerticalSplit = 1D,
                        TopLeftCell = "A2",
                        ActivePane = PaneValues.BottomLeft,
                        State = PaneStateValues.Frozen,
                    },
                }));

                writer.WriteStartElement(new SheetData());

                var rowIndex = 1U;
                WriteHeaderRow(writer, preview.OutputHeaders, rowIndex);

                var completedSources = 0;
                foreach (var source in preview.Sources)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var metadata = BuildMetadataValues(options, source);
                    foreach (var values in WorksheetTableScanner.ReadDataRows(
                        source.FilePath, source.SheetName, source.Headers.Count, cancellationToken))
                    {
                        rowIndex++;
                        WriteDataRow(writer, rowIndex, metadata, values, source.ColumnMap, preview.DataHeaders.Count);
                        writtenDataRows++;
                    }

                    completedSources++;
                    progress?.Report(new MergeProgress(completedSources, preview.Sources.Count, writtenDataRows));
                }

                writer.WriteEndElement(); // SheetData

                if (preview.OutputHeaders.Count > 0)
                {
                    var lastColumn = CellRangeParser.ColumnIndexToLetters(preview.OutputHeaders.Count);
                    writer.WriteElement(new AutoFilter { Reference = $"A1:{lastColumn}{rowIndex}" });
                }

                writer.WriteEndElement(); // Worksheet
            }

            workbookPart.Workbook = new Workbook(new Sheets(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1U,
                Name = options.OutputSheetName,
            }));
            workbookPart.Workbook.Save();
        }

        return writtenDataRows;
    }

    private static void WriteHeaderRow(OpenXmlWriter writer, IReadOnlyList<string> headers, uint rowIndex)
    {
        writer.WriteStartElement(new Row { RowIndex = rowIndex });
        for (var i = 0; i < headers.Count; i++)
        {
            writer.WriteElement(TextCell(Reference(i + 1, rowIndex), headers[i], StyleHeader));
        }

        writer.WriteEndElement();
    }

    private static void WriteDataRow(
        OpenXmlWriter writer,
        uint rowIndex,
        IReadOnlyList<string> metadata,
        MergeCellValue[] values,
        IReadOnlyList<int> columnMap,
        int dataColumnCount)
    {
        var output = new MergeCellValue[dataColumnCount];
        for (var i = 0; i < values.Length && i < columnMap.Count; i++)
        {
            var target = columnMap[i];
            if (target >= 0 && target < dataColumnCount)
            {
                output[target] = values[i];
            }
        }

        writer.WriteStartElement(new Row { RowIndex = rowIndex });

        var column = 0;
        foreach (var text in metadata)
        {
            column++;
            writer.WriteElement(TextCell(Reference(column, rowIndex), text, StyleDefault));
        }

        for (var i = 0; i < output.Length; i++)
        {
            column++;
            if (output[i].IsBlank)
            {
                continue;
            }

            writer.WriteElement(BuildCell(Reference(column, rowIndex), output[i]));
        }

        writer.WriteEndElement();
    }

    private static List<string> BuildMetadataValues(MergeOptions options, MergeSourcePlan source)
    {
        var values = new List<string>(2);
        if (options.IncludeSourceFileColumn)
        {
            values.Add(source.FileName);
        }

        if (options.IncludeSourceSheetColumn)
        {
            values.Add(source.SheetName);
        }

        return values;
    }

    private static string Reference(int column, uint rowIndex)
        => $"{CellRangeParser.ColumnIndexToLetters(column)}{rowIndex}";

    private static Cell TextCell(string reference, string text, uint styleIndex) => new()
    {
        CellReference = reference,
        StyleIndex = styleIndex,
        DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text(text) { Space = SpaceProcessingModeValues.Preserve }),
    };

    private static Cell BuildCell(string reference, MergeCellValue value) => value.Kind switch
    {
        MergeValueKind.Text => TextCell(reference, value.Text ?? string.Empty, StyleDefault),

        MergeValueKind.Boolean => new Cell
        {
            CellReference = reference,
            DataType = CellValues.Boolean,
            CellValue = new CellValue(value.Boolean ? "1" : "0"),
        },

        MergeValueKind.Date => NumericCell(reference, value.Number, StyleDate),
        MergeValueKind.DateTime => NumericCell(reference, value.Number, StyleDateTime),
        MergeValueKind.Time => NumericCell(reference, value.Number, StyleTime),

        _ => NumericCell(reference, value.Number, StyleDefault),
    };

    private static Cell NumericCell(string reference, double number, uint styleIndex) => new()
    {
        CellReference = reference,
        StyleIndex = styleIndex,
        CellValue = new CellValue(number.ToString(CultureInfo.InvariantCulture)),
    };

    private static Stylesheet BuildStylesheet() => new(
        new NumberingFormats(
            new NumberingFormat { NumberFormatId = 164U, FormatCode = "yyyy/mm/dd hh:mm:ss" })
        { Count = 1U },
        new Fonts(
            new Font(new FontSize { Val = 11D }),
            new Font(new Bold(), new FontSize { Val = 11D }))
        { Count = 2U },
        new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }))
        { Count = 2U },
        new Borders(new Border()) { Count = 1U },
        new CellStyleFormats(
            new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U })
        { Count = 1U },
        new CellFormats(
            new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U },
            new CellFormat { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyFont = true },
            new CellFormat { NumberFormatId = 14U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyNumberFormat = true },
            new CellFormat { NumberFormatId = 164U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyNumberFormat = true },
            new CellFormat { NumberFormatId = 21U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyNumberFormat = true })
        { Count = 5U });

    /// <summary>生成した .xlsx を読み取り専用で開き直し、最低限の整合性を確認する。</summary>
    private static string? Validate(
        string path,
        IReadOnlyList<string> expectedHeaders,
        int expectedDataRows,
        string expectedSheetName)
    {
        try
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var document = SpreadsheetDocument.Open(stream, isEditable: false))
            {
                var workbookPart = document.WorkbookPart;
                if (workbookPart is null)
                {
                    return "Workbook 情報がありません。";
                }

                var sheets = workbookPart.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
                if (sheets.Count != 1
                    || !string.Equals(sheets[0].Name?.Value, expectedSheetName, StringComparison.Ordinal))
                {
                    return "出力シートの構成が想定と異なります。";
                }
            }

            var headerScan = WorksheetTableScanner.Scan(path, expectedSheetName, CancellationToken.None);
            if (headerScan.IsBlocked)
            {
                return string.Join(" / ", headerScan.BlockReasons);
            }

            if (!headerScan.Headers.SequenceEqual(expectedHeaders, StringComparer.Ordinal))
            {
                return "出力ヘッダーが想定と異なります。";
            }

            if (headerScan.DataRowCount != expectedDataRows)
            {
                return $"出力行数が想定と異なります(想定 {expectedDataRows:N0} 行 / 実際 {headerScan.DataRowCount:N0} 行)。";
            }

            return null;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
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
}
