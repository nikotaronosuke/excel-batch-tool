using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBatchTool.Core;

/// <summary>
/// .xlsx を読み取り専用で解析する。対象ファイルは一切変更しない。
/// ファイルは FileAccess.Read で開くため、OS レベルで書き込み不能。
/// このクラスは例外を投げず、失敗は <see cref="WorkbookAnalysisResult"/> に記録する
/// (キャンセルによる OperationCanceledException のみ伝播する)。
/// </summary>
public static class WorkbookAnalyzer
{
    /// <summary>1 つの .xlsx ファイルを解析する。</summary>
    public static WorkbookAnalysisResult Analyze(string path, CancellationToken cancellationToken = default)
    {
        var fileName = SafeGetFileName(path);
        long? fileSize = null;

        try
        {
            var info = new FileInfo(path);
            if (!info.Exists)
            {
                return Failed(path, fileName, null, "ファイルが見つかりません。");
            }

            fileSize = info.Length;

            if (!string.Equals(info.Extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return new WorkbookAnalysisResult
                {
                    FilePath = path,
                    FileName = fileName,
                    FileSizeBytes = fileSize,
                    Status = AnalysisStatus.Failed,
                    Level = SafetyLevel.UnsupportedForNow,
                    ErrorMessage = "現在のバージョンで扱えるのは .xlsx のみです。",
                    Findings = [NewFinding(FindingType.UnsupportedFileType, 1, [])],
                };
            }

            // 読み取り専用 + 共有読み取りで開く。FileAccess.Read のため書き込みは物理的に不可能。
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var document = SpreadsheetDocument.Open(stream, isEditable: false);

            return AnalyzeOpenedDocument(document, path, fileName, fileSize.Value, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex) when (
            ex is InvalidDataException or FileFormatException or OpenXmlPackageException)
        {
            // zip / OPC として開けない: パスワード保護(暗号化)または破損の可能性。
            return new WorkbookAnalysisResult
            {
                FilePath = path,
                FileName = fileName,
                FileSizeBytes = fileSize,
                Status = AnalysisStatus.Failed,
                Level = SafetyLevel.UnsupportedForNow,
                ErrorMessage = "ファイルを読み取れません。パスワード保護(暗号化)されているか、破損している可能性があります。",
                Findings = [NewFinding(FindingType.OpenFailed, 1, [])],
            };
        }
        catch (Exception ex)
        {
            return Failed(path, fileName, fileSize, $"解析中にエラーが発生しました: {ex.Message}");
        }
    }

    private static WorkbookAnalysisResult AnalyzeOpenedDocument(
        SpreadsheetDocument document,
        string path,
        string fileName,
        long fileSize,
        CancellationToken cancellationToken)
    {
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            return Failed(path, fileName, fileSize, "Workbook 情報が見つかりません。");
        }

        var findings = new FindingAccumulator();
        var sheets = new List<SheetInfo>();

        DetectWorkbookLevelFindings(workbookPart, findings);

        foreach (var sheet in workbookPart.Workbook?.Sheets?.Elements<Sheet>() ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();

            var name = sheet.Name?.Value ?? "(名称不明)";
            var isHidden = sheet.State is not null
                && (sheet.State.Value == SheetStateValues.Hidden || sheet.State.Value == SheetStateValues.VeryHidden);
            var relId = sheet.Id?.Value;
            var part = relId is null ? null : TryGetPart(workbookPart, relId);

            switch (part)
            {
                case WorksheetPart worksheetPart:
                    sheets.Add(AnalyzeWorksheet(worksheetPart, name, isHidden, findings));
                    break;

                case ChartsheetPart:
                    findings.Add(FindingType.Chart, name);
                    sheets.Add(new SheetInfo { Name = name, Kind = SheetKind.Chartsheet, IsHidden = isHidden });
                    break;

                case MacroSheetPart:
                case InternationalMacroSheetPart:
                    findings.Add(FindingType.MacroRelated, name);
                    sheets.Add(new SheetInfo { Name = name, Kind = SheetKind.MacroSheet, IsHidden = isHidden });
                    break;

                case DialogsheetPart:
                    sheets.Add(new SheetInfo { Name = name, Kind = SheetKind.Dialogsheet, IsHidden = isHidden });
                    break;

                default:
                    sheets.Add(new SheetInfo { Name = name, Kind = SheetKind.Unknown, IsHidden = isHidden });
                    break;
            }
        }

        var findingList = findings.ToList();
        return new WorkbookAnalysisResult
        {
            FilePath = path,
            FileName = fileName,
            FileSizeBytes = fileSize,
            Status = AnalysisStatus.Succeeded,
            Level = FindingCatalog.OverallLevel(findingList),
            Sheets = sheets,
            Findings = findingList,
        };
    }

    private static void DetectWorkbookLevelFindings(WorkbookPart workbookPart, FindingAccumulator findings)
    {
        // マクロ関連: VBA パート、またはマクロ有効ブックのコンテンツタイプ
        // (.xlsm の中身が .xlsx に改名されたケースを含む)。
        if (workbookPart.VbaProjectPart is not null
            || workbookPart.ContentType.Contains("macroEnabled", StringComparison.OrdinalIgnoreCase))
        {
            findings.Add(FindingType.MacroRelated);
        }

        var workbook = workbookPart.Workbook;
        if (workbook is null)
        {
            return;
        }

        if (workbook.WorkbookProtection is not null)
        {
            findings.Add(FindingType.WorkbookProtection);
        }

        var definedNameCount = workbook.DefinedNames?.Elements<DefinedName>().Count() ?? 0;
        if (definedNameCount > 0)
        {
            findings.Add(FindingType.DefinedName, count: definedNameCount);
        }

        var externalCount = workbookPart.ExternalWorkbookParts.Count();
        if (externalCount == 0)
        {
            externalCount = workbook.GetFirstChild<ExternalReferences>()?.Elements<ExternalReference>().Count() ?? 0;
        }

        if (workbookPart.ConnectionsPart is not null)
        {
            externalCount += 1; // 外部データ接続も外部参照として扱う。
        }

        if (externalCount > 0)
        {
            findings.Add(FindingType.ExternalLink, count: externalCount);
        }

        if (workbookPart.CustomXmlParts.Any())
        {
            findings.Add(FindingType.CustomXml, count: workbookPart.CustomXmlParts.Count());
        }

        // ピボットキャッシュ(シート側パートが無くてもキャッシュ定義で検出できる)。
        if (workbookPart.PivotTableCacheDefinitionParts.Any())
        {
            findings.Add(FindingType.PivotTable);
        }
    }

    private static SheetInfo AnalyzeWorksheet(
        WorksheetPart worksheetPart,
        string sheetName,
        bool isHidden,
        FindingAccumulator findings)
    {
        DetectWorksheetPartFindings(worksheetPart, sheetName, findings);

        // シート XML はストリーミングで 1 回だけ走査する(巨大シートでも DOM を構築しない)。
        string? dimensionRef = null;
        var formulaCount = 0;
        var mergedCellCount = 0;
        var dataValidationCount = 0;
        var conditionalFormattingCount = 0;
        var hyperlinkCount = 0;
        var sheetProtected = false;
        var maxRow = 0;
        var maxColumn = 0;

        using (var reader = OpenXmlReader.Create(worksheetPart))
        {
            while (reader.Read())
            {
                if (!reader.IsStartElement)
                {
                    continue;
                }

                var elementType = reader.ElementType;

                if (elementType == typeof(CellFormula))
                {
                    formulaCount++;
                }
                else if (elementType == typeof(Row))
                {
                    foreach (var attribute in reader.Attributes)
                    {
                        if (attribute.LocalName == "r"
                            && int.TryParse(attribute.Value, out var rowIndex)
                            && rowIndex > maxRow)
                        {
                            maxRow = rowIndex;
                        }
                        else if (attribute.LocalName == "spans")
                        {
                            var upper = ParseSpansUpperBound(attribute.Value);
                            if (upper > maxColumn)
                            {
                                maxColumn = upper;
                            }
                        }
                    }
                }
                else if (elementType == typeof(MergeCell))
                {
                    mergedCellCount++;
                }
                else if (elementType == typeof(SheetDimension))
                {
                    dimensionRef = ((SheetDimension)reader.LoadCurrentElement()!).Reference?.Value;
                }
                else if (elementType == typeof(SheetProtection))
                {
                    sheetProtected = true;
                }
                else if (elementType == typeof(DataValidation))
                {
                    dataValidationCount++;
                }
                else if (elementType == typeof(ConditionalFormatting))
                {
                    conditionalFormattingCount++;
                }
                else if (elementType == typeof(Hyperlink))
                {
                    hyperlinkCount++;
                }
            }
        }

        if (formulaCount > 0)
        {
            findings.Add(FindingType.Formula, sheetName, formulaCount);
        }

        if (mergedCellCount > 0)
        {
            findings.Add(FindingType.MergedCell, sheetName, mergedCellCount);
        }

        if (sheetProtected)
        {
            findings.Add(FindingType.SheetProtection, sheetName);
        }

        if (dataValidationCount > 0)
        {
            findings.Add(FindingType.DataValidation, sheetName, dataValidationCount);
        }

        if (conditionalFormattingCount > 0)
        {
            findings.Add(FindingType.ConditionalFormatting, sheetName, conditionalFormattingCount);
        }

        if (hyperlinkCount > 0)
        {
            findings.Add(FindingType.Hyperlink, sheetName, hyperlinkCount);
        }

        return BuildSheetInfo(sheetName, isHidden, dimensionRef, maxRow, maxColumn);
    }

    private static void DetectWorksheetPartFindings(
        WorksheetPart worksheetPart,
        string sheetName,
        FindingAccumulator findings)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart is not null)
        {
            if (drawingsPart.ImageParts.Any())
            {
                findings.Add(FindingType.Image, sheetName, drawingsPart.ImageParts.Count());
            }

            if (drawingsPart.ChartParts.Any())
            {
                findings.Add(FindingType.Chart, sheetName, drawingsPart.ChartParts.Count());
            }

            if (ContainsShape(drawingsPart))
            {
                findings.Add(FindingType.Drawing, sheetName);
            }
        }

        if (worksheetPart.VmlDrawingParts.Any())
        {
            findings.Add(FindingType.Drawing, sheetName);
        }

        if (worksheetPart.ImageParts.Any())
        {
            // シート背景画像など。
            findings.Add(FindingType.Image, sheetName, worksheetPart.ImageParts.Count());
        }

        if (worksheetPart.PivotTableParts.Any())
        {
            findings.Add(FindingType.PivotTable, sheetName, worksheetPart.PivotTableParts.Count());
        }

        var tableCount = worksheetPart.TableDefinitionParts.Count();
        if (tableCount > 0)
        {
            findings.Add(FindingType.Table, sheetName, tableCount);
        }

        if (worksheetPart.WorksheetCommentsPart is not null)
        {
            findings.Add(FindingType.Comment, sheetName);
        }

        if (worksheetPart.GetPartsOfType<WorksheetThreadedCommentsPart>().Any())
        {
            findings.Add(FindingType.ThreadedComment, sheetName);
        }

        if (worksheetPart.EmbeddedObjectParts.Any() || worksheetPart.EmbeddedPackageParts.Any())
        {
            findings.Add(FindingType.EmbeddedObject, sheetName);
        }

        if (worksheetPart.EmbeddedControlPersistenceParts.Any() || worksheetPart.ControlPropertiesParts.Any())
        {
            findings.Add(FindingType.ActiveXControl, sheetName);
        }
    }

    private static bool ContainsShape(DrawingsPart drawingsPart)
    {
        try
        {
            var drawing = drawingsPart.WorksheetDrawing;
            return drawing is not null
                && (drawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().Any()
                    || drawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape>().Any()
                    || drawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.ConnectionShape>().Any());
        }
        catch
        {
            // Drawing XML を読めない場合は、安全側に倒して図形ありとして扱う。
            return true;
        }
    }

    private static SheetInfo BuildSheetInfo(
        string sheetName,
        bool isHidden,
        string? dimensionRef,
        int scannedMaxRow,
        int scannedMaxColumn)
    {
        string? usedRange = null;
        int? rowCount = null;
        int? columnCount = null;

        if (!string.IsNullOrEmpty(dimensionRef)
            && CellRangeParser.TryParseRange(dimensionRef, out var range))
        {
            usedRange = dimensionRef;
            rowCount = range.LastRow - range.FirstRow + 1;
            columnCount = range.LastColumn - range.FirstColumn + 1;
        }
        else if (scannedMaxRow > 0)
        {
            rowCount = scannedMaxRow;
            if (scannedMaxColumn > 0)
            {
                columnCount = scannedMaxColumn;
                usedRange = $"A1:{CellRangeParser.ColumnIndexToLetters(scannedMaxColumn)}{scannedMaxRow}";
            }
        }

        return new SheetInfo
        {
            Name = sheetName,
            Kind = SheetKind.Worksheet,
            IsHidden = isHidden,
            UsedRange = usedRange,
            EstimatedRowCount = rowCount,
            EstimatedColumnCount = columnCount,
        };
    }

    private static int ParseSpansUpperBound(string? spans)
    {
        // spans は "1:6" のような形式(複数区間の場合は空白区切り)。
        if (string.IsNullOrEmpty(spans))
        {
            return 0;
        }

        var max = 0;
        foreach (var span in spans.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            var colonIndex = span.IndexOf(':');
            var upperText = colonIndex >= 0 ? span[(colonIndex + 1)..] : span;
            if (int.TryParse(upperText, out var upper) && upper > max)
            {
                max = upper;
            }
        }

        return max;
    }

    private static OpenXmlPart? TryGetPart(WorkbookPart workbookPart, string relId)
    {
        try
        {
            return workbookPart.GetPartById(relId);
        }
        catch (ArgumentOutOfRangeException)
        {
            return null;
        }
    }

    private static WorkbookAnalysisResult Failed(string path, string fileName, long? fileSize, string message)
        => new()
        {
            FilePath = path,
            FileName = fileName,
            FileSizeBytes = fileSize,
            Status = AnalysisStatus.Failed,
            Level = SafetyLevel.UnsupportedForNow,
            ErrorMessage = message,
        };

    private static WorkbookFinding NewFinding(FindingType type, int count, IReadOnlyList<string> sheetNames)
        => new(type, FindingCatalog.LevelOf(type), count, sheetNames);

    private static string SafeGetFileName(string path)
    {
        try
        {
            return Path.GetFileName(path);
        }
        catch (ArgumentException)
        {
            return path;
        }
    }

    /// <summary>検出要素を種類ごとに集計する内部ヘルパー。</summary>
    private sealed class FindingAccumulator
    {
        private readonly Dictionary<FindingType, (int Count, List<string> SheetNames)> _entries = new();

        public void Add(FindingType type, string? sheetName = null, int count = 1)
        {
            if (!_entries.TryGetValue(type, out var entry))
            {
                entry = (0, []);
            }

            entry.Count += count;
            if (sheetName is not null && !entry.SheetNames.Contains(sheetName))
            {
                entry.SheetNames.Add(sheetName);
            }

            _entries[type] = entry;
        }

        public List<WorkbookFinding> ToList()
            => _entries
                .Select(pair => new WorkbookFinding(
                    pair.Key, FindingCatalog.LevelOf(pair.Key), pair.Value.Count, pair.Value.SheetNames))
                .OrderByDescending(finding => finding.Level)
                .ThenBy(finding => finding.Type)
                .ToList();
    }
}
