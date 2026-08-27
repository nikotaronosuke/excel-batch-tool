using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBatchTool.Core.Merge;

/// <summary>1 つの Worksheet を走査した結果(データ本体は保持しない)。</summary>
internal sealed class SheetScanResult
{
    /// <summary>Header(trim 後、列 1 から連続)。Block 時は空。</summary>
    public IReadOnlyList<string> Headers { get; init; } = Array.Empty<string>();

    /// <summary>完全空行を除いたデータ行数。</summary>
    public int DataRowCount { get; init; }

    /// <summary>この Sheet 単体で統合できない理由。1 件でもあれば Block。</summary>
    public IReadOnlyList<string> BlockReasons { get; init; } = Array.Empty<string>();

    /// <summary>統合はできるが利用者に伝えるべきこと。</summary>
    public IReadOnlyList<string> WarningReasons { get; init; } = Array.Empty<string>();

    public bool IsBlocked => BlockReasons.Count > 0;
}

/// <summary>
/// Worksheet を OpenXmlReader でストリーミング走査する。
/// 対象ファイルは FileAccess.Read で開くため、書き込みは物理的に発生しない。
/// DOM 全体を展開しないので、大量 Workbook・巨大シートでもメモリを抑えられる。
/// </summary>
internal static class WorksheetTableScanner
{
    /// <summary>Header・データ行数・Block 要因を 1 パスで収集する。</summary>
    public static SheetScanResult Scan(string filePath, string sheetName, CancellationToken cancellationToken)
    {
        var blocks = new List<string>();
        var warnings = new List<string>();

        if (!File.Exists(filePath))
        {
            return Blocked("ファイルが見つかりません。");
        }

        if (!string.Equals(Path.GetExtension(filePath), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return Blocked("現在のバージョンで扱えるのは .xlsx のみです。");
        }

        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var document = SpreadsheetDocument.Open(stream, isEditable: false);

            var workbookPart = document.WorkbookPart;
            if (workbookPart is null)
            {
                return Blocked("Workbook 情報が見つかりません。");
            }

            if (workbookPart.VbaProjectPart is not null
                || workbookPart.ContentType.Contains("macroEnabled", StringComparison.OrdinalIgnoreCase))
            {
                return Blocked("マクロ (VBA) を含むため、現在のバージョンでは統合対象にできません。");
            }

            if (FindWorksheetPart(workbookPart, sheetName) is not { } worksheetPart)
            {
                return Blocked($"ワークシート「{sheetName}」が見つかりません(グラフシート等は統合対象外です)。");
            }

            var context = WorkbookReadContext.Create(workbookPart);
            return ScanWorksheet(worksheetPart, context, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex) when (ex is InvalidDataException or System.IO.FileFormatException or OpenXmlPackageException)
        {
            return Blocked("ファイルを読み取れません。パスワード保護(暗号化)されているか、破損している可能性があります。");
        }
        catch (Exception ex)
        {
            return Blocked($"読み取りエラー: {ex.Message}");
        }

        SheetScanResult Blocked(string reason)
        {
            blocks.Add(reason);
            return new SheetScanResult { BlockReasons = blocks, WarningReasons = warnings };
        }
    }

    private static SheetScanResult ScanWorksheet(
        WorksheetPart worksheetPart,
        WorkbookReadContext context,
        CancellationToken cancellationToken)
    {
        var headerCells = new Dictionary<int, string>();
        var mergeRanges = new List<CellRangeParser.CellRange>();

        var hasFormula = false;
        var hasAmbiguousDateFormat = false;
        var dataRowCount = 0;
        var lastDataRow = 0;
        var maxDataColumn = 0;
        int? headerColumnCount = null;

        var currentRow = 0;
        var currentColumn = 0;
        var currentRowHasData = false;

        using (var reader = OpenXmlReader.Create(worksheetPart))
        {
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (!reader.IsStartElement)
                {
                    continue;
                }

                if (reader.ElementType == typeof(Row))
                {
                    FinishRow();
                    currentRow = ReadRowIndex(reader.Attributes, currentRow + 1);
                    currentColumn = 0;
                    currentRowHasData = false;

                    if (currentRow >= 2)
                    {
                        headerColumnCount ??= ComputeHeaderColumnCount(headerCells);
                    }
                }
                else if (reader.ElementType == typeof(Cell))
                {
                    var cell = (Cell)reader.LoadCurrentElement()!;
                    currentColumn = ResolveColumn(cell, currentColumn);

                    if (cell.CellFormula is not null)
                    {
                        hasFormula = true;
                    }

                    var value = context.ReadCell(cell, out var ambiguous);
                    hasAmbiguousDateFormat |= ambiguous;

                    if (value.IsBlank)
                    {
                        continue;
                    }

                    if (currentRow == 1)
                    {
                        headerCells[currentColumn] = value.ToDisplayString().Trim();
                    }
                    else if (currentRow >= 2)
                    {
                        maxDataColumn = Math.Max(maxDataColumn, currentColumn);
                        if (currentColumn <= (headerColumnCount ?? 0))
                        {
                            currentRowHasData = true;
                        }
                    }
                }
                else if (reader.ElementType == typeof(MergeCell))
                {
                    var mergeCell = (MergeCell)reader.LoadCurrentElement()!;
                    if (mergeCell.Reference?.Value is { } reference
                        && CellRangeParser.TryParseRange(reference, out var range))
                    {
                        mergeRanges.Add(range);
                    }
                }
            }

            FinishRow();
        }

        headerColumnCount ??= ComputeHeaderColumnCount(headerCells);

        var blocks = new List<string>();
        var warnings = new List<string>();
        var headers = new List<string>();

        if (headerColumnCount.Value == 0)
        {
            blocks.Add("1 行目(ヘッダー行)が空です。1 行目に列名が必要です。");
        }
        else
        {
            for (var column = 1; column <= headerColumnCount.Value; column++)
            {
                if (headerCells.TryGetValue(column, out var name) && name.Length > 0)
                {
                    headers.Add(name);
                }
                else
                {
                    blocks.Add($"1 行目に空のヘッダーがあります(列 {CellRangeParser.ColumnIndexToLetters(column)})。");
                    headers.Add(string.Empty);
                }
            }

            foreach (var duplicate in headers
                .Where(name => name.Length > 0)
                .GroupBy(name => name, StringComparer.Ordinal)
                .Where(group => group.Count() > 1))
            {
                blocks.Add($"同じシート内にヘッダー「{duplicate.Key}」が {duplicate.Count()} 個あります。");
            }
        }

        if (hasFormula)
        {
            blocks.Add("数式を含むため、Phase 1A では統合対象にできません(計算結果が最新である保証がないため)。");
        }

        var tableLastRow = Math.Max(lastDataRow, 1);
        if (headerColumnCount.Value > 0
            && mergeRanges.Any(range => Intersects(range, headerColumnCount.Value, tableLastRow)))
        {
            blocks.Add("表の範囲に結合セルがあるため、Phase 1A では統合対象にできません。");
        }

        if (headerColumnCount.Value > 0 && maxDataColumn > headerColumnCount.Value)
        {
            warnings.Add(
                $"ヘッダーのない列({CellRangeParser.ColumnIndexToLetters(headerColumnCount.Value + 1)} 列以降)に" +
                "データがあります。この範囲は統合されません。");
        }

        if (hasAmbiguousDateFormat)
        {
            warnings.Add("日付か数値か判断できない表示形式のセルがあります。誤変換を避けるため数値として出力します。");
        }

        return new SheetScanResult
        {
            Headers = blocks.Count > 0 ? Array.Empty<string>() : headers,
            DataRowCount = dataRowCount,
            BlockReasons = blocks,
            WarningReasons = warnings,
        };

        void FinishRow()
        {
            if (currentRow >= 2 && currentRowHasData)
            {
                dataRowCount++;
                lastDataRow = currentRow;
            }
        }
    }

    /// <summary>
    /// データ行を 1 行ずつ読み出す。値は Header 位置(0 始まり)に対応した配列で返す。
    /// 完全空行(ヘッダー列の範囲がすべて空)は返さない。
    /// </summary>
    public static IEnumerable<MergeCellValue[]> ReadDataRows(
        string filePath,
        string sheetName,
        int headerColumnCount,
        CancellationToken cancellationToken = default)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);

        var workbookPart = document.WorkbookPart
            ?? throw new InvalidOperationException("Workbook 情報が見つかりません。");
        var worksheetPart = FindWorksheetPart(workbookPart, sheetName)
            ?? throw new InvalidOperationException($"ワークシート「{sheetName}」が見つかりません。");

        var context = WorkbookReadContext.Create(workbookPart);

        var values = new MergeCellValue[headerColumnCount];
        var hasData = false;
        var currentRow = 0;
        var currentColumn = 0;

        using var reader = OpenXmlReader.Create(worksheetPart);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!reader.IsStartElement)
            {
                continue;
            }

            if (reader.ElementType == typeof(Row))
            {
                if (hasData)
                {
                    yield return values;
                    values = new MergeCellValue[headerColumnCount];
                    hasData = false;
                }

                currentRow = ReadRowIndex(reader.Attributes, currentRow + 1);
                currentColumn = 0;
            }
            else if (reader.ElementType == typeof(Cell) && currentRow >= 2)
            {
                var cell = (Cell)reader.LoadCurrentElement()!;
                currentColumn = ResolveColumn(cell, currentColumn);
                if (currentColumn < 1 || currentColumn > headerColumnCount)
                {
                    continue;
                }

                var value = context.ReadCell(cell, out _);
                if (value.IsBlank)
                {
                    continue;
                }

                values[currentColumn - 1] = value;
                hasData = true;
            }
        }

        if (hasData)
        {
            yield return values;
        }
    }

    private static WorksheetPart? FindWorksheetPart(WorkbookPart workbookPart, string sheetName)
    {
        var sheet = workbookPart.Workbook?.Sheets?.Elements<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal));

        if (sheet?.Id?.Value is not { } relationshipId)
        {
            return null;
        }

        try
        {
            return workbookPart.GetPartById(relationshipId) as WorksheetPart;
        }
        catch (ArgumentOutOfRangeException)
        {
            return null;
        }
    }

    private static int ComputeHeaderColumnCount(Dictionary<int, string> headerCells)
    {
        var max = 0;
        foreach (var pair in headerCells)
        {
            if (pair.Value.Length > 0 && pair.Key > max)
            {
                max = pair.Key;
            }
        }

        return max;
    }

    private static bool Intersects(CellRangeParser.CellRange range, int headerColumnCount, int lastRow)
        => range.FirstColumn <= headerColumnCount
            && range.LastColumn >= 1
            && range.FirstRow <= lastRow
            && range.LastRow >= 1;

    private static int ReadRowIndex(IEnumerable<OpenXmlAttribute> attributes, int fallback)
    {
        foreach (var attribute in attributes)
        {
            if (attribute.LocalName == "r" && int.TryParse(attribute.Value, out var index))
            {
                return index;
            }
        }

        return fallback;
    }

    private static int ResolveColumn(Cell cell, int previousColumn)
        => cell.CellReference?.Value is { } reference
            && CellRangeParser.TryParseCell(reference, out var column, out _)
                ? column
                : previousColumn + 1;
}
