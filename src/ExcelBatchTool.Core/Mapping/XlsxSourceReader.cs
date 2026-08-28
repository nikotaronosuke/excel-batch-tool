using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Mapping;

/// <summary>
/// データ元の .xlsx を読む。Worksheet 全体を DOM へ載せず、行を 1 つずつ読み進める。
/// 必要なキーが先に分かっている場合は、その行だけを保持する。
/// </summary>
internal static class XlsxSourceReader
{
    /// <summary>読み取れるシート名の一覧。</summary>
    public static IReadOnlyList<string> ReadSheetNames(string filePath)
    {
        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var document = SpreadsheetDocument.Open(stream, isEditable: false);

            return [.. document.WorkbookPart?.Workbook?.Sheets?.Elements<Sheet>()
                .Where(sheet => sheet.Id?.Value is not null)
                .Select(sheet => sheet.Name?.Value ?? string.Empty)
                .Where(name => name.Length > 0) ?? []];
        }
        catch (Exception)
        {
            return [];
        }
    }

    /// <summary>指定した行を項目名として読む。</summary>
    public static SourceHeaderResult ReadHeader(string filePath, string sheetName, int headerRow)
    {
        return Open(filePath, sheetName, (worksheetPart, context) =>
        {
            var cells = ReadRowCells(worksheetPart, (uint)headerRow, CancellationToken.None);
            if (cells.Count == 0)
            {
                return SourceHeaderResult.Failed(
                    $"データ元のシート「{sheetName}」の {headerRow} 行目に項目名がありません。");
            }

            // 項目名の範囲は「1 列目 〜 その行にある最も右の列」。途中の抜けは空の項目名として扱う。
            var lastColumn = cells.Keys.Max();
            var raw = new List<string?>(lastColumn);
            for (var column = 1; column <= lastColumn; column++)
            {
                raw.Add(cells.TryGetValue(column, out var cell)
                    ? HeaderText(cell, context)
                    : null);
            }

            return SourceHeaders.Validate(raw, out var columns, out var error)
                ? new SourceHeaderResult { Columns = columns! }
                : SourceHeaderResult.Failed(error!);
        },
        SourceHeaderResult.Failed);
    }

    /// <summary>
    /// 項目名の行より後の行を順に渡す(CSV 変換で使う)。行は保持しない。
    /// <paramref name="onRecord"/> が false を返したところで読み終える。
    /// 読み取れないときは理由を返す。
    /// </summary>
    public static string? ReadRecords(
        string filePath,
        string sheetName,
        int headerRow,
        int columnCount,
        Func<int, IReadOnlyList<SourceValue>, bool> onRecord,
        CancellationToken cancellationToken)
    {
        return Open<string?>(filePath, sheetName, (worksheetPart, context) =>
        {
            var numberFormats = context.NumberFormats;

            using var reader = OpenXmlReader.Create(worksheetPart);
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (!reader.IsStartElement || reader.ElementType != typeof(Row))
                {
                    continue;
                }

                var row = (Row)reader.LoadCurrentElement()!;
                if (row.RowIndex?.Value is not { } rowIndex || rowIndex <= (uint)headerRow)
                {
                    continue;
                }

                var cells = IndexCells(row);
                var values = new SourceValue[columnCount];
                for (var column = 1; column <= columnCount; column++)
                {
                    values[column - 1] = ReadValue(
                        cells.GetValueOrDefault(column), context, numberFormats);
                }

                if (!onRecord((int)rowIndex, values))
                {
                    break;
                }
            }

            return null;
        },
        error => error);
    }

    /// <summary>
    /// 必要なキーに一致する行だけを集める。行そのものは必要な分しか保持しないが、
    /// 重複キーの検出のためキーの一覧だけは通して見る。
    /// </summary>
    public static SourceMatchResult ReadRows(
        string filePath,
        string sheetName,
        int headerRow,
        int keyColumn,
        IReadOnlyList<int> valueColumns,
        IReadOnlySet<string> requiredKeys,
        CancellationToken cancellationToken)
    {
        return Open(filePath, sheetName, (worksheetPart, context) =>
        {
            var rowsByKey = new Dictionary<string, SourceRow>(StringComparer.Ordinal);
            var duplicates = new HashSet<string>(StringComparer.Ordinal);
            var seenKeys = new HashSet<string>(StringComparer.Ordinal);
            var blankRows = 0;
            var blankKeyWithValue = 0;
            var unused = 0;

            var numberFormats = context.NumberFormats;

            using var reader = OpenXmlReader.Create(worksheetPart);
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (!reader.IsStartElement || reader.ElementType != typeof(Row))
                {
                    continue;
                }

                var row = (Row)reader.LoadCurrentElement()!;
                if (row.RowIndex?.Value is not { } rowIndex || rowIndex <= (uint)headerRow)
                {
                    continue;
                }

                var cells = IndexCells(row);
                var keyValue = ReadValue(cells.GetValueOrDefault(keyColumn), context, numberFormats);

                if (keyValue.IsBlank)
                {
                    var hasAnyValue = valueColumns.Any(column =>
                        !ReadValue(cells.GetValueOrDefault(column), context, numberFormats).IsBlank);

                    if (hasAnyValue)
                    {
                        blankKeyWithValue++;
                    }
                    else
                    {
                        blankRows++;
                    }

                    continue;
                }

                if (keyValue.Kind != SourceValueKind.Text)
                {
                    return SourceMatchResult.Failed(
                        $"データ元の {rowIndex} 行目のキーが文字列ではありません。"
                            + "「00123」と「123」を取り違えないよう、キーの列は文字列のセルだけを対象にします。");
                }

                var key = keyValue.Text!;

                if (!seenKeys.Add(key))
                {
                    duplicates.Add(key);
                    rowsByKey.Remove(key);
                    continue;
                }

                if (!requiredKeys.Contains(key))
                {
                    unused++;
                    continue;
                }

                rowsByKey[key] = new SourceRow(
                    (int)rowIndex,
                    [.. valueColumns.Select(column =>
                        ReadValue(cells.GetValueOrDefault(column), context, numberFormats))]);
            }

            return new SourceMatchResult
            {
                RowsByKey = rowsByKey,
                DuplicateKeys = duplicates,
                BlankRowCount = blankRows,
                BlankKeyWithValueCount = blankKeyWithValue,
                UnusedRowCount = unused,
            };
        },
        SourceMatchResult.Failed);
    }

    /// <summary>
    /// キー列だけを読む(表同士の突合更新の 1 パス目)。
    /// 値は保持せず、キーの集合・重複・空欄の数だけを集める。
    /// </summary>
    public static SourceKeyScan ReadKeys(
        string filePath,
        string sheetName,
        int headerRow,
        int keyColumn,
        IReadOnlyList<int> valueColumns,
        CancellationToken cancellationToken)
    {
        return Open(filePath, sheetName, (worksheetPart, context) =>
        {
            var keys = new HashSet<string>(StringComparer.Ordinal);
            var duplicates = new HashSet<string>(StringComparer.Ordinal);
            var keyedRows = 0;
            var blankRows = 0;
            var blankKeyWithValue = 0;

            var numberFormats = context.NumberFormats;

            using var reader = OpenXmlReader.Create(worksheetPart);
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (!reader.IsStartElement || reader.ElementType != typeof(Row))
                {
                    continue;
                }

                var row = (Row)reader.LoadCurrentElement()!;
                if (row.RowIndex?.Value is not { } rowIndex || rowIndex <= (uint)headerRow)
                {
                    continue;
                }

                var cells = IndexCells(row);
                var keyValue = ReadValue(cells.GetValueOrDefault(keyColumn), context, numberFormats);

                if (keyValue.IsBlank)
                {
                    var hasAnyValue = valueColumns.Any(column =>
                        !ReadValue(cells.GetValueOrDefault(column), context, numberFormats).IsBlank);

                    if (hasAnyValue)
                    {
                        blankKeyWithValue++;
                    }
                    else
                    {
                        blankRows++;
                    }

                    continue;
                }

                if (keyValue.Kind != SourceValueKind.Text)
                {
                    return SourceKeyScan.Failed(
                        $"データ元の {rowIndex} 行目のキーが文字列ではありません。"
                            + "「00123」と「123」を取り違えないよう、キーの列は文字列のセルだけを対象にします。");
                }

                keyedRows++;
                if (!keys.Add(keyValue.Text!))
                {
                    duplicates.Add(keyValue.Text!);
                }
            }

            return new SourceKeyScan
            {
                Keys = keys,
                DuplicateKeys = duplicates,
                KeyedRowCount = keyedRows,
                BlankRowCount = blankRows,
                BlankKeyWithValueCount = blankKeyWithValue,
            };
        },
        SourceKeyScan.Failed);
    }

    /// <summary>
    /// データ元の .xlsx を開いて処理する。Excel の形式として壊れているものは扱わない
    /// (転記元としての条件であり、書き換え対象の条件とは別)。
    /// </summary>
    private static TResult Open<TResult>(
        string filePath,
        string sheetName,
        Func<WorksheetPart, SourceReadContext, TResult> read,
        Func<string, TResult> fail)
    {
        if (!File.Exists(filePath))
        {
            return fail("データ元のファイルが見つかりません。");
        }

        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var document = SpreadsheetDocument.Open(stream, isEditable: false);

            var workbookPart = document.WorkbookPart;
            if (workbookPart is null)
            {
                return fail("データ元のブック情報が見つかりません。");
            }

            if (HasStructuralProblem(workbookPart))
            {
                return fail(
                    "データ元のファイルは Excel の形式として問題がある箇所を含むため、"
                        + "現在のバージョンでは転記元にできません。");
            }

            var sheet = workbookPart.Workbook?.Sheets?.Elements<Sheet>()
                .FirstOrDefault(item => string.Equals(item.Name?.Value, sheetName, StringComparison.Ordinal));

            if (sheet?.Id?.Value is not { } relationshipId)
            {
                return fail($"データ元にシート「{sheetName}」が見つかりません。");
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
                return fail("データ元に指定できるのは通常のワークシートだけです。");
            }

            return read(worksheetPart, SourceReadContext.Create(workbookPart));
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex) when (ex is InvalidDataException or FileFormatException or OpenXmlPackageException)
        {
            return fail(
                "データ元のファイルを読み取れません。"
                    + "パスワード保護(暗号化)されているか、破損している可能性があります。");
        }
        catch (Exception ex)
        {
            return fail($"データ元のファイルを読み取れません: {ex.Message}");
        }
    }

    /// <summary>
    /// セルの意味を決めるために読むパート(シート一覧・共有文字列・表示形式)だけを検証する。
    ///
    /// 転記元は読むだけなので、書き換え対象(D-023: 壊れたファイルを書き換えない)とは
    /// 目的が違う。ここで防ぎたいのは「値を取り違えて読むこと」。
    /// ブック全体の検証は、行数に比例して非常に重い(実測: 10 万行で約 26 秒・約 1.4 GB)ため、
    /// 読み取りの解釈に関わるパートに限定する。セルそのものの想定外は、
    /// 読み取り側が「読み取れません」として Block するので黙って通ることはない。
    /// </summary>
    private static bool HasStructuralProblem(WorkbookPart workbookPart)
    {
        var validator = new OpenXmlValidator();

        if (validator.Validate(workbookPart).Take(1).Any())
        {
            return true;
        }

        if (workbookPart.SharedStringTablePart is { } sharedStrings
            && validator.Validate(sharedStrings).Take(1).Any())
        {
            return true;
        }

        return workbookPart.WorkbookStylesPart is { } styles
            && validator.Validate(styles).Take(1).Any();
    }

    /// <summary>指定行のセルを列番号で引けるようにして返す。</summary>
    private static Dictionary<int, Cell> ReadRowCells(
        WorksheetPart worksheetPart, uint rowIndex, CancellationToken cancellationToken)
    {
        using var reader = OpenXmlReader.Create(worksheetPart);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!reader.IsStartElement || reader.ElementType != typeof(Row))
            {
                continue;
            }

            var row = (Row)reader.LoadCurrentElement()!;
            if (row.RowIndex?.Value == rowIndex)
            {
                return IndexCells(row);
            }
        }

        return [];
    }

    private static Dictionary<int, Cell> IndexCells(Row row)
    {
        var cells = new Dictionary<int, Cell>();
        foreach (var cell in row.Elements<Cell>())
        {
            if (cell.CellReference?.Value is { } reference
                && CellRangeParser.TryParseCell(reference, out var column, out _))
            {
                cells[column] = cell;
            }
        }

        return cells;
    }

    /// <summary>項目名として読む文字列(数値の見出しも文字として扱う)。</summary>
    private static string? HeaderText(Cell cell, SourceReadContext context)
    {
        if (cell.CellFormula is not null)
        {
            return null; // 空扱いにして「項目名が空です」で止める。
        }

        var value = context.Context.ReadCell(cell, out _);
        return value.Kind == MergeValueKind.Blank ? null : value.ToDisplayString();
    }

    /// <summary>データ元のセル 1 つを読む。意味が推測になるものは Unsupported にする。</summary>
    private static SourceValue ReadValue(
        Cell? cell, SourceReadContext context, NumberFormatCompatibility numberFormats)
    {
        if (cell is null)
        {
            return SourceValue.Blank();
        }

        if (cell.CellFormula is not null)
        {
            return SourceValue.Unsupported("数式のため、計算結果を転記元にはできません");
        }

        if (context.Context.ReferencesRichText(cell))
        {
            return SourceValue.Unsupported("文字ごとに書式が設定されているため転記できません");
        }

        var value = context.Context.ReadCell(cell, out _);
        switch (value.Kind)
        {
            case MergeValueKind.Blank:
                return SourceValue.Blank();

            case MergeValueKind.Text:
                return SourceValue.OfText(value.Text ?? string.Empty);

            case MergeValueKind.Number:
                // 日付・通貨・% など、表示と生の値が食い違う書式は転記元にしない。
                return numberFormats.IsPlainNumber(cell.StyleIndex?.Value)
                    ? SourceValue.OfNumber(value.Number)
                    : SourceValue.Unsupported(
                        "日付・通貨・パーセントなどの表示形式のため、そのままの数値を転記できません");

            case MergeValueKind.Boolean:
                return SourceValue.Unsupported("TRUE / FALSE は現在のバージョンでは転記できません");

            default:
                return SourceValue.Unsupported("日付・時刻は現在のバージョンでは転記できません");
        }
    }
}

/// <summary>データ元の読み取りに使う共有情報(共有文字列・表示形式)。</summary>
internal sealed record SourceReadContext(WorkbookReadContext Context, NumberFormatCompatibility NumberFormats)
{
    public static SourceReadContext Create(WorkbookPart workbookPart) => new(
        WorkbookReadContext.Create(workbookPart),
        NumberFormatCompatibility.Create(workbookPart));
}
