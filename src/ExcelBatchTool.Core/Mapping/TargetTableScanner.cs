using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Mapping;

/// <summary>表として読んだ転記先シートで、キーが一致した 1 行。</summary>
internal sealed record TargetTableMatch(string Key, int RowNumber);

/// <summary>転記先シートを表として読んだ結果。</summary>
internal sealed record TargetTableSheetScan
{
    /// <summary>シート全体として更新できない理由(保護・ピボット・表の読み取り不能など)。</summary>
    public string? BlockReason { get; init; }

    /// <summary>対応付けごとの列番号(1 始まり)。ヘッダーを読めた場合のみ。</summary>
    public IReadOnlyList<int> MappedColumns { get; init; } = Array.Empty<int>();

    /// <summary>データ元に存在するキーと一致した行(シート内でキーが一意のものだけ)。</summary>
    public IReadOnlyList<TargetTableMatch> Matches { get; init; } = Array.Empty<TargetTableMatch>();

    /// <summary>一致した行 × 対応付けのセルの安全確認結果(参照 → 結果)。</summary>
    public IReadOnlyDictionary<string, TargetCellScan> Cells { get; init; }
        = new Dictionary<string, TargetCellScan>(StringComparer.Ordinal);

    /// <summary>このシートで 2 行以上あったキーのうち、データ元にも存在するもの(Block)。</summary>
    public IReadOnlyCollection<string> UsedDuplicateKeys { get; init; } = Array.Empty<string>();

    /// <summary>このシートで 2 行以上あったキーのうち、今回使わないもの(Warning)。</summary>
    public int UnusedDuplicateKeyCount { get; init; }

    /// <summary>このシートに存在するがデータ元に無いキー。</summary>
    public IReadOnlyCollection<string> TargetOnlyKeys { get; init; } = Array.Empty<string>();

    /// <summary>キーが入っていた行の数(重複行も数える)。</summary>
    public int KeyedRowCount { get; init; }

    /// <summary>キーも更新対象の項目もすべて空欄だった行の数。</summary>
    public int BlankRowCount { get; init; }

    /// <summary>キーが空欄なのに更新対象の項目には値があった行の数。</summary>
    public int BlankKeyWithValueCount { get; init; }
}

/// <summary>転記先 Workbook を表として読んだ結果。</summary>
internal sealed record TargetTableWorkbookScan
{
    public IReadOnlyList<string> BlockReasons { get; init; } = Array.Empty<string>();

    public IReadOnlyDictionary<string, TargetTableSheetScan> Sheets { get; init; }
        = new Dictionary<string, TargetTableSheetScan>(StringComparer.Ordinal);

    /// <summary>共通の変更計画(MutationPlanBuilder)が読める形へ変換する。</summary>
    public WorkbookMutationScan ToMutationScan() => new()
    {
        BlockReasons = BlockReasons,
        Sheets = Sheets.ToDictionary(
            entry => entry.Key,
            entry => new SheetMutationScan
            {
                BlockReason = entry.Value.BlockReason,
                Cells = entry.Value.Cells,
            },
            StringComparer.Ordinal),
    };
}

/// <summary>
/// 転記先の Worksheet を「ヘッダー付きの表」として読む(表同士の突合更新用)。
/// Workbook を開くのは 1 ファイル 1 回、シートの走査も 1 回で、
/// ヘッダー・キー列・一致した行・更新対象セルの安全確認までまとめて行う。
///
/// 更新対象セルの guard そのものは <see cref="CellMutationScanner.ScanTargetCell"/> を共有し、
/// 書き換えエンジンを分岐させない。
/// </summary>
internal static class TargetTableScanner
{
    /// <summary>
    /// 1 シートあたりのキー行数の上限。動作を実測で確認した範囲
    /// (10 万行の転記先で preflight / 出力検証を含めて完走)を超えるものは、
    /// 未確認のまま書き換えないために Block する。根拠は D-027。
    /// </summary>
    public const int MaxKeyedRowsPerSheet = 100_000;

    /// <summary>1 つの転記先 Workbook を開き、指定シートを表として読む。</summary>
    public static TargetTableWorkbookScan Scan(
        string filePath,
        IReadOnlyList<string> sheetNames,
        int headerRow,
        string keyColumnName,
        IReadOnlyList<TableColumnMapping> mappings,
        IReadOnlySet<string> sourceKeys,
        CancellationToken cancellationToken,
        int maxKeyedRows = MaxKeyedRowsPerSheet)
    {
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

            // 書き換え対象としての条件(数式・接続・外部参照・Validator)は既存判定と同じ。
            var blocks = new List<string>();
            CellMutationScanner.AddWorkbookBlocks(document, workbookPart, blocks, cancellationToken);

            var context = WorkbookReadContext.Create(workbookPart);
            var numberFormats = NumberFormatCompatibility.Create(workbookPart);

            var sheets = new Dictionary<string, TargetTableSheetScan>(StringComparer.Ordinal);
            foreach (var sheetName in sheetNames)
            {
                cancellationToken.ThrowIfCancellationRequested();
                sheets[sheetName] = ScanSheet(
                    workbookPart, sheetName, headerRow, keyColumnName, mappings,
                    sourceKeys, context, numberFormats, maxKeyedRows, cancellationToken);
            }

            return new TargetTableWorkbookScan { BlockReasons = blocks, Sheets = sheets };
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex) when (ex is InvalidDataException or FileFormatException or OpenXmlPackageException)
        {
            return Blocked(
                "ファイルを読み取れません。パスワード保護(暗号化)されているか、破損している可能性があります。");
        }
        catch (Exception ex)
        {
            return Blocked($"読み取りエラー: {ex.Message}");
        }

        static TargetTableWorkbookScan Blocked(string reason) => new() { BlockReasons = [reason] };
    }

    private static TargetTableSheetScan ScanSheet(
        WorkbookPart workbookPart,
        string sheetName,
        int headerRow,
        string keyColumnName,
        IReadOnlyList<TableColumnMapping> mappings,
        IReadOnlySet<string> sourceKeys,
        WorkbookReadContext context,
        NumberFormatCompatibility numberFormats,
        int maxKeyedRows,
        CancellationToken cancellationToken)
    {
        var sheet = workbookPart.Workbook?.Sheets?.Elements<Sheet>()
            .FirstOrDefault(item => string.Equals(item.Name?.Value, sheetName, StringComparison.Ordinal));

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
            return Blocked("グラフシート・マクロシート等は変更対象にできません(通常のワークシートのみ)。");
        }

        if (worksheetPart.PivotTableParts.Any())
        {
            return Blocked(
                "ピボットテーブルがあるシートは、更新で値が入れ替わる可能性があるため変更できません。");
        }

        var protectionFound = false;
        var headerSeen = false;
        var keyColumn = 0;
        int[]? mappedColumns = null;

        // 一致した行(キー → 行番号と、対応付けごとのセル)。重複したキーは取り除く。
        var matched = new Dictionary<string, (int Row, Cell?[] Cells)>(StringComparer.Ordinal);
        var seenKeys = new HashSet<string>(StringComparer.Ordinal);
        var duplicates = new HashSet<string>(StringComparer.Ordinal);
        var targetOnly = new HashSet<string>(StringComparer.Ordinal);
        var keyedRows = 0;
        var blankRows = 0;
        var blankKeyWithValue = 0;

        // 結合・リンク・入力規則は範囲だけ集め、一致したセルに対してだけ後で照合する。
        var mergeRanges = new List<string>();
        var hyperlinkRanges = new List<string>();
        var validationRanges = new List<string>();

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

                if (type == typeof(SheetProtection))
                {
                    protectionFound = true;
                }
                else if (type == typeof(MergeCell))
                {
                    if (((MergeCell)reader.LoadCurrentElement()!).Reference?.Value is { } merge)
                    {
                        mergeRanges.Add(merge);
                    }
                }
                else if (type == typeof(Hyperlink))
                {
                    if (((Hyperlink)reader.LoadCurrentElement()!).Reference?.Value is { } link)
                    {
                        hyperlinkRanges.Add(link);
                    }
                }
                else if (type == typeof(DataValidation))
                {
                    AddRanges(
                        ((DataValidation)reader.LoadCurrentElement()!).SequenceOfReferences?.InnerText,
                        validationRanges);
                }
                else if (type == typeof(DocumentFormat.OpenXml.Office2010.Excel.DataValidation))
                {
                    var element = (DocumentFormat.OpenXml.Office2010.Excel.DataValidation)
                        reader.LoadCurrentElement()!;
                    AddRanges(element.ReferenceSequence?.Text, validationRanges);
                }
                else if (type == typeof(Row))
                {
                    var row = (Row)reader.LoadCurrentElement()!;
                    if (row.RowIndex?.Value is not { } rowIndex)
                    {
                        continue;
                    }

                    if (rowIndex < (uint)headerRow)
                    {
                        continue;
                    }

                    if (rowIndex == (uint)headerRow)
                    {
                        if (ResolveHeader(row, keyColumnName, mappings, context,
                            out keyColumn, out mappedColumns) is { } headerError)
                        {
                            return Blocked(headerError);
                        }

                        headerSeen = true;
                        continue;
                    }

                    // データ行。行は昇順で並んでいる前提とし、崩れていたら黙って読み違えず止める。
                    if (!headerSeen)
                    {
                        return Blocked(
                            $"項目名の行({headerRow} 行目)より前にデータ行が現れたため、"
                                + "表として読み取れません。行の並びを確認してください。");
                    }

                    var cells = IndexCells(row);
                    var keyCell = cells.GetValueOrDefault(keyColumn);
                    var keyValue = keyCell is null
                        ? MergeCellValue.Blank
                        : context.ReadCell(keyCell, out _);

                    if (keyCell is null || keyValue.Kind == MergeValueKind.Blank)
                    {
                        // キーが空欄の行は表の一部とみなさない。値の有無だけ数える。
                        var hasAnyValue = mappedColumns!.Any(column =>
                            cells.GetValueOrDefault(column) is { } cell
                            && context.ReadCell(cell, out _).Kind != MergeValueKind.Blank);

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

                    // キー列は素の文字列だけの表として扱う(「A001」と数値 1 を推測で比べない)。
                    if (keyCell.CellFormula is not null)
                    {
                        return Blocked(
                            $"キー列({keyColumnName})の {rowIndex} 行目が数式です。計算結果は使いません。");
                    }

                    if (keyCell.CellMetaIndex is not null || keyCell.ValueMetaIndex is not null)
                    {
                        return Blocked(
                            $"キー列({keyColumnName})の {rowIndex} 行目に特別なデータが紐づいています。");
                    }

                    if (context.ReferencesRichText(keyCell))
                    {
                        return Blocked(
                            $"キー列({keyColumnName})の {rowIndex} 行目は文字ごとに書式が設定されています。");
                    }

                    if (keyValue.Kind != MergeValueKind.Text)
                    {
                        return Blocked(
                            $"キー列({keyColumnName})の {rowIndex} 行目が文字列ではありません。"
                                + "「00123」と「123」を取り違えないよう、キー列は文字列のセルだけを対象にします。");
                    }

                    keyedRows++;
                    if (keyedRows > maxKeyedRows)
                    {
                        return Blocked(
                            $"表の行数が動作を確認した範囲({maxKeyedRows:N0} 行)を超えています。"
                                + "現在のバージョンではこの大きさの表は更新できません。");
                    }

                    var key = keyValue.Text!;

                    if (!seenKeys.Add(key))
                    {
                        duplicates.Add(key);
                        matched.Remove(key);
                        continue;
                    }

                    if (sourceKeys.Contains(key))
                    {
                        matched[key] = (
                            (int)rowIndex,
                            [.. mappedColumns!.Select(column => cells.GetValueOrDefault(column))]);
                    }
                    else
                    {
                        targetOnly.Add(key);
                    }
                }
            }
        }

        if (!headerSeen)
        {
            return Blocked($"項目名の行({headerRow} 行目)がこのシートにありません。");
        }

        if (protectionFound)
        {
            return Blocked(
                "シートが保護されているため変更できません。Excel の保護を迂回して書き換えることはしません。");
        }

        // 一致した行 × 対応付けのセルだけに、既存の guard(結合・リンク・入力規則・
        // 数式・書式の適合など)をまとめて掛ける。
        var cellsByReference = new Dictionary<string, TargetCellScan>(StringComparer.Ordinal);
        var matches = new List<TargetTableMatch>(matched.Count);

        var merged = BuildCoverage(mergeRanges, matched, mappedColumns!);
        var hyperlinked = BuildCoverage(hyperlinkRanges, matched, mappedColumns!);
        var validated = BuildCoverage(validationRanges, matched, mappedColumns!);

        foreach (var (key, (rowNumber, cells)) in matched.OrderBy(entry => entry.Value.Row))
        {
            matches.Add(new TargetTableMatch(key, rowNumber));

            for (var index = 0; index < mappings.Count; index++)
            {
                var column = mappedColumns![index];
                var reference =
                    $"{CellRangeParser.ColumnIndexToLetters(column)}{rowNumber}";

                var address = new TargetCellAddress(reference, column, rowNumber);
                cellsByReference[reference] = CellMutationScanner.ScanTargetCell(
                    new ScanTarget(address, mappings[index].WriteKind),
                    cells[index], merged, hyperlinked, validated, context, numberFormats);
            }
        }

        return new TargetTableSheetScan
        {
            MappedColumns = mappedColumns ?? [],
            Matches = matches,
            Cells = cellsByReference,
            UsedDuplicateKeys = [.. duplicates.Where(sourceKeys.Contains)],
            UnusedDuplicateKeyCount = duplicates.Count(key => !sourceKeys.Contains(key)),
            TargetOnlyKeys = targetOnly,
            KeyedRowCount = keyedRows,
            BlankRowCount = blankRows,
            BlankKeyWithValueCount = blankKeyWithValue,
        };

        static TargetTableSheetScan Blocked(string reason) => new() { BlockReason = reason };
    }

    /// <summary>
    /// 項目名の行から、キー列と対応付けの列番号を決める。
    /// ヘッダーの範囲は「その行に実在する最も右のセルまで」。途中の抜けは空の項目名として扱う。
    /// </summary>
    private static string? ResolveHeader(
        Row row,
        string keyColumnName,
        IReadOnlyList<TableColumnMapping> mappings,
        WorkbookReadContext context,
        out int keyColumn,
        out int[]? mappedColumns)
    {
        keyColumn = 0;
        mappedColumns = null;

        var cells = IndexCells(row);
        if (cells.Count == 0)
        {
            return "項目名の行が空です。";
        }

        var lastColumn = cells.Keys.Max();
        var raw = new List<string?>(lastColumn);
        for (var column = 1; column <= lastColumn; column++)
        {
            raw.Add(cells.TryGetValue(column, out var cell) ? HeaderText(cell, context) : null);
        }

        if (!SourceHeaders.Validate(raw, "転記先", out var columns, out var error))
        {
            return error;
        }

        var names = columns!.ToList();

        keyColumn = names.FindIndex(name => string.Equals(name, keyColumnName, StringComparison.Ordinal)) + 1;
        if (keyColumn == 0)
        {
            return $"転記先に項目「{keyColumnName}」がありません。項目名の行と列名を確認してください。";
        }

        var resolved = new int[mappings.Count];
        for (var index = 0; index < mappings.Count; index++)
        {
            var target = mappings[index].TargetColumn;
            var found = names.FindIndex(name => string.Equals(name, target, StringComparison.Ordinal)) + 1;
            if (found == 0)
            {
                return $"転記先に項目「{target}」がありません。項目名の行と列名を確認してください。";
            }

            resolved[index] = found;
        }

        mappedColumns = resolved;
        return null;
    }

    /// <summary>項目名として読む文字列(数式の見出しは空扱いにして止める)。</summary>
    private static string? HeaderText(Cell cell, WorkbookReadContext context)
    {
        if (cell.CellFormula is not null)
        {
            return null;
        }

        var value = context.ReadCell(cell, out _);
        return value.Kind == MergeValueKind.Blank ? null : value.ToDisplayString();
    }

    /// <summary>集めた範囲のうち、一致した行の更新対象セルに掛かるものだけを参照の集合にする。</summary>
    private static HashSet<string> BuildCoverage(
        IReadOnlyList<string> ranges,
        Dictionary<string, (int Row, Cell?[] Cells)> matched,
        int[] mappedColumns)
    {
        var covered = new HashSet<string>(StringComparer.Ordinal);
        if (ranges.Count == 0 || matched.Count == 0)
        {
            return covered;
        }

        foreach (var reference in ranges)
        {
            var normalized = reference.Replace("$", string.Empty, StringComparison.Ordinal);
            if (!CellRangeParser.TryParseRange(normalized, out var range))
            {
                continue;
            }

            foreach (var (_, (rowNumber, _)) in matched)
            {
                if (rowNumber < range.FirstRow || rowNumber > range.LastRow)
                {
                    continue;
                }

                foreach (var column in mappedColumns)
                {
                    if (column >= range.FirstColumn && column <= range.LastColumn)
                    {
                        covered.Add($"{CellRangeParser.ColumnIndexToLetters(column)}{rowNumber}");
                    }
                }
            }
        }

        return covered;
    }

    private static void AddRanges(string? references, List<string> ranges)
    {
        if (string.IsNullOrWhiteSpace(references))
        {
            return;
        }

        ranges.AddRange(references.Split(' ', StringSplitOptions.RemoveEmptyEntries));
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
}

/// <summary>解釈済みの列対応 1 件(転記先スキャンに渡す)。</summary>
internal readonly record struct TableColumnMapping(
    string SourceColumn, int SourceColumnIndex, string TargetColumn, CellWriteKind WriteKind);
