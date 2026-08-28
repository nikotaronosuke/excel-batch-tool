using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Aggregation;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Mutation;

/// <summary>1 つのシートを変更できるか調べた結果。</summary>
internal sealed record SheetMutationScan
{
    public string? BlockReason { get; init; }

    /// <summary>対象セルの現在値(数式は評価しない)。</summary>
    public MergeCellValue CurrentValue { get; init; }
}

/// <summary>1 つの Workbook を変更できるか調べた結果。</summary>
internal sealed record WorkbookMutationScan
{
    /// <summary>Workbook 全体として変更できない理由。</summary>
    public IReadOnlyList<string> BlockReasons { get; init; } = Array.Empty<string>();

    /// <summary>シート名 → そのシートの走査結果。</summary>
    public IReadOnlyDictionary<string, SheetMutationScan> Sheets { get; init; }
        = new Dictionary<string, SheetMutationScan>(StringComparer.Ordinal);
}

/// <summary>
/// 一括変更の対象を検証する。対象ファイルは読み取り専用でしか開かない。
/// 「変更しても意味が変わらないと言い切れるもの」だけを通し、
/// 判断できないものは黙って書き換えず理由付きで Block する。
/// </summary>
internal static class CellMutationScanner
{
    /// <summary>1 つの Workbook を開き、Workbook 全体と指定シートをまとめて調べる。</summary>
    public static WorkbookMutationScan Scan(
        string filePath,
        IReadOnlyList<string> sheetNames,
        TargetCellAddress address,
        CellWriteKind writeKind,
        CancellationToken cancellationToken)
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

            var blocks = new List<string>();
            AddWorkbookBlocks(document, workbookPart, blocks, cancellationToken);

            var context = WorkbookReadContext.Create(workbookPart);
            var numberFormats = NumberFormatCompatibility.Create(workbookPart);

            var sheets = new Dictionary<string, SheetMutationScan>(StringComparer.Ordinal);
            foreach (var sheetName in sheetNames)
            {
                cancellationToken.ThrowIfCancellationRequested();
                sheets[sheetName] = ScanSheet(
                    workbookPart, sheetName, address, writeKind, context, numberFormats, cancellationToken);
            }

            return new WorkbookMutationScan { BlockReasons = blocks, Sheets = sheets };
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

        static WorkbookMutationScan Blocked(string reason) => new() { BlockReasons = [reason] };
    }

    /// <summary>Workbook 全体として変更を止めるべき条件。</summary>
    private static void AddWorkbookBlocks(
        SpreadsheetDocument document,
        WorkbookPart workbookPart,
        List<string> blocks,
        CancellationToken cancellationToken)
    {
        // 定数セルを書き換えると、その値を参照する数式の計算結果が古いまま残る。
        // このツールは計算エンジンを持たないため、数式を含むブックは対象外にする。
        if (ContainsFormula(workbookPart, cancellationToken))
        {
            blocks.Add(
                "このファイルには数式が含まれているため、変更後の計算結果を保証できません。"
                    + "現在のバージョンでは一括変更の対象外です。");
        }

        if (workbookPart.ConnectionsPart is not null)
        {
            blocks.Add(
                "外部データ接続を含むため、変更した値が更新で上書きされる可能性があります。"
                    + "現在のバージョンでは一括変更の対象外です。");
        }

        var hasExternalLink = workbookPart.ExternalWorkbookParts.Any()
            || (workbookPart.Workbook?.GetFirstChild<ExternalReferences>()?
                .Elements<ExternalReference>().Any() ?? false);

        if (hasExternalLink)
        {
            blocks.Add(
                "他のブックへの外部参照(外部リンク)を含むため、現在のバージョンでは一括変更の対象外です。");
        }

        // 既に壊れているファイルを書き換えると、原因の切り分けができなくなる。
        var errors = new OpenXmlValidator().Validate(document).Take(1).ToList();
        if (errors.Count > 0)
        {
            blocks.Add(
                "このファイルは Excel の形式として問題がある箇所を含むため、"
                    + "現在のバージョンでは一括変更の対象外です。");
        }
    }

    /// <summary>Workbook 内のどこかに数式があるか(1 件見つかった時点で打ち切る)。</summary>
    private static bool ContainsFormula(WorkbookPart workbookPart, CancellationToken cancellationToken)
    {
        if (workbookPart.CalculationChainPart is not null)
        {
            return true;
        }

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            using var reader = OpenXmlReader.Create(worksheetPart);
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (reader.IsStartElement && reader.ElementType == typeof(CellFormula))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static SheetMutationScan ScanSheet(
        WorkbookPart workbookPart,
        string sheetName,
        TargetCellAddress address,
        CellWriteKind writeKind,
        WorkbookReadContext context,
        NumberFormatCompatibility numberFormats,
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

        Cell? target = null;
        var protectionFound = false;
        var mergedFound = false;
        var hyperlinkFound = false;
        var validationFound = false;

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
                else if (type == typeof(Cell))
                {
                    var cell = (Cell)reader.LoadCurrentElement()!;
                    if (Matches(cell.CellReference?.Value, address))
                    {
                        target = cell;
                    }
                }
                else if (type == typeof(MergeCell))
                {
                    var reference = ((MergeCell)reader.LoadCurrentElement()!).Reference?.Value;
                    mergedFound |= Contains(reference, address);
                }
                else if (type == typeof(Hyperlink))
                {
                    var reference = ((Hyperlink)reader.LoadCurrentElement()!).Reference?.Value;
                    hyperlinkFound |= Contains(reference, address);
                }
                else if (type == typeof(DataValidation))
                {
                    var sqref = ((DataValidation)reader.LoadCurrentElement()!).SequenceOfReferences?.InnerText;
                    validationFound |= ContainsAny(sqref, address);
                }
                else if (type == typeof(DocumentFormat.OpenXml.Office2010.Excel.DataValidation))
                {
                    var element = (DocumentFormat.OpenXml.Office2010.Excel.DataValidation)
                        reader.LoadCurrentElement()!;
                    validationFound |= ContainsAny(element.ReferenceSequence?.Text, address);
                }
            }
        }

        if (protectionFound)
        {
            return Blocked(
                "シートが保護されているため変更できません。Excel の保護を迂回して書き換えることはしません。");
        }

        if (mergedFound)
        {
            return Blocked($"{address.Reference} は結合セルの一部のため、現在のバージョンでは変更できません。");
        }

        if (validationFound)
        {
            return Blocked(
                $"{address.Reference} には入力規則が設定されています。新しい値が規則を満たすか判定できないため、"
                    + "現在のバージョンでは変更できません。");
        }

        if (hyperlinkFound)
        {
            return Blocked(
                $"{address.Reference} にはハイパーリンクが設定されています。表示だけを変えるとリンク先と"
                    + "食い違う可能性があるため、現在のバージョンでは変更できません。");
        }

        if (target is null)
        {
            return Blocked(
                $"{address.Reference} はこのシートに存在しないため、現在のバージョンでは変更できません"
                    + "(空のセルを新しく作ることはしません)。");
        }

        if (target.CellFormula is not null)
        {
            return Blocked($"{address.Reference} は数式のため変更できません。数式を値に置き換えることはしません。");
        }

        // cm / vm はセル値に紐づくメタデータ(リンクされたデータ型など)への参照。
        // 値だけ差し替えると参照先と食い違うため触らない。
        if (target.CellMetaIndex is not null || target.ValueMetaIndex is not null)
        {
            return Blocked(
                $"{address.Reference} には特別なデータ(リンクされたデータ型など)が紐づいているため、"
                    + "現在のバージョンでは変更できません。");
        }

        if (context.ReferencesRichText(target))
        {
            return Blocked(
                $"{address.Reference} は文字ごとに書式が設定されているため、"
                    + "書き換えると書式が失われます。現在のバージョンでは変更できません。");
        }

        if (numberFormats.Check(target.StyleIndex?.Value, writeKind, address.Reference) is { } formatError)
        {
            return Blocked(formatError);
        }

        return new SheetMutationScan { CurrentValue = context.ReadCell(target, out _) };

        static SheetMutationScan Blocked(string reason) => new() { BlockReason = reason };
    }

    private static bool Matches(string? cellReference, TargetCellAddress address)
        => cellReference is not null
            && string.Equals(cellReference, address.Reference, StringComparison.OrdinalIgnoreCase);

    /// <summary>1 つの範囲(A1 / A1:B5)に対象セルが含まれるか。</summary>
    private static bool Contains(string? reference, TargetCellAddress address)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            return false;
        }

        var normalized = reference.Replace("$", string.Empty, StringComparison.Ordinal);
        return CellRangeParser.TryParseRange(normalized, out var range)
            && address.Column >= range.FirstColumn && address.Column <= range.LastColumn
            && address.Row >= range.FirstRow && address.Row <= range.LastRow;
    }

    /// <summary>空白区切りの範囲リスト(sqref)に対象セルが含まれるか。</summary>
    private static bool ContainsAny(string? references, TargetCellAddress address)
        => !string.IsNullOrWhiteSpace(references)
            && references.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Any(token => Contains(token, address));
}
