using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelBatchTool.Core.Aggregation;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Mutation;

/// <summary>走査対象の 1 セル(位置と、書き込む値の種類)。</summary>
internal readonly record struct ScanTarget(TargetCellAddress Address, CellWriteKind WriteKind);

/// <summary>1 つの対象セルを変更できるか調べた結果。</summary>
internal sealed record TargetCellScan
{
    public string? BlockReason { get; init; }

    /// <summary>対象セルの現在値(数式は評価しない)。</summary>
    public MergeCellValue CurrentValue { get; init; }
}

/// <summary>照合キーとして読んだセル。</summary>
internal sealed record KeyCellScan(string? Key, string? BlockReason);

/// <summary>1 つのシートを変更できるか調べた結果。</summary>
internal sealed record SheetMutationScan
{
    /// <summary>シート全体として変更できない理由(保護・ピボットなど)。</summary>
    public string? BlockReason { get; init; }

    /// <summary>正規化済みセル参照 → そのセルの走査結果。</summary>
    public IReadOnlyDictionary<string, TargetCellScan> Cells { get; init; }
        = new Dictionary<string, TargetCellScan>(StringComparer.Ordinal);

    /// <summary>照合キーのセル(指定された場合のみ)。</summary>
    public KeyCellScan? KeyCell { get; init; }
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
///
/// 入力セットが何百件あっても Workbook を開くのは 1 ファイル 1 回、
/// シートの走査も 1 シート 1 回で、すべての対象セルをまとめて確かめる。
/// </summary>
internal static class CellMutationScanner
{
    /// <summary>1 つの Workbook を開き、Workbook 全体と指定シートをまとめて調べる。</summary>
    public static WorkbookMutationScan Scan(
        string filePath,
        IReadOnlyList<string> sheetNames,
        IReadOnlyList<ScanTarget> targets,
        TargetCellAddress? keyCell,
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
                    workbookPart, sheetName, targets, keyCell, context, numberFormats, cancellationToken);
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

    /// <summary>Workbook 全体として変更を止めるべき条件(表の突合更新からも同じ判定を使う)。</summary>
    internal static void AddWorkbookBlocks(
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

    /// <summary>1 シートを 1 回だけ走査し、すべての対象セルを確かめる。</summary>
    private static SheetMutationScan ScanSheet(
        WorkbookPart workbookPart,
        string sheetName,
        IReadOnlyList<ScanTarget> targets,
        TargetCellAddress? keyCell,
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

        var found = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        var merged = new HashSet<string>(StringComparer.Ordinal);
        var hyperlinked = new HashSet<string>(StringComparer.Ordinal);
        var validated = new HashSet<string>(StringComparer.Ordinal);
        var protectionFound = false;

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
                    if (cell.CellReference?.Value is { } reference
                        && !found.ContainsKey(reference)
                        && (IsTarget(targets, reference) || IsKeyCell(keyCell, reference)))
                    {
                        found[reference] = cell;
                    }
                }
                else if (type == typeof(MergeCell))
                {
                    // 照合キーのセルも、結合されていれば読み取り位置が曖昧になるので見る。
                    var reference = ((MergeCell)reader.LoadCurrentElement()!).Reference?.Value;
                    MarkCovered(reference, targets, merged);
                    MarkCoveredCell(reference, keyCell, merged);
                }
                else if (type == typeof(Hyperlink))
                {
                    MarkCovered(((Hyperlink)reader.LoadCurrentElement()!).Reference?.Value, targets, hyperlinked);
                }
                else if (type == typeof(DataValidation))
                {
                    MarkCoveredAll(
                        ((DataValidation)reader.LoadCurrentElement()!).SequenceOfReferences?.InnerText,
                        targets, validated);
                }
                else if (type == typeof(DocumentFormat.OpenXml.Office2010.Excel.DataValidation))
                {
                    var element = (DocumentFormat.OpenXml.Office2010.Excel.DataValidation)
                        reader.LoadCurrentElement()!;
                    MarkCoveredAll(element.ReferenceSequence?.Text, targets, validated);
                }
            }
        }

        if (protectionFound)
        {
            return Blocked(
                "シートが保護されているため変更できません。Excel の保護を迂回して書き換えることはしません。");
        }

        var cells = new Dictionary<string, TargetCellScan>(StringComparer.Ordinal);
        foreach (var target in targets)
        {
            cells[target.Address.Reference] = ScanTargetCell(
                target,
                found.GetValueOrDefault(target.Address.Reference),
                merged, hyperlinked, validated, context, numberFormats);
        }

        return new SheetMutationScan
        {
            Cells = cells,
            KeyCell = keyCell is { } key
                ? ReadKeyCell(key, found.GetValueOrDefault(key.Reference), merged, context)
                : null,
        };

        static SheetMutationScan Blocked(string reason) => new() { BlockReason = reason };
    }

    /// <summary>
    /// 照合キーのセルを読む。読むだけなので入力規則やハイパーリンクは問題にしないが、
    /// 「表示形式から 00123 と 123 を推測しない」ため、素の文字列であることを要求する。
    /// </summary>
    private static KeyCellScan ReadKeyCell(
        TargetCellAddress address,
        Cell? cell,
        HashSet<string> merged,
        WorkbookReadContext context)
    {
        var reference = address.Reference;

        if (merged.Contains(reference))
        {
            return Blocked($"照合キーのセル {reference} が結合セルの一部です。");
        }

        if (cell is null)
        {
            return Blocked($"照合キーのセル {reference} がこのシートにありません。");
        }

        if (cell.CellFormula is not null)
        {
            return Blocked($"照合キーのセル {reference} が数式です。計算結果は使いません。");
        }

        if (cell.CellMetaIndex is not null || cell.ValueMetaIndex is not null)
        {
            return Blocked($"照合キーのセル {reference} に特別なデータが紐づいています。");
        }

        if (context.ReferencesRichText(cell))
        {
            return Blocked($"照合キーのセル {reference} は文字ごとに書式が設定されています。");
        }

        var value = context.ReadCell(cell, out _);
        if (value.Kind != MergeValueKind.Text)
        {
            return value.Kind == MergeValueKind.Blank
                ? Blocked($"照合キーのセル {reference} が空欄です。")
                : Blocked(
                    $"照合キーのセル {reference} が文字列ではありません。"
                        + "「00123」と「123」を取り違えないよう、キーは文字列のセルだけを対象にします。");
        }

        return new KeyCellScan(value.Text, null);

        static KeyCellScan Blocked(string reason) => new(null, reason);
    }

    private static bool IsKeyCell(TargetCellAddress? keyCell, string cellReference)
        => keyCell is { } key
            && string.Equals(key.Reference, cellReference, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// 対象セル 1 件の可否を、走査で集めた情報から判定する。
    /// 位置が先に分かっている場合(2A/2B/2C1)も、表の行から見つける場合(2C2)も、
    /// guard はこの 1 つを通す。
    /// </summary>
    internal static TargetCellScan ScanTargetCell(
        ScanTarget target,
        Cell? cell,
        HashSet<string> merged,
        HashSet<string> hyperlinked,
        HashSet<string> validated,
        WorkbookReadContext context,
        NumberFormatCompatibility numberFormats)
    {
        var reference = target.Address.Reference;

        if (merged.Contains(reference))
        {
            return Blocked($"{reference} は結合セルの一部のため、現在のバージョンでは変更できません。");
        }

        if (validated.Contains(reference))
        {
            return Blocked(
                $"{reference} には入力規則が設定されています。新しい値が規則を満たすか判定できないため、"
                    + "現在のバージョンでは変更できません。");
        }

        if (hyperlinked.Contains(reference))
        {
            return Blocked(
                $"{reference} にはハイパーリンクが設定されています。表示だけを変えるとリンク先と"
                    + "食い違う可能性があるため、現在のバージョンでは変更できません。");
        }

        if (cell is null)
        {
            return Blocked(
                $"{reference} はこのシートに存在しないため、現在のバージョンでは変更できません"
                    + "(空のセルを新しく作ることはしません)。");
        }

        if (cell.CellFormula is not null)
        {
            return Blocked($"{reference} は数式のため変更できません。数式を値に置き換えることはしません。");
        }

        // cm / vm はセル値に紐づくメタデータ(リンクされたデータ型など)への参照。
        // 値だけ差し替えると参照先と食い違うため触らない。
        if (cell.CellMetaIndex is not null || cell.ValueMetaIndex is not null)
        {
            return Blocked(
                $"{reference} には特別なデータ(リンクされたデータ型など)が紐づいているため、"
                    + "現在のバージョンでは変更できません。");
        }

        if (context.ReferencesRichText(cell))
        {
            return Blocked(
                $"{reference} は文字ごとに書式が設定されているため、"
                    + "書き換えると書式が失われます。現在のバージョンでは変更できません。");
        }

        if (numberFormats.Check(cell.StyleIndex?.Value, target.WriteKind, reference) is { } formatError)
        {
            return Blocked(formatError);
        }

        return new TargetCellScan { CurrentValue = context.ReadCell(cell, out _) };

        static TargetCellScan Blocked(string reason) => new() { BlockReason = reason };
    }

    /// <summary>このセル参照が対象セルのどれかと一致するか。</summary>
    private static bool IsTarget(IReadOnlyList<ScanTarget> targets, string cellReference)
    {
        foreach (var target in targets)
        {
            if (string.Equals(target.Address.Reference, cellReference, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>1 つの範囲(A1 / A1:B5)に含まれる対象セルを記録する。</summary>
    private static void MarkCovered(
        string? reference, IReadOnlyList<ScanTarget> targets, HashSet<string> covered)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            return;
        }

        var normalized = reference.Replace("$", string.Empty, StringComparison.Ordinal);
        if (!CellRangeParser.TryParseRange(normalized, out var range))
        {
            return;
        }

        foreach (var target in targets)
        {
            if (target.Address.Column >= range.FirstColumn && target.Address.Column <= range.LastColumn
                && target.Address.Row >= range.FirstRow && target.Address.Row <= range.LastRow)
            {
                covered.Add(target.Address.Reference);
            }
        }
    }

    /// <summary>1 つの範囲に照合キーのセルが含まれるかを記録する。</summary>
    private static void MarkCoveredCell(
        string? reference, TargetCellAddress? keyCell, HashSet<string> covered)
    {
        if (keyCell is { } key)
        {
            MarkCovered(reference, [new ScanTarget(key, CellWriteKind.Text)], covered);
        }
    }

    /// <summary>空白区切りの範囲リスト(sqref)に含まれる対象セルを記録する。</summary>
    private static void MarkCoveredAll(
        string? references, IReadOnlyList<ScanTarget> targets, HashSet<string> covered)
    {
        if (string.IsNullOrWhiteSpace(references))
        {
            return;
        }

        foreach (var token in references.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            MarkCovered(token, targets, covered);
        }
    }
}
