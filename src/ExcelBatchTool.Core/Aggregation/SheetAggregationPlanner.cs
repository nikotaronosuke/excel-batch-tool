using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>
/// Sheet 集約前の検証(プレビュー)を作る。セルの中身は読み込まず、
/// 集約できるかどうかと出力シート名・並び順だけを決める。
/// 出力シートの並び順は、渡された選択の順序をそのまま使う。
/// </summary>
public sealed class SheetAggregationPlanner
{
    /// <summary>1 つの Workbook に作れるシート数の実用上限(極端な指定を弾くための保険)。</summary>
    private const int MaxOutputSheets = 1000;

    public SheetAggregationPreview CreatePreview(
        IReadOnlyList<SheetSelection> selections,
        CancellationToken cancellationToken = default)
    {
        var issues = new List<MergeIssue>();

        if (selections.Count == 0)
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, "集約するシートが選択されていません。"));
            return new SheetAggregationPreview { Sheets = Array.Empty<SheetAggregationPlan>(), Issues = issues };
        }

        if (selections.Count > MaxOutputSheets)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"一度に集約できるシートは {MaxOutputSheets:N0} 枚までです(選択 {selections.Count:N0} 枚)。"));
        }

        foreach (var duplicate in selections
            .GroupBy(selection => (Path.GetFullPath(selection.FilePath).ToLowerInvariant(), selection.SheetName))
            .Where(group => group.Count() > 1))
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                "同じシートが複数回選択されています。",
                Path.GetFileName(duplicate.Key.Item1),
                duplicate.Key.SheetName));
        }

        // Workbook 単位の検証は 1 ファイルにつき 1 回。
        var workbookBlocks = new Dictionary<string, IReadOnlyList<string>>(StringComparer.OrdinalIgnoreCase);
        var themeFingerprints = new HashSet<string>(StringComparer.Ordinal);

        foreach (var filePath in selections.Select(s => s.FilePath).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var scan = WorksheetCopyScanner.ScanWorkbook(filePath);
            workbookBlocks[filePath] = scan.BlockReasons;

            foreach (var reason in scan.BlockReasons)
            {
                issues.Add(new MergeIssue(MergeIssueSeverity.Block, reason, Path.GetFileName(filePath)));
            }

            if (scan.BlockReasons.Count == 0 && scan.ThemeFingerprint is { } fingerprint)
            {
                themeFingerprints.Add(fingerprint);
            }
        }

        if (themeFingerprints.Count > 1)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Warning,
                "配色テーマが異なるブックが混在しています。出力には最初のブックのテーマを使うため、" +
                "テーマ色で指定された色は見え方が変わることがあります。"));
        }

        var outputNames = ResolveOutputNames(selections);
        var usedNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var sheets = new List<SheetAggregationPlan>(selections.Count);

        for (var index = 0; index < selections.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var selection = selections[index];
            var fileName = Path.GetFileName(selection.FilePath);
            var outputName = outputNames[index];
            var sheetBlocked = workbookBlocks.TryGetValue(selection.FilePath, out var reasons) && reasons.Count > 0;
            var visibility = SheetVisibility.Visible;

            if (!sheetBlocked)
            {
                var scan = WorksheetCopyScanner.ScanSheet(selection.FilePath, selection.SheetName, cancellationToken);
                visibility = scan.Visibility;
                sheetBlocked = scan.IsBlocked;

                foreach (var reason in scan.BlockReasons)
                {
                    issues.Add(new MergeIssue(MergeIssueSeverity.Block, reason, fileName, selection.SheetName));
                }

                foreach (var reason in scan.WarningReasons)
                {
                    issues.Add(new MergeIssue(MergeIssueSeverity.Warning, reason, fileName, selection.SheetName));
                }
            }

            if (OutputSheetNameResolver.Validate(outputName) is { } nameError)
            {
                issues.Add(new MergeIssue(MergeIssueSeverity.Block, nameError, fileName, selection.SheetName));
                sheetBlocked = true;
            }
            else if (usedNames.TryGetValue(outputName, out var owner))
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"出力シート名「{outputName}」が「{owner}」と重複しています。別の名前を指定してください。",
                    fileName,
                    selection.SheetName));
                sheetBlocked = true;
            }
            else
            {
                usedNames[outputName] = $"{fileName} / {selection.SheetName}";
            }

            sheets.Add(new SheetAggregationPlan
            {
                FilePath = selection.FilePath,
                FileName = fileName,
                SheetName = selection.SheetName,
                OutputSheetName = outputName,
                Visibility = visibility,
                IsBlocked = sheetBlocked,
                Order = index + 1,
            });
        }

        // Excel は「すべてのシートが非表示」の Workbook を開けない。
        if (sheets.Count > 0 && sheets.All(sheet => sheet.IsHidden))
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                "選択したシートがすべて非表示です。Excel は表示シートが 1 枚も無いブックを開けないため、" +
                "表示されているシートを 1 枚以上選んでください。"));
        }

        return new SheetAggregationPreview { Sheets = sheets, Issues = issues };
    }

    /// <summary>
    /// 出力シート名を決める。利用者が指定した名前はそのまま使い(勝手に置き換えない)、
    /// 未指定のものだけ元シート名から決定的に提案する。
    /// </summary>
    private static string[] ResolveOutputNames(IReadOnlyList<SheetSelection> selections)
    {
        var names = new string[selections.Count];
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // 明示指定を先に確保してから、残りを提案する。
        // 空文字なども「利用者が指定した名前」として扱い、勝手に別名へ置き換えず検証で弾く。
        for (var index = 0; index < selections.Count; index++)
        {
            if (selections[index].OutputSheetName is { } specified)
            {
                names[index] = specified;
                used.Add(specified);
            }
        }

        for (var index = 0; index < selections.Count; index++)
        {
            if (names[index] is not null)
            {
                continue;
            }

            var proposed = OutputSheetNameResolver.Propose(selections[index].SheetName, used);
            names[index] = proposed;
            used.Add(proposed);
        }

        return names;
    }
}
