using DocumentFormat.OpenXml.Spreadsheet;
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

        // 候補一覧の名前定義を出力へ作るための集約(名前 → 参照先)。
        var requiredDefinedNames = new Dictionary<string, OutputDefinedName>(StringComparer.OrdinalIgnoreCase);

        // ブック内リンクの参照先を解決するための「元シート → 出力シート名」表。
        var outputNameBySourceSheet = new Dictionary<(string File, string Sheet), string>();
        for (var index = 0; index < selections.Count; index++)
        {
            var key = (NormalizePath(selections[index].FilePath), selections[index].SheetName);
            outputNameBySourceSheet[key] = outputNames[index];
        }

        for (var index = 0; index < selections.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var selection = selections[index];
            var fileName = Path.GetFileName(selection.FilePath);
            var outputName = outputNames[index];
            var sheetBlocked = workbookBlocks.TryGetValue(selection.FilePath, out var reasons) && reasons.Count > 0;
            var visibility = SheetVisibility.Visible;
            var printLayout = new PrintLayoutSummary();
            IReadOnlyList<ResolvedHyperlink> hyperlinks = Array.Empty<ResolvedHyperlink>();
            IReadOnlyList<DataValidationSummary> dataValidations = Array.Empty<DataValidationSummary>();
            IReadOnlyList<ResolvedX14ListValidation> x14Validations = Array.Empty<ResolvedX14ListValidation>();

            if (!sheetBlocked)
            {
                var scan = WorksheetCopyScanner.ScanSheet(selection.FilePath, selection.SheetName, cancellationToken);
                visibility = scan.Visibility;
                printLayout = BuildPrintLayoutSummary(scan);
                sheetBlocked = scan.IsBlocked;

                hyperlinks = ResolveHyperlinks(
                    scan, selection, outputNameBySourceSheet, fileName, issues, ref sheetBlocked);
                dataValidations = SummarizeDataValidations(scan);
                x14Validations = ResolveX14ListValidations(
                    scan, selection, outputNameBySourceSheet, fileName, issues, ref sheetBlocked);
                CollectDefinedNames(
                    scan, selection, outputNameBySourceSheet, fileName,
                    requiredDefinedNames, issues, ref sheetBlocked);

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
                PrintLayout = printLayout,
                Hyperlinks = hyperlinks,
                DataValidations = dataValidations,
                X14ListValidations = x14Validations,
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

        return new SheetAggregationPreview
        {
            Sheets = sheets,
            Issues = issues,
            DefinedNames = [.. requiredDefinedNames.Values.OrderBy(name => name.Name, StringComparer.Ordinal)],
        };
    }

    /// <summary>x14 リスト入力規則の参照元を、出力シート名で組み立て直す。</summary>
    private static List<ResolvedX14ListValidation> ResolveX14ListValidations(
        SheetCopyScan scan,
        SheetSelection selection,
        Dictionary<(string File, string Sheet), string> outputNameBySourceSheet,
        string fileName,
        List<MergeIssue> issues,
        ref bool sheetBlocked)
    {
        var resolved = new List<ResolvedX14ListValidation>(scan.X14ListValidations.Count);

        foreach (var validation in scan.X14ListValidations)
        {
            if (validation.BlockReason is not null || validation.Element is null)
            {
                continue; // 走査側で Block 済み。
            }

            // 名前定義を経由するものは formula1 をそのまま使う(名前は出力にも作る)。
            if (validation.DefinedName is { } definedName)
            {
                resolved.Add(new ResolvedX14ListValidation(validation.Sqref, definedName));
                continue;
            }

            if (validation.TargetSheetName is not { } targetSheet)
            {
                // シート名なしの範囲・直接指定はそのまま。
                resolved.Add(new ResolvedX14ListValidation(validation.Sqref, validation.Range));
                continue;
            }

            if (!TryResolveOutputSheet(
                outputNameBySourceSheet, selection.FilePath, targetSheet, out var outputName))
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"セル {validation.Sqref} の入力規則の参照先「{targetSheet}」シートが"
                        + "集約対象に含まれていないため、プルダウンを安全に維持できません。",
                    fileName,
                    selection.SheetName));
                sheetBlocked = true;
                continue;
            }

            resolved.Add(new ResolvedX14ListValidation(
                validation.Sqref,
                $"{SheetReferenceSyntax.Quote(outputName!)}!{validation.Range}"));
        }

        return resolved;
    }

    /// <summary>候補一覧に使われている名前定義を、出力ブック用に集める。</summary>
    private static void CollectDefinedNames(
        SheetCopyScan scan,
        SheetSelection selection,
        Dictionary<(string File, string Sheet), string> outputNameBySourceSheet,
        string fileName,
        Dictionary<string, OutputDefinedName> requiredDefinedNames,
        List<MergeIssue> issues,
        ref bool sheetBlocked)
    {
        var used = scan.DataValidations
            .Where(item => item.DefinedName is not null)
            .Select(item => (item.Sqref, item.DefinedName!, item.TargetSheetName!, item.TargetRange!))
            .Concat(scan.X14ListValidations
                .Where(item => item.DefinedName is not null)
                .Select(item => (item.Sqref, item.DefinedName!, item.TargetSheetName!, item.Range)));

        foreach (var (sqref, name, targetSheet, range) in used)
        {
            if (!TryResolveOutputSheet(
                outputNameBySourceSheet, selection.FilePath, targetSheet, out var outputName))
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"セル {sqref} の入力規則が使う名前定義「{name}」の参照先「{targetSheet}」シートが"
                        + "集約対象に含まれていないため、プルダウンを安全に維持できません。",
                    fileName,
                    selection.SheetName));
                sheetBlocked = true;
                continue;
            }

            var refersTo = $"{SheetReferenceSyntax.Quote(outputName!)}!{range}";

            if (requiredDefinedNames.TryGetValue(name, out var existing))
            {
                // 同名でも参照先が同じなら 1 つにまとめる。違えば安全側に倒して Block。
                if (!string.Equals(existing.RefersTo, refersTo, StringComparison.Ordinal)
                    || !string.Equals(existing.Name, name, StringComparison.Ordinal))
                {
                    issues.Add(new MergeIssue(
                        MergeIssueSeverity.Block,
                        $"名前定義「{name}」が複数のブックで別々の内容を指しているため、"
                            + "現在のバージョンでは安全に集約できません。",
                        fileName,
                        selection.SheetName));
                    sheetBlocked = true;
                }

                continue;
            }

            requiredDefinedNames[name] = new OutputDefinedName(name, refersTo);
        }
    }

    private static bool TryResolveOutputSheet(
        Dictionary<(string File, string Sheet), string> outputNameBySourceSheet,
        string filePath,
        string sourceSheetName,
        out string? outputName)
        => outputNameBySourceSheet.TryGetValue((NormalizePath(filePath), sourceSheetName), out outputName);

    /// <summary>
    /// ハイパーリンクを出力用に解決する。ブック内リンクは参照先シートの出力名で
    /// 組み立て直し、参照先が集約対象に無ければ黙って落とさず Block する。
    /// </summary>
    private static List<ResolvedHyperlink> ResolveHyperlinks(
        SheetCopyScan scan,
        SheetSelection selection,
        Dictionary<(string File, string Sheet), string> outputNameBySourceSheet,
        string fileName,
        List<MergeIssue> issues,
        ref bool sheetBlocked)
    {
        var resolved = new List<ResolvedHyperlink>(scan.Hyperlinks.Count);

        foreach (var link in scan.Hyperlinks)
        {
            if (link.BlockReason is { } reason)
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"セル {link.Reference}: {reason}",
                    fileName,
                    selection.SheetName));
                sheetBlocked = true;
                continue;
            }

            if (link.Kind == HyperlinkKind.InternalOtherSheet)
            {
                var key = (NormalizePath(selection.FilePath), link.TargetSheetName!);
                if (!outputNameBySourceSheet.TryGetValue(key, out var targetOutputName))
                {
                    issues.Add(new MergeIssue(
                        MergeIssueSeverity.Block,
                        $"セル {link.Reference}: リンク先の「{link.TargetSheetName}」シートが"
                            + "集約対象に含まれていないため、リンクを安全に維持できません。",
                        fileName,
                        selection.SheetName));
                    sheetBlocked = true;
                    continue;
                }

                resolved.Add(new ResolvedHyperlink
                {
                    Reference = link.Reference,
                    Location = $"{SheetReferenceSyntax.Quote(targetOutputName)}!{link.Location}",
                    Tooltip = link.Tooltip,
                    Display = link.Display,
                });
                continue;
            }

            resolved.Add(new ResolvedHyperlink
            {
                Reference = link.Reference,
                ExternalTarget = link.ExternalTarget,
                Location = link.Location,
                Tooltip = link.Tooltip,
                Display = link.Display,
            });
        }

        return resolved;
    }

    /// <summary>出力後の検証で照合するため、移植する入力規則の概要を控える。</summary>
    private static List<DataValidationSummary> SummarizeDataValidations(SheetCopyScan scan)
        => [.. scan.DataValidations
            .Select(item => item.Element)
            .OfType<DataValidation>()
            .Select(element => new DataValidationSummary(
                element.SequenceOfReferences?.InnerText ?? string.Empty,
                element.Type?.InnerText ?? "none",
                element.Formula1?.Text,
                element.Formula2?.Text))];

    private static string NormalizePath(string filePath)
    {
        try
        {
            return Path.GetFullPath(filePath).ToLowerInvariant();
        }
        catch (Exception)
        {
            return filePath.ToLowerInvariant();
        }
    }

    internal static PrintLayoutSummary BuildPrintLayoutSummary(SheetCopyScan scan) => new()
    {
        HasPageSetupProperties = scan.PageSetupProperties is not null,
        HasPrintOptions = scan.PrintOptions is not null,
        HasPageMargins = scan.PageMargins is not null,
        HasPageSetup = scan.PageSetup is not null,
        HasHeaderFooter = scan.HeaderFooter is not null,
        RowBreakCount = scan.RowBreaks?.Elements<Break>().Count() ?? 0,
        ColumnBreakCount = scan.ColumnBreaks?.Elements<Break>().Count() ?? 0,
        PrintAreaRanges = RangesOf(scan, PrintDefinedNameKind.PrintArea),
        PrintTitleRanges = RangesOf(scan, PrintDefinedNameKind.PrintTitles),
    };

    private static IReadOnlyList<string> RangesOf(SheetCopyScan scan, PrintDefinedNameKind kind)
        => scan.PrintDefinedNames.FirstOrDefault(name => name.Kind == kind)?.Ranges ?? Array.Empty<string>();

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
