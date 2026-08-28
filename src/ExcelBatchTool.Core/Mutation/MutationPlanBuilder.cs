using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Mutation;

/// <summary>
/// 「このファイルのこのシートのこのセルへ、この値を書く」という解決済みの 1 件。
/// 手入力の入力セット(Phase 2B)でも、データ元からの転記(Phase 2C)でも、
/// ここまで解決できれば以降の安全確認・出力計画・書き込みは共通になる。
/// </summary>
internal sealed record ResolvedCellMutation
{
    public required string FilePath { get; init; }

    public required string SheetName { get; init; }

    public required TargetCellAddress Address { get; init; }

    public required NewCellValue Value { get; init; }

    /// <summary>この値をデータ元のどこから取ったか(手入力なら null)。</summary>
    public MutationProvenance? Provenance { get; init; }
}

/// <summary>
/// 解決済みの変更から、安全確認済みのプレビューを組み立てる。
/// Phase 2B と Phase 2C で、対象の走査・guard・No-op 判定・出力計画を二重に実装しないための共通部分。
/// </summary>
internal static class MutationPlanBuilder
{
    /// <summary>1 回に扱えるファイル数の実用上限(極端な指定を弾くための保険)。</summary>
    public const int MaxFiles = 500;

    /// <summary>
    /// 解決済みの変更と、その対象ファイルの走査結果からプレビューを作る。
    /// <paramref name="mutations"/> の並び順は保持する(プレビューと控えを利用者の入力順に合わせるため)。
    /// </summary>
    public static CellMutationPreview Build(
        IReadOnlyList<ResolvedCellMutation> mutations,
        IReadOnlyDictionary<string, WorkbookMutationScan> scansByFile,
        string outputSuffix,
        List<MergeIssue> issues,
        MutationDataSourceInfo? dataSource = null,
        MutationDataSourceCheck? dataSourceCheck = null,
        bool reportScanIssues = true,
        MutationTargetTableInfo? targetTable = null)
    {
        if (reportScanIssues)
        {
            ReportScanIssues(mutations, scansByFile, issues);
        }

        var sourcePaths = new HashSet<string>(
            mutations.Select(mutation => MutationPaths.Normalize(mutation.FilePath)), StringComparer.Ordinal);

        var targets = new List<CellMutationTargetPlan>();
        var files = new List<CellMutationFilePlan>();
        var usedOutputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // ファイル単位でまとめる(1 ファイルにつき出力は 1 つ)。GroupBy は最初に現れた順を保つ。
        foreach (var fileGroup in mutations.GroupBy(
            mutation => mutation.FilePath, StringComparer.OrdinalIgnoreCase))
        {
            var filePath = fileGroup.Key;
            var fileName = Path.GetFileName(filePath);
            var scan = scansByFile.GetValueOrDefault(MutationPaths.Normalize(filePath));

            var outputError = TryResolveOutput(
                filePath, outputSuffix, sourcePaths, usedOutputs, out var outputPath, out var auditPath);

            var outputFileName = outputPath is null ? "-" : Path.GetFileName(outputPath);

            // Workbook 全体の問題は、そのファイルのすべての対象に付ける。
            var fileBlock = scan?.BlockReasons.FirstOrDefault() ?? outputError;

            if (outputError is not null)
            {
                issues.Add(new MergeIssue(MergeIssueSeverity.Block, outputError, fileName));
            }

            var changes = new List<CellMutationTargetPlan>();

            foreach (var mutation in fileGroup)
            {
                var sheetScan = scan?.Sheets.GetValueOrDefault(mutation.SheetName);
                var cellScan = sheetScan?.Cells.GetValueOrDefault(mutation.Address.Reference);
                var reason = fileBlock ?? sheetScan?.BlockReason ?? cellScan?.BlockReason;

                var current = cellScan?.CurrentValue ?? MergeCellValue.Blank;
                var isNoOp = reason is null && IsSameValue(current, mutation.Value);

                var plan = new CellMutationTargetPlan
                {
                    FilePath = filePath,
                    FileName = fileName,
                    SheetName = mutation.SheetName,
                    CellReference = mutation.Address.Reference,
                    NewValue = mutation.Value,
                    CurrentValueDisplay = reason is null ? CellValueDisplay.Of(current) : "-",
                    CurrentTypeName = CellValueDisplay.TypeNameOf(current),
                    NewValueDisplay = mutation.Value.Display,
                    NewTypeName = mutation.Value.TypeName,
                    Provenance = mutation.Provenance,
                    OutputFileName = outputFileName,
                    BlockReason = reason,
                    IsNoOp = isNoOp,
                };

                targets.Add(plan);

                if (reason is null && !isNoOp)
                {
                    changes.Add(plan);
                }

                // セル単位の問題だけを、シート単位・ファイル単位の重複なしに報告する。
                if (cellScan?.BlockReason is { } cellReason
                    && fileBlock is null && sheetScan?.BlockReason is null)
                {
                    issues.Add(new MergeIssue(
                        MergeIssueSeverity.Block, cellReason, fileName, mutation.SheetName));
                }
            }

            if (changes.Count == 0 || outputPath is null)
            {
                // 変更が 1 件も無いファイルには出力を作らない。
                continue;
            }

            files.Add(new CellMutationFilePlan
            {
                FilePath = filePath,
                FileName = fileName,
                OutputFileName = outputFileName,
                OutputPath = outputPath,
                AuditPath = auditPath!,
                Snapshot = MutationSnapshot.Take(filePath),
                DataSource = dataSource,
                TargetTable = targetTable,
                Changes = changes,
            });
        }

        var noOpCount = targets.Count(target => !target.IsBlocked && target.IsNoOp);
        if (noOpCount > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Warning,
                $"{noOpCount:N0} 件は現在の値と同じため、変更しません。"));
        }

        if (issues.All(issue => issue.Severity != MergeIssueSeverity.Block) && files.Count == 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                "変更が必要なセルがありません。新しいファイルは作成しません。"));
        }

        return new CellMutationPreview
        {
            Targets = targets,
            Files = files,
            Issues = issues,
            DataSourceCheck = dataSourceCheck,
        };
    }

    /// <summary>
    /// 走査で見つかったファイル単位・シート単位の問題を報告する。
    /// 値が 1 つも解決できなかった場合でも「なぜ止まったか」が出るよう、
    /// 変更の生成とは切り離して呼べるようにしている。
    /// </summary>
    public static void ReportScanIssues(
        IReadOnlyList<ResolvedCellMutation> mutations,
        IReadOnlyDictionary<string, WorkbookMutationScan> scansByFile,
        List<MergeIssue> issues)
        => ReportScanIssues(
            [.. mutations.Select(mutation => (mutation.FilePath, mutation.SheetName))],
            scansByFile, issues);

    /// <summary>ファイル × シートの組だけで報告する版(値が 1 つも解決できない場合に使う)。</summary>
    public static void ReportScanIssues(
        IReadOnlyList<(string FilePath, string SheetName)> locations,
        IReadOnlyDictionary<string, WorkbookMutationScan> scansByFile,
        List<MergeIssue> issues)
    {
        var reportedSheets = new HashSet<(string, string)>();

        foreach (var fileGroup in locations.GroupBy(
            location => location.FilePath, StringComparer.OrdinalIgnoreCase))
        {
            var fileName = Path.GetFileName(fileGroup.Key);
            var scan = scansByFile.GetValueOrDefault(MutationPaths.Normalize(fileGroup.Key));

            foreach (var reason in scan?.BlockReasons ?? [])
            {
                issues.Add(new MergeIssue(MergeIssueSeverity.Block, reason, fileName));
            }

            if (scan is null || scan.BlockReasons.Count > 0)
            {
                continue;
            }

            foreach (var sheetName in fileGroup
                .Select(location => location.SheetName)
                .Distinct(StringComparer.Ordinal))
            {
                if (scan.Sheets.GetValueOrDefault(sheetName)?.BlockReason is { } sheetReason
                    && reportedSheets.Add((fileGroup.Key, sheetName)))
                {
                    issues.Add(new MergeIssue(MergeIssueSeverity.Block, sheetReason, fileName, sheetName));
                }
            }
        }
    }

    /// <summary>対象ファイルごとに 1 回だけ開いて走査する。</summary>
    public static Dictionary<string, WorkbookMutationScan> ScanTargets(
        IReadOnlyList<ResolvedCellMutation> mutations,
        TargetCellAddress? keyCell,
        CancellationToken cancellationToken)
    {
        var scans = new Dictionary<string, WorkbookMutationScan>(StringComparer.Ordinal);

        foreach (var fileGroup in mutations.GroupBy(
            mutation => mutation.FilePath, StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var sheetNames = fileGroup
                .Select(mutation => mutation.SheetName)
                .Distinct(StringComparer.Ordinal)
                .ToList();

            // 同じセルは種類も同じ(重複指定は事前に Block 済み)なので、参照で一意にできる。
            var scanTargets = fileGroup
                .GroupBy(mutation => mutation.Address.Reference, StringComparer.Ordinal)
                .Select(group => new ScanTarget(group.First().Address, group.First().Value.Kind))
                .ToList();

            scans[MutationPaths.Normalize(fileGroup.Key)] = CellMutationScanner.Scan(
                fileGroup.Key, sheetNames, scanTargets, keyCell, cancellationToken);
        }

        return scans;
    }

    /// <summary>出力先を決める。既にファイルがある場合は勝手に名前を変えず Block する。</summary>
    private static string? TryResolveOutput(
        string filePath,
        string suffix,
        HashSet<string> sourcePaths,
        Dictionary<string, string> usedOutputs,
        out string? outputPath,
        out string? auditPath)
    {
        outputPath = null;
        auditPath = null;

        string full;
        try
        {
            full = Path.GetFullPath(filePath);
        }
        catch (Exception ex)
        {
            return $"ファイルのパスを解釈できません: {ex.Message}";
        }

        var directory = Path.GetDirectoryName(full);
        if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
        {
            return "元ファイルのフォルダーが見つかりません。";
        }

        var candidate = Path.Combine(
            directory, OutputNaming.BuildFileName(Path.GetFileName(full), suffix));
        var audit = candidate + CellMutationDefaults.AuditExtension;
        var name = Path.GetFileName(candidate);

        if (sourcePaths.Contains(MutationPaths.Normalize(candidate)))
        {
            return $"出力先「{name}」が変更対象のファイルと同じです。入力ファイルは変更しません。";
        }

        if (usedOutputs.TryGetValue(candidate, out var owner))
        {
            return $"出力先「{name}」が「{owner}」と重複しています。";
        }

        if (File.Exists(candidate))
        {
            return $"「{name}」は既にあります。既存ファイルは上書きしません。"
                + "別の名前を指定するか、既存のファイルを移動してください。";
        }

        if (File.Exists(audit))
        {
            return $"「{Path.GetFileName(audit)}」は既にあります。既存ファイルは上書きしません。";
        }

        usedOutputs[candidate] = Path.GetFileName(full);
        outputPath = candidate;
        auditPath = audit;
        return null;
    }

    /// <summary>現在の値と新しい値が、型を含めて同じか。</summary>
    private static bool IsSameValue(MergeCellValue current, NewCellValue newValue)
        => newValue.Kind switch
        {
            CellWriteKind.Blank => current.Kind == MergeValueKind.Blank,
            CellWriteKind.Text => current.Kind == MergeValueKind.Text
                && string.Equals(current.Text, newValue.Text, StringComparison.Ordinal),
            _ => current.Kind == MergeValueKind.Number && current.Number.Equals(newValue.Number),
        };
}

/// <summary>パスの比較用の正規化。</summary>
internal static class MutationPaths
{
    public static string Normalize(string filePath)
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
}

/// <summary>実行直前の照合に使う、ファイルの控え。</summary>
internal static class MutationSnapshot
{
    public static SourceSnapshot Take(string filePath)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
        var info = new FileInfo(filePath);
        return new SourceSnapshot(hash, info.Length, info.LastWriteTimeUtc);
    }
}
