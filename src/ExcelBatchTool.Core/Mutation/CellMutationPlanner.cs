using System.Globalization;
using System.Security.Cryptography;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Mutation;

/// <summary>
/// 一括変更の実行前検証(プレビュー)を作る。対象ファイルは読み取りしかしない。
/// ここで Block が 1 件でもあれば実行させない(入力セットの一部だけ適用することもしない)。
/// </summary>
public sealed class CellMutationPlanner
{
    /// <summary>1 回に扱えるファイル数の実用上限(極端な指定を弾くための保険)。</summary>
    private const int MaxFiles = 500;

    public CellMutationPreview CreatePreview(
        CellMutationRequest request,
        CancellationToken cancellationToken = default)
    {
        var issues = new List<MergeIssue>();

        if (request.Targets.Count == 0)
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, "変更するシートが選択されていません。"));
            return Empty(issues);
        }

        if (request.Operations.Count == 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, "変更するセルが指定されていません。行を追加してください。"));
            return Empty(issues);
        }

        // 入力セットをすべて解釈する。1 件でも解釈できなければファイルを開かずに終える。
        var operations = ResolveOperations(request.Operations, issues);
        if (operations is null)
        {
            return Empty(issues);
        }

        if (OutputNaming.ValidateSuffix(request.OutputSuffix) is { } suffixError)
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, suffixError));
            return Empty(issues);
        }

        foreach (var duplicate in request.Targets
            .GroupBy(target => (NormalizePath(target.FilePath), target.SheetName))
            .Where(group => group.Count() > 1))
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                "同じシートが複数回選択されています。",
                Path.GetFileName(duplicate.Key.Item1),
                duplicate.Key.SheetName));
        }

        // ファイル単位でまとめる(1 ファイルにつき出力は 1 つ)。
        var groups = request.Targets
            .GroupBy(target => target.FilePath, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (groups.Count > MaxFiles)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"一度に変更できるファイルは {MaxFiles:N0} 個までです(選択 {groups.Count:N0} 個)。"));
        }

        var sourcePaths = new HashSet<string>(
            groups.Select(group => NormalizePath(group.Key)), StringComparer.Ordinal);

        var scanTargets = operations
            .Select(operation => new ScanTarget(operation.Address, operation.Value.Kind))
            .ToList();

        var targets = new List<CellMutationTargetPlan>();
        var files = new List<CellMutationFilePlan>();
        var usedOutputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var group in groups)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var filePath = group.Key;
            var fileName = Path.GetFileName(filePath);
            var sheetNames = group.Select(target => target.SheetName).Distinct(StringComparer.Ordinal).ToList();

            var scan = CellMutationScanner.Scan(filePath, sheetNames, scanTargets, cancellationToken);

            var outputError = TryResolveOutput(
                filePath, request.OutputSuffix, sourcePaths, usedOutputs,
                out var outputPath, out var auditPath);

            var outputFileName = outputPath is null ? "-" : Path.GetFileName(outputPath);

            // Workbook 全体の問題は、そのファイルのすべての対象に付ける。
            var fileBlock = scan.BlockReasons.FirstOrDefault() ?? outputError;

            foreach (var reason in scan.BlockReasons)
            {
                issues.Add(new MergeIssue(MergeIssueSeverity.Block, reason, fileName));
            }

            if (outputError is not null)
            {
                issues.Add(new MergeIssue(MergeIssueSeverity.Block, outputError, fileName));
            }

            var changes = new List<CellMutationTargetPlan>();

            foreach (var sheetName in sheetNames)
            {
                var sheetScan = scan.Sheets.GetValueOrDefault(sheetName);

                if (fileBlock is null && sheetScan?.BlockReason is { } sheetReason)
                {
                    issues.Add(new MergeIssue(MergeIssueSeverity.Block, sheetReason, fileName, sheetName));
                }

                // 入力セットの並び順のまま計画を作る(プレビューと控えの順序を利用者の入力に合わせる)。
                foreach (var operation in operations)
                {
                    var cellScan = sheetScan?.Cells.GetValueOrDefault(operation.Address.Reference);
                    var reason = fileBlock ?? sheetScan?.BlockReason ?? cellScan?.BlockReason;

                    var current = cellScan?.CurrentValue ?? MergeCellValue.Blank;
                    var isNoOp = reason is null && IsSameValue(current, operation.Value);

                    var plan = new CellMutationTargetPlan
                    {
                        FilePath = filePath,
                        FileName = fileName,
                        SheetName = sheetName,
                        CellReference = operation.Address.Reference,
                        NewValue = operation.Value,
                        CurrentValueDisplay = reason is null ? CellValueDisplay.Of(current) : "-",
                        CurrentTypeName = CellValueDisplay.TypeNameOf(current),
                        NewValueDisplay = operation.Value.Display,
                        NewTypeName = operation.Value.TypeName,
                        OutputFileName = outputFileName,
                        BlockReason = reason,
                        IsNoOp = isNoOp,
                    };

                    targets.Add(plan);

                    if (reason is null && !isNoOp)
                    {
                        changes.Add(plan);
                    }

                    // セル単位の問題だけを、シート単位の重複なしに報告する。
                    if (cellScan?.BlockReason is { } cellReason
                        && fileBlock is null && sheetScan?.BlockReason is null)
                    {
                        issues.Add(new MergeIssue(MergeIssueSeverity.Block, cellReason, fileName, sheetName));
                    }
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
                Snapshot = TakeSnapshot(filePath),
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

        if (issues.All(issue => issue.Severity != MergeIssueSeverity.Block)
            && files.Count == 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                "変更が必要なセルがありません。新しいファイルは作成しません。"));
        }

        return new CellMutationPreview { Targets = targets, Files = files, Issues = issues };

        static CellMutationPreview Empty(List<MergeIssue> issues) => new()
        {
            Targets = Array.Empty<CellMutationTargetPlan>(),
            Files = Array.Empty<CellMutationFilePlan>(),
            Issues = issues,
        };
    }

    /// <summary>
    /// 入力セットを解釈する(位置・値・重複)。問題があれば issues へ理由を足して null を返す。
    /// どの行が悪いかまとめて分かるよう、途中で打ち切らずすべて確かめる。
    /// </summary>
    private static List<ResolvedOperation>? ResolveOperations(
        IReadOnlyList<CellMutationOperationRequest> requests,
        List<MergeIssue> issues)
    {
        var resolved = new List<ResolvedOperation>(requests.Count);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var reportedDuplicates = new HashSet<string>(StringComparer.Ordinal);
        var failed = false;

        foreach (var request in requests)
        {
            if (!TargetCellAddress.TryParse(request.CellReference, out var address, out var addressError))
            {
                issues.Add(new MergeIssue(MergeIssueSeverity.Block, addressError!));
                failed = true;
                continue;
            }

            // $B$2 と B2 は同じセルとして扱う(正規化済みの参照で比べる)。
            if (!seen.Add(address.Reference))
            {
                if (reportedDuplicates.Add(address.Reference))
                {
                    issues.Add(new MergeIssue(
                        MergeIssueSeverity.Block,
                        $"セル「{address.Reference}」が入力セット内で重複しています。"
                            + "同じセルには 1 つの値だけを指定してください。"));
                }

                failed = true;
                continue;
            }

            if (!TryResolveNewValue(request, address.Reference, out var value, out var valueError))
            {
                issues.Add(new MergeIssue(MergeIssueSeverity.Block, valueError!));
                failed = true;
                continue;
            }

            resolved.Add(new ResolvedOperation(address, value));
        }

        return failed ? null : resolved;
    }

    /// <summary>元ファイルが実行直前に変わっていないか確かめるための控えを取る。</summary>
    internal static SourceSnapshot TakeSnapshot(string filePath)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        var info = new FileInfo(filePath);
        return new SourceSnapshot(hash, info.Length, info.LastWriteTimeUtc);
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

        if (sourcePaths.Contains(NormalizePath(candidate)))
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

    /// <summary>新しい値を解釈する。</summary>
    private static bool TryResolveNewValue(
        CellMutationOperationRequest request,
        string reference,
        out NewCellValue value,
        out string? error)
    {
        value = default;
        error = null;

        switch (request.WriteKind)
        {
            case CellWriteKind.Blank:
                value = NewCellValue.Blank();
                return true;

            case CellWriteKind.Text:
                if (string.IsNullOrEmpty(request.TextValue))
                {
                    error = $"「{reference}」の新しい値を入力してください"
                        + "(空欄にする場合は種類を「空欄」にしてください)。";
                    return false;
                }

                value = NewCellValue.OfText(request.TextValue);
                return true;

            default:
                var text = request.NumberText?.Trim();
                if (string.IsNullOrEmpty(text))
                {
                    error = $"「{reference}」の新しい数値を入力してください。";
                    return false;
                }

                if (!double.TryParse(
                        text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                    || !double.IsFinite(number))
                {
                    error = $"「{reference}」の値「{text}」を数値として読み取れません(例: 100、-1.5)。";
                    return false;
                }

                value = NewCellValue.OfNumber(number);
                return true;
        }
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

    /// <summary>解釈済みの入力セット 1 項目。</summary>
    private readonly record struct ResolvedOperation(TargetCellAddress Address, NewCellValue Value);
}

/// <summary>解釈済みの「新しい値」。</summary>
internal readonly record struct NewCellValue(CellWriteKind Kind, string? Text, double Number)
{
    public static NewCellValue Blank() => new(CellWriteKind.Blank, null, 0);

    public static NewCellValue OfText(string text) => new(CellWriteKind.Text, text, 0);

    public static NewCellValue OfNumber(double number) => new(CellWriteKind.Number, null, number);

    public string Display => Kind switch
    {
        CellWriteKind.Blank => CellValueDisplay.Blank,
        CellWriteKind.Text => Text ?? string.Empty,
        _ => Number.ToString(CultureInfo.InvariantCulture),
    };

    /// <summary>控えファイルに書く型名。</summary>
    public string TypeName => Kind switch
    {
        CellWriteKind.Blank => "blank",
        CellWriteKind.Text => "text",
        _ => "number",
    };
}

/// <summary>出力ファイル名の組み立てと検証。</summary>
internal static class OutputNaming
{
    public static string BuildFileName(string sourceFileName, string suffix)
        => Path.GetFileNameWithoutExtension(sourceFileName) + suffix + Path.GetExtension(sourceFileName);

    /// <summary>接尾辞が使えるか。使えない場合は理由を返す。</summary>
    public static string? ValidateSuffix(string? suffix)
    {
        if (string.IsNullOrEmpty(suffix))
        {
            return "出力ファイル名に付ける文字を入力してください(既定: "
                + $"{CellMutationDefaults.OutputSuffix})。元のファイルは上書きしません。";
        }

        var invalid = Path.GetInvalidFileNameChars();
        return suffix.Any(invalid.Contains)
            ? "出力ファイル名に付ける文字に、ファイル名として使えない記号が含まれています。"
            : null;
    }
}
