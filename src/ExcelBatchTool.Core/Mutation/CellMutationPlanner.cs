using System.Globalization;
using System.Security.Cryptography;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Mutation;

/// <summary>
/// 一括変更の実行前検証(プレビュー)を作る。対象ファイルは読み取りしかしない。
/// ここで Block が 1 件でもあれば実行させない。
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

        if (!TargetCellAddress.TryParse(request.CellReference, out var address, out var addressError))
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, addressError!));
            return Empty(issues);
        }

        if (!TryResolveNewValue(request, out var newValue, out var valueError))
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, valueError!));
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

        var targets = new List<CellMutationTargetPlan>();
        var files = new List<CellMutationFilePlan>();
        var usedOutputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var group in groups)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var filePath = group.Key;
            var fileName = Path.GetFileName(filePath);
            var sheetNames = group.Select(target => target.SheetName).Distinct(StringComparer.Ordinal).ToList();

            var scan = CellMutationScanner.Scan(
                filePath, sheetNames, address, request.WriteKind, cancellationToken);

            var outputError = TryResolveOutput(
                filePath, request.OutputSuffix, sourcePaths, usedOutputs,
                out var outputPath, out var auditPath);

            var outputFileName = outputPath is null ? "-" : Path.GetFileName(outputPath);

            // Workbook 全体の問題は、そのファイルのすべてのシートに付ける。
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
                var reason = fileBlock ?? sheetScan?.BlockReason;

                var current = sheetScan?.CurrentValue ?? MergeCellValue.Blank;
                var isNoOp = reason is null && IsSameValue(current, newValue, request);

                var plan = new CellMutationTargetPlan
                {
                    FilePath = filePath,
                    FileName = fileName,
                    SheetName = sheetName,
                    CellReference = address.Reference,
                    CurrentValueDisplay = reason is null ? CellValueDisplay.Of(current) : "-",
                    CurrentTypeName = CellValueDisplay.TypeNameOf(current),
                    NewValueDisplay = newValue.Display,
                    NewTypeName = newValue.TypeName,
                    OutputFileName = outputFileName,
                    BlockReason = reason,
                    IsNoOp = isNoOp,
                };

                targets.Add(plan);

                if (reason is null && !isNoOp)
                {
                    changes.Add(plan);
                }

                if (reason is not null && scan.BlockReasons.Count == 0 && outputError is null)
                {
                    issues.Add(new MergeIssue(MergeIssueSeverity.Block, reason, fileName, sheetName));
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
                "変更が必要なシートがありません。新しいファイルは作成しません。"));
        }

        return new CellMutationPreview
        {
            Targets = targets,
            Files = files,
            Issues = issues,
            Address = address,
            NewValue = newValue,
        };

        static CellMutationPreview Empty(List<MergeIssue> issues) => new()
        {
            Targets = Array.Empty<CellMutationTargetPlan>(),
            Files = Array.Empty<CellMutationFilePlan>(),
            Issues = issues,
        };
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
        CellMutationRequest request,
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
                    error = "新しい値を入力してください(空欄にする場合は「空欄にする」を選んでください)。";
                    return false;
                }

                value = NewCellValue.OfText(request.TextValue);
                return true;

            default:
                var text = request.NumberText?.Trim();
                if (string.IsNullOrEmpty(text))
                {
                    error = "新しい数値を入力してください。";
                    return false;
                }

                if (!double.TryParse(
                        text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                    || !double.IsFinite(number))
                {
                    error = $"「{text}」を数値として読み取れません(例: 100、-1.5)。";
                    return false;
                }

                value = NewCellValue.OfNumber(number);
                return true;
        }
    }

    /// <summary>現在の値と新しい値が、型を含めて同じか。</summary>
    private static bool IsSameValue(
        MergeCellValue current, NewCellValue newValue, CellMutationRequest request)
        => request.WriteKind switch
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
