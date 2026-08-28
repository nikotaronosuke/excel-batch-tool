using System.Globalization;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.Core.Mutation;

/// <summary>
/// 手入力の入力セット(Phase 2B)から、実行前検証(プレビュー)を作る。
/// 対象ファイルは読み取りしかしない。Block が 1 件でもあれば実行させない
/// (入力セットの一部だけ適用することもしない)。
///
/// 対象の走査・guard・No-op 判定・出力計画は <see cref="MutationPlanBuilder"/> と共有する。
/// このクラスの役割は「新しい値をどこから得るか」= 利用者の手入力の解釈だけ。
/// </summary>
public sealed class CellMutationPlanner
{
    public CellMutationPreview CreatePreview(
        CellMutationRequest request,
        CancellationToken cancellationToken = default)
    {
        var issues = new List<MergeIssue>();

        if (request.Targets.Count == 0)
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, "変更するシートが選択されていません。"));
            return MutationPreviewFactory.Empty(issues);
        }

        if (request.Operations.Count == 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, "変更するセルが指定されていません。行を追加してください。"));
            return MutationPreviewFactory.Empty(issues);
        }

        // 入力セットをすべて解釈する。1 件でも解釈できなければファイルを開かずに終える。
        var operations = ResolveOperations(request.Operations, issues);
        if (operations is null)
        {
            return MutationPreviewFactory.Empty(issues);
        }

        if (OutputNaming.ValidateSuffix(request.OutputSuffix) is { } suffixError)
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, suffixError));
            return MutationPreviewFactory.Empty(issues);
        }

        if (MutationTargets.Validate(request.Targets, issues) is not { } targets)
        {
            return MutationPreviewFactory.Empty(issues);
        }

        // 同じ入力セットを、選択したすべてのシートへ広げる。
        var mutations = new List<ResolvedCellMutation>(targets.Count * operations.Count);
        foreach (var target in targets)
        {
            foreach (var operation in operations)
            {
                mutations.Add(new ResolvedCellMutation
                {
                    FilePath = target.FilePath,
                    SheetName = target.SheetName,
                    Address = operation.Address,
                    Value = operation.Value,
                });
            }
        }

        var scans = MutationPlanBuilder.ScanTargets(mutations, keyCell: null, cancellationToken);
        return MutationPlanBuilder.Build(mutations, scans, request.OutputSuffix, issues);
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
    internal static SourceSnapshot TakeSnapshot(string filePath) => MutationSnapshot.Take(filePath);

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

    /// <summary>解釈済みの入力セット 1 項目。</summary>
    private readonly record struct ResolvedOperation(TargetCellAddress Address, NewCellValue Value);
}

/// <summary>対象シートの選択内容の検証(Phase 2B / 2C 共通)。</summary>
internal static class MutationTargets
{
    /// <summary>重複選択と件数上限を確かめる。問題があれば issues へ足して null を返す。</summary>
    public static IReadOnlyList<CellMutationTarget>? Validate(
        IReadOnlyList<CellMutationTarget> targets,
        List<MergeIssue> issues)
    {
        var failed = false;

        foreach (var duplicate in targets
            .GroupBy(target => (MutationPaths.Normalize(target.FilePath), target.SheetName))
            .Where(group => group.Count() > 1))
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                "同じシートが複数回選択されています。",
                Path.GetFileName(duplicate.Key.Item1),
                duplicate.Key.SheetName));
            failed = true;
        }

        var fileCount = targets
            .Select(target => MutationPaths.Normalize(target.FilePath))
            .Distinct(StringComparer.Ordinal)
            .Count();

        if (fileCount > MutationPlanBuilder.MaxFiles)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"一度に変更できるファイルは {MutationPlanBuilder.MaxFiles:N0} 個までです"
                    + $"(選択 {fileCount:N0} 個)。"));
            failed = true;
        }

        return failed ? null : targets;
    }
}

/// <summary>中身の無いプレビュー(解釈の段階で止まったとき)。</summary>
internal static class MutationPreviewFactory
{
    public static CellMutationPreview Empty(List<MergeIssue> issues) => new()
    {
        Targets = Array.Empty<CellMutationTargetPlan>(),
        Files = Array.Empty<CellMutationFilePlan>(),
        Issues = issues,
    };
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
