using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace ExcelBatchTool.Core.Mutation;

/// <summary>
/// プレビュー済みの計画にしたがって、選択したシートの同じセルへ同じ値を書き込む。
///
/// 元ファイルは絶対に書き換えない。元ファイルをコピーし、コピー側の
/// 対象 WorksheetPart だけを変更する。理解できない Part(グラフ・図・テーブル等)は
/// 触らないことで、そのまま保持する。
///
/// 出力は「一時ファイルへ作成 → 開き直して検証 → すべて成功したら確定」で決める。
/// 途中で 1 件でも失敗したら、出力を 1 つも残さない。
/// </summary>
public sealed class CellMutator
{
    /// <summary>検証エラーを表示する最大件数。</summary>
    private const int MaxReportedValidationErrors = 3;

    public CellMutationResult Execute(
        CellMutationPreview preview,
        IProgress<CellMutationProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (!preview.CanExecute)
        {
            return CellMutationResult.Failed("解決していない問題があるため、一括変更を実行できません。");
        }

        // 1. 実行直前の再確認。プレビュー後に元ファイルや出力先が変わっていたら全体を中止する。
        if (Preflight(preview) is { } preflightError)
        {
            return CellMutationResult.Failed(preflightError);
        }

        var pending = new List<PendingOutput>();
        var moved = new List<string>();

        try
        {
            var completed = 0;
            foreach (var file in preview.Files)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // 作る前に一時ファイルの場所を控える。途中で失敗しても取り消し対象から漏れない。
                var output = ReserveOutput(file);
                pending.Add(output);
                BuildOutput(preview, file, output, cancellationToken);

                completed++;
                progress?.Report(new CellMutationProgress(completed, preview.Files.Count));
            }

            // 2. すべての一時ファイルが検証を通ったあとで確定する。
            foreach (var output in pending)
            {
                if (File.Exists(output.WorkbookPath) || File.Exists(output.AuditPath))
                {
                    throw new InvalidOperationException(
                        $"「{Path.GetFileName(output.WorkbookPath)}」が実行中に作成されました。"
                            + "既存ファイルは上書きしません。");
                }

                File.Move(output.TempWorkbookPath, output.WorkbookPath);
                moved.Add(output.WorkbookPath);

                File.Move(output.TempAuditPath, output.AuditPath);
                moved.Add(output.AuditPath);
            }
        }
        catch (OperationCanceledException)
        {
            return CellMutationResult.Failed(
                Describe("一括変更を中止しました。", RollBack(pending, moved)));
        }
        catch (Exception ex)
        {
            return CellMutationResult.Failed(
                Describe($"一括変更に失敗しました: {ex.Message}", RollBack(pending, moved)));
        }

        var changeCount = preview.ChangeCount;
        return new CellMutationResult
        {
            Success = true,
            OutputFileNames = [.. preview.Files.Select(file => file.OutputFileName)],
            ChangedCellCount = changeCount,
            Message = $"{preview.Files.Count:N0} ファイルに {changeCount:N0} 件の変更を書き込みました。"
                + "元のファイルは変更していません。",
        };
    }

    /// <summary>元ファイルと出力先が、プレビュー時と同じ状態か確かめる。</summary>
    private static string? Preflight(CellMutationPreview preview)
    {
        foreach (var file in preview.Files)
        {
            if (!File.Exists(file.FilePath))
            {
                return $"「{file.FileName}」が見つかりません。もう一度プレビューを更新してください。";
            }

            SourceSnapshot current;
            try
            {
                current = CellMutationPlanner.TakeSnapshot(file.FilePath);
            }
            catch (Exception ex)
            {
                return $"「{file.FileName}」を読み取れません: {ex.Message}";
            }

            if (current != file.Snapshot)
            {
                return $"「{file.FileName}」がプレビュー後に変更されました。"
                    + "安全のため、どのファイルも変更していません。もう一度プレビューを更新してください。";
            }

            if (File.Exists(file.OutputPath))
            {
                return $"「{file.OutputFileName}」が既にあります。既存ファイルは上書きしません。";
            }

            if (File.Exists(file.AuditPath))
            {
                return $"「{Path.GetFileName(file.AuditPath)}」が既にあります。既存ファイルは上書きしません。";
            }
        }

        return null;
    }

    /// <summary>この出力に使う一時ファイルと確定先の組を決める(まだ何も作らない)。</summary>
    private static PendingOutput ReserveOutput(CellMutationFilePlan file)
    {
        var directory = Path.GetDirectoryName(file.OutputPath)!;
        var id = Guid.NewGuid().ToString("N");

        return new PendingOutput(
            Path.Combine(directory, $"~ebt-mut-{id}.xlsx"),
            file.OutputPath,
            Path.Combine(directory, $"~ebt-mut-{id}.json"),
            file.AuditPath);
    }

    /// <summary>元ファイルをコピーし、コピー側だけを変更して検証する。</summary>
    private static void BuildOutput(
        CellMutationPreview preview,
        CellMutationFilePlan file,
        PendingOutput output,
        CancellationToken cancellationToken)
    {
        // 元ファイルはここでしか読まない(読み取り専用のコピー)。
        File.Copy(file.FilePath, output.TempWorkbookPath);

        var applied = ApplyChanges(preview, file, output.TempWorkbookPath, cancellationToken);

        if (Validate(output.TempWorkbookPath, preview, applied) is { } validationError)
        {
            throw new InvalidOperationException(
                $"{file.FileName}: 出力ファイルの検証に失敗しました: {validationError}");
        }

        MutationAuditLog.Write(output.TempAuditPath, file, applied, Hash(output.TempWorkbookPath));
    }

    /// <summary>
    /// コピーした Workbook の対象セルだけを書き換える。
    /// <see cref="OpenSettings.AutoSave"/> を false にして、保存するのは対象 WorksheetPart だけにする
    /// (既定のままだと、シートを探すために読んだ workbook.xml まで書き戻されてしまう)。
    /// </summary>
    private static List<AppliedChange> ApplyChanges(
        CellMutationPreview preview,
        CellMutationFilePlan file,
        string tempPath,
        CancellationToken cancellationToken)
    {
        var applied = new List<AppliedChange>(file.Changes.Count);

        using var document = SpreadsheetDocument.Open(
            tempPath, isEditable: true, new OpenSettings { AutoSave = false });

        var workbookPart = document.WorkbookPart
            ?? throw new InvalidOperationException($"{file.FileName}: Workbook 情報が見つかりません。");

        foreach (var change in file.Changes)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var sheet = workbookPart.Workbook?.Sheets?.Elements<Sheet>()
                .FirstOrDefault(item => string.Equals(item.Name?.Value, change.SheetName, StringComparison.Ordinal));

            if (sheet?.Id?.Value is not { } relationshipId
                || workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                throw new InvalidOperationException(
                    $"{file.FileName}: ワークシート「{change.SheetName}」が見つかりません。");
            }

            var worksheet = worksheetPart.Worksheet
                ?? throw new InvalidOperationException(
                    $"{file.FileName}: シート「{change.SheetName}」の内容を読み取れません。");

            var cell = worksheet.Descendants<Cell>().FirstOrDefault(item =>
                string.Equals(item.CellReference?.Value, change.CellReference, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException(
                    $"{file.FileName} / {change.SheetName}: {change.CellReference} が見つかりません。");

            // 書式(StyleIndex)は触らない。値の中身だけを入れ替える。
            var styleIndex = cell.StyleIndex?.Value;
            SetValue(cell, preview.NewValue);

            worksheet.Save();
            applied.Add(new AppliedChange(change, styleIndex));
        }

        return applied;
    }

    /// <summary>セルの値の中身だけを入れ替える(書式・位置はそのまま)。</summary>
    private static void SetValue(Cell cell, NewCellValue value)
    {
        cell.CellFormula = null;
        cell.CellValue = null;
        cell.InlineString = null;
        cell.DataType = null;

        switch (value.Kind)
        {
            case CellWriteKind.Text:
                // 共有文字列表(sharedStrings)を増やさないよう、セルの中に直接書く。
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(
                    new Text(value.Text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
                break;

            case CellWriteKind.Number:
                cell.CellValue = new CellValue(value.Number.ToString(CultureInfo.InvariantCulture));
                break;

            // Blank: 値の中身を消したままにする。
        }
    }

    /// <summary>作った一時ファイルを開き直し、対象セルと Open XML の妥当性を確かめる。</summary>
    private static string? Validate(
        string path,
        CellMutationPreview preview,
        IReadOnlyList<AppliedChange> applied)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var document = SpreadsheetDocument.Open(stream, isEditable: false);

            var workbookPart = document.WorkbookPart;
            if (workbookPart is null)
            {
                return "Workbook 情報がありません。";
            }

            foreach (var (change, styleIndex) in applied)
            {
                var sheet = workbookPart.Workbook?.Sheets?.Elements<Sheet>()
                    .FirstOrDefault(item =>
                        string.Equals(item.Name?.Value, change.SheetName, StringComparison.Ordinal));

                if (sheet?.Id?.Value is not { } relationshipId
                    || workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
                {
                    return $"シート「{change.SheetName}」がありません。";
                }

                var cell = worksheetPart.Worksheet?.Descendants<Cell>().FirstOrDefault(item =>
                    string.Equals(
                        item.CellReference?.Value, change.CellReference, StringComparison.OrdinalIgnoreCase));

                if (cell is null)
                {
                    return $"シート「{change.SheetName}」の {change.CellReference} がありません。";
                }

                if (cell.StyleIndex?.Value != styleIndex)
                {
                    return $"シート「{change.SheetName}」の {change.CellReference} の書式が変わっています。";
                }

                if (CheckValue(cell, preview.NewValue) is { } valueError)
                {
                    return $"シート「{change.SheetName}」の {change.CellReference}: {valueError}";
                }
            }

            var errors = new OpenXmlValidator().Validate(document).ToList();
            if (errors.Count > 0)
            {
                var details = errors
                    .Take(MaxReportedValidationErrors)
                    .Select(error => $"{error.Path?.XPath}: {error.Description}");
                return $"Open XML の検証エラーが {errors.Count} 件あります。{string.Join(" / ", details)}";
            }

            return null;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }

    private static string? CheckValue(Cell cell, NewCellValue value)
    {
        switch (value.Kind)
        {
            case CellWriteKind.Text:
                if (cell.DataType?.Value != CellValues.InlineString)
                {
                    return "文字として書き込まれていません。";
                }

                var text = cell.InlineString?.Text?.Text;
                return string.Equals(text, value.Text, StringComparison.Ordinal)
                    ? null
                    : $"値が想定と異なります(想定「{value.Text}」/ 実際「{text}」)。";

            case CellWriteKind.Number:
                if (cell.DataType?.Value is { } type && type != CellValues.Number)
                {
                    return "数値として書き込まれていません。";
                }

                return double.TryParse(
                        cell.CellValue?.InnerText, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                        && number.Equals(value.Number)
                    ? null
                    : $"値が想定と異なります(想定 {value.Number.ToString(CultureInfo.InvariantCulture)}"
                        + $" / 実際「{cell.CellValue?.InnerText}」)。";

            default:
                return cell.CellValue is null && cell.InlineString is null && cell.CellFormula is null
                    ? null
                    : "空欄になっていません。";
        }
    }

    /// <summary>
    /// 失敗時は一時ファイルを消し、既に確定したものも取り消す。
    /// 取り消しは best effort なので、消せたつもりで終わらせず、
    /// 最後に本当に残っていないかを確かめて残存分を返す。
    /// 実行前から在ったファイルには触れない(消したのは自分が作ったものだけ)。
    /// </summary>
    private RollbackOutcome RollBack(IReadOnlyList<PendingOutput> pending, IReadOnlyList<string> moved)
    {
        var candidates = new List<string>(moved);
        foreach (var output in pending)
        {
            candidates.Add(output.TempWorkbookPath);
            candidates.Add(output.TempAuditPath);
        }

        var remaining = new List<string>();
        foreach (var path in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            FileDeleter(path);

            // 削除処理の戻り値ではなく、実際に残っているかどうかを最終判断にする。
            if (File.Exists(path))
            {
                remaining.Add(Path.GetFileName(path));
            }
        }

        return new RollbackOutcome(remaining);
    }

    /// <summary>
    /// 取り消しの結果を、断定しすぎずに伝える。残ったファイルはファイル名だけを出す
    /// (利用者向けのメッセージに絶対パスを載せない)。
    /// 元ファイルは読み取りしかしていないので、そこだけは断定してよい。
    /// </summary>
    private static string Describe(string reason, RollbackOutcome outcome)
        => outcome.IsClean
            ? $"{reason}作成途中のファイルは取り消しました。元のファイルは変更していません。"
            : $"{reason}取り消せなかったファイルが残っている可能性があります。"
                + $"次のファイルを確認してください: {string.Join(" / ", outcome.RemainingFileNames)}。"
                + "元のファイルは変更していません。";

    internal static string Hash(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    /// <summary>
    /// ファイルを 1 つ消す。消せたか(元から無かった場合も含む)を返す。
    /// 取り消しが本当に効かない状況をテストで再現できるよう、ここだけ差し替えられるようにしている。
    /// </summary>
    internal Func<string, bool> FileDeleter { get; init; } = TryDeleteFile;

    internal static bool TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            return !File.Exists(path);
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>確定前の出力 1 組(Workbook と控えファイル)。</summary>
    private sealed record PendingOutput(
        string TempWorkbookPath,
        string WorkbookPath,
        string TempAuditPath,
        string AuditPath);

    /// <summary>取り消しの結果。残ったものはファイル名だけを持つ。</summary>
    private sealed record RollbackOutcome(IReadOnlyList<string> RemainingFileNames)
    {
        public bool IsClean => RemainingFileNames.Count == 0;
    }
}

/// <summary>書き換えた 1 セルと、書き換え前の書式。</summary>
internal readonly record struct AppliedChange(CellMutationTargetPlan Change, uint? StyleIndex);
