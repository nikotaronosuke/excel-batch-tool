namespace ExcelBatchTool.Core.Merge;

/// <summary>
/// 統合前の検証(プレビュー)を作る。実データは読み込まず、Header と行数・Block 要因だけを集める。
/// 入力ファイルは読み取り専用でしか開かない。
/// </summary>
public sealed class MergePlanner
{
    /// <summary>Excel のワークシート最大行数。</summary>
    private const int MaxWorksheetRows = 1_048_576;

    public MergePreview CreatePreview(
        IReadOnlyList<MergeSourceSelection> selections,
        MergeOptions options,
        CancellationToken cancellationToken = default)
    {
        var issues = new List<MergeIssue>();

        if (selections.Count == 0)
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, "統合対象が選択されていません。"));
            return new MergePreview { Sources = Array.Empty<MergeSourcePlan>(), Issues = issues };
        }

        var duplicated = selections
            .GroupBy(selection => Path.GetFullPath(selection.FilePath), StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1);
        foreach (var group in duplicated)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                "同じファイルが複数回選択されています。1 ファイルにつき 1 シートを選択してください。",
                Path.GetFileName(group.Key)));
        }

        // 各シートを走査して Header とデータ行数を集める。
        var scans = new List<(MergeSourceSelection Selection, SheetScanResult Scan)>();
        foreach (var selection in selections)
        {
            cancellationToken.ThrowIfCancellationRequested();
            scans.Add((selection, WorksheetTableScanner.Scan(selection.FilePath, selection.SheetName, cancellationToken)));
        }

        foreach (var (selection, scan) in scans)
        {
            var fileName = Path.GetFileName(selection.FilePath);
            foreach (var reason in scan.BlockReasons)
            {
                issues.Add(new MergeIssue(MergeIssueSeverity.Block, reason, fileName, selection.SheetName));
            }

            foreach (var reason in scan.WarningReasons)
            {
                issues.Add(new MergeIssue(MergeIssueSeverity.Warning, reason, fileName, selection.SheetName));
            }
        }

        // 基準は「最初に Header を読めた選択シート」。出力列順はこの Header 順を使う。
        var baseScan = scans.FirstOrDefault(entry => !entry.Scan.IsBlocked && entry.Scan.Headers.Count > 0);
        var dataHeaders = baseScan.Scan?.Headers ?? Array.Empty<string>();

        var sources = new List<MergeSourcePlan>();
        foreach (var (selection, scan) in scans)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var fileName = Path.GetFileName(selection.FilePath);
            var columnMap = Array.Empty<int>();

            if (!scan.IsBlocked && dataHeaders.Count > 0)
            {
                columnMap = BuildColumnMap(scan.Headers, dataHeaders, out var missing, out var extra);

                if (missing.Count > 0)
                {
                    issues.Add(new MergeIssue(
                        MergeIssueSeverity.Block,
                        $"ヘッダーが一致しません。不足ヘッダー: {string.Join("、", missing)}",
                        fileName,
                        selection.SheetName));
                }

                if (extra.Count > 0)
                {
                    issues.Add(new MergeIssue(
                        MergeIssueSeverity.Block,
                        $"ヘッダーが一致しません。余分なヘッダー: {string.Join("、", extra)}",
                        fileName,
                        selection.SheetName));
                }
            }

            sources.Add(new MergeSourcePlan
            {
                FilePath = selection.FilePath,
                FileName = fileName,
                SheetName = selection.SheetName,
                Headers = scan.Headers,
                DataRowCount = scan.DataRowCount,
                IsBlocked = scan.IsBlocked,
                ColumnMap = columnMap,
            });
        }

        var metadataNames = BuildMetadataColumnNames(options);
        foreach (var name in metadataNames)
        {
            if (dataHeaders.Contains(name, StringComparer.Ordinal))
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"追加しようとしている列「{name}」が元データのヘッダーと重複しています。" +
                    "この列の追加をオフにするか、元データのヘッダー名を変更してください。"));
            }
        }

        var inputDataRowCount = sources.Sum(source => source.DataRowCount);
        if (inputDataRowCount + 1 > MaxWorksheetRows)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"出力予定の行数({inputDataRowCount + 1:N0} 行)が 1 シートの上限({MaxWorksheetRows:N0} 行)を超えます。"));
        }

        if (dataHeaders.Count == 0 && issues.All(issue => issue.Severity != MergeIssueSeverity.Block))
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, "統合できるヘッダーが見つかりませんでした。"));
        }

        return new MergePreview
        {
            Sources = sources,
            DataHeaders = dataHeaders,
            OutputHeaders = [.. metadataNames, .. dataHeaders],
            MetadataColumnCount = metadataNames.Count,
            InputDataRowCount = inputDataRowCount,
            Issues = issues,
        };
    }

    internal static List<string> BuildMetadataColumnNames(MergeOptions options)
    {
        var names = new List<string>(2);
        if (options.IncludeSourceFileColumn)
        {
            names.Add(options.SourceFileColumnName);
        }

        if (options.IncludeSourceSheetColumn)
        {
            names.Add(options.SourceSheetColumnName);
        }

        return names;
    }

    /// <summary>
    /// Header 名(trim 済み)で列を対応付ける。列順が違っても同じ名前の集合なら統合できる。
    /// 大文字小文字・Unicode 正規化などの推測はせず、文字列として厳密に比較する。
    /// </summary>
    private static int[] BuildColumnMap(
        IReadOnlyList<string> sourceHeaders,
        IReadOnlyList<string> baseHeaders,
        out List<string> missing,
        out List<string> extra)
    {
        var baseIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var i = 0; i < baseHeaders.Count; i++)
        {
            baseIndex[baseHeaders[i]] = i;
        }

        var map = new int[sourceHeaders.Count];
        var sourceNames = new HashSet<string>(sourceHeaders, StringComparer.Ordinal);

        extra = [];
        for (var i = 0; i < sourceHeaders.Count; i++)
        {
            if (baseIndex.TryGetValue(sourceHeaders[i], out var target))
            {
                map[i] = target;
            }
            else
            {
                map[i] = -1;
                extra.Add(sourceHeaders[i]);
            }
        }

        missing = [.. baseHeaders.Where(name => !sourceNames.Contains(name))];
        return map;
    }
}
