using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Mapping;

/// <summary>
/// データ元の表と転記先の表をキーで突合し、既存の転記先行の指定列を更新する(Phase 2C2)。
///
/// 更新するのは「両側に存在するキー」の行だけ。データ元にしか無いキーで行を追加せず、
/// 転記先にしか無いキーの行も削除・空欄化しない。片側だけのキーは件数として知らせる。
///
/// このクラスがやるのは「どの行のどのセルへ、どの値を書くか」の解決まで。
/// 以降の安全確認・出力計画・書き込み・検証は Phase 2B / 2C1 と同じ
/// <see cref="MutationPlanBuilder"/> と <see cref="CellMutator"/> をそのまま使う。
/// </summary>
public sealed class TableUpdatePlanner
{
    private readonly int _maxKeyedRowsPerSheet;

    public TableUpdatePlanner()
        : this(TargetTableScanner.MaxKeyedRowsPerSheet)
    {
    }

    /// <summary>テスト用: 行数の上限を小さくして上限動作を確かめられるようにする。</summary>
    internal TableUpdatePlanner(int maxKeyedRowsPerSheet)
        => _maxKeyedRowsPerSheet = maxKeyedRowsPerSheet;

    public TableUpdatePreview CreatePreview(
        TableUpdateBatchRequest request,
        CancellationToken cancellationToken = default)
    {
        var issues = new List<MergeIssue>();
        var summary = new TableMatchSummary();

        if (Validate(request, issues) is not { } setup)
        {
            return Empty(issues, summary);
        }

        // 1. データ元のキー列だけを 1 パスで読む(値はまだ読まない)。
        var keyScan = ReadSourceKeys(request, setup, cancellationToken);
        if (!keyScan.IsSuccess)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, keyScan.Error!, setup.SourceFileName));
            return Empty(issues, summary);
        }

        if (keyScan.BlankKeyWithValueCount > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Warning,
                $"データ元の {keyScan.BlankKeyWithValueCount:N0} 行はキーが空欄のため使いません"
                    + "(他の項目には値があります)。",
                setup.SourceFileName));
        }

        // 2. 転記先を 1 ファイル 1 回だけ開き、表として読む(ヘッダー・キー・一致行・guard)。
        var scans = ScanTargets(request, setup, keyScan.Keys, _maxKeyedRowsPerSheet, cancellationToken);
        var locations = setup.Targets
            .Select(target => (target.FilePath, target.SheetName))
            .ToList();

        var mutationScans = scans.ToDictionary(
            entry => entry.Key, entry => entry.Value.ToMutationScan(), StringComparer.Ordinal);

        // 値が 1 つも解決できなかった場合でも、ファイル・シート単位の理由が出るよう先に報告する。
        MutationPlanBuilder.ReportScanIssues(locations, mutationScans, issues);

        // 3. 突合の結果を集計し、重複キーの Block / Warning を出す。
        var collected = CollectMatches(setup, scans, keyScan, issues);
        summary = collected.Summary;

        if (collected.MatchedKeys.Count == 0)
        {
            if (collected.HadUsableSheet
                && issues.All(issue => issue.Severity != MergeIssueSeverity.Block))
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    "一致する行がありません。キー列や表記(大文字小文字・空白)を確認してください。"));
            }

            return Empty(issues, summary);
        }

        // 4. 一致したキーの分だけ、データ元の値を 2 パス目で読む。
        var valueIndexes = setup.Mappings.Select(mapping => mapping.SourceColumnIndex).ToList();
        var match = setup.Kind == SourceFileKind.Csv
            ? CsvSourceReader.ReadRows(
                setup.SourceFilePath, setup.SourceColumns.Count, setup.SourceKeyColumnIndex,
                valueIndexes, collected.MatchedKeys, cancellationToken)
            : XlsxSourceReader.ReadRows(
                setup.SourceFilePath, request.SourceSheetName!, request.SourceHeaderRow,
                setup.SourceKeyColumnIndex + 1, [.. valueIndexes.Select(index => index + 1)],
                collected.MatchedKeys, cancellationToken);

        if (!match.IsSuccess)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, match.Error!, setup.SourceFileName));
            return Empty(issues, summary);
        }

        // 5. 一致した行 × 対応付けから、解決済みの変更を組み立てる。
        var mutations = BuildMutations(setup, scans, keyScan, match, issues);

        var dataSource = new MutationDataSourceInfo
        {
            FileName = setup.SourceFileName,
            Sha256 = setup.SourceSnapshot.Sha256,
            Type = setup.Kind == SourceFileKind.Csv ? "csv" : "xlsx",
            SheetName = setup.Kind == SourceFileKind.Xlsx ? request.SourceSheetName : null,
            HeaderRow = request.SourceHeaderRow,
            KeyColumn = request.SourceKeyColumn,
        };

        var check = new MutationDataSourceCheck(
            setup.SourceFilePath, setup.SourceFileName, setup.SourceSnapshot);

        if (mutations.Count == 0)
        {
            return new TableUpdatePreview
            {
                Mutation = new CellMutationPreview
                {
                    Targets = Array.Empty<CellMutationTargetPlan>(),
                    Files = Array.Empty<CellMutationFilePlan>(),
                    Issues = issues,
                    DataSourceCheck = check,
                },
                Summary = summary,
            };
        }

        var preview = MutationPlanBuilder.Build(
            mutations, mutationScans, request.OutputSuffix, issues,
            dataSource, check, reportScanIssues: false,
            targetTable: new MutationTargetTableInfo(request.TargetHeaderRow, request.TargetKeyColumn));

        return new TableUpdatePreview { Mutation = preview, Summary = summary };
    }

    /// <summary>指定内容を解釈する。問題があれば issues へ足して null を返す。</summary>
    private static UpdateSetup? Validate(TableUpdateBatchRequest request, List<MergeIssue> issues)
    {
        if (string.IsNullOrWhiteSpace(request.SourceFilePath) || !File.Exists(request.SourceFilePath))
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, "データ元のファイルを選んでください。"));
            return null;
        }

        if (SourceMappingPlanner.KindOf(request.SourceFilePath) is not { } kind)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, "データ元にできるのは .xlsx または .csv です。"));
            return null;
        }

        if (kind == SourceFileKind.Xlsx && string.IsNullOrEmpty(request.SourceSheetName))
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, "データ元のシートを選んでください。"));
            return null;
        }

        if (request.SourceHeaderRow < 1 || request.TargetHeaderRow < 1)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, "項目名の行は 1 以上で指定してください。"));
            return null;
        }

        if (request.Targets.Count == 0)
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, "転記先のシートが選択されていません。"));
            return null;
        }

        if (request.Mappings.Count == 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, "対応付けが指定されていません。行を追加してください。"));
            return null;
        }

        if (string.IsNullOrWhiteSpace(request.TargetKeyColumn))
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, "転記先のキーにする項目を選んでください。"));
            return null;
        }

        if (OutputNaming.ValidateSuffix(request.OutputSuffix) is { } suffixError)
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, suffixError));
            return null;
        }

        var failed = false;

        // データ元を転記先にもすると、読みながら書くことになる。
        var sourceFull = MutationPaths.Normalize(request.SourceFilePath);
        foreach (var target in request.Targets)
        {
            if (MutationPaths.Normalize(target.FilePath) == sourceFull)
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    "データ元のファイルを転記先にはできません。別のファイルを選んでください。",
                    Path.GetFileName(target.FilePath)));
                failed = true;
                break;
            }
        }

        if (MutationTargets.Validate(request.Targets, issues) is null)
        {
            failed = true;
        }

        var header = SourceMappingPlanner.ReadColumns(
            request.SourceFilePath, request.SourceSheetName, request.SourceHeaderRow);

        if (!header.IsSuccess)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, header.Error!, Path.GetFileName(request.SourceFilePath)));
            return null;
        }

        var sourceColumns = header.Columns.ToList();
        var keyIndex = sourceColumns.FindIndex(
            name => string.Equals(name, request.SourceKeyColumn, StringComparison.Ordinal));

        if (keyIndex < 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"データ元に項目「{request.SourceKeyColumn}」がありません。"
                    + "もう一度項目を読み込んでください。"));
            failed = true;
        }

        var mappings = ResolveMappings(request, sourceColumns, issues);
        if (mappings is null)
        {
            failed = true;
        }

        if (failed || mappings is null || keyIndex < 0)
        {
            return null;
        }

        return new UpdateSetup
        {
            SourceFilePath = request.SourceFilePath,
            SourceFileName = Path.GetFileName(request.SourceFilePath),
            SourceSnapshot = MutationSnapshot.Take(request.SourceFilePath),
            Kind = kind,
            SourceColumns = header.Columns,
            SourceKeyColumnIndex = keyIndex,
            Mappings = mappings,
            Targets = request.Targets,
        };
    }

    /// <summary>対応付けを解釈する(項目名の存在・重複・キー列の更新禁止)。</summary>
    private static List<TableColumnMapping>? ResolveMappings(
        TableUpdateBatchRequest request,
        List<string> sourceColumns,
        List<MergeIssue> issues)
    {
        var resolved = new List<TableColumnMapping>(request.Mappings.Count);
        var seenTargets = new HashSet<string>(StringComparer.Ordinal);
        var reportedDuplicates = new HashSet<string>(StringComparer.Ordinal);
        var failed = false;

        foreach (var mapping in request.Mappings)
        {
            var sourceIndex = sourceColumns.FindIndex(
                name => string.Equals(name, mapping.SourceColumn, StringComparison.Ordinal));

            if (sourceIndex < 0)
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"データ元に項目「{mapping.SourceColumn}」がありません。"
                        + "もう一度項目を読み込むか、対応付けを選び直してください。"));
                failed = true;
                continue;
            }

            var target = mapping.TargetColumn.Trim();
            if (target.Length == 0)
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block, "転記先の項目が選ばれていない対応付けがあります。"));
                failed = true;
                continue;
            }

            if (mapping.WriteKind == CellWriteKind.Blank)
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"「{target}」の種類に「空欄」は指定できません。"
                        + "セルを空欄にしたい場合は「4. セルをまとめて変更」を使ってください。"));
                failed = true;
                continue;
            }

            // 照合に使っている列を、実行中に書き換えない。
            if (string.Equals(target, request.TargetKeyColumn, StringComparison.Ordinal))
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"「{target}」はキーの列です。照合に使っている列は更新できません。"));
                failed = true;
                continue;
            }

            if (!seenTargets.Add(target))
            {
                if (reportedDuplicates.Add(target))
                {
                    issues.Add(new MergeIssue(
                        MergeIssueSeverity.Block,
                        $"転記先の項目「{target}」が対応付けの中で重複しています。"
                            + "同じ項目には 1 つのデータ元だけを指定してください。"));
                }

                failed = true;
                continue;
            }

            resolved.Add(new TableColumnMapping(
                mapping.SourceColumn, sourceIndex, target, mapping.WriteKind));
        }

        return failed ? null : resolved;
    }

    private static SourceKeyScan ReadSourceKeys(
        TableUpdateBatchRequest request, UpdateSetup setup, CancellationToken cancellationToken)
    {
        var valueIndexes = setup.Mappings.Select(mapping => mapping.SourceColumnIndex).ToList();

        return setup.Kind == SourceFileKind.Csv
            ? CsvSourceReader.ReadKeys(
                setup.SourceFilePath, setup.SourceColumns.Count, setup.SourceKeyColumnIndex,
                valueIndexes, cancellationToken)
            : XlsxSourceReader.ReadKeys(
                setup.SourceFilePath, request.SourceSheetName!, request.SourceHeaderRow,
                setup.SourceKeyColumnIndex + 1, [.. valueIndexes.Select(index => index + 1)],
                cancellationToken);
    }

    /// <summary>転記先を 1 ファイル 1 回だけ開いて表として読む。</summary>
    private static Dictionary<string, TargetTableWorkbookScan> ScanTargets(
        TableUpdateBatchRequest request,
        UpdateSetup setup,
        IReadOnlySet<string> sourceKeys,
        int maxKeyedRows,
        CancellationToken cancellationToken)
    {
        var scans = new Dictionary<string, TargetTableWorkbookScan>(StringComparer.Ordinal);

        foreach (var fileGroup in setup.Targets.GroupBy(
            target => target.FilePath, StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var sheetNames = fileGroup
                .Select(target => target.SheetName)
                .Distinct(StringComparer.Ordinal)
                .ToList();

            scans[MutationPaths.Normalize(fileGroup.Key)] = TargetTableScanner.Scan(
                fileGroup.Key, sheetNames, request.TargetHeaderRow, request.TargetKeyColumn,
                setup.Mappings, sourceKeys, cancellationToken, maxKeyedRows);
        }

        return scans;
    }

    /// <summary>突合の集計と、重複キーの Block / Warning。</summary>
    private static CollectedMatches CollectMatches(
        UpdateSetup setup,
        Dictionary<string, TargetTableWorkbookScan> scans,
        SourceKeyScan keyScan,
        List<MergeIssue> issues)
    {
        var matchedKeys = new HashSet<string>(StringComparer.Ordinal);
        var targetKeyUniverse = new HashSet<string>(StringComparer.Ordinal);
        var targetDuplicates = new HashSet<string>(StringComparer.Ordinal);
        var reportedSourceDuplicates = new HashSet<string>(StringComparer.Ordinal);
        var targetKeyedRows = 0;
        var targetBlankRows = 0;
        var hadUsableSheet = false;
        var unusedTargetDuplicates = 0;

        foreach (var target in setup.Targets)
        {
            var fileName = Path.GetFileName(target.FilePath);
            var scan = scans.GetValueOrDefault(MutationPaths.Normalize(target.FilePath));

            if (scan is null || scan.BlockReasons.Count > 0)
            {
                continue;
            }

            var sheetScan = scan.Sheets.GetValueOrDefault(target.SheetName);
            if (sheetScan is null || sheetScan.BlockReason is not null)
            {
                continue;
            }

            hadUsableSheet = true;
            targetKeyedRows += sheetScan.KeyedRowCount;
            targetBlankRows += sheetScan.BlankRowCount + sheetScan.BlankKeyWithValueCount;
            targetKeyUniverse.UnionWith(sheetScan.TargetOnlyKeys);
            unusedTargetDuplicates += sheetScan.UnusedDuplicateKeyCount;

            if (sheetScan.BlankKeyWithValueCount > 0)
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Warning,
                    $"キーが空欄の {sheetScan.BlankKeyWithValueCount:N0} 行は更新しません"
                        + "(他の項目には値があります)。",
                    fileName,
                    target.SheetName));
            }

            foreach (var duplicate in sheetScan.UsedDuplicateKeys)
            {
                targetKeyUniverse.Add(duplicate);
                targetDuplicates.Add(duplicate);
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"キー「{duplicate}」がこのシートに 2 行以上あります。どの行を更新するか判断できません。",
                    fileName,
                    target.SheetName));
            }

            foreach (var match in sheetScan.Matches)
            {
                targetKeyUniverse.Add(match.Key);

                if (keyScan.DuplicateKeys.Contains(match.Key))
                {
                    // データ元側の重複。どの行の値を使うか決められない。
                    if (reportedSourceDuplicates.Add(match.Key))
                    {
                        issues.Add(new MergeIssue(
                            MergeIssueSeverity.Block,
                            $"キー「{match.Key}」がデータ元に 2 件以上あります。どの行を使うか判断できません。",
                            setup.SourceFileName));
                    }

                    continue;
                }

                matchedKeys.Add(match.Key);
            }
        }

        var unusedSourceDuplicates = keyScan.DuplicateKeys
            .Count(key => !reportedSourceDuplicates.Contains(key));

        if (unusedSourceDuplicates > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Warning,
                $"データ元の {unusedSourceDuplicates:N0} 件のキーは重複していますが、今回は使いません。",
                setup.SourceFileName));
        }

        if (unusedTargetDuplicates > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Warning,
                $"転記先の {unusedTargetDuplicates:N0} 件のキーは重複していますが、"
                    + "データ元に無いため今回は更新しません。"));
        }

        var sourceOnly = keyScan.Keys.Count(key => !targetKeyUniverse.Contains(key));
        var targetOnly = targetKeyUniverse.Count(key => !keyScan.Keys.Contains(key));

        if (sourceOnly > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Warning,
                $"データ元の {sourceOnly:N0} 件のキーは転記先に無いため使いません(行の追加はしません)。",
                setup.SourceFileName));
        }

        if (targetOnly > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Warning,
                $"転記先の {targetOnly:N0} 件のキーはデータ元に無いため、そのまま残します。"));
        }

        var summary = new TableMatchSummary
        {
            SourceKeyedRowCount = keyScan.KeyedRowCount,
            TargetKeyedRowCount = targetKeyedRows,
            MatchedKeyCount = matchedKeys.Count,
            SourceOnlyKeyCount = sourceOnly,
            TargetOnlyKeyCount = targetOnly,
            DuplicateKeyCount = keyScan.DuplicateKeys.Count + targetDuplicates.Count,
            BlankKeyRowCount = keyScan.BlankRowCount + keyScan.BlankKeyWithValueCount + targetBlankRows,
        };

        return new CollectedMatches(matchedKeys, summary, hadUsableSheet);
    }

    /// <summary>一致した行 × 対応付けから、解決済みの変更を組み立てる。</summary>
    private static List<ResolvedCellMutation> BuildMutations(
        UpdateSetup setup,
        Dictionary<string, TargetTableWorkbookScan> scans,
        SourceKeyScan keyScan,
        SourceMatchResult match,
        List<MergeIssue> issues)
    {
        var mutations = new List<ResolvedCellMutation>();
        var reportedValues = new HashSet<string>(StringComparer.Ordinal);

        foreach (var target in setup.Targets)
        {
            var fileName = Path.GetFileName(target.FilePath);
            var scan = scans.GetValueOrDefault(MutationPaths.Normalize(target.FilePath));

            if (scan is null || scan.BlockReasons.Count > 0)
            {
                continue;
            }

            var sheetScan = scan.Sheets.GetValueOrDefault(target.SheetName);
            if (sheetScan is null || sheetScan.BlockReason is not null)
            {
                continue;
            }

            foreach (var tableMatch in sheetScan.Matches)
            {
                if (keyScan.DuplicateKeys.Contains(tableMatch.Key))
                {
                    continue; // データ元側の重複として報告済み。
                }

                if (!match.RowsByKey.TryGetValue(tableMatch.Key, out var row))
                {
                    issues.Add(new MergeIssue(
                        MergeIssueSeverity.Block,
                        $"キー「{tableMatch.Key}」のデータを読み直せませんでした。"
                            + "データ元が変更された可能性があります。もう一度プレビューしてください。",
                        setup.SourceFileName));
                    continue;
                }

                for (var index = 0; index < setup.Mappings.Count; index++)
                {
                    var mapping = setup.Mappings[index];
                    var value = row.Values[index];

                    if (!SourceValueConversion.TryConvert(
                        value, mapping.WriteKind, setup.Kind, out var newValue, out var reason))
                    {
                        // 同じキー × 項目の理由はシートが違っても同じなので 1 回だけ報告する。
                        if (reportedValues.Add($"{tableMatch.Key}\n{mapping.SourceColumn}"))
                        {
                            issues.Add(new MergeIssue(
                                MergeIssueSeverity.Block,
                                $"キー「{tableMatch.Key}」の「{mapping.SourceColumn}」{reason}",
                                setup.SourceFileName));
                        }

                        continue;
                    }

                    var column = sheetScan.MappedColumns[index];
                    var reference =
                        $"{CellRangeParser.ColumnIndexToLetters(column)}{tableMatch.RowNumber}";

                    mutations.Add(new ResolvedCellMutation
                    {
                        FilePath = target.FilePath,
                        SheetName = target.SheetName,
                        Address = new TargetCellAddress(reference, column, tableMatch.RowNumber),
                        Value = newValue,
                        Provenance = new MutationProvenance(
                            mapping.SourceColumn,
                            tableMatch.Key,
                            row.RowNumber,
                            mapping.TargetColumn,
                            tableMatch.RowNumber),
                    });
                }
            }
        }

        return mutations;
    }

    private static TableUpdatePreview Empty(List<MergeIssue> issues, TableMatchSummary summary) => new()
    {
        Mutation = new CellMutationPreview
        {
            Targets = Array.Empty<CellMutationTargetPlan>(),
            Files = Array.Empty<CellMutationFilePlan>(),
            Issues = issues,
        },
        Summary = summary,
    };

    /// <summary>解釈済みの指定内容。</summary>
    private sealed record UpdateSetup
    {
        public required string SourceFilePath { get; init; }

        public required string SourceFileName { get; init; }

        public required SourceSnapshot SourceSnapshot { get; init; }

        public required SourceFileKind Kind { get; init; }

        public required IReadOnlyList<string> SourceColumns { get; init; }

        public required int SourceKeyColumnIndex { get; init; }

        public required IReadOnlyList<TableColumnMapping> Mappings { get; init; }

        public required IReadOnlyList<CellMutationTarget> Targets { get; init; }
    }

    /// <summary>突合の結果。</summary>
    private readonly record struct CollectedMatches(
        HashSet<string> MatchedKeys, TableMatchSummary Summary, bool HadUsableSheet);
}
