using System.Globalization;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.Mapping;

/// <summary>
/// データ元の表からキーで 1 行を特定し、転記先の決まったセルへ入れる(Phase 2C1)。
///
/// このクラスがやるのは「新しい値をどこから得るか」だけ。
/// 転記先の安全確認・No-op 判定・出力計画・書き込み・検証は、
/// Phase 2B と同じ <see cref="MutationPlanBuilder"/> と <see cref="CellMutator"/> をそのまま使う。
/// </summary>
public sealed class SourceMappingPlanner
{
    /// <summary>データ元の項目名を読む(利用者が対応付けを作る前の下準備)。</summary>
    public static IReadOnlyList<string> ReadSourceSheetNames(string filePath)
        => XlsxSourceReader.ReadSheetNames(filePath);

    /// <summary>データ元の項目名(ヘッダー)を読む。</summary>
    internal static SourceHeaderResult ReadHeader(
        string filePath, SourceFileKind kind, string? sheetName, int headerRow)
        => kind == SourceFileKind.Csv
            ? CsvSourceReader.ReadHeader(filePath)
            : XlsxSourceReader.ReadHeader(filePath, sheetName ?? string.Empty, headerRow);

    /// <summary>
    /// データ元の項目名を読む(画面から使う入口)。
    /// 対応付けを作る前に、利用者へ項目名の一覧を見せるために呼ぶ。
    /// </summary>
    public static SourceColumnsResult ReadColumns(string filePath, string? sheetName, int headerRow)
    {
        if (KindOf(filePath) is not { } kind)
        {
            return new SourceColumnsResult { Error = "データ元にできるのは .xlsx または .csv です。" };
        }

        if (headerRow < 1)
        {
            return new SourceColumnsResult { Error = "項目名の行は 1 以上で指定してください。" };
        }

        var header = ReadHeader(filePath, kind, sheetName, headerRow);
        return new SourceColumnsResult
        {
            Columns = header.Columns,
            EncodingName = header.EncodingName,
            Error = header.Error,
        };
    }

    /// <summary>拡張子から種類を決める。扱えない拡張子なら null。</summary>
    public static SourceFileKind? KindOf(string filePath) => Path.GetExtension(filePath).ToLowerInvariant() switch
    {
        ".xlsx" => SourceFileKind.Xlsx,
        ".csv" => SourceFileKind.Csv,
        _ => null,
    };

    public CellMutationPreview CreatePreview(
        SourceMappingBatchRequest request,
        CancellationToken cancellationToken = default)
    {
        var issues = new List<MergeIssue>();

        if (Validate(request, issues) is not { } setup)
        {
            return MutationPreviewFactory.Empty(issues);
        }

        // 1. 転記先を 1 ファイル 1 回だけ開き、対象セルの安全確認とキーの読み取りをまとめて行う。
        var probes = BuildProbeMutations(setup);
        var scans = MutationPlanBuilder.ScanTargets(probes, setup.KeyCell, cancellationToken);

        // 値が 1 つも解決できなかった場合でも理由が出るよう、ここで先に報告しておく。
        MutationPlanBuilder.ReportScanIssues(probes, scans, issues);

        // 2. 転記先が必要とするキーを集めてから、データ元を 1 回だけ読む。
        var keys = CollectKeys(setup, scans, issues);
        var match = ReadSource(request, setup, keys.Required, cancellationToken);
        if (match is null || !match.IsSuccess)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, match?.Error ?? "データ元を読み取れません。", setup.SourceFileName));
            return MutationPreviewFactory.Empty(issues);
        }

        ReportSourceNotes(match, issues, setup.SourceFileName);

        // 3. キーごとの行から、解決済みの変更を組み立てる。
        var mutations = BuildMutations(setup, keys, match, issues);

        var dataSource = new MutationDataSourceInfo
        {
            FileName = setup.SourceFileName,
            Sha256 = setup.SourceSnapshot.Sha256,
            Type = setup.Kind == SourceFileKind.Csv ? "csv" : "xlsx",
            SheetName = setup.Kind == SourceFileKind.Xlsx ? request.SourceSheetName : null,
            HeaderRow = request.HeaderRow,
            KeyColumn = request.KeyColumn,
        };

        var check = new MutationDataSourceCheck(
            setup.SourceFilePath, setup.SourceFileName, setup.SourceSnapshot);

        if (mutations.Count == 0)
        {
            // 転記できる組み合わせが 1 件も無い(理由は issues に入っている)。
            return new CellMutationPreview
            {
                Targets = Array.Empty<CellMutationTargetPlan>(),
                Files = Array.Empty<CellMutationFilePlan>(),
                Issues = issues,
                DataSourceCheck = check,
            };
        }

        return MutationPlanBuilder.Build(
            mutations, scans, request.OutputSuffix, issues, dataSource, check, reportScanIssues: false);
    }

    /// <summary>指定内容を解釈する。問題があれば issues へ足して null を返す。</summary>
    private static MappingSetup? Validate(SourceMappingBatchRequest request, List<MergeIssue> issues)
    {
        var failed = false;

        if (string.IsNullOrWhiteSpace(request.SourceFilePath) || !File.Exists(request.SourceFilePath))
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, "データ元のファイルを選んでください。"));
            return null;
        }

        if (KindOf(request.SourceFilePath) is not { } kind)
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

        if (request.HeaderRow < 1)
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

        if (OutputNaming.ValidateSuffix(request.OutputSuffix) is { } suffixError)
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, suffixError));
            return null;
        }

        if (!TargetCellAddress.TryParse(request.TargetKeyCell, out var keyCell, out var keyCellError))
        {
            issues.Add(new MergeIssue(MergeIssueSeverity.Block, $"キーのセル: {keyCellError}"));
            return null;
        }

        var sourceFull = MutationPaths.Normalize(request.SourceFilePath);

        // データ元を転記先にもすると、読みながら書くことになる。
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

        var header = ReadHeader(
            request.SourceFilePath, kind, request.SourceSheetName, request.HeaderRow);

        if (!header.IsSuccess)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block, header.Error!, Path.GetFileName(request.SourceFilePath)));
            return null;
        }

        var keyIndex = header.Columns.ToList().FindIndex(
            name => string.Equals(name, request.KeyColumn, StringComparison.Ordinal));

        if (keyIndex < 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Block,
                $"データ元に項目「{request.KeyColumn}」がありません。もう一度項目を読み込んでください。"));
            failed = true;
        }

        var mappings = ResolveMappings(request.Mappings, header.Columns, keyCell, issues);
        if (mappings is null)
        {
            failed = true;
        }

        if (failed || mappings is null || keyIndex < 0)
        {
            return null;
        }

        return new MappingSetup
        {
            SourceFilePath = request.SourceFilePath,
            SourceFileName = Path.GetFileName(request.SourceFilePath),
            SourceSnapshot = MutationSnapshot.Take(request.SourceFilePath),
            Kind = kind,
            Columns = header.Columns,
            KeyColumnIndex = keyIndex,
            KeyCell = keyCell,
            Mappings = mappings,
            Targets = request.Targets,
        };
    }

    /// <summary>対応付けを解釈する(項目名の存在・転記先セル・重複)。</summary>
    private static List<ResolvedMapping>? ResolveMappings(
        IReadOnlyList<SourceMappingRequest> requests,
        IReadOnlyList<string> columns,
        TargetCellAddress keyCell,
        List<MergeIssue> issues)
    {
        var resolved = new List<ResolvedMapping>(requests.Count);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var reportedDuplicates = new HashSet<string>(StringComparer.Ordinal);
        var failed = false;

        foreach (var request in requests)
        {
            var columnIndex = columns.ToList().FindIndex(
                name => string.Equals(name, request.SourceColumn, StringComparison.Ordinal));

            if (columnIndex < 0)
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"データ元に項目「{request.SourceColumn}」がありません。"
                        + "もう一度項目を読み込むか、対応付けを選び直してください。"));
                failed = true;
                continue;
            }

            if (!TargetCellAddress.TryParse(request.TargetCell, out var address, out var addressError))
            {
                issues.Add(new MergeIssue(MergeIssueSeverity.Block, addressError!));
                failed = true;
                continue;
            }

            if (request.WriteKind == CellWriteKind.Blank)
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"{address.Reference} の種類に「空欄」は指定できません。"
                        + "セルを空欄にしたい場合は「4. セルをまとめて変更」を使ってください。"));
                failed = true;
                continue;
            }

            // 照合に使っているセルを、実行中に書き換えない。
            if (string.Equals(address.Reference, keyCell.Reference, StringComparison.Ordinal))
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"{address.Reference} はキーのセルです。照合に使っているセルは転記先にできません。"));
                failed = true;
                continue;
            }

            // $D$5 と D5 は同じセルとして扱う。
            if (!seen.Add(address.Reference))
            {
                if (reportedDuplicates.Add(address.Reference))
                {
                    issues.Add(new MergeIssue(
                        MergeIssueSeverity.Block,
                        $"転記先のセル「{address.Reference}」が対応付けの中で重複しています。"
                            + "同じセルには 1 つの項目だけを指定してください。"));
                }

                failed = true;
                continue;
            }

            resolved.Add(new ResolvedMapping(request.SourceColumn, columnIndex, address, request.WriteKind));
        }

        return failed ? null : resolved;
    }

    /// <summary>転記先を走査するための、値がまだ決まっていない仮の変更一覧。</summary>
    private static List<ResolvedCellMutation> BuildProbeMutations(MappingSetup setup)
    {
        var probes = new List<ResolvedCellMutation>(setup.Targets.Count * setup.Mappings.Count);
        foreach (var target in setup.Targets)
        {
            foreach (var mapping in setup.Mappings)
            {
                probes.Add(new ResolvedCellMutation
                {
                    FilePath = target.FilePath,
                    SheetName = target.SheetName,
                    Address = mapping.Address,
                    Value = mapping.WriteKind == CellWriteKind.Number
                        ? NewCellValue.OfNumber(0)
                        : NewCellValue.OfText(string.Empty),
                });
            }
        }

        return probes;
    }

    /// <summary>転記先のキーセルを読み、必要なキーの集合を作る。</summary>
    private static CollectedKeys CollectKeys(
        MappingSetup setup,
        IReadOnlyDictionary<string, WorkbookMutationScan> scans,
        List<MergeIssue> issues)
    {
        var required = new HashSet<string>(StringComparer.Ordinal);
        var keyByTarget = new Dictionary<(string File, string Sheet), string>();

        foreach (var target in setup.Targets)
        {
            var fileName = Path.GetFileName(target.FilePath);
            var scan = scans.GetValueOrDefault(MutationPaths.Normalize(target.FilePath));

            // ファイル・シート単位の問題は共通の計画側で報告されるので、ここでは触れない。
            if (scan is null || scan.BlockReasons.Count > 0)
            {
                continue;
            }

            var sheetScan = scan.Sheets.GetValueOrDefault(target.SheetName);
            if (sheetScan is null || sheetScan.BlockReason is not null)
            {
                continue;
            }

            if (sheetScan.KeyCell is not { } keyCell)
            {
                continue;
            }

            if (keyCell.BlockReason is { } reason)
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block, reason, fileName, target.SheetName));
                continue;
            }

            var key = keyCell.Key!;
            required.Add(key);
            keyByTarget[(target.FilePath, target.SheetName)] = key;
        }

        return new CollectedKeys(required, keyByTarget);
    }

    private static SourceMatchResult? ReadSource(
        SourceMappingBatchRequest request,
        MappingSetup setup,
        IReadOnlySet<string> requiredKeys,
        CancellationToken cancellationToken)
    {
        var valueIndexes = setup.Mappings.Select(mapping => mapping.ColumnIndex).ToList();

        return setup.Kind == SourceFileKind.Csv
            ? CsvSourceReader.ReadRows(
                setup.SourceFilePath, setup.Columns.Count, setup.KeyColumnIndex,
                valueIndexes, requiredKeys, cancellationToken)
            : XlsxSourceReader.ReadRows(
                setup.SourceFilePath, request.SourceSheetName!, request.HeaderRow,
                setup.KeyColumnIndex + 1, [.. valueIndexes.Select(index => index + 1)],
                requiredKeys, cancellationToken);
    }

    /// <summary>データ元側の「知らせておくこと」を issues へ足す。</summary>
    private static void ReportSourceNotes(
        SourceMatchResult match, List<MergeIssue> issues, string sourceFileName)
    {
        if (match.UnusedRowCount > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Warning,
                $"データ元の {match.UnusedRowCount:N0} 行は、今回の転記先と一致しないため使いません。",
                sourceFileName));
        }

        if (match.BlankKeyWithValueCount > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Warning,
                $"データ元の {match.BlankKeyWithValueCount:N0} 行はキーが空欄のため使いません"
                    + "(他の項目には値があります)。",
                sourceFileName));
        }
    }

    /// <summary>キーごとに一致した行から、解決済みの変更を組み立てる。</summary>
    private static List<ResolvedCellMutation> BuildMutations(
        MappingSetup setup,
        CollectedKeys keys,
        SourceMatchResult match,
        List<MergeIssue> issues)
    {
        var mutations = new List<ResolvedCellMutation>();
        var reportedKeys = new HashSet<string>(StringComparer.Ordinal);

        foreach (var target in setup.Targets)
        {
            var fileName = Path.GetFileName(target.FilePath);

            if (!keys.KeyByTarget.TryGetValue((target.FilePath, target.SheetName), out var key))
            {
                continue; // キーを読めなかった(理由は報告済み)。
            }

            if (match.DuplicateKeys.Contains(key))
            {
                if (reportedKeys.Add($"dup:{key}"))
                {
                    issues.Add(new MergeIssue(
                        MergeIssueSeverity.Block,
                        $"キー「{key}」がデータ元に 2 件以上あります。どの行を使うか判断できません。",
                        setup.SourceFileName));
                }

                continue;
            }

            if (!match.RowsByKey.TryGetValue(key, out var row))
            {
                issues.Add(new MergeIssue(
                    MergeIssueSeverity.Block,
                    $"キー「{key}」に一致するデータがデータ元にありません。",
                    fileName,
                    target.SheetName));
                continue;
            }

            for (var index = 0; index < setup.Mappings.Count; index++)
            {
                var mapping = setup.Mappings[index];
                var value = row.Values[index];

                if (!SourceValueConversion.TryConvert(
                    value, mapping.WriteKind, setup.Kind, out var newValue, out var reason))
                {
                    issues.Add(new MergeIssue(
                        MergeIssueSeverity.Block,
                        $"キー「{key}」の「{mapping.SourceColumn}」{reason}",
                        fileName,
                        target.SheetName));
                    continue;
                }

                mutations.Add(new ResolvedCellMutation
                {
                    FilePath = target.FilePath,
                    SheetName = target.SheetName,
                    Address = mapping.Address,
                    Value = newValue,
                    Provenance = new MutationProvenance(mapping.SourceColumn, key, row.RowNumber),
                });
            }
        }

        return mutations;
    }

    /// <summary>解釈済みの指定内容。</summary>
    private sealed record MappingSetup
    {
        public required string SourceFilePath { get; init; }

        public required string SourceFileName { get; init; }

        public required SourceSnapshot SourceSnapshot { get; init; }

        public required SourceFileKind Kind { get; init; }

        public required IReadOnlyList<string> Columns { get; init; }

        public required int KeyColumnIndex { get; init; }

        public required TargetCellAddress KeyCell { get; init; }

        public required IReadOnlyList<ResolvedMapping> Mappings { get; init; }

        public required IReadOnlyList<CellMutationTarget> Targets { get; init; }
    }

    /// <summary>解釈済みの対応付け 1 件。</summary>
    private readonly record struct ResolvedMapping(
        string SourceColumn, int ColumnIndex, TargetCellAddress Address, CellWriteKind WriteKind);

    /// <summary>転記先から集めたキー。</summary>
    private readonly record struct CollectedKeys(
        HashSet<string> Required, Dictionary<(string File, string Sheet), string> KeyByTarget);
}
