using System.Globalization;
using ExcelBatchTool.Core.Mapping;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.Core.CsvTransform;

/// <summary>
/// 「データ元の表を、指定した列構成の新しい CSV にする」計画を立てる。
///
/// データ元は読み取りのみ。プレビューでは 1 回だけ流し読みして、
/// 行数と読み取れないセルを確かめ、先頭の数行だけを見本として持つ。
/// 全行を画面やメモリへ展開しない。
/// </summary>
public sealed class CsvTransformPlanner
{
    /// <summary>データ元の項目名を読む(画面の候補づくり用)。</summary>
    public static SourceColumnsResult ReadColumns(string filePath, string? sheetName, int headerRow)
        => SourceMappingPlanner.ReadColumns(filePath, sheetName, headerRow);

    public static IReadOnlyList<string> ReadSheetNames(string filePath)
        => SourceMappingPlanner.ReadSourceSheetNames(filePath);

    public static SourceFileKind? KindOf(string filePath) => SourceMappingPlanner.KindOf(filePath);

    /// <summary>現在の指定でプレビューを作る。</summary>
    public CsvTransformPreview CreatePreview(
        CsvTransformRequest request, CancellationToken cancellationToken = default)
    {
        var issues = new List<MergeIssue>();

        if (KindOf(request.SourceFilePath) is not { } kind)
        {
            return Blocked(issues, "データ元には .xlsx または .csv のファイルを選んでください。");
        }

        if (!File.Exists(request.SourceFilePath))
        {
            return Blocked(issues, "データ元のファイルが見つかりません。");
        }

        if (request.HeaderRow < 1)
        {
            return Blocked(issues, "項目名の行は 1 以上の数字で指定してください。");
        }

        var header = ReadColumns(request.SourceFilePath, request.SourceSheetName, request.HeaderRow);
        if (!header.IsSuccess)
        {
            return Blocked(issues, header.Error!);
        }

        var columns = ValidateColumns(request.Columns, header.Columns, issues);
        var sourceFileName = Path.GetFileName(request.SourceFilePath);

        if (columns is null)
        {
            return new CsvTransformPreview
            {
                Columns = [],
                SourceColumns = header.Columns,
                Issues = issues,
                SourceFileName = sourceFileName,
                SourceEncodingName = header.EncodingName,
            };
        }

        var (outputPath, auditPath, outputFileName) = ResolveOutput(
            request.SourceFilePath, request.OutputSuffix, issues);

        SourceSnapshot snapshot;
        try
        {
            snapshot = MutationSnapshot.Take(request.SourceFilePath);
        }
        catch (Exception ex)
        {
            return Blocked(issues, $"データ元のファイルを読み取れません: {ex.Message}");
        }

        var scan = Scan(request, kind, header.Columns, columns, issues, cancellationToken);

        return new CsvTransformPreview
        {
            Columns = columns,
            SourceColumns = header.Columns,
            SampleRows = scan.Sample,
            Issues = issues,
            SourceRowCount = scan.RowCount,
            OutputRowCount = scan.RowCount,
            BlankRowCount = scan.BlankRowCount,
            SourceFileName = sourceFileName,
            SourceEncodingName = header.EncodingName,
            OutputFileName = outputFileName,
            OutputPath = outputPath,
            AuditPath = auditPath,
            Snapshot = snapshot,
            Request = request,
        };
    }

    /// <summary>出力する列の指定を確かめる。1 つでも問題があれば列を作らない。</summary>
    private static IReadOnlyList<CsvOutputColumnPlan>? ValidateColumns(
        IReadOnlyList<CsvOutputColumnRequest> requested,
        IReadOnlyList<string> sourceColumns,
        List<MergeIssue> issues)
    {
        if (requested.Count == 0)
        {
            issues.Add(Block("出力する項目がありません。作る CSV の項目を 1 つ以上追加してください。"));
            return null;
        }

        var plans = new List<CsvOutputColumnPlan>(requested.Count);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var ok = true;

        foreach (var column in requested)
        {
            var name = (column.OutputName ?? string.Empty).Trim();

            if (name.Length == 0)
            {
                issues.Add(Block("出力する項目名が空の行があります。名前を入れてください。"));
                ok = false;
                continue;
            }

            if (name.Any(char.IsControl))
            {
                issues.Add(Block($"出力する項目名「{name}」に改行や特殊な文字は使えません。"));
                ok = false;
                continue;
            }

            if (!names.Add(name))
            {
                // 大文字小文字だけの違いも、取り違えを避けるため重複として止める。
                issues.Add(Block(
                    $"出力する項目名「{name}」が重複しています。CSV の項目名は 1 つずつ分けてください。"));
                ok = false;
                continue;
            }

            if (column.ValueSourceKind == CsvValueSourceKind.SourceColumn)
            {
                var source = (column.SourceColumn ?? string.Empty).Trim();
                if (source.Length == 0)
                {
                    issues.Add(Block($"「{name}」に入れるデータ元の項目が選ばれていません。"));
                    ok = false;
                    continue;
                }

                if (!sourceColumns.Contains(source, StringComparer.Ordinal))
                {
                    issues.Add(Block(
                        $"データ元に項目「{source}」がありません({name} 用)。項目を選び直してください。"));
                    ok = false;
                    continue;
                }

                plans.Add(new CsvOutputColumnPlan
                {
                    OutputName = name,
                    ValueSourceKind = CsvValueSourceKind.SourceColumn,
                    SourceColumn = source,
                });
                continue;
            }

            plans.Add(new CsvOutputColumnPlan
            {
                OutputName = name,
                ValueSourceKind = column.ValueSourceKind,
                FixedValue = column.ValueSourceKind == CsvValueSourceKind.FixedText
                    ? column.FixedValue ?? string.Empty
                    : null,
            });
        }

        return ok ? plans : null;
    }

    /// <summary>出力先を決める。上書きはしない。</summary>
    private static (string OutputPath, string AuditPath, string OutputFileName) ResolveOutput(
        string sourceFilePath, string suffix, List<MergeIssue> issues)
    {
        var directory = Path.GetDirectoryName(Path.GetFullPath(sourceFilePath)) ?? string.Empty;
        var baseName = Path.GetFileNameWithoutExtension(sourceFilePath);
        var outputFileName = baseName + suffix + ".csv";
        var outputPath = Path.Combine(directory, outputFileName);
        var auditPath = outputPath + ".audit.json";

        if (suffix.Length == 0)
        {
            issues.Add(Block("出力名が空です。元のファイルを上書きしないよう、付ける文字を入れてください。"));
        }
        else if (outputFileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        {
            issues.Add(Block($"出力名「{suffix}」にファイル名として使えない文字が含まれています。"));
        }
        else if (File.Exists(outputPath))
        {
            issues.Add(Block($"「{outputFileName}」はすでにあります。既存のファイルは上書きしません。"));
        }
        else if (File.Exists(auditPath))
        {
            issues.Add(Block(
                $"「{Path.GetFileName(auditPath)}」はすでにあります。既存のファイルは上書きしません。"));
        }

        return (outputPath, auditPath, outputFileName);
    }

    /// <summary>
    /// データ元を 1 回だけ流し読みして、行数・読み取れないセル・見本を集める。
    /// 実行時に初めて失敗しないよう、ここですべての行を確かめる。
    /// </summary>
    private static ScanResult Scan(
        CsvTransformRequest request,
        SourceFileKind kind,
        IReadOnlyList<string> sourceColumns,
        IReadOnlyList<CsvOutputColumnPlan> columns,
        List<MergeIssue> issues,
        CancellationToken cancellationToken)
    {
        var indexes = ColumnIndexes(columns, sourceColumns);
        var sample = new List<CsvSampleRow>(CsvTransformDefaults.SampleRowCount);
        var rowCount = 0;
        var blankRowCount = 0;
        string? unsupported = null;

        bool OnValues(int rowNumber, Func<int, (bool Ok, string Text, string? Reason)> read, bool allBlank)
        {
            if (allBlank)
            {
                blankRowCount++;
                return true;
            }

            var values = new string[columns.Count];
            for (var index = 0; index < columns.Count; index++)
            {
                var column = columns[index];
                if (column.ValueSourceKind != CsvValueSourceKind.SourceColumn)
                {
                    values[index] = column.ValueSourceKind == CsvValueSourceKind.FixedText
                        ? column.FixedValue ?? string.Empty
                        : string.Empty;
                    continue;
                }

                var (ok, text, reason) = read(indexes[index]);
                if (!ok)
                {
                    unsupported = $"データ元の {rowNumber} 行目「{column.SourceColumn}」{reason}";
                    return false;
                }

                values[index] = text;
            }

            rowCount++;
            if (sample.Count < CsvTransformDefaults.SampleRowCount)
            {
                sample.Add(new CsvSampleRow(rowNumber, values));
            }

            return true;
        }

        string? error;
        if (kind == SourceFileKind.Csv)
        {
            error = CsvSourceReader.ReadRecords(
                request.SourceFilePath,
                sourceColumns.Count,
                (rowNumber, fields) => OnValues(
                    rowNumber,
                    index => (true, fields[index], null),
                    fields.All(field => field.Length == 0)),
                cancellationToken);
        }
        else
        {
            error = XlsxSourceReader.ReadRecords(
                request.SourceFilePath,
                request.SourceSheetName ?? string.Empty,
                request.HeaderRow,
                sourceColumns.Count,
                (rowNumber, values) => OnValues(
                    rowNumber,
                    index => Render(values[index]),
                    values.All(value => value.IsBlank)),
                cancellationToken);
        }

        if (error is not null)
        {
            issues.Add(Block(error));
        }
        else if (unsupported is not null)
        {
            issues.Add(Block(unsupported));
        }
        else if (rowCount == 0)
        {
            issues.Add(Block("データ元に行がありません。作る CSV の中身が空になるため実行しません。"));
        }

        if (blankRowCount > 0)
        {
            issues.Add(new MergeIssue(
                MergeIssueSeverity.Warning,
                $"すべての項目が空欄の {blankRowCount:N0} 行は読み飛ばします。"));
        }

        return new ScanResult(rowCount, blankRowCount, sample);
    }

    /// <summary>出力する列ごとに、データ元の何列目から取るかを決める。</summary>
    private static int[] ColumnIndexes(
        IReadOnlyList<CsvOutputColumnPlan> columns, IReadOnlyList<string> sourceColumns)
    {
        var indexes = new int[columns.Count];
        for (var index = 0; index < columns.Count; index++)
        {
            indexes[index] = columns[index].ValueSourceKind == CsvValueSourceKind.SourceColumn
                ? sourceColumns.ToList().IndexOf(columns[index].SourceColumn!)
                : -1;
        }

        return indexes;
    }

    /// <summary>.xlsx のセル 1 つを CSV の項目にする。意味が推測になるものは通さない。</summary>
    internal static (bool Ok, string Text, string? Reason) Render(SourceValue value) => value.Kind switch
    {
        SourceValueKind.Blank => (true, string.Empty, null),
        SourceValueKind.Text => (true, value.Text ?? string.Empty, null),
        SourceValueKind.Number => (true, value.Number.ToString(CultureInfo.InvariantCulture), null),
        _ => (false, string.Empty, value.Reason is { } reason ? $"は{reason}。" : "を読み取れません。"),
    };

    internal static CsvTransformPreview Blocked(List<MergeIssue> issues, string message)
    {
        issues.Add(Block(message));
        return new CsvTransformPreview { Columns = [], Issues = issues };
    }

    private static MergeIssue Block(string message) => new(MergeIssueSeverity.Block, message);

    private sealed record ScanResult(int RowCount, int BlankRowCount, IReadOnlyList<CsvSampleRow> Sample);
}
