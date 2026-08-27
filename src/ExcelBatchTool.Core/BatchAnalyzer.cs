namespace ExcelBatchTool.Core;

/// <summary>
/// 複数ファイルをバックグラウンドで並列解析する。
/// 1 ファイルの失敗は結果(<see cref="AnalysisStatus.Failed"/>)として記録され、
/// 他のファイルの解析は継続する。
/// </summary>
public sealed class BatchAnalyzer
{
    /// <summary>並列解析の最大同時実行数。UI の応答性を保つため 1 コア分は空ける。</summary>
    public int MaxDegreeOfParallelism { get; init; } = Math.Max(1, Environment.ProcessorCount - 1);

    /// <summary>
    /// 指定ファイルをすべて解析し、入力順で結果を返す。
    /// 各ファイルの完了ごとに <paramref name="progress"/> へ結果を通知する。
    /// </summary>
    public async Task<IReadOnlyList<WorkbookAnalysisResult>> AnalyzeAsync(
        IReadOnlyList<string> paths,
        IProgress<WorkbookAnalysisResult>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var results = new WorkbookAnalysisResult[paths.Count];

        var options = new ParallelOptions
        {
            MaxDegreeOfParallelism = MaxDegreeOfParallelism,
            CancellationToken = cancellationToken,
        };

        await Parallel.ForEachAsync(
            Enumerable.Range(0, paths.Count),
            options,
            (index, ct) =>
            {
                var result = WorkbookAnalyzer.Analyze(paths[index], ct);
                results[index] = result;
                progress?.Report(result);
                return ValueTask.CompletedTask;
            }).ConfigureAwait(false);

        return results;
    }
}
