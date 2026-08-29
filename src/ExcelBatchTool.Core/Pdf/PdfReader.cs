using System.Globalization;
using ExcelBatchTool.Core.CsvTransform;
using ExcelBatchTool.Core.Mutation;
using Microsoft.VisualBasic.FileIO;

namespace ExcelBatchTool.Core.Pdf;

/// <summary>
/// プレビューどおりのファイルを作る。
///
/// 1. 実行直前に PDF が変わっていないか確かめる
/// 2. この実行だけが持つ名前で作業用ファイルを新規作成する(Phase 2E.1 と同じ所有権の決まり)
/// 3. 書いたものを読み直し、行数・列数・各項目が一致するか確かめる
/// 4. そのあとで本来の名前へ移す。既存ファイルは上書きしない
/// </summary>
public sealed class PdfReader
{
    /// <summary>取り消しの検証用に差し替えられるようにしている(Phase 2A.1 と同じ考え方)。</summary>
    internal Func<string, bool> FileDeleter { get; init; } = CellMutator.TryDeleteFile;

    public PdfReadResult Execute(PdfReadPreview preview, CancellationToken cancellationToken = default)
    {
        if (!preview.CanExecute || preview.Request is not { } request
            || preview.Snapshot is not { } snapshot)
        {
            return PdfReadResult.Failed("解決していない問題があるため、ファイルを作成できません。");
        }

        if (Preflight(request.SourceFilePath, snapshot) is { } preflightError)
        {
            return PdfReadResult.Failed(preflightError);
        }

        if (File.Exists(preview.OutputPath) || File.Exists(preview.AuditPath))
        {
            return PdfReadResult.Failed(
                $"「{preview.OutputFileName}」はすでにあります。既存のファイルは上書きしません。");
        }

        var rows = ToRows(preview);
        var created = new List<string>();

        try
        {
            cancellationToken.ThrowIfCancellationRequested();

            var tempOutput = WriteOutput(preview, request, rows, created);

            if (Verify(preview, request, tempOutput, rows) is { } verifyError)
            {
                return PdfReadResult.Failed(Describe(verifyError, RollBack(created)));
            }

            var tempAudit = WriteAudit(preview, request, snapshot, rows, created);

            if (File.Exists(preview.OutputPath) || File.Exists(preview.AuditPath))
            {
                return PdfReadResult.Failed(Describe(
                    $"「{preview.OutputFileName}」が実行中に作成されました。既存ファイルは上書きしません。",
                    RollBack(created)));
            }

            File.Move(tempOutput, preview.OutputPath);
            created.Remove(tempOutput);
            created.Add(preview.OutputPath);

            File.Move(tempAudit, preview.AuditPath);
            created.Remove(tempAudit);
            created.Add(preview.AuditPath);

            var dataRows = preview.Kind == PdfDocumentKind.Table
                ? Math.Max(rows.Count - 1, 0)
                : rows.Count - 1;

            return new PdfReadResult
            {
                Success = true,
                OutputFileNames = [preview.OutputFileName],
                Message = $"{dataRows:N0} 行を取り出しました。元の PDF は変更していません。",
            };
        }
        catch (OperationCanceledException)
        {
            return PdfReadResult.Failed(Describe("読み取りを中止しました。", RollBack(created)));
        }
        catch (Exception ex)
        {
            return PdfReadResult.Failed(
                Describe($"ファイルの作成に失敗しました: {ex.Message}", RollBack(created)));
        }
    }

    /// <summary>プレビューの内容を、出力する行の形(1 行目は見出し)にする。</summary>
    internal static IReadOnlyList<string[]> ToRows(PdfReadPreview preview)
    {
        if (preview.Kind == PdfDocumentKind.Table)
        {
            return preview.TableRows;
        }

        var rows = new List<string[]> { new[] { "ページ", "行", "内容" } };
        rows.AddRange(preview.Lines.Select(line => new[]
        {
            line.Page.ToString(CultureInfo.InvariantCulture),
            line.Line.ToString(CultureInfo.InvariantCulture),
            line.Text,
        }));

        return rows;
    }

    private static string? Preflight(string sourceFilePath, SourceSnapshot snapshot)
    {
        if (!File.Exists(sourceFilePath))
        {
            return "PDF ファイルが見つかりません。もう一度プレビューを更新してください。";
        }

        SourceSnapshot current;
        try
        {
            current = MutationSnapshot.Take(sourceFilePath);
        }
        catch (Exception ex)
        {
            return $"PDF ファイルを読み取れません: {ex.Message}";
        }

        return current == snapshot
            ? null
            : "PDF ファイルがプレビュー後に変更されています。もう一度プレビューを更新してください。";
    }

    private static string WriteOutput(
        PdfReadPreview preview,
        PdfReadRequest request,
        IReadOnlyList<string[]> rows,
        List<string> created)
    {
        var owned = CreateOwnedFile(preview.OutputPath);
        created.Add(owned.Path);

        // 書き手(OpenXML / CsvWriter)がストリームを閉じるので、
        // 閉じたあとのストリームには触らない。
        if (request.OutputFormat == PdfOutputFormat.Xlsx)
        {
            using var stream = owned.Stream;
            PdfWorkbookWriter.Write(stream, rows);
        }
        else
        {
            using var writer = new CsvWriter(owned.Stream, request.CsvEncoding, request.CsvQuoteMode);
            foreach (var row in rows)
            {
                writer.WriteRow(row);
            }

            writer.Flush();
        }

        return owned.Path;
    }

    private static string WriteAudit(
        PdfReadPreview preview,
        PdfReadRequest request,
        SourceSnapshot snapshot,
        IReadOnlyList<string[]> rows,
        List<string> created)
    {
        var owned = CreateOwnedFile(preview.AuditPath);
        created.Add(owned.Path);

        using (var stream = owned.Stream)
        {
            PdfReadAuditLog.Write(
                stream, preview, request, snapshot,
                rows.Count,
                rows.Count == 0 ? 0 : rows.Max(row => row.Length));
            stream.Flush(flushToDisk: true);
        }

        return owned.Path;
    }

    /// <summary>作ったファイルを読み直し、行数・列数・各項目が指定どおりか確かめる。</summary>
    private static string? Verify(
        PdfReadPreview preview,
        PdfReadRequest request,
        string tempPath,
        IReadOnlyList<string[]> rows)
        => request.OutputFormat == PdfOutputFormat.Xlsx
            ? PdfWorkbookWriter.Verify(tempPath, rows)
            : VerifyCsv(tempPath, request, rows);

    private static string? VerifyCsv(
        string path, PdfReadRequest request, IReadOnlyList<string[]> rows)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var reader = new StreamReader(
                stream, CsvEncodings.ForReading(request.CsvEncoding),
                detectEncodingFromByteOrderMarks: true);
            using var parser = new TextFieldParser(reader)
            {
                TextFieldType = FieldType.Delimited,
                HasFieldsEnclosedInQuotes = true,
                TrimWhiteSpace = false,
            };

            parser.SetDelimiters(",");

            var index = 0;
            while (!parser.EndOfData)
            {
                var fields = parser.ReadFields();
                if (fields is null)
                {
                    continue;
                }

                if (index >= rows.Count)
                {
                    return "作成した CSV の行数が予定より多くなりました。作成を取り消しました。";
                }

                if (!fields.SequenceEqual(rows[index], StringComparer.Ordinal))
                {
                    return $"作成した CSV の {index + 1} 行目の内容が指定と違います。作成を取り消しました。";
                }

                index++;
            }

            return index == rows.Count
                ? null
                : $"作成した CSV の行数({index:N0})が予定({rows.Count:N0})と違います。"
                    + "作成を取り消しました。";
        }
        catch (Exception ex)
        {
            return $"作成した CSV を確認できませんでした: {ex.Message}";
        }
    }

    /// <summary>
    /// この実行だけが持つ作業用ファイルを新しく作る(Phase 2E.1 / D-031 と同じ決まり)。
    /// 名前が既に使われていたら、そのファイルには一切触れず別の名前で作り直す。
    /// </summary>
    private static OwnedFile CreateOwnedFile(string finalPath)
    {
        for (var attempt = 0; attempt < 8; attempt++)
        {
            var candidate = $"{finalPath}.{Guid.NewGuid():n}.tmp";
            try
            {
                return new OwnedFile(
                    candidate,
                    // OpenXML は書いたあと読み直すので、読み書き両方で開く。
                    new FileStream(candidate, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None));
            }
            catch (IOException) when (File.Exists(candidate))
            {
                // 既にあるものには一切触らず、別の名前を試す。
            }
        }

        throw new IOException("作業用のファイルを作れませんでした。");
    }

    private IReadOnlyList<string> RollBack(IEnumerable<string> created)
    {
        var remaining = new List<string>();
        foreach (var path in created.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            FileDeleter(path);

            // 削除処理の戻り値ではなく、実際に残っているかどうかを最終判断にする。
            if (File.Exists(path))
            {
                remaining.Add(Path.GetFileName(path));
            }
        }

        return remaining;
    }

    private static string Describe(string reason, IReadOnlyList<string> remaining)
        => remaining.Count == 0
            ? $"{reason}作成途中のファイルは取り消しました。元の PDF は変更していません。"
            : $"{reason}取り消せなかったファイルが残っている可能性があります。"
                + $"次のファイルを確認してください: {string.Join(" / ", remaining)}。"
                + "元の PDF は変更していません。";

    private sealed record OwnedFile(string Path, FileStream Stream);
}
