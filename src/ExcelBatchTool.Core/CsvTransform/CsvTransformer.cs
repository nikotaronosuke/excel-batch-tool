using System.Security.Cryptography;
using System.Text;
using ExcelBatchTool.Core.Mapping;
using ExcelBatchTool.Core.Mutation;
using Microsoft.VisualBasic.FileIO;

namespace ExcelBatchTool.Core.CsvTransform;

/// <summary>
/// プレビューどおりの CSV を作る。
///
/// 1. 実行直前にデータ元が変わっていないか確かめる
/// 2. 同じフォルダーの一時ファイルへ 1 行ずつ書き出す(全行をメモリへ載せない)
/// 3. 書いたものを読み直し、項目名・行数・列数・各項目が一致するか確かめる
/// 4. そのあとで本来の名前へ移す。既存ファイルは上書きしない
///
/// 途中で失敗したら一時ファイルを片付け、消せなかったものは名前で知らせる。
/// </summary>
public sealed class CsvTransformer
{
    /// <summary>
    /// ファイルを 1 つ消す。取り消しが本当に効かない状況をテストで再現できるよう、
    /// ここだけ差し替えられるようにしている(Phase 2A.1 と同じ考え方)。
    /// </summary>
    internal Func<string, bool> FileDeleter { get; init; } = CellMutator.TryDeleteFile;

    public CsvTransformResult Execute(
        CsvTransformPreview preview, CancellationToken cancellationToken = default)
    {
        if (!preview.CanExecute || preview.Request is not { } request || preview.Snapshot is not { } snapshot)
        {
            return CsvTransformResult.Failed("解決していない問題があるため、CSV を作成できません。");
        }

        if (Preflight(request.SourceFilePath, snapshot) is { } preflightError)
        {
            return CsvTransformResult.Failed(preflightError);
        }

        if (File.Exists(preview.OutputPath) || File.Exists(preview.AuditPath))
        {
            return CsvTransformResult.Failed(
                $"「{preview.OutputFileName}」はすでにあります。既存のファイルは上書きしません。");
        }

        // 作業用ファイルは、この実行だけが持つ名前で新規作成する。
        // 実行前からあったファイルは、上書きも取り消しの対象にもしない。
        var created = new List<string>();

        try
        {
            var written = Write(preview, request, created, out var tempOutput, cancellationToken);

            if (Verify(preview, request, tempOutput, written) is { } verifyError)
            {
                return CsvTransformResult.Failed(Describe(verifyError, RollBack(created)));
            }

            var tempAudit = WriteAudit(preview, request, snapshot, written.RowCount, created);

            // 直前にもう一度確かめてから確定する。
            if (File.Exists(preview.OutputPath) || File.Exists(preview.AuditPath))
            {
                return CsvTransformResult.Failed(Describe(
                    $"「{preview.OutputFileName}」が実行中に作成されました。既存ファイルは上書きしません。",
                    RollBack(created)));
            }

            File.Move(tempOutput, preview.OutputPath);
            created.Remove(tempOutput);
            created.Add(preview.OutputPath);

            File.Move(tempAudit, preview.AuditPath);
            created.Remove(tempAudit);
            created.Add(preview.AuditPath);

            return new CsvTransformResult
            {
                Success = true,
                RowCount = written.RowCount,
                OutputFileNames = [preview.OutputFileName],
                Message = $"{written.RowCount:N0} 行の CSV を作成しました。元のファイルは変更していません。",
            };
        }
        catch (OperationCanceledException)
        {
            return CsvTransformResult.Failed(
                Describe("CSV の作成を中止しました。", RollBack(created)));
        }
        catch (Exception ex)
        {
            return CsvTransformResult.Failed(
                Describe($"CSV の作成に失敗しました: {ex.Message}", RollBack(created)));
        }
    }

    /// <summary>データ元がプレビュー時と同じか確かめる。</summary>
    private static string? Preflight(string sourceFilePath, SourceSnapshot snapshot)
    {
        if (!File.Exists(sourceFilePath))
        {
            return "データ元のファイルが見つかりません。もう一度プレビューを更新してください。";
        }

        SourceSnapshot current;
        try
        {
            current = MutationSnapshot.Take(sourceFilePath);
        }
        catch (Exception ex)
        {
            return $"データ元のファイルを読み取れません: {ex.Message}";
        }

        return current == snapshot
            ? null
            : "データ元のファイルがプレビュー後に変更されています。もう一度プレビューを更新してください。";
    }

    /// <summary>一時ファイルへ 1 行ずつ書く。書いた内容の照合用に指紋も作る。</summary>
    private static WrittenCsv Write(
        CsvTransformPreview preview,
        CsvTransformRequest request,
        List<string> created,
        out string tempPath,
        CancellationToken cancellationToken)
    {
        var header = preview.Columns.Select(column => column.OutputName).ToArray();
        using var fingerprint = new CsvFingerprint();
        var rowCount = 0;

        // 新規作成できた時点で「この実行が持つファイル」になる。
        // 取り消しの対象に入れてから、はじめて中身を書く。
        var owned = CreateOwnedFile(preview.OutputPath);
        tempPath = owned.Path;
        created.Add(owned.Path);

        using (var stream = owned.Stream)
        using (var writer = new CsvWriter(stream, request.Encoding, request.QuoteMode))
        {
            writer.WriteRow(header);
            fingerprint.Add(header);

            var error = ReadSource(preview, request, values =>
            {
                writer.WriteRow(values);
                fingerprint.Add(values);
                rowCount++;
                return true;
            },
            cancellationToken);

            if (error is not null)
            {
                throw new InvalidOperationException(error);
            }

            writer.Flush();
            stream.Flush(flushToDisk: true);
        }

        return new WrittenCsv(header, rowCount, fingerprint.Value);
    }

    /// <summary>控えを一時ファイルへ書く。既存ファイルは決して上書きしない。</summary>
    private static string WriteAudit(
        CsvTransformPreview preview,
        CsvTransformRequest request,
        SourceSnapshot snapshot,
        int rowCount,
        List<string> created)
    {
        var owned = CreateOwnedFile(preview.AuditPath);
        created.Add(owned.Path);

        using (var stream = owned.Stream)
        {
            CsvTransformAuditLog.Write(stream, preview, request, snapshot, rowCount);
            stream.Flush(flushToDisk: true);
        }

        return owned.Path;
    }

    /// <summary>
    /// この実行だけが持つ作業用ファイルを、同じフォルダーに新しく作る。
    ///
    /// 名前が既に使われていたら、そのファイルには一切触れず別の名前で作り直す。
    /// 固定名にすると、実行前からあった同名ファイルを取り消しの対象にしてしまうため、
    /// 実行ごとに違う名前にしている。名前は作業用ファイルにしか使わないので、
    /// 出来上がるファイル名にも控えにも残らない。
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
                    new FileStream(candidate, FileMode.CreateNew, FileAccess.Write, FileShare.None));
            }
            catch (IOException) when (File.Exists(candidate))
            {
                // 既にあるものには一切触らず、別の名前を試す。
            }
        }

        throw new IOException("作業用のファイルを作れませんでした。");
    }

    /// <summary>
    /// 書いた CSV を読み直し、項目名・行数・列数・各項目がそのままか確かめる。
    /// 書きながら作った指紋と突き合わせるので、全行をメモリに置かずに全項目を照合できる。
    /// </summary>
    private static string? Verify(
        CsvTransformPreview preview, CsvTransformRequest request, string tempPath, WrittenCsv written)
    {
        if (request.Encoding == CsvOutputEncoding.Utf8Bom && !HasUtf8Bom(tempPath))
        {
            return "作成した CSV に BOM が付いていません。作成を取り消しました。";
        }

        if (request.Encoding == CsvOutputEncoding.Utf8 && HasUtf8Bom(tempPath))
        {
            return "作成した CSV に余分な BOM が付いています。作成を取り消しました。";
        }

        using var fingerprint = new CsvFingerprint();
        var rowCount = 0;

        try
        {
            using var stream = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var reader = new StreamReader(
                stream, CsvEncodings.ForReading(request.Encoding), detectEncodingFromByteOrderMarks: true);
            using var parser = new TextFieldParser(reader)
            {
                TextFieldType = FieldType.Delimited,
                HasFieldsEnclosedInQuotes = true,
                TrimWhiteSpace = false,
            };

            parser.SetDelimiters(",");

            var recordNumber = 0;
            while (!parser.EndOfData)
            {
                var fields = parser.ReadFields();
                recordNumber++;

                if (fields is null)
                {
                    continue;
                }

                if (fields.Length != written.Header.Count)
                {
                    return $"作成した CSV の {recordNumber} 行目の列数が違います。作成を取り消しました。";
                }

                if (recordNumber == 1)
                {
                    if (!fields.SequenceEqual(written.Header, StringComparer.Ordinal))
                    {
                        return "作成した CSV の項目名が指定と違います。作成を取り消しました。";
                    }
                }
                else
                {
                    rowCount++;
                }

                fingerprint.Add(fields);
            }
        }
        catch (Exception ex)
        {
            return $"作成した CSV を確認できませんでした: {ex.Message}";
        }

        if (rowCount != written.RowCount)
        {
            return $"作成した CSV の行数({rowCount:N0})が予定({written.RowCount:N0})と違います。"
                + "作成を取り消しました。";
        }

        return fingerprint.Value == written.Fingerprint
            ? null
            : "作成した CSV の内容が指定と一致しません。作成を取り消しました。";
    }

    /// <summary>データ元を 1 回流し読みして、出力する 1 行ずつを渡す。</summary>
    private static string? ReadSource(
        CsvTransformPreview preview,
        CsvTransformRequest request,
        Func<IReadOnlyList<string>, bool> onRow,
        CancellationToken cancellationToken)
    {
        var columns = preview.Columns;
        var sourceColumns = preview.SourceColumns;
        var indexes = columns
            .Select(column => column.ValueSourceKind == CsvValueSourceKind.SourceColumn
                ? sourceColumns.ToList().IndexOf(column.SourceColumn!)
                : -1)
            .ToArray();

        var values = new string[columns.Count];

        bool Emit(Func<int, (bool Ok, string Text, string? Reason)> read, bool allBlank)
        {
            if (allBlank)
            {
                return true; // すべて空欄の行は出力しない(プレビューと同じ扱い)。
            }

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

                var (ok, text, _) = read(indexes[index]);
                if (!ok)
                {
                    // プレビューで全行を確かめているので、ここへ来るのは想定外。
                    throw new InvalidOperationException(
                        "データ元に読み取れない値があります。もう一度プレビューを更新してください。");
                }

                values[index] = text;
            }

            return onRow(values);
        }

        return CsvTransformPlanner.KindOf(request.SourceFilePath) == SourceFileKind.Csv
            ? CsvSourceReader.ReadRecords(
                request.SourceFilePath,
                sourceColumns.Count,
                (_, fields) => Emit(index => (true, fields[index], null),
                    fields.All(field => field.Length == 0)),
                cancellationToken)
            : XlsxSourceReader.ReadRecords(
                request.SourceFilePath,
                request.SourceSheetName ?? string.Empty,
                request.HeaderRow,
                sourceColumns.Count,
                (_, cells) => Emit(index => CsvTransformPlanner.Render(cells[index]),
                    cells.All(value => value.IsBlank)),
                cancellationToken);
    }

    private static bool HasUtf8Bom(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        Span<byte> head = stackalloc byte[3];
        return stream.Read(head) == 3 && head[0] == 0xEF && head[1] == 0xBB && head[2] == 0xBF;
    }

    /// <summary>作りかけを片付ける。消せなかったものはファイル名で返す。</summary>
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
            ? $"{reason}作成途中のファイルは取り消しました。元のファイルは変更していません。"
            : $"{reason}取り消せなかったファイルが残っている可能性があります。"
                + $"次のファイルを確認してください: {string.Join(" / ", remaining)}。"
                + "元のファイルは変更していません。";

    private sealed record WrittenCsv(IReadOnlyList<string> Header, int RowCount, string Fingerprint);

    /// <summary>新規作成できた作業用ファイル(この実行が持つもの)。</summary>
    private sealed record OwnedFile(string Path, FileStream Stream);
}

/// <summary>
/// 書いた内容と読み直した内容を突き合わせるための指紋。
/// 全項目を順に流し込むので、全行をメモリへ置かずに 1 項目ずつ照合できる。
/// </summary>
internal sealed class CsvFingerprint : IDisposable
{
    private readonly IncrementalHash _hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);

    public void Add(IReadOnlyList<string> fields)
    {
        foreach (var field in fields)
        {
            // 長さも混ぜて、項目の切れ目が変わったことも見つけられるようにする。
            _hash.AppendData(BitConverter.GetBytes(field.Length));
            _hash.AppendData(Encoding.UTF8.GetBytes(field));
        }

        _hash.AppendData([0x1E]); // レコード区切り。
    }

    public string Value => Convert.ToHexString(_hash.GetCurrentHash());

    public void Dispose() => _hash.Dispose();
}
