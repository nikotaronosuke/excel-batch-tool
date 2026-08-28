using System.Text;
using Microsoft.VisualBasic.FileIO;

namespace ExcelBatchTool.Core.Mapping;

/// <summary>
/// データ元の CSV を読む。区切りはカンマのみ。引用符・引用符内のカンマ・
/// 二重化した引用符・引用符内の改行・CRLF / LF をすべて正しく扱うため、
/// 自前の split ではなく <see cref="TextFieldParser"/> を使う。
///
/// 行番号は「CSV レコード番号」で数える(引用符内の改行があるため物理行番号は使わない)。
/// ヘッダーが 1 番目のレコード、最初のデータ行が 2 番目。
/// </summary>
internal static class CsvSourceReader
{
    /// <summary>ヘッダー(1 レコード目)だけを読む。</summary>
    public static SourceHeaderResult ReadHeader(string filePath)
    {
        if (SourceEncoding.Detect(filePath, out var encoding, out var encodingName, out var detectError)
            is false)
        {
            return SourceHeaderResult.Failed(detectError!);
        }

        try
        {
            using var parser = CreateParser(filePath, encoding!);
            if (parser.EndOfData)
            {
                return SourceHeaderResult.Failed("データ元の CSV が空です。");
            }

            var fields = parser.ReadFields();
            if (fields is null || fields.Length == 0)
            {
                return SourceHeaderResult.Failed("データ元の CSV に項目名の行がありません。");
            }

            if (SourceHeaders.Validate(fields, out var columns, out var headerError) is false)
            {
                return SourceHeaderResult.Failed(headerError!);
            }

            return new SourceHeaderResult { Columns = columns!, EncodingName = encodingName };
        }
        catch (MalformedLineException ex)
        {
            return SourceHeaderResult.Failed(
                $"データ元の CSV を読み取れません({ex.LineNumber} 行目付近)。引用符の対応を確認してください。");
        }
        catch (Exception ex)
        {
            return SourceHeaderResult.Failed($"データ元の CSV を読み取れません: {ex.Message}");
        }
    }

    /// <summary>
    /// 必要なキーに一致する行だけを集める。行そのものは必要な分しか保持しないが、
    /// 重複キーの検出のためキーの一覧だけは通して見る。
    /// </summary>
    public static SourceMatchResult ReadRows(
        string filePath,
        int columnCount,
        int keyIndex,
        IReadOnlyList<int> valueIndexes,
        IReadOnlySet<string> requiredKeys,
        CancellationToken cancellationToken)
    {
        if (SourceEncoding.Detect(filePath, out var encoding, out _, out var detectError) is false)
        {
            return SourceMatchResult.Failed(detectError!);
        }

        var rowsByKey = new Dictionary<string, SourceRow>(StringComparer.Ordinal);
        var duplicates = new HashSet<string>(StringComparer.Ordinal);
        var seenKeys = new HashSet<string>(StringComparer.Ordinal);
        var blankRows = 0;
        var blankKeyWithValue = 0;
        var unused = 0;

        try
        {
            using var parser = CreateParser(filePath, encoding!);

            var recordNumber = 0;
            while (!parser.EndOfData)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var fields = parser.ReadFields();
                recordNumber++;

                if (fields is null)
                {
                    continue;
                }

                if (recordNumber == 1)
                {
                    continue; // ヘッダー。
                }

                if (fields.Length != columnCount)
                {
                    return SourceMatchResult.Failed(
                        $"データ元の CSV の {recordNumber} 行目の列数({fields.Length})が"
                            + $"項目名の行({columnCount})と違います。読み取り位置がずれるため中止します。");
                }

                var key = fields[keyIndex];

                if (key.Length == 0)
                {
                    // キーが空欄の行は転記先と対応付けられない。値の有無だけ知らせる。
                    if (valueIndexes.Any(index => fields[index].Length > 0))
                    {
                        blankKeyWithValue++;
                    }
                    else
                    {
                        blankRows++;
                    }

                    continue;
                }

                if (!seenKeys.Add(key))
                {
                    duplicates.Add(key);
                    rowsByKey.Remove(key);
                    continue;
                }

                if (!requiredKeys.Contains(key))
                {
                    unused++;
                    continue;
                }

                rowsByKey[key] = new SourceRow(
                    recordNumber,
                    [.. valueIndexes.Select(index =>
                        fields[index].Length == 0 ? SourceValue.Blank() : SourceValue.OfText(fields[index]))]);
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (MalformedLineException ex)
        {
            return SourceMatchResult.Failed(
                $"データ元の CSV を読み取れません({ex.LineNumber} 行目付近)。引用符の対応を確認してください。");
        }
        catch (Exception ex)
        {
            return SourceMatchResult.Failed($"データ元の CSV を読み取れません: {ex.Message}");
        }

        return new SourceMatchResult
        {
            RowsByKey = rowsByKey,
            DuplicateKeys = duplicates,
            BlankRowCount = blankRows,
            BlankKeyWithValueCount = blankKeyWithValue,
            UnusedRowCount = unused,
        };
    }

    private static TextFieldParser CreateParser(string filePath, Encoding encoding)
    {
        var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);

        // BOM を先頭の項目名に混ぜないよう StreamReader に取り除かせる。
        // UTF-16 は判定の時点で弾いているので、ここで別の文字コードへ切り替わることはない。
        var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true);

        var parser = new TextFieldParser(reader)
        {
            TextFieldType = FieldType.Delimited,
            HasFieldsEnclosedInQuotes = true,
            TrimWhiteSpace = false,
        };

        parser.SetDelimiters(",");
        return parser;
    }
}

/// <summary>
/// CSV の文字コードを判定する。日本の実務では UTF-8 と CP932 が混在するため、
/// BOM → 厳密な UTF-8 → CP932 の順で決める(推測でどちらかに寄せない)。
/// </summary>
internal static class SourceEncoding
{
    public const string Utf8Name = "UTF-8";

    public const string ShiftJisName = "Shift_JIS";

    /// <summary>Windows-31J(CP932)。CodePagesEncodingProvider が必要。</summary>
    private const int Cp932 = 932;

    private static readonly object RegistrationLock = new();

    private static bool _registered;

    public static bool Detect(
        string filePath,
        out Encoding? encoding,
        out string? encodingName,
        out string? error)
    {
        encoding = null;
        encodingName = null;
        error = null;

        byte[] bytes;
        try
        {
            bytes = File.ReadAllBytes(filePath);
        }
        catch (Exception ex)
        {
            error = $"データ元の CSV を読み取れません: {ex.Message}";
            return false;
        }

        if (StartsWith(bytes, [0xEF, 0xBB, 0xBF]))
        {
            encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
            encodingName = Utf8Name;
            return true;
        }

        if (StartsWith(bytes, [0xFF, 0xFE]) || StartsWith(bytes, [0xFE, 0xFF]))
        {
            error = "データ元の CSV が UTF-16 で保存されています。"
                + "現在のバージョンでは UTF-8 または Shift_JIS の CSV のみ扱えます。";
            return false;
        }

        // BOM が無い場合は、まず厳密な UTF-8 として読めるかを試す。
        try
        {
            _ = new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes);
            encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
            encodingName = Utf8Name;
            return true;
        }
        catch (DecoderFallbackException)
        {
            // UTF-8 として不正なバイトがあるので CP932 とみなす。
        }

        try
        {
            encoding = GetCp932();
            encodingName = ShiftJisName;
            return true;
        }
        catch (Exception ex)
        {
            error = $"データ元の CSV の文字コードを判定できません: {ex.Message}";
            return false;
        }
    }

    private static Encoding GetCp932()
    {
        // .NET Core 以降は CodePagesEncodingProvider を登録しないと CP932 を取得できない。
        lock (RegistrationLock)
        {
            if (!_registered)
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                _registered = true;
            }
        }

        return Encoding.GetEncoding(Cp932);
    }

    private static bool StartsWith(byte[] bytes, byte[] prefix)
        => bytes.Length >= prefix.Length && bytes.Take(prefix.Length).SequenceEqual(prefix);
}

/// <summary>項目名(ヘッダー)の共通検証。</summary>
internal static class SourceHeaders
{
    /// <summary>前後の空白だけ落とす。空・重複は勝手に直さず Block する。</summary>
    public static bool Validate(
        IReadOnlyList<string?> raw,
        out IReadOnlyList<string>? columns,
        out string? error)
    {
        columns = null;
        error = null;

        var trimmed = raw.Select(value => (value ?? string.Empty).Trim()).ToList();

        for (var index = 0; index < trimmed.Count; index++)
        {
            if (trimmed[index].Length == 0)
            {
                error = $"データ元の {index + 1} 列目の項目名が空です。"
                    + "項目名の行をすべて埋めてから読み込んでください。";
                return false;
            }
        }

        foreach (var duplicate in trimmed
            .GroupBy(name => name, StringComparer.Ordinal)
            .Where(group => group.Count() > 1))
        {
            error = $"データ元の項目名「{duplicate.Key}」が重複しています。"
                + "どの列を使うか決められないため、項目名を分けてください。";
            return false;
        }

        columns = trimmed;
        return true;
    }
}
