using System.Text;

namespace ExcelBatchTool.Core.CsvTransform;

/// <summary>
/// CSV を書き出す。RFC 4180 相当の決まりに従う。
///
/// - 区切りはカンマ、行の終わりは CRLF
/// - 項目にカンマ・引用符・CR・LF が入っていれば引用符で囲む
/// - 項目の中の引用符は 2 つに重ねる
/// - 「すべて引用符」を選んだときは、中身にかかわらず全項目を囲む
///
/// string.Join(",") のような単純な連結はしない(カンマや改行を含む値で壊れるため)。
/// </summary>
internal sealed class CsvWriter : IDisposable
{
    /// <summary>行の終わり。Windows の業務ツールに合わせて CRLF で固定する。</summary>
    public const string LineEnding = "\r\n";

    private const char Delimiter = ',';

    private const char Quote = '"';

    private static readonly char[] MustQuote = [Delimiter, Quote, '\r', '\n'];

    private readonly StreamWriter _writer;
    private readonly CsvQuoteMode _quoteMode;

    public CsvWriter(Stream stream, CsvOutputEncoding encoding, CsvQuoteMode quoteMode)
    {
        _quoteMode = quoteMode;
        _writer = new StreamWriter(stream, CsvEncodings.For(encoding))
        {
            NewLine = LineEnding,
        };
    }

    /// <summary>1 行書く。</summary>
    public void WriteRow(IReadOnlyList<string> fields)
    {
        for (var index = 0; index < fields.Count; index++)
        {
            if (index > 0)
            {
                _writer.Write(Delimiter);
            }

            Write(fields[index]);
        }

        _writer.Write(LineEnding);
    }

    public void Flush() => _writer.Flush();

    public void Dispose() => _writer.Dispose();

    private void Write(string field)
    {
        if (_quoteMode == CsvQuoteMode.Minimal && field.IndexOfAny(MustQuote) < 0)
        {
            _writer.Write(field);
            return;
        }

        _writer.Write(Quote);
        foreach (var character in field)
        {
            if (character == Quote)
            {
                _writer.Write(Quote); // 引用符は 2 つに重ねる。
            }

            _writer.Write(character);
        }

        _writer.Write(Quote);
    }
}

/// <summary>出力用の文字コード。CP932 は CodePagesEncodingProvider を使う。</summary>
internal static class CsvEncodings
{
    private const int Cp932 = 932;

    private static readonly object RegistrationLock = new();

    private static bool _registered;

    public static Encoding For(CsvOutputEncoding encoding) => encoding switch
    {
        CsvOutputEncoding.Utf8Bom => new UTF8Encoding(encoderShouldEmitUTF8Identifier: true),
        CsvOutputEncoding.Utf8 => new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        _ => ShiftJis(),
    };

    /// <summary>読み直すとき用(BOM は書かない)。</summary>
    public static Encoding ForReading(CsvOutputEncoding encoding) => encoding switch
    {
        CsvOutputEncoding.Utf8Bom or CsvOutputEncoding.Utf8
            => new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        _ => ShiftJis(),
    };

    public static string DisplayName(CsvOutputEncoding encoding) => encoding switch
    {
        CsvOutputEncoding.Utf8Bom => "UTF-8(BOM あり)",
        CsvOutputEncoding.Utf8 => "UTF-8(BOM なし)",
        _ => "Shift_JIS",
    };

    private static Encoding ShiftJis()
    {
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
}
