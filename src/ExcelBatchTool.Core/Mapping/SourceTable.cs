using System.Globalization;

namespace ExcelBatchTool.Core.Mapping;

/// <summary>データ元のセル 1 つの種類。</summary>
internal enum SourceValueKind
{
    /// <summary>値が無い。</summary>
    Blank = 0,

    /// <summary>素の文字列。</summary>
    Text,

    /// <summary>そのまま読み取ってよい数値。</summary>
    Number,

    /// <summary>読み取れない(数式・日付書式・リッチテキストなど)。</summary>
    Unsupported,
}

/// <summary>データ元のセル 1 つ。読み取れないものは理由を持つ。</summary>
internal readonly record struct SourceValue(
    SourceValueKind Kind, string? Text, double Number, string? Reason)
{
    public static SourceValue Blank() => new(SourceValueKind.Blank, null, 0, null);

    public static SourceValue OfText(string text) => new(SourceValueKind.Text, text, 0, null);

    public static SourceValue OfNumber(double number) => new(SourceValueKind.Number, null, number, null);

    public static SourceValue Unsupported(string reason) => new(SourceValueKind.Unsupported, null, 0, reason);

    public bool IsBlank => Kind == SourceValueKind.Blank;

    /// <summary>プレビュー用の表示文字列。</summary>
    public string Display => Kind switch
    {
        SourceValueKind.Text => Text ?? string.Empty,
        SourceValueKind.Number => Number.ToString(CultureInfo.InvariantCulture),
        SourceValueKind.Unsupported => "(読み取れません)",
        _ => "(空欄)",
    };
}

/// <summary>データ元の 1 行(必要な列だけ)。</summary>
internal sealed record SourceRow(int RowNumber, IReadOnlyList<SourceValue> Values);

/// <summary>項目名(ヘッダー)の読み取り結果。</summary>
internal sealed record SourceHeaderResult
{
    public IReadOnlyList<string> Columns { get; init; } = Array.Empty<string>();

    /// <summary>CSV の場合に判定した文字コード(表示用)。</summary>
    public string? EncodingName { get; init; }

    public string? Error { get; init; }

    public bool IsSuccess => Error is null;

    public static SourceHeaderResult Failed(string error) => new() { Error = error };
}

/// <summary>必要なキーに一致する行を集めた結果。</summary>
internal sealed record SourceMatchResult
{
    /// <summary>キー → 一致した 1 行(重複していたキーは含めない)。</summary>
    public IReadOnlyDictionary<string, SourceRow> RowsByKey { get; init; }
        = new Dictionary<string, SourceRow>(StringComparer.Ordinal);

    /// <summary>データ元で 2 件以上あったキー。</summary>
    public IReadOnlyCollection<string> DuplicateKeys { get; init; } = Array.Empty<string>();

    /// <summary>キーが空欄で、他の項目にも値が無かった行の数(読み飛ばし)。</summary>
    public int BlankRowCount { get; init; }

    /// <summary>キーが空欄なのに他の項目には値があった行の数(読み飛ばすが知らせる)。</summary>
    public int BlankKeyWithValueCount { get; init; }

    /// <summary>今回どの転記先にも使わなかった行の数。</summary>
    public int UnusedRowCount { get; init; }

    public string? Error { get; init; }

    public bool IsSuccess => Error is null;

    public static SourceMatchResult Failed(string error) => new() { Error = error };
}
