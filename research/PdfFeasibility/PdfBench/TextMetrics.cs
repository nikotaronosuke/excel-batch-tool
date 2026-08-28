using System.Text;

namespace PdfBench;

/// <summary>評価指標。exact match を最重要にする(1200 → 120O は不合格)。</summary>
public static class TextMetrics
{
    /// <summary>空白だけを除いた完全一致(納品可能かの判定)。大文字小文字は区別する。</summary>
    public static bool ExactIgnoringSpaces(string expected, string actual)
        => Strip(expected) == Strip(actual);

    /// <summary>
    /// 空白を除き、NFKC で正規化する。
    ///
    /// PDF の埋め込みテキストは、フォントの逆引きの都合で康熙部首(⽉ U+2F49)などの
    /// 互換コードポイントとして返ることがある(Skia + Yu Gothic で実測)。見た目は
    /// 同じでも文字コードが違うため、正規化せずに比較すると「読めているのに不一致」に
    /// なる。製品でも同じ正規化が必要になる、という research の所見。
    /// GT と抽出結果の両方に等しく適用する。
    /// </summary>
    public static string Strip(string text)
    {
        var normalized = text.Normalize(NormalizationForm.FormKC);
        var builder = new StringBuilder(normalized.Length);
        foreach (var character in normalized)
        {
            if (!char.IsWhiteSpace(character))
            {
                builder.Append(character);
            }
        }

        return builder.ToString();
    }

    /// <summary>文字正解率 = 1 - (編集距離 / 正解の文字数)。空白は除いて比べる。</summary>
    public static double CharacterAccuracy(string expected, string actual)
    {
        var e = Strip(expected);
        var a = Strip(actual);
        if (e.Length == 0)
        {
            return a.Length == 0 ? 1 : 0;
        }

        return Math.Max(0, 1.0 - (double)Levenshtein(e, a) / e.Length);
    }

    public static int Levenshtein(string left, string right)
    {
        var previous = new int[right.Length + 1];
        var current = new int[right.Length + 1];
        for (var j = 0; j <= right.Length; j++)
        {
            previous[j] = j;
        }

        for (var i = 1; i <= left.Length; i++)
        {
            current[0] = i;
            for (var j = 1; j <= right.Length; j++)
            {
                var cost = left[i - 1] == right[j - 1] ? 0 : 1;
                current[j] = Math.Min(
                    Math.Min(current[j - 1] + 1, previous[j] + 1),
                    previous[j - 1] + cost);
            }

            (previous, current) = (current, previous);
        }

        return previous[right.Length];
    }

    /// <summary>「ページ全文の中に、その値が空白を除いて 1 回以上そのまま現れるか」。</summary>
    public static bool AppearsExactly(string pageText, string value)
        => Strip(pageText).Contains(Strip(value), StringComparison.Ordinal);
}

/// <summary>1 つの測定のまとめ(JSON へ保存する)。</summary>
public sealed class StageResult
{
    public string Stage { get; init; } = string.Empty;

    public Dictionary<string, object> Metrics { get; } = [];

    public List<string> Failures { get; } = [];

    public double Seconds { get; set; }

    public int Pages { get; set; }

    public long PeakWorkingSetMb { get; set; }

    public void Save(string outDir)
    {
        Directory.CreateDirectory(outDir);
        PeakWorkingSetMb = Environment.WorkingSet / 1024 / 1024;
        try
        {
            using var process = System.Diagnostics.Process.GetCurrentProcess();
            PeakWorkingSetMb = process.PeakWorkingSet64 / 1024 / 1024;
        }
        catch (Exception)
        {
        }

        Json.Save(Path.Combine(outDir, Stage + ".json"), this);
        Console.WriteLine($"[{Stage}] pages={Pages} sec={Seconds:F1} peakMB={PeakWorkingSetMb} "
            + string.Join(" ", Metrics.Select(m => $"{m.Key}={FormatMetric(m.Value)}")));
    }

    private static string FormatMetric(object value) => value switch
    {
        double d => d.ToString("0.###"),
        _ => value.ToString() ?? string.Empty,
    };
}
