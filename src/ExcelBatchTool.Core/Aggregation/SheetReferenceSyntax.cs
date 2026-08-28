using System.Text;

namespace ExcelBatchTool.Core.Aggregation;

/// <summary>「シート名!範囲」形式の参照を分解できなかった理由。</summary>
internal enum SheetReferenceProblem
{
    None = 0,

    /// <summary>引用符が閉じていない、「!」が無いなど、形として壊れている。</summary>
    Malformed,

    /// <summary>Sheet1:Sheet3 のように複数シートにまたがっている。</summary>
    ThreeDimensional,

    /// <summary>引用符なしのシート名に使えない文字がある(数式など)。</summary>
    FormulaLike,
}

/// <summary>
/// Excel の参照文字列に出てくる「シート名」の書き方だけを扱う共通処理。
/// 印刷範囲・印刷タイトル(D-016)とハイパーリンク(D-017)の両方から使う。
/// 数式全般は解釈しない。
/// </summary>
internal static class SheetReferenceSyntax
{
    /// <summary>引用符なしのシート名には現れない文字(数式との判別に使う)。</summary>
    private static readonly System.Buffers.SearchValues<char> FormulaLikeCharacters =
        System.Buffers.SearchValues.Create("()+-*/^&%<>=\"; ");

    /// <summary>
    /// シート名を参照内で使える形にする。アポストロフィは 2 つ重ねてエスケープし、
    /// 全体を引用符で囲む(引用は常に付けてよいので、囲むかどうかの判定はしない)。
    /// </summary>
    public static string Quote(string sheetName)
        => $"'{sheetName.Replace("'", "''", StringComparison.Ordinal)}'";

    /// <summary>
    /// 「シート名!残り」を分解する。引用符あり・なしの両方を受け付け、
    /// 引用符内の '' はエスケープとして解く。
    /// </summary>
    public static bool TrySplit(
        string text,
        out string sheetName,
        out string rest,
        out SheetReferenceProblem problem)
    {
        sheetName = string.Empty;
        rest = string.Empty;
        problem = SheetReferenceProblem.None;

        if (text.Length == 0)
        {
            problem = SheetReferenceProblem.Malformed;
            return false;
        }

        if (text[0] == '\'')
        {
            var builder = new StringBuilder();
            var index = 1;
            var closed = false;

            while (index < text.Length)
            {
                if (text[index] == '\'')
                {
                    if (index + 1 < text.Length && text[index + 1] == '\'')
                    {
                        builder.Append('\'');
                        index += 2;
                        continue;
                    }

                    closed = true;
                    index++;
                    break;
                }

                builder.Append(text[index]);
                index++;
            }

            if (!closed || index >= text.Length || text[index] != '!')
            {
                problem = SheetReferenceProblem.Malformed;
                return false;
            }

            sheetName = builder.ToString();
            rest = text[(index + 1)..];
            return true;
        }

        var separator = text.IndexOf('!', StringComparison.Ordinal);
        if (separator <= 0)
        {
            problem = SheetReferenceProblem.Malformed;
            return false;
        }

        sheetName = text[..separator];
        rest = text[(separator + 1)..];

        if (sheetName.Contains(':'))
        {
            problem = SheetReferenceProblem.ThreeDimensional;
            return false;
        }

        if (sheetName.AsSpan().IndexOfAny(FormulaLikeCharacters) >= 0)
        {
            problem = SheetReferenceProblem.FormulaLike;
            return false;
        }

        return true;
    }

    /// <summary>参照が「!」でシート名を伴っているか(引用符内の「!」は数えない)。</summary>
    public static bool HasSheetName(string text)
    {
        if (text.Length == 0)
        {
            return false;
        }

        if (text[0] == '\'')
        {
            return true;
        }

        return text.Contains('!', StringComparison.Ordinal);
    }
}
