using System.Text;

namespace ExcelBatchTool.Core.Pdf;

/// <summary>
/// PDF から抽出した文字の正規化。
///
/// PDF の埋め込みテキストは、フォントの逆引きの都合で康熙部首
/// (⽉ U+2F49 ≠ 月 U+6708)などの互換コードポイントで返ることがある
/// (Phase 2F-R で実測)。見た目は同じでも文字コードが違うため、
/// **NFKC 正規化だけ**を必ず通す。
///
/// それ以外の意味を変える加工(全角半角の独自変換・大文字小文字・表記揺れ吸収)は
/// 行わない。前後の空白を落とすことだけは行う。
/// </summary>
public static class PdfTextNormalization
{
    public static string Normalize(string text)
        => text.Normalize(NormalizationForm.FormKC).Trim();
}
