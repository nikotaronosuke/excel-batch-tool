using ExcelBatchTool.Core.Pdf;

namespace ExcelBatchTool.Core.Ocr;

/// <summary>読み取り 1 件の状態。</summary>
public enum OcrItemStatus
{
    /// <summary>自動確定。2 つのモデルが一致し、十分な自信がある。</summary>
    AutoAccepted = 0,

    /// <summary>要確認。人が見るまで出力できない。</summary>
    NeedsReview,

    /// <summary>読取不能。文字として読めていない。人が入れるまで出力できない。</summary>
    Unreadable,

    /// <summary>人が確認した(修正した / 元のままで正しいと確認した)。</summary>
    UserConfirmed,
}

public static class OcrItemStatusText
{
    public static string Display(OcrItemStatus status) => status switch
    {
        OcrItemStatus.AutoAccepted => "自動確定",
        OcrItemStatus.NeedsReview => "要確認",
        OcrItemStatus.Unreadable => "読取不能",
        OcrItemStatus.UserConfirmed => "確認済み",
        _ => string.Empty,
    };
}

/// <summary>統合前の、モデルごとの読み取り。画面と控えの根拠として残す。</summary>
public sealed record OcrEngineReading(string Engine, string Text, double Score);

/// <summary>
/// 確認・修正の対象になる読み取り 1 件。
///
/// 状態が変わるので record ではなく class。元の OCR 結果は必ず残し、
/// 修正しても上書きしない(あとから「元は何だったか」を見られるようにする)。
/// </summary>
public sealed class OcrItem
{
    public required int PageNumber { get; init; }

    /// <summary>ページ内での行番号(bbox から組み立てた順)。</summary>
    public required int LineNumber { get; init; }

    /// <summary>同じ行の中での位置(左から)。</summary>
    public required int IndexInLine { get; init; }

    /// <summary>統合後の読み取り結果(修正前)。</summary>
    public required string Text { get; init; }

    public required OcrBox BoundingBox { get; init; }

    public required double Confidence { get; init; }

    /// <summary>どちらのモデルを採用したか、なぜその状態になったかの短い説明。</summary>
    public required string Reason { get; init; }

    /// <summary>統合前の、モデルごとの読み取り。</summary>
    public required IReadOnlyList<OcrEngineReading> OriginalEngineResults { get; init; }

    public OcrItemStatus InitialStatus { get; init; }

    public OcrItemStatus Status { get; set; }

    /// <summary>人が入れ直した文字。修正していなければ null。</summary>
    public string? EditedText { get; set; }

    public bool IsUserEdited => EditedText is not null;

    /// <summary>出力に使う文字。</summary>
    public string FinalText => EditedText ?? Text;

    public bool IsResolved => Status is OcrItemStatus.AutoAccepted or OcrItemStatus.UserConfirmed;

    /// <summary>人が「この内容で正しい」と確認した(修正あり / なしの両方)。</summary>
    public void Confirm(string? correctedText = null)
    {
        if (correctedText is not null)
        {
            EditedText = PdfTextNormalization.Normalize(correctedText);
        }

        Status = OcrItemStatus.UserConfirmed;
    }

    /// <summary>確認を取り消して、元の状態に戻す。</summary>
    public void Unconfirm()
    {
        EditedText = null;
        Status = InitialStatus;
    }
}

/// <summary>1 つの PDF を OCR した結果ぜんぶ。確認が終わるまで出力できない。</summary>
public sealed class OcrDocumentReading
{
    public required IReadOnlyList<OcrItem> Items { get; init; }

    /// <summary>OCR したページ番号。</summary>
    public required IReadOnlyList<int> OcrPages { get; init; }

    public required OcrEngineInfo EngineInfo { get; init; }

    /// <summary>傾きが大きく、この段階では確定させないページ。</summary>
    public IReadOnlyList<int> NeedsDeskewPages { get; init; } = [];

    /// <summary>表らしいスキャンのページ(表としての読み取りは次の段階)。</summary>
    public IReadOnlyList<int> TableLikePages { get; init; } = [];

    /// <summary>この段階では扱えないと分かった理由。1 件でもあれば出力できない。</summary>
    public IReadOnlyList<Merge.MergeIssue> Issues { get; init; } = [];

    public int AutoAcceptedCount => Items.Count(item => item.Status == OcrItemStatus.AutoAccepted);

    public int NeedsReviewCount => Items.Count(item => item.Status == OcrItemStatus.NeedsReview);

    public int UnreadableCount => Items.Count(item => item.Status == OcrItemStatus.Unreadable);

    public int UserConfirmedCount => Items.Count(item => item.Status == OcrItemStatus.UserConfirmed);

    public int UserEditedCount => Items.Count(item => item.IsUserEdited);

    /// <summary>最初の分類でいくつが自動確定だったか(控えに残す件数はこちらを使う)。</summary>
    public int InitiallyAutoAcceptedCount
        => Items.Count(item => item.InitialStatus == OcrItemStatus.AutoAccepted);

    public int InitiallyNeedsReviewCount
        => Items.Count(item => item.InitialStatus == OcrItemStatus.NeedsReview);

    public int InitiallyUnreadableCount
        => Items.Count(item => item.InitialStatus == OcrItemStatus.Unreadable);

    public int UnresolvedCount => Items.Count(item => !item.IsResolved);

    public bool IsComplete => UnresolvedCount == 0;
}
