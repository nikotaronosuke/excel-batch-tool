using ExcelBatchTool.Core.Ocr;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>
/// 確認・修正の一覧に出す 1 行。中身は <see cref="OcrItem"/> そのもので、
/// 画面はその状態を見せて変えるだけ(元の OCR 結果は上書きしない)。
/// </summary>
public sealed class OcrReviewRow(OcrItem item) : ObservableObject
{
    public OcrItem Item { get; } = item;

    public int PageNumber => Item.PageNumber;

    public int LineNumber => Item.LineNumber;

    public string PositionText => $"{Item.PageNumber} / {Item.LineNumber}";

    /// <summary>読み取った内容(修正前)。</summary>
    public string OriginalText => Item.Text;

    public string ConfidenceText => Item.Confidence.ToString("P0");

    public string StatusText => OcrItemStatusText.Display(Item.Status);

    public string ReasonText => Item.Reason;

    /// <summary>モデルごとの読み。どちらが違うのかを見て直せるようにする。</summary>
    public string EngineText => string.Join(
        " / ",
        Item.OriginalEngineResults.Select(
            reading => $"{reading.Engine}「{reading.Text}」{reading.Score:P0}"));

    public bool IsResolved => Item.IsResolved;

    public bool NeedsAttention => !Item.IsResolved;

    public bool IsUserEdited => Item.IsUserEdited;

    /// <summary>修正値。編集しただけでは確認済みにしない(確認は明示の操作)。</summary>
    public string EditedText
    {
        get => Item.EditedText ?? Item.Text;
        set
        {
            if (string.Equals(value, EditedText, StringComparison.Ordinal))
            {
                return;
            }

            Item.EditedText = value;
            RaiseAll();
        }
    }

    /// <summary>この内容で正しいと確認する(修正した場合も、元のままの場合も)。</summary>
    public void Confirm()
    {
        Item.Confirm(Item.EditedText);
        RaiseAll();
    }

    /// <summary>編集中の文字を元の読み取りへ戻す(確認済みにはしない)。</summary>
    public void ResetEdit()
    {
        Item.EditedText = null;
        RaiseAll();
    }

    /// <summary>外から状態が変わったときに、画面の表示を更新する。</summary>
    public void Refresh() => RaiseAll();

    /// <summary>確認を取り消して、読み取った直後の状態へ戻す。</summary>
    public void Unconfirm()
    {
        Item.Unconfirm();
        RaiseAll();
    }

    private void RaiseAll()
    {
        OnPropertyChanged(nameof(EditedText));
        OnPropertyChanged(nameof(StatusText));
        OnPropertyChanged(nameof(IsResolved));
        OnPropertyChanged(nameof(NeedsAttention));
        OnPropertyChanged(nameof(IsUserEdited));
    }
}
