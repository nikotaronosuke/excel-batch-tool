namespace ExcelBatchTool.Core.Ocr;

/// <summary>
/// 確認作業の進め方。画面から切り離してあるので、そのままテストできる。
///
/// ここには「まとめて確認済みにする」操作を置かない。要確認・読取不能は
/// 元のページを見ながら 1 件ずつ確認する、というのがこの段階の安全の要だから。
/// 代わりに「確認して次へ」を用意して、飛ばさずに速く進めるようにしている。
/// </summary>
public sealed class OcrReviewSession(OcrDocumentReading reading)
{
    private bool _showAutoAccepted;

    public OcrDocumentReading Reading { get; } = reading;

    /// <summary>
    /// 自動確定も一覧に出すか。既定は出さない(人が見るべきものだけを並べる)。
    /// 出せるようにしてあるのは、自動確定した内容も元のページと見比べられるようにするため。
    /// </summary>
    public bool ShowAutoAccepted
    {
        get => _showAutoAccepted;
        set
        {
            if (_showAutoAccepted == value)
            {
                return;
            }

            _showAutoAccepted = value;

            // 表示から外れた行を選んだままにしない。
            if (Selected is { } selected && !IsVisible(selected))
            {
                Selected = Visible.FirstOrDefault();
            }
        }
    }

    /// <summary>
    /// いま一覧に出す項目。
    ///
    /// 絞り込みは**最初の分類**で行う。確認済みにした行が消えると、
    /// 何を確認したのか見えなくなり、取り消すこともできなくなるため。
    /// </summary>
    public IReadOnlyList<OcrItem> Visible
        => [.. Reading.Items.Where(IsVisible)];

    public OcrItem? Selected { get; private set; }

    public int UnresolvedCount => Reading.UnresolvedCount;

    public bool IsComplete => Reading.IsComplete;

    public bool Select(OcrItem? item)
    {
        if (item is not null && !IsVisible(item))
        {
            return false;
        }

        Selected = item;
        return true;
    }

    /// <summary>いちばん最初の未確認を選ぶ(読み取り直後の入口)。</summary>
    public bool SelectFirstUnresolved()
    {
        var first = Visible.FirstOrDefault(item => !item.IsResolved) ?? Visible.FirstOrDefault();
        Selected = first;
        return first is not null;
    }

    /// <summary>次の未確認へ。無ければ選択を変えない。</summary>
    public bool MoveToNextUnresolved() => Move(forward: true);

    /// <summary>前の未確認へ。無ければ選択を変えない。</summary>
    public bool MoveToPreviousUnresolved() => Move(forward: false);

    /// <summary>
    /// 選んでいる項目を確認済みにして、次の未確認へ進む。
    /// <paramref name="correctedText"/> が null なら、元の読み取りのままで確認する。
    /// </summary>
    public bool ConfirmSelectedAndAdvance(string? correctedText = null)
    {
        if (Selected is not { } selected)
        {
            return false;
        }

        selected.Confirm(correctedText);
        return MoveToNextUnresolved();
    }

    private bool IsVisible(OcrItem item)
        => ShowAutoAccepted || item.InitialStatus != OcrItemStatus.AutoAccepted;

    private bool Move(bool forward)
    {
        var visible = Visible;
        if (visible.Count == 0)
        {
            return false;
        }

        var start = -1;
        for (var index = 0; index < visible.Count; index++)
        {
            if (ReferenceEquals(visible[index], Selected))
            {
                start = index;
                break;
            }
        }

        if (forward)
        {
            for (var index = start + 1; index < visible.Count; index++)
            {
                if (!visible[index].IsResolved)
                {
                    Selected = visible[index];
                    return true;
                }
            }
        }
        else
        {
            for (var index = start - 1; index >= 0; index--)
            {
                if (!visible[index].IsResolved)
                {
                    Selected = visible[index];
                    return true;
                }
            }
        }

        return false;
    }
}
