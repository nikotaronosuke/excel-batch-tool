namespace ExcelBatchTool.Core.Ocr;

/// <summary>読み取る値の種類。</summary>
public enum FormFieldKind
{
    /// <summary>そのままの文字。</summary>
    Text = 0,

    /// <summary>数値らしい文字(金額・件数など)。数値にできるときだけ数値にする。</summary>
    NumberLike,

    /// <summary>英数字のコード。数値にしない。</summary>
    Code,

    /// <summary>チェック / 丸などの印。文字として読ませない。</summary>
    Choice,
}

/// <summary>
/// 帳票の 1 項目。基準ページ上のどこを読むかと、どう扱うかを持つ。
/// レシピへ入れられるよう、画面の都合を混ぜない型付きの指定にしてある。
/// </summary>
public sealed record FormField
{
    /// <summary>出力の列名になる。</summary>
    public required string Name { get; init; }

    /// <summary>基準ページ上の読み取り領域(元のページの座標)。</summary>
    public required OcrBox Area { get; init; }

    public FormFieldKind Kind { get; init; } = FormFieldKind.Text;

    /// <summary>必須の項目。読めなければ必ず人へ回す。</summary>
    public bool IsRequired { get; init; } = true;

    /// <summary>
    /// <see cref="FormFieldKind.Choice"/> のときの選択肢。
    /// それぞれの印の領域を持つ。
    /// </summary>
    public IReadOnlyList<FormChoice> Choices { get; init; } = [];
}

/// <summary>選択肢 1 つと、その印の領域。</summary>
public sealed record FormChoice(string Label, OcrBox Area);

/// <summary>
/// 位置合わせの手がかり。ページごとの小さなずれを吸収するために使う。
/// 「店舗コード」のような、どのページにも同じ位置に出る文字を指定する。
/// </summary>
public sealed record FormAnchor(string Text, OcrBox Area);

/// <summary>同じ様式が続く帳票の読み取り指定。</summary>
public sealed record FormTemplate
{
    public required string Name { get; init; }

    /// <summary>この指定を作ったときの基準ページ。</summary>
    public int BasePage { get; init; } = 1;

    public required IReadOnlyList<FormField> Fields { get; init; }

    public IReadOnlyList<FormAnchor> Anchors { get; init; } = [];

    /// <summary>1 ページ 1 件として出す(この段階では複数件のページは扱わない)。</summary>
    public bool OneRecordPerPage { get; init; } = true;
}
