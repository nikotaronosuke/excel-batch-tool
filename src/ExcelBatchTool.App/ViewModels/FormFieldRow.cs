using ExcelBatchTool.Core.Ocr;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>
/// 「同じ様式の帳票として読む」ときの項目 1 つ。
///
/// 利用者が触るのは名前・種類・必須かどうかだけ。読み取る場所は
/// 1 ページ目の読み取り結果から自動で拾う(自分で座標を入れさせない)。
/// </summary>
public sealed class FormFieldRow(string name, OcrBox area) : ObservableObject
{
    private const string KindText = "そのままの文字";
    private const string KindNumber = "数値";
    private const string KindCode = "コード";

    private string _name = name;
    private string _kind = KindText;
    private bool _isRequired = true;
    private bool _useAsAnchor = true;

    public static IReadOnlyList<string> KindOptions { get; } = [KindText, KindNumber, KindCode];

    /// <summary>値を読み取る場所(元のページの座標)。</summary>
    public OcrBox Area { get; } = area;

    /// <summary>項目名そのものが出ている場所。位置合わせの手がかりに使う。</summary>
    public OcrBox LabelArea { get; init; }

    public string Name
    {
        get => _name;
        set => SetProperty(ref _name, value);
    }

    public string Kind
    {
        get => _kind;
        set => SetProperty(ref _kind, value);
    }

    public bool IsRequired
    {
        get => _isRequired;
        set => SetProperty(ref _isRequired, value);
    }

    /// <summary>この項目名をページのずれ合わせに使うか。</summary>
    public bool UseAsAnchor
    {
        get => _useAsAnchor;
        set => SetProperty(ref _useAsAnchor, value);
    }

    public string AreaText
        => $"{Area.X:F0}, {Area.Y:F0}({Area.Width:F0}×{Area.Height:F0})";

    public FormField ToField() => new()
    {
        Name = Name,
        Area = Area,
        IsRequired = IsRequired,
        Kind = Kind switch
        {
            KindNumber => FormFieldKind.NumberLike,
            KindCode => FormFieldKind.Code,
            _ => FormFieldKind.Text,
        },
    };
}
