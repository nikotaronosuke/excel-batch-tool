using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelBatchTool.Ocr;

/// <summary>
/// Pack がどのモデルを持っているかの目録(`models/models.json`)。
///
/// Pack を組み直すだけで別のモデル構成を試せるようにしてある。
/// Phase 2F-B3 で「今の v4 二重読み」と「PP-OCRv5 + 日本語 v4」を
/// **同じ製品経路のまま**比べるために入れた。
/// 比較のためにベンチ専用の近道を作ると、測った数値が製品と一致しなくなる。
/// </summary>
public sealed record OcrModelSet
{
    /// <summary>1 つ目の認識モデル(数字・英字に強いほう)。</summary>
    public string PrimaryDirectory { get; init; } = "rec-multi";

    public string PrimaryName { get; init; } = "ch_PP-OCRv4_rec";

    public int PrimaryVersion { get; init; } = 4;

    /// <summary>2 つ目の認識モデル(かな漢字に強いほう)。</summary>
    public string SecondaryDirectory { get; init; } = "rec-japan";

    public string SecondaryName { get; init; } = "japan_PP-OCRv4_rec";

    public int SecondaryVersion { get; init; } = 4;

    public string DetectionDirectory { get; init; } = "det";

    public int DetectionVersion { get; init; } = 4;

    public string ClassificationDirectory { get; init; } = "cls";

    /// <summary>Paddle Inference の版。表示と控えに使う。</summary>
    public string Runtime { get; init; } = "Paddle Inference 2.6.1";

    /// <summary>
    /// 2 つ目のモデルを、1 つ目が迷ったときだけ動かすか。
    /// false なら常に両方へ通す(既定。B1 で選んだ二重読み)。
    /// </summary>
    public bool SecondaryOnDemand { get; init; }

    /// <summary>
    /// <see cref="SecondaryOnDemand"/> のとき、1 つ目の自信がこれ未満なら
    /// 2 つ目も動かす。日本語を含む読みは自信によらず必ず両方へ通す。
    /// </summary>
    public double SecondaryTrigger { get; init; } = 0.95;

    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
        WriteIndented = true,
    };

    /// <summary>
    /// 目録を読む。無ければ既定(v4 二重読み)。
    /// Pack を作った側と本体の食い違いは <see cref="OcrPackMismatchException"/> ではなく
    /// ここで既定へ倒さない — 壊れた目録は読み取り前に気づきたいので投げる。
    /// </summary>
    public static OcrModelSet Load(string modelsDirectory)
    {
        var path = Path.Combine(modelsDirectory, "models.json");
        if (!File.Exists(path))
        {
            return new OcrModelSet();
        }

        var text = File.ReadAllText(path);
        return JsonSerializer.Deserialize<OcrModelSet>(text, Options)
            ?? throw new InvalidDataException(
                "OCR Pack のモデル目録(models.json)を読めませんでした。");
    }

    public string ToJson() => JsonSerializer.Serialize(this, Options);
}
