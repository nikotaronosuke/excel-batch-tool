using System.Text.Json;
using System.Text.Json.Serialization;

namespace PdfBench;

/// <summary>
/// 生成する架空データと、その正解(Ground Truth)。
/// すべて架空で、実在の人物・企業・案件のデータは含まない。
/// 座標は PDF ポイント(1pt = 1/72 インチ)。A4 = 595.28 x 841.89。
/// </summary>
public static class Layouts
{
    public const float PageWidth = 595.28f;
    public const float PageHeight = 841.89f;

    // ---- 固定帳票のフィールド位置(ラベルの右に値を書く) ----
    // name, labelX, valueX, y, valueWidth
    public static readonly (string Name, float LabelX, float ValueX, float Y, float Width)[] FormFields =
    [
        ("店舗コード", 60, 170, 150, 180),
        ("担当者", 60, 170, 195, 220),
        ("日付", 60, 170, 240, 180),
        ("売上", 60, 170, 285, 200),
        ("Q1", 60, 170, 330, 100),
        ("Q2", 60, 170, 375, 100),
        ("備考", 60, 170, 420, 360),
    ];

    // チェックボックス(Q3): はい / いいえ / 未回答
    public const float CheckY = 470;
    public static readonly (string Label, float BoxX)[] CheckOptions =
    [
        ("はい", 170),
        ("いいえ", 280),
        ("未回答", 400),
    ];
    public const float CheckBoxSize = 14;

    // ---- 表(4 列) ----
    public static readonly string[] TableHeaders = ["商品コード", "商品名", "単価", "在庫"];
    public static readonly float[] TableColumnX = [60, 170, 380, 470];
    public static readonly float TableRight = 540;
    public const float TableTop = 110;
    public const float TableRowHeight = 17;
    public const int TableRowsPerPage = 40;

    /// <summary>決定的な乱数(再現可能にする)。</summary>
    public static Random NewRandom(int seed) => new(seed);

    private static readonly string[] ProductNames =
    [
        "架空りんご", "架空みかん", "架空ぶどう", "架空の緑茶", "架空クッキー",
        "架空ノート", "架空ボールペン", "架空マグカップ", "架空タオル", "架空石けん",
        "架空カレー(中辛)", "架空スープ", "架空パスタ", "架空ジャム", "架空せんべい",
    ];

    private static readonly string[] PersonNames =
    [
        "架空 太郎", "架空 花子", "見本 一郎", "見本 二葉", "試験 三平",
        "仮名 四月", "仮名 五実", "試作 六郎", "試作 七海", "架空 八雲",
    ];

    private static readonly string[] Remarks =
    [
        "特になし", "棚卸し済み", "要フォロー", "新装開店セール実施", "",
        "備品の補充を依頼", "翌月へ持ち越し", "", "検品済み", "月末に再確認",
    ];

    public static string ProductCode(int index) => $"A{index:D4}";

    public static string ProductName(Random random) => ProductNames[random.Next(ProductNames.Length)];

    public static string Person(Random random) => PersonNames[random.Next(PersonNames.Length)];

    public static string Remark(Random random) => Remarks[random.Next(Remarks.Length)];

    public static string StoreCode(Random random)
        => $"S{random.Next(100, 999)}-{random.Next(10, 99)}";

    public static string Money(Random random)
        => random.Next(10_000, 9_999_999).ToString("N0");

    public static string DateText(Random random)
        => $"2026/{random.Next(1, 13):D2}/{random.Next(1, 29):D2}";

    public static string Phone(Random random)
        => $"000-{random.Next(1000, 9999)}-{random.Next(1000, 9999)}";
}

public sealed record TextPageGt(int Page, Dictionary<string, string> Fields);

public sealed record TablePageGt(int Page, List<string[]> Rows);

public sealed record FormPageGt(
    int Page,
    Dictionary<string, string> Fields,
    string Checkbox); // はい / いいえ / 未回答 / none

public static class Json
{
    public static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public static void Save<T>(string path, T value)
        => File.WriteAllText(path, JsonSerializer.Serialize(value, Options));

    public static T Load<T>(string path)
        => JsonSerializer.Deserialize<T>(File.ReadAllText(path), Options)!;
}
