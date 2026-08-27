namespace ExcelBatchTool.Core;

/// <summary>
/// 検出要素ごとの安全性分類・表示名・説明の定義。
/// 分類の基準: 「将来の書き換え時に、単純なセル値書き換えだけでは
/// 壊れる・ずれる・意味が変わる可能性がある要素」を ⚠、
/// 現バージョンで扱えないものを ✖ とする。
/// </summary>
public static class FindingCatalog
{
    private sealed record Entry(SafetyLevel Level, string DisplayName, string Description);

    private static readonly IReadOnlyDictionary<FindingType, Entry> Entries =
        new Dictionary<FindingType, Entry>
        {
            [FindingType.Formula] = new(SafetyLevel.NeedsAttention, "数式",
                "書き換え後に数式と計算結果の不整合が起きる可能性があります。参照先の変更にも注意が必要です。"),
            [FindingType.MergedCell] = new(SafetyLevel.NeedsAttention, "結合セル",
                "行・列の挿入や転記で結合範囲がずれたり、値の書き込み位置が想定と変わる可能性があります。"),
            [FindingType.Drawing] = new(SafetyLevel.NeedsAttention, "図形 (Drawing)",
                "図形はセル位置を基準に配置されるため、行・列の変更で位置がずれる可能性があります。"),
            [FindingType.Chart] = new(SafetyLevel.NeedsAttention, "グラフ",
                "グラフはデータ範囲への参照を持つため、シート名や範囲の変更で壊れる可能性があります。"),
            [FindingType.Image] = new(SafetyLevel.NeedsAttention, "画像",
                "画像はセル位置を基準に配置されるため、行・列の変更で位置がずれる可能性があります。"),
            [FindingType.PivotTable] = new(SafetyLevel.NeedsAttention, "ピボットテーブル",
                "ソース範囲とキャッシュを持つため、元データの書き換えで不整合になる可能性があります。"),
            [FindingType.ExternalLink] = new(SafetyLevel.NeedsAttention, "外部参照 (External Link)",
                "他のファイルへの参照を含みます。参照先が無い環境では値の意味が変わる可能性があります。"),
            [FindingType.SheetProtection] = new(SafetyLevel.NeedsAttention, "シート保護",
                "保護されたシートは書き換えが制限されます。保護設定を壊さない配慮が必要です。"),
            [FindingType.WorkbookProtection] = new(SafetyLevel.NeedsAttention, "ブック保護",
                "ブック構成が保護されています。シート追加・削除などが制限されます。"),
            [FindingType.Table] = new(SafetyLevel.NeedsAttention, "テーブル (ListObject)",
                "テーブルは範囲定義を持つため、行の追加・削除で定義とずれる可能性があります。"),
            [FindingType.DataValidation] = new(SafetyLevel.NeedsAttention, "データ入力規則",
                "入力規則の適用範囲が行・列の変更でずれる可能性があります。"),
            [FindingType.ConditionalFormatting] = new(SafetyLevel.NeedsAttention, "条件付き書式",
                "適用範囲や相対参照が行・列の変更でずれる可能性があります。"),
            [FindingType.Comment] = new(SafetyLevel.NeedsAttention, "コメント",
                "コメントはセルに紐づくため、行・列の変更で位置がずれる可能性があります。"),
            [FindingType.ThreadedComment] = new(SafetyLevel.NeedsAttention, "スレッド形式コメント",
                "新形式のコメントはセルに紐づくため、行・列の変更で位置がずれる可能性があります。"),
            [FindingType.DefinedName] = new(SafetyLevel.NeedsAttention, "定義名",
                "定義名は範囲参照を持つため、シートや範囲の変更で壊れる可能性があります。"),
            [FindingType.Hyperlink] = new(SafetyLevel.NeedsAttention, "ハイパーリンク",
                "リンクの適用範囲が行・列の変更でずれる可能性があります。"),
            [FindingType.EmbeddedObject] = new(SafetyLevel.NeedsAttention, "埋め込みオブジェクト (OLE)",
                "他アプリケーションのデータが埋め込まれています。書き換えで壊しやすい要素です。"),
            [FindingType.ActiveXControl] = new(SafetyLevel.NeedsAttention, "ActiveX コントロール",
                "コントロールはシート上の位置と設定を持ち、書き換えで壊しやすい要素です。"),
            [FindingType.CustomXml] = new(SafetyLevel.NeedsAttention, "Custom XML パート",
                "他システムとの連携データの可能性があるため、書き換え時も保持が必要です。"),
            [FindingType.MacroRelated] = new(SafetyLevel.UnsupportedForNow, "マクロ (VBA) 関連",
                "マクロを含むファイルは現在のバージョンでは対象外です(VBA との整合を保証できないため)。"),
            [FindingType.UnsupportedFileType] = new(SafetyLevel.UnsupportedForNow, "対象外のファイル形式",
                "現在のバージョンで扱えるのは .xlsx のみです。"),
            [FindingType.OpenFailed] = new(SafetyLevel.UnsupportedForNow, "読み取り不能",
                "パスワード保護(暗号化)されているか、ファイルが破損している可能性があります。"),
        };

    public static SafetyLevel LevelOf(FindingType type) => Entries[type].Level;

    public static string DisplayNameOf(FindingType type) => Entries[type].DisplayName;

    public static string DescriptionOf(FindingType type) => Entries[type].Description;

    /// <summary>検出要素から Workbook 全体の分類を求める(最大深刻度)。</summary>
    public static SafetyLevel OverallLevel(IEnumerable<WorkbookFinding> findings)
    {
        var level = SafetyLevel.Normal;
        foreach (var finding in findings)
        {
            if (finding.Level > level)
            {
                level = finding.Level;
            }
        }

        return level;
    }
}
