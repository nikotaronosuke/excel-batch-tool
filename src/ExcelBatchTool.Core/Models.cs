namespace ExcelBatchTool.Core;

/// <summary>Workbook の安全性分類。</summary>
public enum SafetyLevel
{
    /// <summary>✅ 通常: 書き換え時に特別な注意を要する要素が検出されなかった。</summary>
    Normal = 0,

    /// <summary>⚠ 注意が必要: 将来の書き換え時に追加の注意・処理が必要な要素を含む。</summary>
    NeedsAttention = 1,

    /// <summary>✖ 現在非対応: 現バージョンでは扱えない(マクロ関連・読み取り不能など)。</summary>
    UnsupportedForNow = 2,
}

/// <summary>1 ファイルの解析状態。</summary>
public enum AnalysisStatus
{
    /// <summary>解析が完了した。</summary>
    Succeeded = 0,

    /// <summary>ファイルを解析できなかった(形式外・破損・読み取り失敗など)。</summary>
    Failed = 1,
}

/// <summary>解析で検出される要素の種類。</summary>
public enum FindingType
{
    Formula,
    MergedCell,
    Drawing,
    Chart,
    Image,
    PivotTable,
    ExternalLink,
    SheetProtection,
    WorkbookProtection,
    Table,
    DataValidation,
    ConditionalFormatting,
    Comment,
    ThreadedComment,
    DefinedName,
    Hyperlink,
    EmbeddedObject,
    ActiveXControl,
    CustomXml,
    MacroRelated,
    UnsupportedFileType,
    OpenFailed,
}

/// <summary>Workbook から検出された 1 種類の要素。</summary>
/// <param name="Type">要素の種類。</param>
/// <param name="Level">この要素が与える安全性分類。</param>
/// <param name="Count">検出数(個数を数えない要素は 1)。</param>
/// <param name="SheetNames">検出されたシート名。ブック全体の要素の場合は空。</param>
public sealed record WorkbookFinding(
    FindingType Type,
    SafetyLevel Level,
    int Count,
    IReadOnlyList<string> SheetNames)
{
    /// <summary>UI 表示用の名称。</summary>
    public string DisplayName => FindingCatalog.DisplayNameOf(Type);

    /// <summary>この要素に注意が必要な理由(書き換え時の観点)。</summary>
    public string Description => FindingCatalog.DescriptionOf(Type);
}

/// <summary>シートの種類。</summary>
public enum SheetKind
{
    Worksheet,
    Chartsheet,
    MacroSheet,
    Dialogsheet,
    Unknown,
}

/// <summary>1 シートの解析情報。</summary>
public sealed record SheetInfo
{
    public required string Name { get; init; }

    public SheetKind Kind { get; init; } = SheetKind.Worksheet;

    /// <summary>シートの表示状態(visible / hidden / veryHidden)。</summary>
    public bool IsHidden { get; init; }

    /// <summary>使用範囲(A1 形式)。取得できない場合は null。</summary>
    public string? UsedRange { get; init; }

    /// <summary>概算行数。取得できない場合は null。</summary>
    public int? EstimatedRowCount { get; init; }

    /// <summary>概算列数。取得できない場合は null。</summary>
    public int? EstimatedColumnCount { get; init; }
}

/// <summary>1 つの .xlsx ファイルの解析結果。解析は読み取り専用で、対象ファイルを変更しない。</summary>
public sealed record WorkbookAnalysisResult
{
    public required string FilePath { get; init; }

    public required string FileName { get; init; }

    /// <summary>ファイルサイズ(バイト)。取得できなかった場合は null。</summary>
    public long? FileSizeBytes { get; init; }

    public required AnalysisStatus Status { get; init; }

    /// <summary>ブック全体の安全性分類(検出要素の最大深刻度)。</summary>
    public required SafetyLevel Level { get; init; }

    /// <summary>解析失敗時のメッセージ。</summary>
    public string? ErrorMessage { get; init; }

    public IReadOnlyList<SheetInfo> Sheets { get; init; } = [];

    public IReadOnlyList<WorkbookFinding> Findings { get; init; } = [];
}
