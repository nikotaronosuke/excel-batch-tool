using System.IO;
using ExcelBatchTool.Core;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>ファイル一覧の 1 行(1 Workbook)。</summary>
public sealed class WorkbookItemViewModel : ObservableObject
{
    private WorkbookAnalysisResult? _result;

    public WorkbookItemViewModel(string filePath)
    {
        FilePath = filePath;
        FileName = Path.GetFileName(filePath);
    }

    public string FilePath { get; }

    public string FileName { get; }

    public WorkbookAnalysisResult? Result => _result;

    public bool IsAnalyzed => _result is not null;

    public string LevelGlyph => _result switch
    {
        null => "…",
        { Level: SafetyLevel.Normal } => "✅",
        { Level: SafetyLevel.NeedsAttention } => "⚠",
        _ => "✖",
    };

    public string LevelDisplay => _result switch
    {
        null => "解析中",
        { Status: AnalysisStatus.Failed, ErrorMessage: not null } => "現在非対応",
        { Level: SafetyLevel.Normal } => "通常",
        { Level: SafetyLevel.NeedsAttention } => "注意が必要",
        _ => "現在非対応",
    };

    public string SheetCountDisplay => _result is { Status: AnalysisStatus.Succeeded }
        ? _result.Sheets.Count.ToString()
        : "-";

    public string FileSizeDisplay => FormatFileSize(_result?.FileSizeBytes);

    public string StatusDisplay => _result switch
    {
        null => "解析中…",
        { Status: AnalysisStatus.Succeeded } => "完了",
        _ => "失敗",
    };

    public string? ErrorMessage => _result?.ErrorMessage;

    public bool HasError => _result?.ErrorMessage is not null;

    public IReadOnlyList<SheetRowViewModel> Sheets { get; private set; } = [];

    public IReadOnlyList<FindingRowViewModel> Findings { get; private set; } = [];

    public bool HasFindings => Findings.Count > 0;

    public bool HasNoFindings => IsAnalyzed && Findings.Count == 0 && !HasError;

    /// <summary>解析結果を反映する(UI スレッドから呼ぶこと)。</summary>
    public void Apply(WorkbookAnalysisResult result)
    {
        _result = result;
        Sheets = result.Sheets.Select(sheet => new SheetRowViewModel(sheet)).ToList();
        Findings = result.Findings.Select(finding => new FindingRowViewModel(finding)).ToList();
        OnPropertyChanged(string.Empty); // 全プロパティ更新
    }

    private static string FormatFileSize(long? bytes) => bytes switch
    {
        null => "-",
        < 1024 => $"{bytes} B",
        < 1024 * 1024 => $"{bytes / 1024.0:0.#} KB",
        < 1024L * 1024 * 1024 => $"{bytes / (1024.0 * 1024):0.#} MB",
        _ => $"{bytes / (1024.0 * 1024 * 1024):0.##} GB",
    };
}

/// <summary>シート情報の 1 行。</summary>
public sealed class SheetRowViewModel(SheetInfo sheet)
{
    public string Name { get; } = sheet.Name;

    public string KindDisplay { get; } = BuildKindDisplay(sheet);

    public string RangeDisplay { get; } = BuildRangeDisplay(sheet);

    private static string BuildKindDisplay(SheetInfo sheet)
    {
        var kind = sheet.Kind switch
        {
            SheetKind.Worksheet => "ワークシート",
            SheetKind.Chartsheet => "グラフシート",
            SheetKind.MacroSheet => "マクロシート",
            SheetKind.Dialogsheet => "ダイアログシート",
            _ => "不明",
        };

        return sheet.Visibility switch
        {
            SheetVisibility.Hidden => $"{kind}(非表示)",
            SheetVisibility.VeryHidden => $"{kind}(非常に非表示)",
            _ => kind,
        };
    }

    private static string BuildRangeDisplay(SheetInfo sheet)
    {
        if (sheet.UsedRange is null)
        {
            return sheet.Kind == SheetKind.Worksheet ? "(空)" : "-";
        }

        var detail = (sheet.EstimatedRowCount, sheet.EstimatedColumnCount) switch
        {
            (int rows, int cols) => $"(約 {rows:N0} 行 × {cols:N0} 列)",
            (int rows, null) => $"(約 {rows:N0} 行)",
            _ => string.Empty,
        };

        return $"{sheet.UsedRange} {detail}".TrimEnd();
    }
}

/// <summary>安全性チェック詳細の 1 行。</summary>
public sealed class FindingRowViewModel(WorkbookFinding finding)
{
    public string Glyph { get; } = finding.Level switch
    {
        SafetyLevel.NeedsAttention => "⚠",
        SafetyLevel.UnsupportedForNow => "✖",
        _ => "✅",
    };

    public SafetyLevel Level { get; } = finding.Level;

    public string DisplayName { get; } = finding.DisplayName;

    public string CountDisplay { get; } = finding.Count > 1 ? $"{finding.Count:N0} 件" : string.Empty;

    public string SheetsDisplay { get; } = finding.SheetNames.Count > 0
        ? "シート: " + string.Join("、", finding.SheetNames)
        : string.Empty;

    public string Description { get; } = finding.Description;
}
