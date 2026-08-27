using System.Collections.ObjectModel;
using System.IO;
using ExcelBatchTool.Core;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>統合対象として選べる 1 つの Workbook。</summary>
public sealed class MergeSourceItemViewModel : ObservableObject
{
    private readonly Action _onSelectionChanged;
    private bool _isIncluded;
    private string? _selectedSheetName;

    public MergeSourceItemViewModel(WorkbookItemViewModel workbook, Action onSelectionChanged)
    {
        _onSelectionChanged = onSelectionChanged;
        FilePath = workbook.FilePath;
        FileName = workbook.FileName;

        var result = workbook.Result;
        SheetNames = result?.Sheets
            .Where(sheet => sheet.Kind == SheetKind.Worksheet)
            .Select(sheet => sheet.Name)
            .ToList() ?? [];

        UnavailableReason = result switch
        {
            null => "解析中です。",
            { Status: AnalysisStatus.Failed } => result.ErrorMessage ?? "解析できませんでした。",
            _ when result.Findings.Any(f => f.Type == FindingType.MacroRelated)
                => "マクロ (VBA) を含むため統合対象にできません。",
            _ when SheetNames.Count == 0 => "ワークシートがありません(グラフシート等のみ)。",
            _ => null,
        };

        CanInclude = UnavailableReason is null;
        _isIncluded = CanInclude;
        _selectedSheetName = SheetNames.FirstOrDefault();
    }

    public string FilePath { get; }

    public string FileName { get; }

    public IReadOnlyList<string> SheetNames { get; }

    public bool CanInclude { get; }

    public string? UnavailableReason { get; }

    public bool HasUnavailableReason => UnavailableReason is not null;

    public bool IsIncluded
    {
        get => _isIncluded;
        set
        {
            if (SetProperty(ref _isIncluded, value && CanInclude))
            {
                _onSelectionChanged();
            }
        }
    }

    public string? SelectedSheetName
    {
        get => _selectedSheetName;
        set
        {
            if (SetProperty(ref _selectedSheetName, value))
            {
                _onSelectionChanged();
            }
        }
    }

    /// <summary>一覧を作り直したときに、以前の選択を引き継ぐ。</summary>
    internal void Restore(bool isIncluded, string? sheetName)
    {
        _isIncluded = isIncluded && CanInclude;
        if (sheetName is not null && SheetNames.Contains(sheetName))
        {
            _selectedSheetName = sheetName;
        }

        OnPropertyChanged(nameof(IsIncluded));
        OnPropertyChanged(nameof(SelectedSheetName));
    }
}

/// <summary>「表をまとめる」(Phase 1A: 表データの縦結合)の ViewModel。</summary>
public sealed class MergeViewModel : ObservableObject
{
    private readonly Func<string, string?> _pickSavePath;
    private readonly MergePlanner _planner = new();
    private readonly TableMerger _merger = new();

    private bool _includeSourceFileColumn = true;
    private bool _includeSourceSheetColumn = true;
    private bool _isBusy;
    private bool _isPreviewStale = true;
    private MergePreview? _preview;
    private string _statusText = "解析済みのファイルから統合するシートを選び、プレビューを更新してください。";
    private string? _resultText;
    private string? _resultPath;

    public MergeViewModel(Func<string, string?> pickSavePath)
    {
        _pickSavePath = pickSavePath;
        RefreshPreviewCommand = new RelayCommand(
            () => _ = RefreshPreviewAsync(),
            () => !IsBusy && Sources.Any(source => source.IsIncluded));
        CreateCommand = new RelayCommand(() => _ = CreateAsync(), () => CanCreate);
    }

    public ObservableCollection<MergeSourceItemViewModel> Sources { get; } = [];

    public RelayCommand RefreshPreviewCommand { get; }

    public RelayCommand CreateCommand { get; }

    public bool HasSources => Sources.Count > 0;

    public bool IncludeSourceFileColumn
    {
        get => _includeSourceFileColumn;
        set
        {
            if (SetProperty(ref _includeSourceFileColumn, value))
            {
                InvalidatePreview();
            }
        }
    }

    public bool IncludeSourceSheetColumn
    {
        get => _includeSourceSheetColumn;
        set
        {
            if (SetProperty(ref _includeSourceSheetColumn, value))
            {
                InvalidatePreview();
            }
        }
    }

    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            if (SetProperty(ref _isBusy, value))
            {
                RaiseCommandStates();
            }
        }
    }

    public bool IsPreviewStale
    {
        get => _isPreviewStale;
        private set
        {
            if (SetProperty(ref _isPreviewStale, value))
            {
                OnPropertyChanged(nameof(HasPreview));
                OnPropertyChanged(nameof(CanCreate));
                RaiseCommandStates();
            }
        }
    }

    public string StatusText
    {
        get => _statusText;
        private set => SetProperty(ref _statusText, value);
    }

    public string? ResultText
    {
        get => _resultText;
        private set
        {
            if (SetProperty(ref _resultText, value))
            {
                OnPropertyChanged(nameof(HasResult));
            }
        }
    }

    public string? ResultPath
    {
        get => _resultPath;
        private set
        {
            if (SetProperty(ref _resultPath, value))
            {
                OnPropertyChanged(nameof(HasResultPath));
            }
        }
    }

    public bool HasResult => !string.IsNullOrEmpty(ResultText);

    public bool HasResultPath => !string.IsNullOrEmpty(ResultPath);

    public MergePreview? Preview => _preview;

    public bool HasPreview => _preview is not null && !IsPreviewStale;

    public bool CanCreate => !IsBusy && HasPreview && _preview!.CanExecute;

    public string TargetSummaryText => _preview is null
        ? "-"
        : $"{_preview.WorkbookCount:N0} Workbook / {_preview.SheetCount:N0} Worksheet";

    public string PlannedRowsText => _preview is null
        ? "-"
        : $"{_preview.InputDataRowCount:N0} データ行 + ヘッダー 1 行"
            + (_preview.MetadataColumnCount > 0 ? $" + metadata {_preview.MetadataColumnCount} 列" : string.Empty);

    public string IssueSummaryText => _preview is null
        ? "-"
        : $"注意 {_preview.WarningCount:N0} 件 / 実行できない問題 {_preview.BlockCount:N0} 件";

    public IReadOnlyList<string> OutputHeaders => _preview?.OutputHeaders ?? [];

    public IReadOnlyList<MergeIssue> Issues => _preview is null
        ? []
        : [.. _preview.Issues.OrderByDescending(issue => issue.Severity)];

    public bool HasIssues => Issues.Count > 0;

    public bool HasNoIssues => HasPreview && Issues.Count == 0;

    /// <summary>解析済みファイル一覧から統合対象の候補を作り直す(以前の選択は引き継ぐ)。</summary>
    public void Sync(IEnumerable<WorkbookItemViewModel> files)
    {
        var previous = Sources.ToDictionary(source => source.FilePath, StringComparer.OrdinalIgnoreCase);

        Sources.Clear();
        foreach (var file in files)
        {
            var item = new MergeSourceItemViewModel(file, InvalidatePreview);
            if (previous.TryGetValue(file.FilePath, out var old))
            {
                item.Restore(old.IsIncluded, old.SelectedSheetName);
            }

            Sources.Add(item);
        }

        OnPropertyChanged(nameof(HasSources));
        InvalidatePreview();
    }

    private void InvalidatePreview()
    {
        ResultText = null;
        ResultPath = null;
        IsPreviewStale = true;
        StatusText = _preview is null
            ? "統合するシートを選び、「プレビューを更新」を押してください。"
            : "設定が変更されました。「プレビューを更新」を押してください。";
        RaiseCommandStates();
    }

    private async Task RefreshPreviewAsync()
    {
        var selections = BuildSelections();
        if (selections.Count == 0)
        {
            StatusText = "統合するシートが選択されていません。";
            return;
        }

        IsBusy = true;
        StatusText = "プレビューを作成しています…(入力ファイルは読み取りのみ)";
        try
        {
            var options = BuildOptions();
            _preview = await Task.Run(() => _planner.CreatePreview(selections, options));
            IsPreviewStale = false;
            StatusText = _preview.CanExecute
                ? "プレビューを確認して「統合ファイルを作成」を押してください。"
                : "実行できない問題があります。下の一覧を確認してください。";
            RaisePreviewProperties();
        }
        catch (Exception ex)
        {
            _preview = null;
            StatusText = $"プレビューの作成に失敗しました: {ex.Message}";
            RaisePreviewProperties();
        }
        finally
        {
            IsBusy = false;
        }
    }

    private async Task CreateAsync()
    {
        if (_preview is null || !CanCreate)
        {
            return;
        }

        var outputPath = _pickSavePath("統合結果.xlsx");
        if (string.IsNullOrEmpty(outputPath))
        {
            return;
        }

        IsBusy = true;
        StatusText = "統合ファイルを作成しています…";
        try
        {
            var options = BuildOptions();
            var preview = _preview;
            var result = await Task.Run(() => _merger.Execute(preview, options, outputPath));

            ResultText = result.Message;
            ResultPath = result.Success ? $"保存先: {result.OutputPath}" : null;
            StatusText = result.Success ? "統合が完了しました。" : "統合を実行できませんでした。";
        }
        catch (Exception ex)
        {
            ResultText = $"統合に失敗しました: {ex.Message}";
            ResultPath = null;
            StatusText = "統合を実行できませんでした。";
        }
        finally
        {
            IsBusy = false;
        }
    }

    private List<MergeSourceSelection> BuildSelections()
        => [.. Sources
            .Where(source => source is { IsIncluded: true, SelectedSheetName: not null })
            .Select(source => new MergeSourceSelection(source.FilePath, source.SelectedSheetName!))];

    private MergeOptions BuildOptions() => new()
    {
        IncludeSourceFileColumn = IncludeSourceFileColumn,
        IncludeSourceSheetColumn = IncludeSourceSheetColumn,
    };

    private void RaisePreviewProperties()
    {
        OnPropertyChanged(nameof(Preview));
        OnPropertyChanged(nameof(HasPreview));
        OnPropertyChanged(nameof(CanCreate));
        OnPropertyChanged(nameof(TargetSummaryText));
        OnPropertyChanged(nameof(PlannedRowsText));
        OnPropertyChanged(nameof(IssueSummaryText));
        OnPropertyChanged(nameof(OutputHeaders));
        OnPropertyChanged(nameof(Issues));
        OnPropertyChanged(nameof(HasIssues));
        OnPropertyChanged(nameof(HasNoIssues));
        RaiseCommandStates();
    }

    private void RaiseCommandStates()
    {
        RefreshPreviewCommand.RaiseCanExecuteChanged();
        CreateCommand.RaiseCanExecuteChanged();
    }
}
