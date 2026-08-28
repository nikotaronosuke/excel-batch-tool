using System.Collections.ObjectModel;
using ExcelBatchTool.Core;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>一括変更の対象として選べる 1 つの Worksheet。</summary>
public sealed class MutationSheetItemViewModel : ObservableObject
{
    private readonly Action _onChanged;
    private bool _isSelected;

    public MutationSheetItemViewModel(SheetInfo sheet, Action onChanged)
    {
        _onChanged = onChanged;
        SheetName = sheet.Name;
        Visibility = sheet.Visibility;
    }

    public string SheetName { get; }

    public SheetVisibility Visibility { get; }

    public string DisplayName => Visibility switch
    {
        SheetVisibility.Hidden => $"{SheetName}(非表示)",
        SheetVisibility.VeryHidden => $"{SheetName}(非常に非表示)",
        _ => SheetName,
    };

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (SetProperty(ref _isSelected, value))
            {
                _onChanged();
            }
        }
    }

    internal void Restore(bool isSelected)
    {
        _isSelected = isSelected;
        OnPropertyChanged(nameof(IsSelected));
    }
}

/// <summary>一括変更の対象候補となる 1 つの Workbook。</summary>
public sealed class MutationWorkbookViewModel : ObservableObject
{
    public MutationWorkbookViewModel(WorkbookItemViewModel workbook, Action onChanged)
    {
        FilePath = workbook.FilePath;
        FileName = workbook.FileName;

        var result = workbook.Result;
        Sheets = [.. result?.Sheets
            .Where(sheet => sheet.Kind == SheetKind.Worksheet)
            .Select(sheet => new MutationSheetItemViewModel(sheet, onChanged)) ?? []];

        UnavailableReason = result switch
        {
            null => "解析中です。",
            { Status: AnalysisStatus.Failed } => result.ErrorMessage ?? "解析できませんでした。",
            _ when result.Findings.Any(f => f.Type == FindingType.Formula)
                => "数式を含むため、現在のバージョンでは一括変更の対象にできません。",
            _ when result.Findings.Any(f => f.Type == FindingType.ExternalLink)
                => "他のブックへの外部参照を含むため、一括変更の対象にできません。",
            _ when Sheets.Count == 0 => "ワークシートがありません(グラフシート等のみ)。",
            _ => null,
        };

        CanSelect = UnavailableReason is null;
    }

    public string FilePath { get; }

    public string FileName { get; }

    public ObservableCollection<MutationSheetItemViewModel> Sheets { get; }

    public bool CanSelect { get; }

    public string? UnavailableReason { get; }

    public bool HasUnavailableReason => UnavailableReason is not null;
}

/// <summary>「セルをまとめて変更」(Phase 2A)の ViewModel。</summary>
public sealed class CellMutationViewModel : ObservableObject
{
    private readonly CellMutationPlanner _planner = new();
    private readonly CellMutator _mutator = new();

    private bool _isBusy;
    private bool _isPreviewStale = true;
    private CellMutationPreview? _preview;
    private string _cellReference = string.Empty;
    private string _newValueText = string.Empty;
    private string _outputSuffix = CellMutationDefaults.OutputSuffix;
    private CellWriteKind _writeKind = CellWriteKind.Text;
    private string _statusText = "変更するシートとセルを指定して、プレビューを更新してください。";
    private string? _resultText;
    private bool _lastRunSucceeded;

    public CellMutationViewModel()
    {
        RefreshPreviewCommand = new RelayCommand(
            () => _ = RefreshPreviewAsync(),
            () => !IsBusy && SelectedSheetCount > 0);
        ExecuteCommand = new RelayCommand(() => _ = ExecuteAsync(), () => CanExecute);
    }

    public ObservableCollection<MutationWorkbookViewModel> Workbooks { get; } = [];

    public RelayCommand RefreshPreviewCommand { get; }

    public RelayCommand ExecuteCommand { get; }

    public bool HasWorkbooks => Workbooks.Count > 0;

    public int SelectedSheetCount => Workbooks.Sum(workbook => workbook.Sheets.Count(sheet => sheet.IsSelected));

    /// <summary>変更する位置(A1 形式の単一セル)。</summary>
    public string CellReference
    {
        get => _cellReference;
        set
        {
            if (SetProperty(ref _cellReference, value))
            {
                OnSettingsChanged();
            }
        }
    }

    /// <summary>新しい値。「空欄にする」を選んだ場合は使わない。</summary>
    public string NewValueText
    {
        get => _newValueText;
        set
        {
            if (SetProperty(ref _newValueText, value))
            {
                OnSettingsChanged();
            }
        }
    }

    public string OutputSuffix
    {
        get => _outputSuffix;
        set
        {
            if (SetProperty(ref _outputSuffix, value))
            {
                OnSettingsChanged();
            }
        }
    }

    public bool IsTextKind
    {
        get => _writeKind == CellWriteKind.Text;
        set
        {
            if (value)
            {
                SetWriteKind(CellWriteKind.Text);
            }
        }
    }

    public bool IsNumberKind
    {
        get => _writeKind == CellWriteKind.Number;
        set
        {
            if (value)
            {
                SetWriteKind(CellWriteKind.Number);
            }
        }
    }

    public bool IsBlankKind
    {
        get => _writeKind == CellWriteKind.Blank;
        set
        {
            if (value)
            {
                SetWriteKind(CellWriteKind.Blank);
            }
        }
    }

    /// <summary>「空欄にする」以外では新しい値を入力する。</summary>
    public bool IsValueInputEnabled => _writeKind != CellWriteKind.Blank;

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
                OnPropertyChanged(nameof(CanExecute));
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

    public bool HasResult => !string.IsNullOrEmpty(ResultText);

    /// <summary>直前の実行が成功したか。失敗の内容を成功と同じ見た目で出さないために使う。</summary>
    public bool LastRunSucceeded
    {
        get => _lastRunSucceeded;
        private set => SetProperty(ref _lastRunSucceeded, value);
    }

    public CellMutationPreview? Preview => _preview;

    public bool HasPreview => _preview is not null && !IsPreviewStale;

    public bool CanExecute => !IsBusy && HasPreview && _preview!.CanExecute;

    public IReadOnlyList<CellMutationTargetPlan> Targets => _preview?.Targets ?? [];

    public IReadOnlyList<MergeIssue> Issues => _preview is null
        ? []
        : [.. _preview.Issues.OrderByDescending(issue => issue.Severity)];

    public bool HasIssues => Issues.Count > 0;

    public bool HasNoIssues => HasPreview && Issues.Count == 0;

    public string TargetSummaryText => _preview is null
        ? "-"
        : $"{_preview.Targets.Count:N0} シート";

    public string PlannedChangesText => _preview is null
        ? "-"
        : $"{_preview.OutputFileCount:N0} ファイルを作成 / {_preview.ChangeCount:N0} セルを変更";

    public string IssueSummaryText => _preview is null
        ? "-"
        : $"変更なし {_preview.NoOpCount:N0} 件 / 実行できない問題 {_preview.BlockCount:N0} 件";

    /// <summary>解析済みファイル一覧から候補を作り直す(以前の選択は引き継ぐ)。</summary>
    public void Sync(IEnumerable<WorkbookItemViewModel> files)
    {
        var previous = Workbooks.ToDictionary(workbook => workbook.FilePath, StringComparer.OrdinalIgnoreCase);

        Workbooks.Clear();
        foreach (var file in files)
        {
            var workbook = new MutationWorkbookViewModel(file, OnSettingsChanged);

            if (previous.TryGetValue(file.FilePath, out var old))
            {
                foreach (var sheet in workbook.Sheets)
                {
                    var match = old.Sheets.FirstOrDefault(item => item.SheetName == sheet.SheetName);
                    if (match is not null)
                    {
                        sheet.Restore(match.IsSelected);
                    }
                }
            }

            Workbooks.Add(workbook);
        }

        OnPropertyChanged(nameof(HasWorkbooks));
        OnSettingsChanged();
    }

    /// <summary>現在の指定でプレビューを作り直す。</summary>
    public async Task RefreshPreviewAsync()
    {
        var request = BuildRequest();
        if (request.Targets.Count == 0)
        {
            StatusText = "変更するシートが選択されていません。";
            return;
        }

        IsBusy = true;
        StatusText = "プレビューを作成しています…(元のファイルは読み取りのみ)";
        try
        {
            _preview = await Task.Run(() => _planner.CreatePreview(request));
            IsPreviewStale = false;
            StatusText = _preview.CanExecute
                ? "内容を確認して「変更したファイルを作成」を押してください。"
                : "実行できない問題があります。下の一覧を確認してください。";
        }
        catch (Exception ex)
        {
            _preview = null;
            StatusText = $"プレビューの作成に失敗しました: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
            RaisePreviewProperties();
        }
    }

    private CellMutationRequest BuildRequest() => new()
    {
        Targets = [.. Workbooks
            .Where(workbook => workbook.CanSelect)
            .SelectMany(workbook => workbook.Sheets
                .Where(sheet => sheet.IsSelected)
                .Select(sheet => new CellMutationTarget(workbook.FilePath, sheet.SheetName)))],
        CellReference = CellReference,
        WriteKind = _writeKind,
        TextValue = _writeKind == CellWriteKind.Text ? NewValueText : null,
        NumberText = _writeKind == CellWriteKind.Number ? NewValueText : null,
        OutputSuffix = OutputSuffix,
    };

    private async Task ExecuteAsync()
    {
        if (_preview is null || !CanExecute)
        {
            return;
        }

        IsBusy = true;
        StatusText = "変更したファイルを作成しています…";
        try
        {
            var preview = _preview;
            var result = await Task.Run(() => _mutator.Execute(preview));

            LastRunSucceeded = result.Success;
            ResultText = result.Success
                ? $"{result.Message}\n作成: {string.Join(" / ", result.OutputFileNames)}"
                : result.Message;

            StatusText = result.Success ? "作成が完了しました。" : "作成を実行できませんでした。";

            if (result.Success)
            {
                // 作った直後に同じ指定で押し直せないよう、プレビューを作り直させる。
                IsPreviewStale = true;
            }
        }
        catch (Exception ex)
        {
            LastRunSucceeded = false;
            ResultText = $"作成に失敗しました: {ex.Message}";
            StatusText = "作成を実行できませんでした。";
        }
        finally
        {
            IsBusy = false;
            RaiseCommandStates();
        }
    }

    private void SetWriteKind(CellWriteKind kind)
    {
        if (_writeKind == kind)
        {
            return;
        }

        _writeKind = kind;
        OnPropertyChanged(nameof(IsTextKind));
        OnPropertyChanged(nameof(IsNumberKind));
        OnPropertyChanged(nameof(IsBlankKind));
        OnPropertyChanged(nameof(IsValueInputEnabled));
        OnSettingsChanged();
    }

    /// <summary>指定が変わったらプレビューを無効にする(古い内容のまま実行させない)。</summary>
    private void OnSettingsChanged()
    {
        ResultText = null;
        IsPreviewStale = true;
        StatusText = SelectedSheetCount == 0
            ? "変更するシートを選んでください。"
            : "設定が変更されました。「プレビューを更新」を押してください。";
        OnPropertyChanged(nameof(SelectedSheetCount));
        RaiseCommandStates();
    }

    private void RaisePreviewProperties()
    {
        OnPropertyChanged(nameof(Preview));
        OnPropertyChanged(nameof(HasPreview));
        OnPropertyChanged(nameof(CanExecute));
        OnPropertyChanged(nameof(TargetSummaryText));
        OnPropertyChanged(nameof(PlannedChangesText));
        OnPropertyChanged(nameof(IssueSummaryText));
        OnPropertyChanged(nameof(Targets));
        OnPropertyChanged(nameof(Issues));
        OnPropertyChanged(nameof(HasIssues));
        OnPropertyChanged(nameof(HasNoIssues));
        RaiseCommandStates();
    }

    private void RaiseCommandStates()
    {
        RefreshPreviewCommand.RaiseCanExecuteChanged();
        ExecuteCommand.RaiseCanExecuteChanged();
    }
}
