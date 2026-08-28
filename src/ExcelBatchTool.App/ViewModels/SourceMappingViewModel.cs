using System.Collections.ObjectModel;
using System.IO;
using ExcelBatchTool.Core;
using ExcelBatchTool.Core.Mapping;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>対応付けの 1 行(データ元の項目 → 転記先セル)。</summary>
public sealed class SourceMappingRowViewModel : ObservableObject
{
    public const string KindText = "文字";
    public const string KindNumber = "数値";

    /// <summary>「種類」列の選択肢。データ元から供給する値なので「空欄」は無い。</summary>
    public static IReadOnlyList<string> KindOptions { get; } = [KindText, KindNumber];

    private readonly Action _onChanged;
    private string _sourceColumn = string.Empty;
    private string _targetCell = string.Empty;
    private string _kindDisplay = KindText;

    public SourceMappingRowViewModel(Action onChanged, IReadOnlyList<string> columns)
    {
        _onChanged = onChanged;
        AvailableColumns = columns;
    }

    /// <summary>データ元の項目名の候補(読み込んだヘッダー)。</summary>
    public IReadOnlyList<string> AvailableColumns { get; private set; }

    public string SourceColumn
    {
        get => _sourceColumn;
        set
        {
            if (SetProperty(ref _sourceColumn, value))
            {
                _onChanged();
            }
        }
    }

    public string TargetCell
    {
        get => _targetCell;
        set
        {
            if (SetProperty(ref _targetCell, value))
            {
                _onChanged();
            }
        }
    }

    public string KindDisplay
    {
        get => _kindDisplay;
        set
        {
            if (SetProperty(ref _kindDisplay, value))
            {
                _onChanged();
            }
        }
    }

    internal void UpdateColumns(IReadOnlyList<string> columns)
    {
        AvailableColumns = columns;
        OnPropertyChanged(nameof(AvailableColumns));

        if (_sourceColumn.Length > 0 && !columns.Contains(_sourceColumn))
        {
            // 読み込み直しで消えた項目は選び直してもらう。
            _sourceColumn = string.Empty;
            OnPropertyChanged(nameof(SourceColumn));
        }
    }

    internal SourceMappingRequest ToRequest() => new()
    {
        SourceColumn = SourceColumn,
        TargetCell = TargetCell,
        WriteKind = string.Equals(KindDisplay, KindNumber, StringComparison.Ordinal)
            ? CellWriteKind.Number
            : CellWriteKind.Text,
    };
}

/// <summary>「5. 表から転記」(Phase 2C1)の ViewModel。</summary>
public sealed class SourceMappingViewModel : ObservableObject
{
    private readonly SourceMappingPlanner _planner = new();
    private readonly CellMutator _mutator = new();
    private readonly Func<string?> _pickSourceFile;

    private string _sourceFilePath = string.Empty;
    private string? _sourceSheetName;
    private string _headerRowText = "1";
    private string _keyColumn = string.Empty;
    private string _targetKeyCell = "A1";
    private string _outputSuffix = SourceMappingDefaults.OutputSuffix;

    private bool _isBusy;
    private bool _isPreviewStale = true;
    private CellMutationPreview? _preview;
    private string _statusText = "データ元のファイルを選び、「項目を読み込む」を押してください。";
    private string? _sourceInfoText;
    private string? _resultText;
    private bool _lastRunSucceeded;
    private SourceMappingRowViewModel? _selectedMapping;

    public SourceMappingViewModel()
        : this(() => null)
    {
    }

    /// <summary>テスト用: ファイル選択を差し替えられるようにする。</summary>
    internal SourceMappingViewModel(Func<string?> pickSourceFile)
    {
        _pickSourceFile = pickSourceFile;

        SelectSourceCommand = new RelayCommand(SelectSource, () => !IsBusy);
        LoadColumnsCommand = new RelayCommand(LoadColumns, () => !IsBusy && SourceFilePath.Length > 0);
        AddMappingCommand = new RelayCommand(AddMapping, () => !IsBusy && HasColumns);
        RemoveMappingCommand = new RelayCommand(
            RemoveSelectedMapping, () => !IsBusy && SelectedMapping is not null);
        RefreshPreviewCommand = new RelayCommand(
            () => _ = RefreshPreviewAsync(), () => !IsBusy && HasColumns && SelectedSheetCount > 0);
        ExecuteCommand = new RelayCommand(() => _ = ExecuteAsync(), () => CanExecute);
    }

    public ObservableCollection<MutationWorkbookViewModel> Workbooks { get; } = [];

    /// <summary>データ元の項目名(読み込み結果)。</summary>
    public ObservableCollection<string> SourceColumns { get; } = [];

    /// <summary>.xlsx のデータ元で選べるシート。</summary>
    public ObservableCollection<string> SourceSheetNames { get; } = [];

    public ObservableCollection<SourceMappingRowViewModel> Mappings { get; } = [];

    public RelayCommand SelectSourceCommand { get; }

    public RelayCommand LoadColumnsCommand { get; }

    public RelayCommand AddMappingCommand { get; }

    public RelayCommand RemoveMappingCommand { get; }

    public RelayCommand RefreshPreviewCommand { get; }

    public RelayCommand ExecuteCommand { get; }

    public bool HasWorkbooks => Workbooks.Count > 0;

    public bool HasColumns => SourceColumns.Count > 0;

    public bool IsSheetSelectionEnabled => SourceSheetNames.Count > 0;

    public int SelectedSheetCount => Workbooks.Sum(workbook => workbook.Sheets.Count(sheet => sheet.IsSelected));

    /// <summary>データ元のファイル名(パスは画面に出さない)。</summary>
    public string SourceFileNameDisplay => SourceFilePath.Length == 0
        ? "(未選択)"
        : Path.GetFileName(SourceFilePath);

    public string SourceFilePath
    {
        get => _sourceFilePath;
        private set
        {
            if (SetProperty(ref _sourceFilePath, value))
            {
                OnPropertyChanged(nameof(SourceFileNameDisplay));
            }
        }
    }

    public string? SourceSheetName
    {
        get => _sourceSheetName;
        set
        {
            if (SetProperty(ref _sourceSheetName, value))
            {
                OnSettingsChanged();
            }
        }
    }

    /// <summary>項目名の行(1 始まり)。</summary>
    public string HeaderRowText
    {
        get => _headerRowText;
        set
        {
            if (SetProperty(ref _headerRowText, value))
            {
                OnSettingsChanged();
            }
        }
    }

    public string KeyColumn
    {
        get => _keyColumn;
        set
        {
            if (SetProperty(ref _keyColumn, value))
            {
                OnSettingsChanged();
            }
        }
    }

    /// <summary>転記先シートで、キーが入っているセル。</summary>
    public string TargetKeyCell
    {
        get => _targetKeyCell;
        set
        {
            if (SetProperty(ref _targetKeyCell, value))
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

    public SourceMappingRowViewModel? SelectedMapping
    {
        get => _selectedMapping;
        set
        {
            if (SetProperty(ref _selectedMapping, value))
            {
                RemoveMappingCommand.RaiseCanExecuteChanged();
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

    /// <summary>読み込んだデータ元の概要(項目数・文字コードなど)。</summary>
    public string? SourceInfoText
    {
        get => _sourceInfoText;
        private set
        {
            if (SetProperty(ref _sourceInfoText, value))
            {
                OnPropertyChanged(nameof(HasSourceInfo));
            }
        }
    }

    public bool HasSourceInfo => !string.IsNullOrEmpty(SourceInfoText);

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
        : $"{_preview.Targets.Select(t => (t.FilePath, t.SheetName)).Distinct().Count():N0} シート"
            + $" × {Mappings.Count:N0} 項目";

    public string PlannedChangesText => _preview is null
        ? "-"
        : $"{_preview.OutputFileCount:N0} ファイルを作成 / {_preview.ChangeCount:N0} セルへ転記";

    public string IssueSummaryText => _preview is null
        ? "-"
        : $"変更なし {_preview.NoOpCount:N0} 件 / 実行できない問題 {_preview.BlockCount:N0} 件";

    /// <summary>解析済みファイル一覧から転記先の候補を作り直す(以前の選択は引き継ぐ)。</summary>
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

    /// <summary>データ元のファイルを選ぶ。</summary>
    private void SelectSource()
    {
        if (_pickSourceFile() is not { } path || path.Length == 0)
        {
            return;
        }

        SetSourceFile(path);
    }

    /// <summary>データ元のファイルを設定し、シート候補を読み直す。</summary>
    public void SetSourceFile(string path)
    {
        SourceFilePath = path;
        SourceColumns.Clear();
        SourceSheetNames.Clear();
        SourceInfoText = null;

        if (SourceMappingPlanner.KindOf(path) == SourceFileKind.Xlsx)
        {
            foreach (var name in SourceMappingPlanner.ReadSourceSheetNames(path))
            {
                SourceSheetNames.Add(name);
            }

            _sourceSheetName = SourceSheetNames.FirstOrDefault();
            OnPropertyChanged(nameof(SourceSheetName));
        }
        else
        {
            _sourceSheetName = null;
            OnPropertyChanged(nameof(SourceSheetName));
        }

        OnPropertyChanged(nameof(IsSheetSelectionEnabled));
        OnPropertyChanged(nameof(HasColumns));
        StatusText = "「項目を読み込む」を押してください。";
        OnSettingsChanged();
    }

    /// <summary>データ元の項目名を読み込む。</summary>
    internal void LoadColumns()
    {
        SourceColumns.Clear();
        SourceInfoText = null;

        if (!TryParseHeaderRow(out var headerRow))
        {
            StatusText = "項目名の行は 1 以上の数字で指定してください。";
            OnPropertyChanged(nameof(HasColumns));
            OnSettingsChanged();
            return;
        }

        var header = SourceMappingPlanner.ReadColumns(SourceFilePath, SourceSheetName, headerRow);

        if (!header.IsSuccess)
        {
            StatusText = header.Error!;
            OnPropertyChanged(nameof(HasColumns));
            OnSettingsChanged();
            return;
        }

        foreach (var column in header.Columns)
        {
            SourceColumns.Add(column);
        }

        SourceInfoText = header.EncodingName is { } encoding
            ? $"{header.Columns.Count:N0} 項目 / 文字コード {encoding}"
            : $"{header.Columns.Count:N0} 項目";

        if (!SourceColumns.Contains(KeyColumn))
        {
            _keyColumn = SourceColumns.FirstOrDefault() ?? string.Empty;
            OnPropertyChanged(nameof(KeyColumn));
        }

        foreach (var mapping in Mappings)
        {
            mapping.UpdateColumns(SourceColumns);
        }

        OnPropertyChanged(nameof(HasColumns));
        StatusText = "対応付けを指定して、プレビューを更新してください。";
        OnSettingsChanged();
    }

    /// <summary>現在の指定でプレビューを作り直す。</summary>
    public async Task RefreshPreviewAsync()
    {
        if (!TryParseHeaderRow(out var headerRow))
        {
            StatusText = "項目名の行は 1 以上の数字で指定してください。";
            return;
        }

        var request = new SourceMappingBatchRequest
        {
            SourceFilePath = SourceFilePath,
            SourceSheetName = SourceSheetName,
            HeaderRow = headerRow,
            KeyColumn = KeyColumn,
            TargetKeyCell = TargetKeyCell,
            Targets = [.. Workbooks
                .Where(workbook => workbook.CanSelect)
                .SelectMany(workbook => workbook.Sheets
                    .Where(sheet => sheet.IsSelected)
                    .Select(sheet => new CellMutationTarget(workbook.FilePath, sheet.SheetName)))],
            Mappings = [.. Mappings.Select(mapping => mapping.ToRequest())],
            OutputSuffix = OutputSuffix,
        };

        IsBusy = true;
        StatusText = "プレビューを作成しています…(データ元と転記先は読み取りのみ)";
        try
        {
            _preview = await Task.Run(() => _planner.CreatePreview(request));
            IsPreviewStale = false;
            StatusText = _preview.CanExecute
                ? "内容を確認して「転記したファイルを作成」を押してください。"
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

    private async Task ExecuteAsync()
    {
        if (_preview is null || !CanExecute)
        {
            return;
        }

        IsBusy = true;
        StatusText = "転記したファイルを作成しています…";
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

    private void AddMapping()
    {
        var mapping = new SourceMappingRowViewModel(OnSettingsChanged, SourceColumns)
        {
            SourceColumn = SourceColumns.FirstOrDefault() ?? string.Empty,
        };

        Mappings.Add(mapping);
        SelectedMapping = mapping;
        OnSettingsChanged();
    }

    private void RemoveSelectedMapping()
    {
        if (SelectedMapping is not { } mapping)
        {
            return;
        }

        Mappings.Remove(mapping);
        SelectedMapping = null;
        OnSettingsChanged();
    }

    private bool TryParseHeaderRow(out int headerRow)
        => int.TryParse(HeaderRowText, out headerRow) && headerRow >= 1;

    /// <summary>指定が変わったらプレビューを無効にする(古い内容のまま実行させない)。</summary>
    private void OnSettingsChanged()
    {
        ResultText = null;
        IsPreviewStale = true;
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
        SelectSourceCommand.RaiseCanExecuteChanged();
        LoadColumnsCommand.RaiseCanExecuteChanged();
        AddMappingCommand.RaiseCanExecuteChanged();
        RemoveMappingCommand.RaiseCanExecuteChanged();
        RefreshPreviewCommand.RaiseCanExecuteChanged();
        ExecuteCommand.RaiseCanExecuteChanged();
    }
}
