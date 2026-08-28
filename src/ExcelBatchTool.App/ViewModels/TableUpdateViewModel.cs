using System.Collections.ObjectModel;
using System.IO;
using ExcelBatchTool.Core.Mapping;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>列の対応付けの 1 行(データ元の項目 → 転記先の項目)。</summary>
public sealed class TableColumnMappingRowViewModel : ObservableObject
{
    public const string KindText = "文字";
    public const string KindNumber = "数値";

    public static IReadOnlyList<string> KindOptions { get; } = [KindText, KindNumber];

    private readonly Action _onChanged;
    private string _sourceColumn = string.Empty;
    private string _targetColumn = string.Empty;
    private string _kindDisplay = KindText;

    public TableColumnMappingRowViewModel(
        Action onChanged,
        IReadOnlyList<string> sourceColumns,
        IReadOnlyList<string> targetColumns)
    {
        _onChanged = onChanged;
        SourceColumns = sourceColumns;
        TargetColumns = targetColumns;
    }

    public IReadOnlyList<string> SourceColumns { get; private set; }

    public IReadOnlyList<string> TargetColumns { get; private set; }

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

    public string TargetColumn
    {
        get => _targetColumn;
        set
        {
            if (SetProperty(ref _targetColumn, value))
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

    internal void UpdateSourceColumns(IReadOnlyList<string> columns)
    {
        SourceColumns = columns;
        OnPropertyChanged(nameof(SourceColumns));

        if (_sourceColumn.Length > 0 && !columns.Contains(_sourceColumn))
        {
            _sourceColumn = string.Empty;
            OnPropertyChanged(nameof(SourceColumn));
        }
    }

    internal void UpdateTargetColumns(IReadOnlyList<string> columns)
    {
        TargetColumns = columns;
        OnPropertyChanged(nameof(TargetColumns));

        if (_targetColumn.Length > 0 && !columns.Contains(_targetColumn))
        {
            _targetColumn = string.Empty;
            OnPropertyChanged(nameof(TargetColumn));
        }
    }

    internal TableColumnMappingRequest ToRequest() => new()
    {
        SourceColumn = SourceColumn,
        TargetColumn = TargetColumn,
        WriteKind = string.Equals(KindDisplay, KindNumber, StringComparison.Ordinal)
            ? CellWriteKind.Number
            : CellWriteKind.Text,
    };
}

/// <summary>「6. 表を突合して更新」(Phase 2C2)の ViewModel。</summary>
public sealed class TableUpdateViewModel : ObservableObject
{
    private readonly TableUpdatePlanner _planner = new();
    private readonly CellMutator _mutator = new();
    private readonly Func<string?> _pickSourceFile;

    private string _sourceFilePath = string.Empty;
    private string? _sourceSheetName;
    private string _sourceHeaderRowText = "1";
    private string _sourceKeyColumn = string.Empty;
    private string _targetHeaderRowText = "1";
    private string _targetKeyColumn = string.Empty;
    private (string FilePath, string SheetName)? _referenceSheet;
    private string _outputSuffix = TableUpdateDefaults.OutputSuffix;

    private bool _isBusy;
    private bool _isPreviewStale = true;
    private TableUpdatePreview? _preview;
    private string _statusText = "データ元と転記先を指定して、プレビューを更新してください。";
    private string? _sourceInfoText;
    private string? _resultText;
    private bool _lastRunSucceeded;
    private TableColumnMappingRowViewModel? _selectedMapping;

    public TableUpdateViewModel()
        : this(() => null)
    {
    }

    /// <summary>テスト用: ファイル選択を差し替えられるようにする。</summary>
    internal TableUpdateViewModel(Func<string?> pickSourceFile)
    {
        _pickSourceFile = pickSourceFile;

        SelectSourceCommand = new RelayCommand(SelectSource, () => !IsBusy);
        LoadSourceColumnsCommand = new RelayCommand(
            LoadSourceColumns, () => !IsBusy && SourceFilePath.Length > 0);
        LoadTargetColumnsCommand = new RelayCommand(
            LoadTargetColumns, () => !IsBusy && ReferenceSheetChoices.Count > 0);
        AddMappingCommand = new RelayCommand(
            AddMapping, () => !IsBusy && SourceColumns.Count > 0 && TargetColumns.Count > 0);
        RemoveMappingCommand = new RelayCommand(
            RemoveSelectedMapping, () => !IsBusy && SelectedMapping is not null);
        RefreshPreviewCommand = new RelayCommand(
            () => _ = RefreshPreviewAsync(),
            () => !IsBusy && SourceColumns.Count > 0 && SelectedSheetCount > 0);
        ExecuteCommand = new RelayCommand(() => _ = ExecuteAsync(), () => CanExecute);
    }

    public ObservableCollection<MutationWorkbookViewModel> Workbooks { get; } = [];

    public ObservableCollection<string> SourceColumns { get; } = [];

    public ObservableCollection<string> SourceSheetNames { get; } = [];

    /// <summary>転記先の項目名(基準シートから読んだもの)。</summary>
    public ObservableCollection<string> TargetColumns { get; } = [];

    /// <summary>「項目を読み込む基準シート」の候補(選択中の転記先シート)。</summary>
    public ObservableCollection<string> ReferenceSheetChoices { get; } = [];

    public ObservableCollection<TableColumnMappingRowViewModel> Mappings { get; } = [];

    public RelayCommand SelectSourceCommand { get; }

    public RelayCommand LoadSourceColumnsCommand { get; }

    public RelayCommand LoadTargetColumnsCommand { get; }

    public RelayCommand AddMappingCommand { get; }

    public RelayCommand RemoveMappingCommand { get; }

    public RelayCommand RefreshPreviewCommand { get; }

    public RelayCommand ExecuteCommand { get; }

    public bool HasWorkbooks => Workbooks.Count > 0;

    public int SelectedSheetCount => Workbooks.Sum(workbook => workbook.Sheets.Count(sheet => sheet.IsSelected));

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

    public string SourceHeaderRowText
    {
        get => _sourceHeaderRowText;
        set
        {
            if (SetProperty(ref _sourceHeaderRowText, value))
            {
                OnSettingsChanged();
            }
        }
    }

    public string SourceKeyColumn
    {
        get => _sourceKeyColumn;
        set
        {
            if (SetProperty(ref _sourceKeyColumn, value))
            {
                OnSettingsChanged();
            }
        }
    }

    public string TargetHeaderRowText
    {
        get => _targetHeaderRowText;
        set
        {
            if (SetProperty(ref _targetHeaderRowText, value))
            {
                OnSettingsChanged();
            }
        }
    }

    public string TargetKeyColumn
    {
        get => _targetKeyColumn;
        set
        {
            if (SetProperty(ref _targetKeyColumn, value))
            {
                OnSettingsChanged();
            }
        }
    }

    /// <summary>「ファイル名 / シート名」形式の基準シート選択。</summary>
    public string? ReferenceSheetDisplay
    {
        get => _referenceSheet is { } reference
            ? $"{Path.GetFileName(reference.FilePath)} / {reference.SheetName}"
            : null;
        set
        {
            _referenceSheet = ResolveReferenceSheet(value);
            OnPropertyChanged(nameof(ReferenceSheetDisplay));
            OnSettingsChanged();
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

    public TableColumnMappingRowViewModel? SelectedMapping
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

    public TableUpdatePreview? Preview => _preview;

    public bool HasPreview => _preview is not null && !IsPreviewStale;

    public bool CanExecute => !IsBusy && HasPreview && _preview!.Mutation.CanExecute;

    public IReadOnlyList<CellMutationTargetPlan> Targets => _preview?.Mutation.Targets ?? [];

    public IReadOnlyList<MergeIssue> Issues => _preview is null
        ? []
        : [.. _preview.Mutation.Issues.OrderByDescending(issue => issue.Severity)];

    public bool HasIssues => Issues.Count > 0;

    public bool HasNoIssues => HasPreview && Issues.Count == 0;

    /// <summary>突合の集計(一致 / 片側のみ / 重複 / 空欄)。</summary>
    public string MatchSummaryText
    {
        get
        {
            if (_preview?.Summary is not { } summary)
            {
                return "-";
            }

            return $"一致 {summary.MatchedKeyCount:N0} 件 / "
                + $"データ元のみ {summary.SourceOnlyKeyCount:N0} 件 / "
                + $"転記先のみ {summary.TargetOnlyKeyCount:N0} 件"
                + (summary.DuplicateKeyCount > 0 ? $" / 重複 {summary.DuplicateKeyCount:N0} 件" : string.Empty)
                + (summary.BlankKeyRowCount > 0 ? $" / 空欄 {summary.BlankKeyRowCount:N0} 行" : string.Empty);
        }
    }

    public string PlannedChangesText => _preview is null
        ? "-"
        : $"{_preview.Mutation.OutputFileCount:N0} ファイルを作成 / {_preview.Mutation.ChangeCount:N0} セルを更新"
            + (_preview.Mutation.NoOpCount > 0 ? $" / 変更なし {_preview.Mutation.NoOpCount:N0} 件" : string.Empty);

    public string IssueSummaryText => _preview is null
        ? "-"
        : $"注意 {_preview.Mutation.WarningCount:N0} 件 / 実行できない問題 {_preview.Mutation.BlockCount:N0} 件";

    /// <summary>解析済みファイル一覧から転記先の候補を作り直す(以前の選択は引き継ぐ)。</summary>
    public void Sync(IEnumerable<WorkbookItemViewModel> files)
    {
        var previous = Workbooks.ToDictionary(workbook => workbook.FilePath, StringComparer.OrdinalIgnoreCase);

        Workbooks.Clear();
        foreach (var file in files)
        {
            var workbook = new MutationWorkbookViewModel(file, OnTargetSelectionChanged);

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
        OnTargetSelectionChanged();
    }

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
        }
        else
        {
            _sourceSheetName = null;
        }

        OnPropertyChanged(nameof(SourceSheetName));
        OnPropertyChanged(nameof(IsSheetSelectionEnabled));
        StatusText = "「データ元の項目を読み込む」を押してください。";
        OnSettingsChanged();
    }

    public bool IsSheetSelectionEnabled => SourceSheetNames.Count > 0;

    /// <summary>データ元の項目名を読み込む。</summary>
    public void LoadSourceColumns()
    {
        SourceColumns.Clear();
        SourceInfoText = null;

        if (!TryParseRow(SourceHeaderRowText, out var headerRow))
        {
            StatusText = "データ元の項目名の行は 1 以上の数字で指定してください。";
            OnSettingsChanged();
            return;
        }

        var header = SourceMappingPlanner.ReadColumns(SourceFilePath, SourceSheetName, headerRow);
        if (!header.IsSuccess)
        {
            StatusText = header.Error!;
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

        if (!SourceColumns.Contains(SourceKeyColumn))
        {
            _sourceKeyColumn = SourceColumns.FirstOrDefault() ?? string.Empty;
            OnPropertyChanged(nameof(SourceKeyColumn));
        }

        foreach (var mapping in Mappings)
        {
            mapping.UpdateSourceColumns(SourceColumns);
        }

        StatusText = "転記先の項目も読み込んで、対応付けを指定してください。";
        OnSettingsChanged();
    }

    /// <summary>基準シートから転記先の項目名を読み込む。</summary>
    public void LoadTargetColumns()
    {
        TargetColumns.Clear();

        if (_referenceSheet is not { } reference)
        {
            StatusText = "項目を読み込む基準シートを選んでください。";
            OnSettingsChanged();
            return;
        }

        if (!TryParseRow(TargetHeaderRowText, out var headerRow))
        {
            StatusText = "転記先の項目名の行は 1 以上の数字で指定してください。";
            OnSettingsChanged();
            return;
        }

        // 基準シートは候補を出すために読むだけ。他のシートはプレビュー時に個別検証する。
        var header = SourceMappingPlanner.ReadColumns(
            reference.FilePath, reference.SheetName, headerRow);

        if (!header.IsSuccess)
        {
            StatusText = header.Error!.Replace("データ元", "転記先", StringComparison.Ordinal);
            OnSettingsChanged();
            return;
        }

        foreach (var column in header.Columns)
        {
            TargetColumns.Add(column);
        }

        if (!TargetColumns.Contains(TargetKeyColumn))
        {
            _targetKeyColumn = TargetColumns.FirstOrDefault() ?? string.Empty;
            OnPropertyChanged(nameof(TargetKeyColumn));
        }

        foreach (var mapping in Mappings)
        {
            mapping.UpdateTargetColumns(TargetColumns);
        }

        StatusText = "対応付けを指定して、プレビューを更新してください。";
        OnSettingsChanged();
    }

    public async Task RefreshPreviewAsync()
    {
        if (!TryParseRow(SourceHeaderRowText, out var sourceHeaderRow)
            || !TryParseRow(TargetHeaderRowText, out var targetHeaderRow))
        {
            StatusText = "項目名の行は 1 以上の数字で指定してください。";
            return;
        }

        var request = new TableUpdateBatchRequest
        {
            SourceFilePath = SourceFilePath,
            SourceSheetName = SourceSheetName,
            SourceHeaderRow = sourceHeaderRow,
            SourceKeyColumn = SourceKeyColumn,
            TargetHeaderRow = targetHeaderRow,
            TargetKeyColumn = TargetKeyColumn,
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
            StatusText = _preview.Mutation.CanExecute
                ? "内容を確認して「更新したファイルを作成」を押してください。"
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
        StatusText = "更新したファイルを作成しています…";
        try
        {
            var preview = _preview;
            var result = await Task.Run(() => _mutator.Execute(preview.Mutation));

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
        var mapping = new TableColumnMappingRowViewModel(OnSettingsChanged, SourceColumns, TargetColumns)
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

    /// <summary>転記先の選択が変わったら、基準シートの候補も作り直す。</summary>
    private void OnTargetSelectionChanged()
    {
        var current = ReferenceSheetDisplay;

        ReferenceSheetChoices.Clear();
        foreach (var workbook in Workbooks.Where(item => item.CanSelect))
        {
            foreach (var sheet in workbook.Sheets.Where(item => item.IsSelected))
            {
                ReferenceSheetChoices.Add($"{workbook.FileName} / {sheet.SheetName}");
            }
        }

        // 以前の基準が選択から外れたら選び直してもらう(先頭を黙って基準にしない)。
        if (current is not null && ReferenceSheetChoices.Contains(current))
        {
            _referenceSheet = ResolveReferenceSheet(current);
        }
        else
        {
            _referenceSheet = null;
        }

        OnPropertyChanged(nameof(ReferenceSheetDisplay));
        OnSettingsChanged();
    }

    private (string FilePath, string SheetName)? ResolveReferenceSheet(string? display)
    {
        if (string.IsNullOrEmpty(display))
        {
            return null;
        }

        foreach (var workbook in Workbooks)
        {
            foreach (var sheet in workbook.Sheets)
            {
                if (string.Equals(
                    $"{workbook.FileName} / {sheet.SheetName}", display, StringComparison.Ordinal))
                {
                    return (workbook.FilePath, sheet.SheetName);
                }
            }
        }

        return null;
    }

    private static bool TryParseRow(string text, out int row)
        => int.TryParse(text, out row) && row >= 1;

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
        OnPropertyChanged(nameof(MatchSummaryText));
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
        LoadSourceColumnsCommand.RaiseCanExecuteChanged();
        LoadTargetColumnsCommand.RaiseCanExecuteChanged();
        AddMappingCommand.RaiseCanExecuteChanged();
        RemoveMappingCommand.RaiseCanExecuteChanged();
        RefreshPreviewCommand.RaiseCanExecuteChanged();
        ExecuteCommand.RaiseCanExecuteChanged();
    }
}
