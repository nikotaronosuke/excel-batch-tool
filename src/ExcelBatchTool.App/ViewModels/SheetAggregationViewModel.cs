using System.Collections.ObjectModel;
using System.IO;
using ExcelBatchTool.Core;
using ExcelBatchTool.Core.Aggregation;
using ExcelBatchTool.Core.Merge;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>集約対象として選べる 1 つの Worksheet。</summary>
public sealed class AggregationSheetItemViewModel : ObservableObject
{
    private readonly Action _onChanged;
    private bool _isSelected;
    private string _outputName = string.Empty;

    public AggregationSheetItemViewModel(SheetInfo sheet, Action onChanged)
    {
        _onChanged = onChanged;
        SheetName = sheet.Name;
        IsHidden = sheet.IsHidden;
    }

    public string SheetName { get; }

    public bool IsHidden { get; }

    public string DisplayName => IsHidden ? $"{SheetName}(非表示)" : SheetName;

    /// <summary>利用者が出力シート名を手入力した(以後、自動提案で上書きしない)。</summary>
    public bool IsNameCustomized { get; private set; }

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

    public string OutputName
    {
        get => _outputName;
        set
        {
            if (SetProperty(ref _outputName, value))
            {
                IsNameCustomized = true;
                _onChanged();
            }
        }
    }

    /// <summary>自動提案で名前を入れる(手入力扱いにしない)。</summary>
    internal void SetProposedName(string name)
    {
        if (_outputName == name)
        {
            return;
        }

        _outputName = name;
        OnPropertyChanged(nameof(OutputName));
    }

    internal void Restore(bool isSelected, string outputName, bool isCustomized)
    {
        _isSelected = isSelected;
        _outputName = outputName;
        IsNameCustomized = isCustomized;
        OnPropertyChanged(nameof(IsSelected));
        OnPropertyChanged(nameof(OutputName));
    }
}

/// <summary>集約対象の候補となる 1 つの Workbook。</summary>
public sealed class AggregationWorkbookViewModel : ObservableObject
{
    public AggregationWorkbookViewModel(WorkbookItemViewModel workbook, Action onChanged)
    {
        FilePath = workbook.FilePath;
        FileName = workbook.FileName;

        var result = workbook.Result;
        Sheets = [.. result?.Sheets
            .Where(sheet => sheet.Kind == SheetKind.Worksheet)
            .Select(sheet => new AggregationSheetItemViewModel(sheet, onChanged)) ?? []];

        UnavailableReason = result switch
        {
            null => "解析中です。",
            { Status: AnalysisStatus.Failed } => result.ErrorMessage ?? "解析できませんでした。",
            _ when result.Findings.Any(f => f.Type == FindingType.MacroRelated)
                => "マクロ (VBA) を含むため集約対象にできません。",
            _ when result.Findings.Any(f => f.Type == FindingType.ExternalLink)
                => "他のブックへの外部参照を含むため集約対象にできません。",
            _ when Sheets.Count == 0 => "ワークシートがありません(グラフシート等のみ)。",
            _ => null,
        };

        CanSelect = UnavailableReason is null;
    }

    public string FilePath { get; }

    public string FileName { get; }

    public ObservableCollection<AggregationSheetItemViewModel> Sheets { get; }

    public bool CanSelect { get; }

    public string? UnavailableReason { get; }

    public bool HasUnavailableReason => UnavailableReason is not null;
}

/// <summary>「シートをまとめる」(Phase 1B.1: 基本要素を保持した Worksheet 集約)の ViewModel。</summary>
public sealed class SheetAggregationViewModel : ObservableObject
{
    private readonly Func<string, string?> _pickSavePath;
    private readonly SheetAggregationPlanner _planner = new();
    private readonly SheetAggregator _aggregator = new();

    private bool _isBusy;
    private bool _isPreviewStale = true;
    private SheetAggregationPreview? _preview;
    private string _statusText = "解析済みのファイルからシートを選び、プレビューを更新してください。";
    private string? _resultText;
    private string? _resultPath;

    public SheetAggregationViewModel(Func<string, string?> pickSavePath)
    {
        _pickSavePath = pickSavePath;
        RefreshPreviewCommand = new RelayCommand(
            () => _ = RefreshPreviewAsync(),
            () => !IsBusy && SelectedSheetCount > 0);
        CreateCommand = new RelayCommand(() => _ = CreateAsync(), () => CanCreate);
    }

    public ObservableCollection<AggregationWorkbookViewModel> Workbooks { get; } = [];

    public RelayCommand RefreshPreviewCommand { get; }

    public RelayCommand CreateCommand { get; }

    public bool HasWorkbooks => Workbooks.Count > 0;

    public int SelectedSheetCount => Workbooks.Sum(workbook => workbook.Sheets.Count(sheet => sheet.IsSelected));

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

    public SheetAggregationPreview? Preview => _preview;

    public bool HasPreview => _preview is not null && !IsPreviewStale;

    public bool CanCreate => !IsBusy && HasPreview && _preview!.CanExecute;

    public string TargetSummaryText => _preview is null
        ? "-"
        : $"{_preview.WorkbookCount:N0} Workbook / {_preview.SheetCount:N0} Worksheet";

    public string PlannedSheetsText => _preview is null
        ? "-"
        : $"{_preview.SheetCount:N0} Worksheet を持つ新しいブックを作成";

    public string IssueSummaryText => _preview is null
        ? "-"
        : $"注意 {_preview.WarningCount:N0} 件 / 実行できない問題 {_preview.BlockCount:N0} 件";

    public IReadOnlyList<SheetAggregationPlan> OrderedSheets => _preview?.Sheets ?? [];

    public IReadOnlyList<MergeIssue> Issues => _preview is null
        ? []
        : [.. _preview.Issues.OrderByDescending(issue => issue.Severity)];

    public bool HasIssues => Issues.Count > 0;

    public bool HasNoIssues => HasPreview && Issues.Count == 0;

    /// <summary>解析済みファイル一覧から候補を作り直す(以前の選択と出力名は引き継ぐ)。</summary>
    public void Sync(IEnumerable<WorkbookItemViewModel> files)
    {
        var previous = Workbooks.ToDictionary(workbook => workbook.FilePath, StringComparer.OrdinalIgnoreCase);

        Workbooks.Clear();
        foreach (var file in files)
        {
            var workbook = new AggregationWorkbookViewModel(file, OnSelectionChanged);

            if (previous.TryGetValue(file.FilePath, out var old))
            {
                foreach (var sheet in workbook.Sheets)
                {
                    var match = old.Sheets.FirstOrDefault(s => s.SheetName == sheet.SheetName);
                    if (match is not null)
                    {
                        sheet.Restore(match.IsSelected, match.OutputName, match.IsNameCustomized);
                    }
                }
            }
            else if (workbook.CanSelect)
            {
                // 初期値は各ブックの最初の表示ワークシート。
                var first = workbook.Sheets.FirstOrDefault(sheet => !sheet.IsHidden) ?? workbook.Sheets.FirstOrDefault();
                if (first is not null)
                {
                    first.Restore(true, first.OutputName, false);
                }
            }

            Workbooks.Add(workbook);
        }

        OnPropertyChanged(nameof(HasWorkbooks));
        OnSelectionChanged();
    }

    private void OnSelectionChanged()
    {
        RefreshProposedNames();
        ResultText = null;
        ResultPath = null;
        IsPreviewStale = true;
        StatusText = SelectedSheetCount == 0
            ? "集約するシートを選んでください。"
            : "設定が変更されました。「プレビューを更新」を押してください。";
        OnPropertyChanged(nameof(SelectedSheetCount));
        RaiseCommandStates();
    }

    /// <summary>
    /// 選択中のシートについて、手入力されていないものの出力シート名を決定的に付け直す。
    /// 手入力された名前は勝手に置き換えない。
    /// </summary>
    private void RefreshProposedNames()
    {
        var selected = SelectedSheetsInOutputOrder().ToList();
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var (_, sheet) in selected.Where(entry => entry.Sheet.IsNameCustomized))
        {
            used.Add(sheet.OutputName);
        }

        foreach (var (_, sheet) in selected.Where(entry => !entry.Sheet.IsNameCustomized))
        {
            var proposed = OutputSheetNameResolver.Propose(sheet.SheetName, used);
            sheet.SetProposedName(proposed);
            used.Add(proposed);
        }
    }

    /// <summary>出力順(ブックの追加順 → ブック内のシート順)で選択中のシートを列挙する。</summary>
    private IEnumerable<(AggregationWorkbookViewModel Workbook, AggregationSheetItemViewModel Sheet)>
        SelectedSheetsInOutputOrder()
        => Workbooks
            .Where(workbook => workbook.CanSelect)
            .SelectMany(workbook => workbook.Sheets
                .Where(sheet => sheet.IsSelected)
                .Select(sheet => (workbook, sheet)));

    /// <summary>現在の選択内容でプレビューを作り直す。</summary>
    public async Task RefreshPreviewAsync()
    {
        var selections = SelectedSheetsInOutputOrder()
            .Select(entry => new SheetSelection(entry.Workbook.FilePath, entry.Sheet.SheetName, entry.Sheet.OutputName))
            .ToList();

        if (selections.Count == 0)
        {
            StatusText = "集約するシートが選択されていません。";
            return;
        }

        IsBusy = true;
        StatusText = "プレビューを作成しています…(入力ファイルは読み取りのみ)";
        try
        {
            _preview = await Task.Run(() => _planner.CreatePreview(selections));
            IsPreviewStale = false;
            StatusText = _preview.CanExecute
                ? "プレビューを確認して「ブックを作成」を押してください。"
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

    private async Task CreateAsync()
    {
        if (_preview is null || !CanCreate)
        {
            return;
        }

        var outputPath = _pickSavePath("まとめたブック.xlsx");
        if (string.IsNullOrEmpty(outputPath))
        {
            return;
        }

        IsBusy = true;
        StatusText = "ブックを作成しています…";
        try
        {
            var preview = _preview;
            var result = await Task.Run(() => _aggregator.Execute(preview, outputPath));

            ResultText = result.Message;
            ResultPath = result.Success ? $"保存先: {result.OutputPath}" : null;
            StatusText = result.Success ? "作成が完了しました。" : "作成を実行できませんでした。";
        }
        catch (Exception ex)
        {
            ResultText = $"作成に失敗しました: {ex.Message}";
            ResultPath = null;
            StatusText = "作成を実行できませんでした。";
        }
        finally
        {
            IsBusy = false;
        }
    }

    private void RaisePreviewProperties()
    {
        OnPropertyChanged(nameof(Preview));
        OnPropertyChanged(nameof(HasPreview));
        OnPropertyChanged(nameof(CanCreate));
        OnPropertyChanged(nameof(TargetSummaryText));
        OnPropertyChanged(nameof(PlannedSheetsText));
        OnPropertyChanged(nameof(IssueSummaryText));
        OnPropertyChanged(nameof(OrderedSheets));
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
