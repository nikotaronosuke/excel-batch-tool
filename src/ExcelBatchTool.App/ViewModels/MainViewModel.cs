using System.Collections.ObjectModel;
using System.IO;
using ExcelBatchTool.Core;
using ExcelBatchTool.Core.Recipes;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>メイン画面の ViewModel。</summary>
public sealed class MainViewModel : ObservableObject
{
    private readonly BatchAnalyzer _analyzer = new();
    private readonly Func<string[]?> _pickFiles;
    private WorkbookItemViewModel? _selectedFile;
    private bool _isAnalyzing;
    private string _statusText = "ファイルを追加すると解析を開始します。対象ファイルは変更しません(読み取り専用)。";

    public MainViewModel(Func<string[]?> pickFiles, Func<string, string?> pickSavePath)
        : this(pickFiles, pickSavePath, () => null)
    {
    }

    /// <param name="recipeStore">
    /// 処理設定(レシピ)の置き場所。既定は %LOCALAPPDATA% の中で、外部へは送らない。
    /// </param>
    public MainViewModel(
        Func<string[]?> pickFiles,
        Func<string, string?> pickSavePath,
        Func<string?> pickSourceFile,
        RecipeStore? recipeStore = null,
        Func<string?>? pickPdfFile = null)
    {
        _pickFiles = pickFiles;

        // 3 つのタブで同じファイルを見る(種類ごとに一覧を分けて表示する)。
        var recipes = recipeStore ?? new RecipeStore();

        Merge = new MergeViewModel(pickSavePath);
        Aggregation = new SheetAggregationViewModel(pickSavePath);
        Mutation = new CellMutationViewModel(recipes);
        Mapping = new SourceMappingViewModel(pickSourceFile, recipes);
        TableUpdate = new TableUpdateViewModel(pickSourceFile, recipes);
        CsvTransform = new CsvTransformViewModel(pickSourceFile, recipes);
        PdfRead = new PdfReadViewModel(pickPdfFile ?? (() => null));
        SelectFilesCommand = new RelayCommand(SelectFiles, () => !IsAnalyzing);
        ClearCommand = new RelayCommand(Clear, () => !IsAnalyzing && Files.Count > 0);

        Mutation.Recipes.Reload();
        Mapping.Recipes.Reload();
        TableUpdate.Recipes.Reload();
        CsvTransform.Recipes.Reload();
    }

    public ObservableCollection<WorkbookItemViewModel> Files { get; } = [];

    /// <summary>「表をまとめる」(Phase 1A)の状態。</summary>
    public MergeViewModel Merge { get; }

    /// <summary>「シートをまとめる」(Phase 1B.1)の状態。</summary>
    public SheetAggregationViewModel Aggregation { get; }

    /// <summary>「セルをまとめて変更」(Phase 2A / 2B)の状態。</summary>
    public CellMutationViewModel Mutation { get; }

    /// <summary>「表から転記」(Phase 2C1)の状態。</summary>
    public SourceMappingViewModel Mapping { get; }

    /// <summary>「表を突合して更新」(Phase 2C2)の状態。</summary>
    public TableUpdateViewModel TableUpdate { get; }

    /// <summary>「CSV を変換」(Phase 2E)の状態。</summary>
    public CsvTransformViewModel CsvTransform { get; }

    /// <summary>「PDF を読み取る」(Phase 2F-A)の状態。</summary>
    public PdfReadViewModel PdfRead { get; }

    public RelayCommand SelectFilesCommand { get; }

    public RelayCommand ClearCommand { get; }

    public WorkbookItemViewModel? SelectedFile
    {
        get => _selectedFile;
        set
        {
            if (SetProperty(ref _selectedFile, value))
            {
                OnPropertyChanged(nameof(HasSelection));
            }
        }
    }

    public bool HasSelection => SelectedFile is not null;

    public bool IsAnalyzing
    {
        get => _isAnalyzing;
        private set
        {
            if (SetProperty(ref _isAnalyzing, value))
            {
                SelectFilesCommand.RaiseCanExecuteChanged();
                ClearCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public string StatusText
    {
        get => _statusText;
        private set => SetProperty(ref _statusText, value);
    }

    public bool HasFiles => Files.Count > 0;

    public int TotalFileCount => Files.Count;

    public int TotalSheetCount => Files.Sum(file => file.Result?.Sheets.Count ?? 0);

    public int NormalCount => CountByLevel(SafetyLevel.Normal);

    public int AttentionCount => CountByLevel(SafetyLevel.NeedsAttention);

    public int UnsupportedCount => CountByLevel(SafetyLevel.UnsupportedForNow);

    public string SummaryText => IsAnalyzing
        ? $"{TotalFileCount} ファイルを解析しています…"
        : $"{TotalFileCount} ファイルを確認しました";

    /// <summary>ファイルを追加して解析する(既に追加済みのパスは無視)。</summary>
    public async Task AddFilesAsync(IEnumerable<string> paths)
    {
        if (IsAnalyzing)
        {
            StatusText = "解析中です。完了までお待ちください。";
            return;
        }

        var known = new HashSet<string>(
            Files.Select(file => file.FilePath), StringComparer.OrdinalIgnoreCase);

        var newPaths = paths
            .Where(path => File.Exists(path)) // フォルダーや存在しないパスは対象外
            .Where(known.Add)
            .ToList();

        if (newPaths.Count == 0)
        {
            StatusText = "追加できるファイルがありませんでした(.xlsx ファイルを追加してください)。";
            return;
        }

        var items = newPaths.Select(path => new WorkbookItemViewModel(path)).ToList();
        foreach (var item in items)
        {
            Files.Add(item);
        }

        SelectedFile ??= items[0];

        IsAnalyzing = true;
        RefreshSummary();

        var completed = 0;
        var itemsByPath = items.ToDictionary(item => item.FilePath, StringComparer.OrdinalIgnoreCase);

        // Progress<T> は UI スレッドの SynchronizationContext 経由で通知される。
        var progress = new Progress<WorkbookAnalysisResult>(result =>
        {
            if (itemsByPath.TryGetValue(result.FilePath, out var item))
            {
                item.Apply(result);
            }

            completed++;
            StatusText = $"解析中… {completed}/{items.Count}";
            RefreshSummary();
        });

        try
        {
            await _analyzer.AnalyzeAsync(newPaths, progress);
        }
        finally
        {
            IsAnalyzing = false;
            RefreshSummary();
            Merge.Sync(Files);
            Aggregation.Sync(Files);
            Mutation.Sync(Files);
            Mapping.Sync(Files);
            TableUpdate.Sync(Files);
            StatusText = $"解析が完了しました({items.Count} ファイル)。対象ファイルは変更していません。";
        }
    }

    private void SelectFiles()
    {
        var paths = _pickFiles();
        if (paths is { Length: > 0 })
        {
            _ = AddFilesAsync(paths);
        }
    }

    private void Clear()
    {
        Files.Clear();
        SelectedFile = null;
        RefreshSummary();
        Merge.Sync(Files);
        Aggregation.Sync(Files);
        Mutation.Sync(Files);
        Mapping.Sync(Files);
        TableUpdate.Sync(Files);
        StatusText = "一覧をクリアしました。";
        ClearCommand.RaiseCanExecuteChanged();
    }

    private int CountByLevel(SafetyLevel level)
        => Files.Count(file => file.Result is { } result && result.Level == level);

    private void RefreshSummary()
    {
        OnPropertyChanged(nameof(HasFiles));
        OnPropertyChanged(nameof(TotalFileCount));
        OnPropertyChanged(nameof(TotalSheetCount));
        OnPropertyChanged(nameof(NormalCount));
        OnPropertyChanged(nameof(AttentionCount));
        OnPropertyChanged(nameof(UnsupportedCount));
        OnPropertyChanged(nameof(SummaryText));
        ClearCommand.RaiseCanExecuteChanged();
    }
}
