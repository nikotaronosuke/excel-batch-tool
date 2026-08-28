using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using ExcelBatchTool.Core.CsvTransform;
using ExcelBatchTool.Core.Mapping;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Recipes;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>出力する CSV の 1 列(項目名・入れる内容)。</summary>
public sealed class CsvOutputColumnRowViewModel : ObservableObject
{
    public const string KindSource = "データ元";
    public const string KindFixed = "固定値";
    public const string KindBlank = "空欄";

    /// <summary>「入れる内容」列の選択肢。</summary>
    public static IReadOnlyList<string> KindOptions { get; } = [KindSource, KindFixed, KindBlank];

    private readonly Action _onChanged;
    private string _outputName = string.Empty;
    private string _kindDisplay = KindSource;
    private string _sourceColumn = string.Empty;
    private string _fixedValue = string.Empty;

    public CsvOutputColumnRowViewModel(Action onChanged, IReadOnlyList<string> sourceColumns)
    {
        _onChanged = onChanged;
        SourceColumns = sourceColumns;
    }

    /// <summary>データ元の項目名の候補。</summary>
    public IReadOnlyList<string> SourceColumns { get; private set; }

    /// <summary>出力する CSV の項目名。</summary>
    public string OutputName
    {
        get => _outputName;
        set
        {
            if (SetProperty(ref _outputName, value))
            {
                _onChanged();
            }
        }
    }

    /// <summary>データ元 / 固定値 / 空欄。</summary>
    public string KindDisplay
    {
        get => _kindDisplay;
        set
        {
            if (SetProperty(ref _kindDisplay, value))
            {
                OnPropertyChanged(nameof(IsSourceColumn));
                OnPropertyChanged(nameof(IsFixedText));
                _onChanged();
            }
        }
    }

    public bool IsSourceColumn => string.Equals(_kindDisplay, KindSource, StringComparison.Ordinal);

    public bool IsFixedText => string.Equals(_kindDisplay, KindFixed, StringComparison.Ordinal);

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

    public string FixedValue
    {
        get => _fixedValue;
        set
        {
            if (SetProperty(ref _fixedValue, value))
            {
                _onChanged();
            }
        }
    }

    internal CsvValueSourceKind Kind => _kindDisplay switch
    {
        KindFixed => CsvValueSourceKind.FixedText,
        KindBlank => CsvValueSourceKind.Blank,
        _ => CsvValueSourceKind.SourceColumn,
    };

    internal static string DisplayOf(CsvValueSourceKind kind) => kind switch
    {
        CsvValueSourceKind.FixedText => KindFixed,
        CsvValueSourceKind.Blank => KindBlank,
        _ => KindSource,
    };

    /// <summary>候補を入れ替える。選べなくなった項目名があれば、その名前を返す。</summary>
    internal string? UpdateSourceColumns(IReadOnlyList<string> columns)
    {
        SourceColumns = columns;
        OnPropertyChanged(nameof(SourceColumns));

        if (IsSourceColumn && _sourceColumn.Length > 0 && !columns.Contains(_sourceColumn))
        {
            var dropped = _sourceColumn;
            _sourceColumn = string.Empty;
            OnPropertyChanged(nameof(SourceColumn));
            return dropped;
        }

        return null;
    }

    internal CsvOutputColumnRequest ToRequest() => new()
    {
        OutputName = OutputName,
        ValueSourceKind = Kind,
        SourceColumn = Kind == CsvValueSourceKind.SourceColumn ? SourceColumn : null,
        FixedValue = Kind == CsvValueSourceKind.FixedText ? FixedValue : null,
    };
}

/// <summary>「7. CSV を変換」(Phase 2E)の ViewModel。</summary>
public sealed class CsvTransformViewModel : ObservableObject, IRecipeHost
{
    private const string EncodingUtf8Bom = "UTF-8(BOM あり)";
    private const string EncodingUtf8 = "UTF-8(BOM なし)";
    private const string EncodingShiftJis = "Shift_JIS";

    private const string QuoteMinimal = "必要なときだけ";
    private const string QuoteAll = "すべての項目";

    private readonly CsvTransformPlanner _planner = new();
    private readonly CsvTransformer _transformer = new();
    private readonly Func<string?> _pickSourceFile;

    private string? _recipeSourceSheetName;
    private SavedRecipe? _lastSuccessfulRecipe;

    private string _sourceFilePath = string.Empty;
    private string? _sourceSheetName;
    private string _headerRowText = "1";
    private string _encodingDisplay = EncodingUtf8Bom;
    private string _quoteDisplay = QuoteMinimal;
    private string _outputSuffix = CsvTransformDefaults.OutputSuffix;

    private bool _isBusy;
    private bool _isPreviewStale = true;
    private CsvTransformPreview? _preview;
    private string _statusText = "データ元のファイルを選び、「項目を読み込む」を押してください。";
    private string? _sourceInfoText;
    private string? _resultText;
    private bool _lastRunSucceeded;
    private CsvOutputColumnRowViewModel? _selectedColumn;

    public CsvTransformViewModel()
        : this(() => null)
    {
    }

    /// <summary>テスト用: ファイル選択・レシピの置き場所を差し替えられるようにする。</summary>
    internal CsvTransformViewModel(
        Func<string?> pickSourceFile,
        RecipeStore? recipeStore = null,
        Func<string, bool>? confirm = null)
    {
        _pickSourceFile = pickSourceFile;
        Recipes = new RecipeAreaViewModel(
            this, recipeStore ?? new RecipeStore(), confirm ?? RecipeSaveGuard.AskInDialog);

        SelectSourceCommand = new RelayCommand(SelectSource, () => !IsBusy);
        LoadColumnsCommand = new RelayCommand(LoadColumns, () => !IsBusy && SourceFilePath.Length > 0);
        AddColumnCommand = new RelayCommand(AddColumn, () => !IsBusy);
        AddAllSourceColumnsCommand = new RelayCommand(
            AddAllSourceColumns, () => !IsBusy && SourceColumns.Count > 0);
        RemoveColumnCommand = new RelayCommand(
            RemoveSelectedColumn, () => !IsBusy && SelectedColumn is not null);
        MoveUpCommand = new RelayCommand(() => Move(-1), () => CanMove(-1));
        MoveDownCommand = new RelayCommand(() => Move(1), () => CanMove(1));
        RefreshPreviewCommand = new RelayCommand(
            () => _ = RefreshPreviewAsync(), () => !IsBusy && Columns.Count > 0);
        ExecuteCommand = new RelayCommand(() => _ = ExecuteAsync(), () => CanExecute);
    }

    /// <summary>データ元の項目名。</summary>
    public ObservableCollection<string> SourceColumns { get; } = [];

    /// <summary>.xlsx のデータ元で選べるシート。</summary>
    public ObservableCollection<string> SourceSheetNames { get; } = [];

    /// <summary>出力する CSV の列(この順に並ぶ)。</summary>
    public ObservableCollection<CsvOutputColumnRowViewModel> Columns { get; } = [];

    /// <summary>このタブの処理設定(レシピ)。</summary>
    public RecipeAreaViewModel Recipes { get; }

    public static IReadOnlyList<string> EncodingOptions { get; }
        = [EncodingUtf8Bom, EncodingUtf8, EncodingShiftJis];

    public static IReadOnlyList<string> QuoteOptions { get; } = [QuoteMinimal, QuoteAll];

    public RelayCommand SelectSourceCommand { get; }

    public RelayCommand LoadColumnsCommand { get; }

    public RelayCommand AddColumnCommand { get; }

    public RelayCommand AddAllSourceColumnsCommand { get; }

    public RelayCommand RemoveColumnCommand { get; }

    public RelayCommand MoveUpCommand { get; }

    public RelayCommand MoveDownCommand { get; }

    public RelayCommand RefreshPreviewCommand { get; }

    public RelayCommand ExecuteCommand { get; }

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
                _recipeSourceSheetName = null;
                OnSettingsChanged();
            }
        }
    }

    public bool IsSheetSelectionEnabled => SourceSheetNames.Count > 0;

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

    public string EncodingDisplay
    {
        get => _encodingDisplay;
        set
        {
            if (SetProperty(ref _encodingDisplay, value))
            {
                OnSettingsChanged();
            }
        }
    }

    public string QuoteDisplay
    {
        get => _quoteDisplay;
        set
        {
            if (SetProperty(ref _quoteDisplay, value))
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

    public CsvOutputColumnRowViewModel? SelectedColumn
    {
        get => _selectedColumn;
        set
        {
            if (SetProperty(ref _selectedColumn, value))
            {
                RemoveColumnCommand.RaiseCanExecuteChanged();
                MoveUpCommand.RaiseCanExecuteChanged();
                MoveDownCommand.RaiseCanExecuteChanged();
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

    public CsvTransformPreview? Preview => _preview;

    public bool HasPreview => _preview is not null && !IsPreviewStale;

    public bool CanExecute => !IsBusy && HasPreview && _preview!.CanExecute;

    public IReadOnlyList<CsvSampleRow> SampleRows => _preview?.SampleRows ?? [];

    public IReadOnlyList<MergeIssue> Issues => _preview is null
        ? []
        : [.. _preview.Issues.OrderByDescending(issue => issue.Severity)];

    public bool HasIssues => Issues.Count > 0;

    public bool HasNoIssues => HasPreview && Issues.Count == 0;

    /// <summary>出力する列の並び(画面で確認するための 1 行表示)。</summary>
    public string ColumnSummaryText => _preview is null
        ? "-"
        : string.Join(" / ", _preview.Columns.Select(
            (column, index) => $"{index + 1} {column.OutputName}"));

    public string RowSummaryText => _preview is null
        ? "-"
        : $"入力 {_preview.SourceColumns.Count:N0} 列 {_preview.SourceRowCount:N0} 行"
            + $" → 出力 {_preview.Columns.Count:N0} 列 {_preview.OutputRowCount:N0} 行";

    public string OutputSummaryText => _preview is null
        ? "-"
        : $"作成 {_preview.OutputFileName} / 文字コード {EncodingName()} / 引用符 {QuoteDisplay}";

    public string IssueSummaryText => _preview is null
        ? "-"
        : $"注意 {_preview.WarningCount:N0} 件 / 実行できない問題 {_preview.BlockCount:N0} 件";

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
        StatusText = "「項目を読み込む」を押してください。";

        if (CsvTransformPlanner.KindOf(path) == SourceFileKind.Xlsx)
        {
            foreach (var name in CsvTransformPlanner.ReadSheetNames(path))
            {
                SourceSheetNames.Add(name);
            }

            if (_recipeSourceSheetName is { } saved)
            {
                var found = SourceSheetNames.Contains(saved);
                _sourceSheetName = found ? saved : null;
                if (!found)
                {
                    StatusText = SourceMappingViewModel.MissingSheetMessage(saved);
                }
            }
            else
            {
                _sourceSheetName = SourceSheetNames.FirstOrDefault();
            }
        }
        else
        {
            _sourceSheetName = null;
        }

        OnPropertyChanged(nameof(SourceSheetName));
        OnPropertyChanged(nameof(IsSheetSelectionEnabled));
        OnSettingsChanged();
    }

    /// <summary>データ元の項目名を読み込む。</summary>
    public void LoadColumns()
    {
        SourceColumns.Clear();
        SourceInfoText = null;

        if (!TryParseHeaderRow(out var headerRow))
        {
            StatusText = "項目名の行は 1 以上の数字で指定してください。";
            OnSettingsChanged();
            return;
        }

        var header = CsvTransformPlanner.ReadColumns(SourceFilePath, SourceSheetName, headerRow);
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

        var notes = new List<string>();
        foreach (var column in Columns)
        {
            if (column.UpdateSourceColumns(SourceColumns) is { } dropped)
            {
                notes.Add($"保存された項目「{dropped}」が今回のデータ元にありません。選び直してください。");
            }
        }

        StatusText = notes.Count > 0
            ? string.Join(" ", notes)
            : "出力する項目を作って、プレビューを更新してください。";
        OnSettingsChanged();
    }

    /// <summary>現在の指定でプレビューを作り直す。</summary>
    public async Task RefreshPreviewAsync()
    {
        if (!TryParseHeaderRow(out _))
        {
            StatusText = "項目名の行は 1 以上の数字で指定してください。";
            return;
        }

        var request = BuildRequest();

        IsBusy = true;
        StatusText = "プレビューを作成しています…(データ元は読み取りのみ)";
        try
        {
            _preview = await Task.Run(() => _planner.CreatePreview(request));
            IsPreviewStale = false;
            StatusText = _preview.CanExecute
                ? "内容を確認して「CSV を作成」を押してください。"
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

    private CsvTransformRequest BuildRequest() => new()
    {
        SourceFilePath = SourceFilePath,
        SourceSheetName = SourceSheetName,
        HeaderRow = TryParseHeaderRow(out var headerRow) ? headerRow : 1,
        Columns = [.. Columns.Select(column => column.ToRequest())],
        Encoding = Encoding(),
        QuoteMode = QuoteMode(),
        OutputSuffix = OutputSuffix,
    };

    private async Task ExecuteAsync()
    {
        if (_preview is null || !CanExecute)
        {
            return;
        }

        IsBusy = true;
        StatusText = "CSV を作成しています…";

        // 実行の途中で画面を触られても、確認済みなのは「実際に流した設定」。
        var executed = BuildRecipe(string.Empty);
        try
        {
            var preview = _preview;
            var result = await Task.Run(() => _transformer.Execute(preview));

            LastRunSucceeded = result.Success;
            ResultText = result.Success
                ? $"{result.Message}\n作成: {string.Join(" / ", result.OutputFileNames)}"
                : result.Message;

            StatusText = result.Success ? "作成が完了しました。" : "作成を実行できませんでした。";

            if (result.Success)
            {
                _lastSuccessfulRecipe = executed;
                IsPreviewStale = true;
                Recipes.ShowSavableAfterRun();
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

    private void AddColumn()
    {
        var column = new CsvOutputColumnRowViewModel(OnSettingsChanged, SourceColumns)
        {
            SourceColumn = SourceColumns.FirstOrDefault() ?? string.Empty,
        };

        Columns.Add(column);
        SelectedColumn = column;
        OnSettingsChanged();
    }

    /// <summary>データ元の項目をすべて、同じ名前・同じ並びで追加する。</summary>
    private void AddAllSourceColumns()
    {
        foreach (var name in SourceColumns)
        {
            Columns.Add(new CsvOutputColumnRowViewModel(OnSettingsChanged, SourceColumns)
            {
                OutputName = name,
                SourceColumn = name,
            });
        }

        SelectedColumn = Columns.LastOrDefault();
        OnSettingsChanged();
    }

    private void RemoveSelectedColumn()
    {
        if (SelectedColumn is not { } column)
        {
            return;
        }

        Columns.Remove(column);
        SelectedColumn = null;
        OnSettingsChanged();
    }

    private bool CanMove(int offset)
    {
        if (IsBusy || SelectedColumn is not { } column)
        {
            return false;
        }

        var index = Columns.IndexOf(column) + offset;
        return index >= 0 && index < Columns.Count;
    }

    /// <summary>選んだ行を 1 つ上/下へ動かす(出力の並び順を決める)。</summary>
    private void Move(int offset)
    {
        if (!CanMove(offset) || SelectedColumn is not { } column)
        {
            return;
        }

        var index = Columns.IndexOf(column);
        Columns.Move(index, index + offset);
        SelectedColumn = column;
        OnSettingsChanged();
    }

    private bool TryParseHeaderRow(out int headerRow)
        => int.TryParse(HeaderRowText, out headerRow) && headerRow >= 1;

    private CsvOutputEncoding Encoding() => EncodingDisplay switch
    {
        EncodingUtf8 => CsvOutputEncoding.Utf8,
        EncodingShiftJis => CsvOutputEncoding.ShiftJis,
        _ => CsvOutputEncoding.Utf8Bom,
    };

    private string EncodingName() => EncodingDisplay;

    private CsvQuoteMode QuoteMode()
        => string.Equals(QuoteDisplay, QuoteAll, StringComparison.Ordinal)
            ? CsvQuoteMode.All
            : CsvQuoteMode.Minimal;

    private static string DisplayOf(CsvOutputEncoding encoding) => encoding switch
    {
        CsvOutputEncoding.Utf8 => EncodingUtf8,
        CsvOutputEncoding.ShiftJis => EncodingShiftJis,
        _ => EncodingUtf8Bom,
    };

    private static string DisplayOf(CsvQuoteMode mode)
        => mode == CsvQuoteMode.All ? QuoteAll : QuoteMinimal;

    RecipeType IRecipeHost.RecipeType => RecipeType.CsvTransform;

    string? IRecipeHost.RecipeSaveBlockedReason
        => RecipeSaveGuard.ReasonFor(_preview, IsPreviewStale, MatchesLastSuccessfulRun);

    private bool MatchesLastSuccessfulRun
        => RecipeConfiguration.AreSame(_lastSuccessfulRecipe, BuildRecipe(string.Empty));

    SavedRecipe IRecipeHost.CreateRecipe(string name) => BuildRecipe(name);

    private SavedRecipe BuildRecipe(string name)
    {
        var kind = CsvTransformPlanner.KindOf(SourceFilePath) ?? SourceFileKind.Csv;

        return new SavedRecipe
        {
            Name = name,
            Type = RecipeType.CsvTransform,
            CsvTransform = new CsvTransformRecipe
            {
                SourceFileKind = kind,
                SourceSheetName = kind == SourceFileKind.Xlsx ? SourceSheetName : null,
                HeaderRow = TryParseHeaderRow(out var headerRow) ? headerRow : 1,
                OutputColumns = [.. Columns.Select(column => new RecipeCsvColumn
                {
                    OutputName = column.OutputName,
                    ValueSourceKind = column.Kind,
                    SourceColumn = column.Kind == CsvValueSourceKind.SourceColumn
                        ? column.SourceColumn
                        : null,
                    FixedValue = column.Kind == CsvValueSourceKind.FixedText
                        ? column.FixedValue
                        : null,
                })],
                Encoding = Encoding(),
                QuoteMode = QuoteMode(),
                OutputSuffix = OutputSuffix,
            },
        };
    }

    /// <summary>レシピを画面へ戻す。今回のファイルの選択はそのまま残し、プレビューはやり直させる。</summary>
    IReadOnlyList<string> IRecipeHost.ApplyRecipe(SavedRecipe recipe)
    {
        var payload = recipe.CsvTransform!;
        var notes = new List<string>();

        _recipeSourceSheetName = payload.SourceFileKind == SourceFileKind.Xlsx
            ? payload.SourceSheetName
            : null;

        if (SourceFilePath.Length > 0
            && CsvTransformPlanner.KindOf(SourceFilePath) != payload.SourceFileKind)
        {
            notes.Add(SourceMappingViewModel.SourceKindMismatchMessage(payload.SourceFileKind));
        }

        HeaderRowText = payload.HeaderRow.ToString(CultureInfo.InvariantCulture);
        EncodingDisplay = DisplayOf(payload.Encoding);
        QuoteDisplay = DisplayOf(payload.QuoteMode);
        OutputSuffix = payload.OutputSuffix;

        Columns.Clear();
        foreach (var column in payload.OutputColumns)
        {
            Columns.Add(new CsvOutputColumnRowViewModel(OnSettingsChanged, SourceColumns)
            {
                OutputName = column.OutputName,
                KindDisplay = CsvOutputColumnRowViewModel.DisplayOf(column.ValueSourceKind),
                SourceColumn = column.SourceColumn ?? string.Empty,
                FixedValue = column.FixedValue ?? string.Empty,
            });
        }

        SelectedColumn = null;

        if (SourceSheetNames.Count > 0 && _recipeSourceSheetName is { } sheet)
        {
            var found = SourceSheetNames.Contains(sheet);
            _sourceSheetName = found ? sheet : null;
            OnPropertyChanged(nameof(SourceSheetName));

            if (!found)
            {
                notes.Add(SourceMappingViewModel.MissingSheetMessage(sheet));
            }
        }

        if (SourceColumns.Count > 0)
        {
            foreach (var column in Columns)
            {
                if (column.UpdateSourceColumns(SourceColumns) is { } dropped)
                {
                    notes.Add($"保存された項目「{dropped}」が今回のデータ元にありません。選び直してください。");
                }
            }
        }

        OnSettingsChanged();
        return notes;
    }

    /// <summary>指定が変わったらプレビューを無効にする(古い内容のまま実行させない)。</summary>
    private void OnSettingsChanged()
    {
        ResultText = null;
        IsPreviewStale = true;
        RaiseCommandStates();
    }

    private void RaisePreviewProperties()
    {
        OnPropertyChanged(nameof(Preview));
        OnPropertyChanged(nameof(HasPreview));
        OnPropertyChanged(nameof(CanExecute));
        OnPropertyChanged(nameof(ColumnSummaryText));
        OnPropertyChanged(nameof(RowSummaryText));
        OnPropertyChanged(nameof(OutputSummaryText));
        OnPropertyChanged(nameof(IssueSummaryText));
        OnPropertyChanged(nameof(SampleRows));
        OnPropertyChanged(nameof(Issues));
        OnPropertyChanged(nameof(HasIssues));
        OnPropertyChanged(nameof(HasNoIssues));
        RaiseCommandStates();
    }

    private void RaiseCommandStates()
    {
        SelectSourceCommand.RaiseCanExecuteChanged();
        LoadColumnsCommand.RaiseCanExecuteChanged();
        AddColumnCommand.RaiseCanExecuteChanged();
        AddAllSourceColumnsCommand.RaiseCanExecuteChanged();
        RemoveColumnCommand.RaiseCanExecuteChanged();
        MoveUpCommand.RaiseCanExecuteChanged();
        MoveDownCommand.RaiseCanExecuteChanged();
        RefreshPreviewCommand.RaiseCanExecuteChanged();
        ExecuteCommand.RaiseCanExecuteChanged();
    }
}
