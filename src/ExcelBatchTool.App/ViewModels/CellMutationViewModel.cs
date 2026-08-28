using System.Collections.ObjectModel;
using ExcelBatchTool.Core;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Mutation;
using ExcelBatchTool.Core.Recipes;

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

/// <summary>入力セットの 1 行(セル位置・種類・新しい値)。</summary>
public sealed class MutationOperationViewModel : ObservableObject
{
    public const string KindText = "文字";
    public const string KindNumber = "数値";
    public const string KindBlank = "空欄";

    /// <summary>「種類」列の選択肢。</summary>
    public static IReadOnlyList<string> KindOptions { get; } = [KindText, KindNumber, KindBlank];

    /// <summary>保存した種類を画面の表示へ戻す。</summary>
    internal static string DisplayOf(CellWriteKind kind) => kind switch
    {
        CellWriteKind.Number => KindNumber,
        CellWriteKind.Blank => KindBlank,
        _ => KindText,
    };

    private readonly Action _onChanged;
    private string _cellReference = string.Empty;
    private string _kindDisplay = KindText;
    private string _valueText = string.Empty;

    public MutationOperationViewModel(Action onChanged) => _onChanged = onChanged;

    /// <summary>変更する位置(A1 形式の単一セル)。</summary>
    public string CellReference
    {
        get => _cellReference;
        set
        {
            if (SetProperty(ref _cellReference, value))
            {
                _onChanged();
            }
        }
    }

    /// <summary>値の種類(文字 / 数値 / 空欄)。</summary>
    public string KindDisplay
    {
        get => _kindDisplay;
        set
        {
            if (SetProperty(ref _kindDisplay, value))
            {
                OnPropertyChanged(nameof(IsValueEnabled));
                _onChanged();
            }
        }
    }

    /// <summary>新しい値。「空欄」の行では使わない。</summary>
    public string ValueText
    {
        get => _valueText;
        set
        {
            if (SetProperty(ref _valueText, value))
            {
                _onChanged();
            }
        }
    }

    public bool IsValueEnabled => !string.Equals(_kindDisplay, KindBlank, StringComparison.Ordinal);

    internal CellWriteKind Kind => _kindDisplay switch
    {
        KindNumber => CellWriteKind.Number,
        KindBlank => CellWriteKind.Blank,
        _ => CellWriteKind.Text,
    };

    internal CellMutationOperationRequest ToRequest() => new()
    {
        CellReference = CellReference,
        WriteKind = Kind,
        TextValue = Kind == CellWriteKind.Text ? ValueText : null,
        NumberText = Kind == CellWriteKind.Number ? ValueText : null,
    };
}

/// <summary>「セルをまとめて変更」(Phase 2A / 2B)の ViewModel。</summary>
public sealed class CellMutationViewModel : ObservableObject, IRecipeHost
{
    private readonly CellMutationPlanner _planner = new();
    private readonly CellMutator _mutator = new();
    private readonly Func<string?> _readClipboardText;

    private bool _isBusy;
    private bool _isPreviewStale = true;
    private CellMutationPreview? _preview;
    private string _outputSuffix = CellMutationDefaults.OutputSuffix;
    private string _statusText = "変更するシートとセルを指定して、プレビューを更新してください。";
    private string? _resultText;
    private bool _lastRunSucceeded;
    private MutationOperationViewModel? _selectedOperation;

    public CellMutationViewModel()
        : this(ReadClipboardText)
    {
    }

    /// <summary>レシピの置き場所だけを指定する(画面から使うときの経路)。</summary>
    internal CellMutationViewModel(RecipeStore recipeStore)
        : this(ReadClipboardText, recipeStore)
    {
    }

    /// <summary>テスト用: クリップボードの読み取り・レシピの置き場所を差し替えられるようにする。</summary>
    internal CellMutationViewModel(
        Func<string?> readClipboardText,
        RecipeStore? recipeStore = null,
        Func<string, bool>? confirm = null)
    {
        _readClipboardText = readClipboardText;
        Recipes = new RecipeAreaViewModel(
            this, recipeStore ?? new RecipeStore(), confirm ?? RecipeSaveGuard.AskInDialog);
        RefreshPreviewCommand = new RelayCommand(
            () => _ = RefreshPreviewAsync(),
            () => !IsBusy && SelectedSheetCount > 0);
        ExecuteCommand = new RelayCommand(() => _ = ExecuteAsync(), () => CanExecute);
        AddOperationCommand = new RelayCommand(AddOperation, () => !IsBusy);
        RemoveOperationCommand = new RelayCommand(
            RemoveSelectedOperation, () => !IsBusy && SelectedOperation is not null);
        PasteOperationsCommand = new RelayCommand(PasteOperations, () => !IsBusy);

        // 初期状態から 1 行編集できるようにしておく。
        Operations.Add(new MutationOperationViewModel(OnSettingsChanged));
    }

    public ObservableCollection<MutationWorkbookViewModel> Workbooks { get; } = [];

    /// <summary>入力セット(変更するセルの一覧)。</summary>
    public ObservableCollection<MutationOperationViewModel> Operations { get; } = [];

    /// <summary>このタブの処理設定(レシピ)。</summary>
    public RecipeAreaViewModel Recipes { get; }

    public RelayCommand RefreshPreviewCommand { get; }

    public RelayCommand ExecuteCommand { get; }

    public RelayCommand AddOperationCommand { get; }

    public RelayCommand RemoveOperationCommand { get; }

    public RelayCommand PasteOperationsCommand { get; }

    public bool HasWorkbooks => Workbooks.Count > 0;

    public int SelectedSheetCount => Workbooks.Sum(workbook => workbook.Sheets.Count(sheet => sheet.IsSelected));

    /// <summary>一覧で選択中の行(削除対象)。</summary>
    public MutationOperationViewModel? SelectedOperation
    {
        get => _selectedOperation;
        set
        {
            if (SetProperty(ref _selectedOperation, value))
            {
                RemoveOperationCommand.RaiseCanExecuteChanged();
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
        : $"{_preview.Targets.Select(target => (target.FilePath, target.SheetName)).Distinct().Count():N0} シート"
            + $" × {_preview.Targets.Select(target => target.CellReference).Distinct().Count():N0} セル";

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
        Operations = [.. Operations.Select(operation => operation.ToRequest())],
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

    private void AddOperation()
    {
        var operation = new MutationOperationViewModel(OnSettingsChanged);
        Operations.Add(operation);
        SelectedOperation = operation;
        OnSettingsChanged();
    }

    private void RemoveSelectedOperation()
    {
        if (SelectedOperation is not { } operation)
        {
            return;
        }

        Operations.Remove(operation);
        SelectedOperation = null;
        OnSettingsChanged();
    }

    /// <summary>
    /// 表計算からコピーした「セル TAB 種類 TAB 値」の行を入力セットへ追加する。
    /// 1 行でも読み取れなければ何も追加しない(一部だけ貼り付けない)。
    /// </summary>
    private void PasteOperations()
    {
        string? text;
        try
        {
            text = _readClipboardText();
        }
        catch (Exception)
        {
            StatusText = "貼り付ける内容を読み取れませんでした。";
            return;
        }

        if (!TryParsePastedOperations(text, out var rows, out var error))
        {
            StatusText = error!;
            return;
        }

        foreach (var (cell, kind, value) in rows)
        {
            Operations.Add(new MutationOperationViewModel(OnSettingsChanged)
            {
                CellReference = cell,
                KindDisplay = kind,
                ValueText = value,
            });
        }

        OnSettingsChanged();
        StatusText = $"{rows.Count:N0} 行を追加しました。「プレビューを更新」を押してください。";
    }

    /// <summary>
    /// 貼り付けテキストを解釈する。列はタブ区切りで「セル、種類(文字/数値/空欄)、値」。
    /// 「空欄」の行は値の列が無くてもよい。
    /// </summary>
    internal static bool TryParsePastedOperations(
        string? text,
        out List<(string Cell, string Kind, string Value)> rows,
        out string? error)
    {
        rows = [];
        error = null;

        if (string.IsNullOrWhiteSpace(text))
        {
            error = "貼り付ける内容がありません。表計算で「セル・種類・値」の 3 列をコピーしてください。";
            return false;
        }

        var lines = text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
        for (var index = 0; index < lines.Length; index++)
        {
            var line = lines[index].TrimEnd('\r');
            if (string.IsNullOrWhiteSpace(line))
            {
                continue; // 末尾などの空行は読み飛ばす。
            }

            var columns = line.Split('\t');
            var lineNumber = index + 1;

            if (columns.Length is < 2 or > 3)
            {
                error = $"{lineNumber} 行目を読み取れません。"
                    + "「セル・種類・値」のタブ区切り 3 列で貼り付けてください。";
                return Fail(rows);
            }

            var cell = columns[0].Trim();
            var kind = columns[1].Trim();
            var value = columns.Length >= 3 ? columns[2] : string.Empty;

            if (cell.Length == 0)
            {
                error = $"{lineNumber} 行目のセルの位置が空です。";
                return Fail(rows);
            }

            if (!MutationOperationViewModel.KindOptions.Contains(kind))
            {
                error = $"{lineNumber} 行目の種類「{kind}」を読み取れません"
                    + "(使えるのは 文字 / 数値 / 空欄 のみです)。";
                return Fail(rows);
            }

            if (string.Equals(kind, MutationOperationViewModel.KindBlank, StringComparison.Ordinal)
                && value.Length > 0)
            {
                error = $"{lineNumber} 行目は「空欄」なのに値が指定されています。値の列を空にしてください。";
                return Fail(rows);
            }

            rows.Add((cell, kind, value));
        }

        if (rows.Count == 0)
        {
            error = "貼り付ける内容がありません。表計算で「セル・種類・値」の 3 列をコピーしてください。";
            return false;
        }

        return true;

        static bool Fail(List<(string, string, string)> rows)
        {
            rows.Clear(); // 一部だけ追加しない。
            return false;
        }
    }

    RecipeType IRecipeHost.RecipeType => RecipeType.CellInputSet;

    string? IRecipeHost.RecipeSaveBlockedReason => RecipeSaveGuard.ReasonFor(_preview, IsPreviewStale);

    /// <summary>今の入力セットをレシピにする。対象のファイル・シートは含めない。</summary>
    SavedRecipe IRecipeHost.CreateRecipe(string name) => new()
    {
        Name = name,
        Type = RecipeType.CellInputSet,
        CellInputSet = new CellInputSetRecipe
        {
            Operations = [.. Operations.Select(operation => new RecipeOperation
            {
                Cell = operation.CellReference.Trim(),
                Kind = operation.Kind,
                Value = operation.Kind == CellWriteKind.Blank ? null : operation.ValueText,
            })],
            OutputSuffix = OutputSuffix,
        },
    };

    /// <summary>レシピを画面へ戻す。対象ファイルの選択はそのまま残し、プレビューはやり直させる。</summary>
    IReadOnlyList<string> IRecipeHost.ApplyRecipe(SavedRecipe recipe)
    {
        var payload = recipe.CellInputSet!;

        Operations.Clear();
        foreach (var operation in payload.Operations)
        {
            Operations.Add(new MutationOperationViewModel(OnSettingsChanged)
            {
                CellReference = operation.Cell,
                KindDisplay = MutationOperationViewModel.DisplayOf(operation.Kind),
                ValueText = operation.Value ?? string.Empty,
            });
        }

        SelectedOperation = null;
        OutputSuffix = payload.OutputSuffix;

        // 値が同じでも必ずプレビューをやり直させる。
        OnSettingsChanged();
        return [];
    }

    /// <summary>クリップボードの文字列を読む(UI からの実行時のみ使う)。</summary>
    private static string? ReadClipboardText()
        => System.Windows.Clipboard.ContainsText() ? System.Windows.Clipboard.GetText() : null;

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
        AddOperationCommand.RaiseCanExecuteChanged();
        RemoveOperationCommand.RaiseCanExecuteChanged();
        PasteOperationsCommand.RaiseCanExecuteChanged();
    }
}
