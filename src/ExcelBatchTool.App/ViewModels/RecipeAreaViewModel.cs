using System.Collections.ObjectModel;
using ExcelBatchTool.Core.Mutation;
using ExcelBatchTool.Core.Recipes;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>
/// 「壊れた設定」を保存しにくくするための共通判定。
/// 最新のプレビューで問題が無いか、直前の正常終了と同じ設定のときだけ保存できる。
/// </summary>
internal static class RecipeSaveGuard
{
    public const string StaleReason = "先にプレビューを更新してください(今の設定で確かめてから保存します)。";

    public const string BlockedReason = "実行できない問題があるため、この設定は保存できません。";

    /// <param name="matchesLastSuccessfulRun">
    /// 今の設定が、最後に正常終了した実行で使った設定とまったく同じか。
    /// 作ったばかりの出力があるとプレビューは同名衝突で止まるが、その実行で
    /// すでに確かめ済みの設定なので、保存だけは認める。
    /// </param>
    public static string? ReasonFor(
        CellMutationPreview? preview, bool isStale, bool matchesLastSuccessfulRun)
    {
        if (matchesLastSuccessfulRun)
        {
            return null;
        }

        if (preview is null || isStale)
        {
            return StaleReason;
        }

        // 今回は変えるところが無い(全部が現在の値と同じ)だけなら、指定そのものは
        // 正しいので保存できる。次回のファイルでは変わるかもしれない。
        return !preview.HasBlocks || preview.IsBlockedOnlyByHavingNothingToChange
            ? null
            : BlockedReason;
    }

    /// <summary>CSV 変換の判定。考え方は上と同じ。</summary>
    public static string? ReasonFor(
        Core.CsvTransform.CsvTransformPreview? preview, bool isStale, bool matchesLastSuccessfulRun)
    {
        if (matchesLastSuccessfulRun)
        {
            return null;
        }

        if (preview is null || isStale)
        {
            return StaleReason;
        }

        return preview.HasBlocks ? BlockedReason : null;
    }

    /// <summary>更新・削除の確認(画面から使うときのみダイアログを出す)。</summary>
    public static bool AskInDialog(string message)
        => System.Windows.MessageBox.Show(
            message,
            "確認",
            System.Windows.MessageBoxButton.OKCancel,
            System.Windows.MessageBoxImage.Question) == System.Windows.MessageBoxResult.OK;
}

/// <summary>
/// レシピ領域から見た「今のタブ」。
/// 設定の取り出しと反映だけを受け持ち、ファイルの読み書きには関わらない。
/// </summary>
internal interface IRecipeHost
{
    /// <summary>このタブで扱えるレシピの種類(他の種類は一覧に出さない)。</summary>
    RecipeType RecipeType { get; }

    /// <summary>今の設定を保存できない理由。保存できるときは null。</summary>
    string? RecipeSaveBlockedReason { get; }

    /// <summary>今の設定からレシピを作る(今回使うファイルは含めない)。</summary>
    SavedRecipe CreateRecipe(string name);

    /// <summary>レシピを画面へ反映する。利用者に確かめてほしいことを返す。</summary>
    IReadOnlyList<string> ApplyRecipe(SavedRecipe recipe);
}

/// <summary>一覧に出す 1 件のレシピ(画面には名前だけを見せる)。</summary>
public sealed class RecipeItemViewModel(string id, string name)
{
    internal string Id { get; } = id;

    public string Name { get; } = name;
}

/// <summary>
/// タブ 4〜6 の上部に置く「処理設定(レシピ)」の領域。
/// 読み込んでも自動では何も実行せず、必ずプレビューからやり直させる。
/// </summary>
public sealed class RecipeAreaViewModel : ObservableObject
{
    private readonly IRecipeHost _host;
    private readonly RecipeStore _store;
    private readonly Func<string, bool> _confirm;

    private RecipeItemViewModel? _selectedRecipe;
    private string _nameText = string.Empty;
    private string? _messageText;
    private bool _isMessageError;

    internal RecipeAreaViewModel(IRecipeHost host, RecipeStore store, Func<string, bool> confirm)
    {
        _host = host;
        _store = store;
        _confirm = confirm;

        LoadCommand = new RelayCommand(LoadSelected, () => SelectedRecipe is not null);
        SaveCommand = new RelayCommand(SaveNew);
        UpdateCommand = new RelayCommand(UpdateSelected, () => SelectedRecipe is not null);
        DeleteCommand = new RelayCommand(DeleteSelected, () => SelectedRecipe is not null);
    }

    /// <summary>このタブで使えるレシピ(名前順)。</summary>
    public ObservableCollection<RecipeItemViewModel> Recipes { get; } = [];

    public RelayCommand LoadCommand { get; }

    public RelayCommand SaveCommand { get; }

    public RelayCommand UpdateCommand { get; }

    public RelayCommand DeleteCommand { get; }

    public RecipeItemViewModel? SelectedRecipe
    {
        get => _selectedRecipe;
        set
        {
            if (SetProperty(ref _selectedRecipe, value))
            {
                LoadCommand.RaiseCanExecuteChanged();
                UpdateCommand.RaiseCanExecuteChanged();
                DeleteCommand.RaiseCanExecuteChanged();
            }
        }
    }

    /// <summary>新しく保存するときの名前。</summary>
    public string NameText
    {
        get => _nameText;
        set => SetProperty(ref _nameText, value);
    }

    public string? MessageText
    {
        get => _messageText;
        private set
        {
            if (SetProperty(ref _messageText, value))
            {
                OnPropertyChanged(nameof(HasMessage));
            }
        }
    }

    public bool HasMessage => !string.IsNullOrEmpty(MessageText);

    /// <summary>直前の操作が問題だったか(表示色を分けるために使う)。</summary>
    public bool IsMessageError
    {
        get => _isMessageError;
        private set => SetProperty(ref _isMessageError, value);
    }

    /// <summary>
    /// 実行が正常終了した直後に、その設定を残せることを知らせる。
    /// 保存そのものは行わない(利用者が名前を入れて押したときだけ保存する)。
    /// </summary>
    internal void ShowSavableAfterRun()
        => SetMessage(
            "今回使った設定は、処理設定として保存できます。"
                + "名前を入れて「現在の設定を保存」を押してください。",
            isError: false);

    /// <summary>保存済みの一覧をファイルから読み直す。</summary>
    public void Reload()
    {
        var loaded = _store.Load();

        if (!loaded.IsSuccess)
        {
            Fill([]);
            SetMessage(
                loaded.HasBackup
                    ? $"{loaded.Error} 1 つ前の内容の控え({RecipeStore.FileName}.bak)は残っています。"
                    : loaded.Error!,
                isError: true);
            return;
        }

        Fill(loaded.Recipes);
    }

    private void LoadSelected()
    {
        if (SelectedRecipe is not { } selected)
        {
            return;
        }

        // 別の操作で変わっているかもしれないので、読み込む直前にファイルを見る。
        var loaded = _store.Load();
        if (!loaded.IsSuccess)
        {
            SetMessage(loaded.Error!, isError: true);
            return;
        }

        var recipe = loaded.Recipes.FirstOrDefault(
            item => item.Id == selected.Id && item.Type == _host.RecipeType);

        if (recipe is null)
        {
            Fill(loaded.Recipes);
            SetMessage("選んだレシピが見つかりません。一覧を選び直してください。", isError: true);
            return;
        }

        var notes = _host.ApplyRecipe(recipe);

        SetMessage(
            $"「{recipe.Name}」を読み込みました。今回使うファイルを確認して、プレビューを更新してください。"
                + string.Concat(notes.Select(note => "\n" + note)),
            isError: notes.Count > 0);
    }

    private void SaveNew()
    {
        if (_host.RecipeSaveBlockedReason is { } blocked)
        {
            SetMessage(blocked, isError: true);
            return;
        }

        if (!RecipeName.TryNormalize(NameText, out var name, out var nameError))
        {
            SetMessage(nameError!, isError: true);
            return;
        }

        var result = _store.Add(_host.CreateRecipe(name));
        if (!result.IsSuccess)
        {
            SetMessage(result.Error!, isError: true);
            return;
        }

        Fill(result.Recipes);
        SelectById(result.Recipe!.Id);
        NameText = string.Empty;
        SetMessage($"「{name}」を保存しました。", isError: false);
    }

    private void UpdateSelected()
    {
        if (SelectedRecipe is not { } selected)
        {
            return;
        }

        if (_host.RecipeSaveBlockedReason is { } blocked)
        {
            SetMessage(blocked, isError: true);
            return;
        }

        if (!_confirm($"「{selected.Name}」の内容を、今の設定で置き換えます。よろしいですか?"))
        {
            SetMessage("更新をやめました。", isError: false);
            return;
        }

        var result = _store.Update(selected.Id, _host.CreateRecipe(selected.Name));
        if (!result.IsSuccess)
        {
            SetMessage(result.Error!, isError: true);
            return;
        }

        Fill(result.Recipes);
        SelectById(result.Recipe!.Id);
        SetMessage($"「{selected.Name}」を今の設定で更新しました。", isError: false);
    }

    private void DeleteSelected()
    {
        if (SelectedRecipe is not { } selected)
        {
            return;
        }

        if (!_confirm($"「{selected.Name}」を削除します。よろしいですか?"))
        {
            SetMessage("削除をやめました。", isError: false);
            return;
        }

        var result = _store.Delete(selected.Id);
        if (!result.IsSuccess)
        {
            SetMessage(result.Error!, isError: true);
            return;
        }

        Fill(result.Recipes);
        SetMessage($"「{selected.Name}」を削除しました。", isError: false);
    }

    /// <summary>このタブで使える種類だけを一覧に並べる。</summary>
    private void Fill(IReadOnlyList<SavedRecipe> recipes)
    {
        var selectedId = SelectedRecipe?.Id;

        Recipes.Clear();
        foreach (var recipe in recipes.Where(item => item.Type == _host.RecipeType))
        {
            Recipes.Add(new RecipeItemViewModel(recipe.Id, recipe.Name));
        }

        SelectedRecipe = selectedId is null
            ? null
            : Recipes.FirstOrDefault(item => item.Id == selectedId);
    }

    private void SelectById(string id)
        => SelectedRecipe = Recipes.FirstOrDefault(item => item.Id == id);

    private void SetMessage(string text, bool isError)
    {
        IsMessageError = isError;
        MessageText = text;
    }
}
