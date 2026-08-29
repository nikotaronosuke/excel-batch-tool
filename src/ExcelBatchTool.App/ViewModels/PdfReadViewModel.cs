using System.Collections.ObjectModel;
using System.IO;
using ExcelBatchTool.Core.CsvTransform;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Ocr;
using ExcelBatchTool.Core.Pdf;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>プレビューに出す 1 行(文字 PDF は ページ/行/内容、表 PDF は列を連結)。</summary>
public sealed record PdfPreviewRow(string First, string Second, string Text);

/// <summary>「8. PDF を読み取る」(Phase 2F-A / 2F-B1)の ViewModel。</summary>
public sealed class PdfReadViewModel : ObservableObject
{
    private const string FormatXlsx = "Excel (.xlsx)";
    private const string FormatCsv = "CSV (.csv)";

    private const string EncodingUtf8Bom = "UTF-8(BOM あり)";
    private const string EncodingUtf8 = "UTF-8(BOM なし)";
    private const string EncodingShiftJis = "Shift_JIS";

    private const string QuoteMinimal = "必要なときだけ";
    private const string QuoteAll = "すべての項目";

    private readonly PdfReadPlanner _planner = new();
    private readonly PdfReader _reader = new();
    private readonly PdfScanReader _scanReader = new();
    private readonly Func<string?> _pickPdfFile;
    private readonly Func<OcrPackStatus> _inspectPack;
    private readonly Func<OcrPackStatus, IOcrEngine> _loadEngine;

    private string _sourceFilePath = string.Empty;
    private string _formatDisplay = FormatXlsx;
    private string _encodingDisplay = EncodingUtf8Bom;
    private string _quoteDisplay = QuoteMinimal;
    private string _outputSuffix = PdfReadDefaults.OutputSuffix;

    private bool _isBusy;
    private bool _isPreviewStale = true;
    private PdfReadPreview? _preview;
    private string _statusText = "PDF を選んで「PDF を解析」を押してください。";
    private string? _resultText;
    private bool _lastRunSucceeded;

    private OcrPackStatus? _packStatus;
    private OcrDocumentReading? _reading;
    private CancellationTokenSource? _ocrCancellation;
    private string _progressText = string.Empty;
    private bool _showOnlyNeedsReview = true;

    public PdfReadViewModel()
        : this(() => null)
    {
    }

    /// <summary>テスト用: ファイル選択と OCR Pack を差し替えられるようにする。</summary>
    internal PdfReadViewModel(
        Func<string?> pickPdfFile,
        Func<OcrPackStatus>? inspectPack = null,
        Func<OcrPackStatus, IOcrEngine>? loadEngine = null)
    {
        _pickPdfFile = pickPdfFile;
        _inspectPack = inspectPack ?? (() => OcrPack.Inspect());
        _loadEngine = loadEngine ?? OcrPack.Load;

        SelectSourceCommand = new RelayCommand(SelectSource, () => !IsBusy);
        AnalyzeCommand = new RelayCommand(
            () => _ = AnalyzeAsync(), () => !IsBusy && SourceFilePath.Length > 0);
        RunOcrCommand = new RelayCommand(() => _ = RunOcrAsync(), () => CanRunOcr);
        CancelOcrCommand = new RelayCommand(CancelOcr, () => IsBusy && _ocrCancellation is not null);
        ConfirmSelectedCommand = new RelayCommand(ConfirmSelected, () => SelectedItem is not null);
        ConfirmAllShownCommand = new RelayCommand(
            ConfirmAllShown, () => ReviewItems.Count > 0 && !IsBusy);
        ExecuteCommand = new RelayCommand(() => _ = ExecuteAsync(), () => CanExecute);
    }

    public static IReadOnlyList<string> FormatOptions { get; } = [FormatXlsx, FormatCsv];

    public static IReadOnlyList<string> EncodingOptions { get; }
        = [EncodingUtf8Bom, EncodingUtf8, EncodingShiftJis];

    public static IReadOnlyList<string> QuoteOptions { get; } = [QuoteMinimal, QuoteAll];

    /// <summary>プレビュー(先頭のみ。全行は出力時に使う)。</summary>
    public ObservableCollection<PdfPreviewRow> PreviewRows { get; } = [];

    public RelayCommand SelectSourceCommand { get; }

    public RelayCommand AnalyzeCommand { get; }

    public RelayCommand RunOcrCommand { get; }

    public RelayCommand CancelOcrCommand { get; }

    public RelayCommand ConfirmSelectedCommand { get; }

    public RelayCommand ConfirmAllShownCommand { get; }

    public RelayCommand ExecuteCommand { get; }

    /// <summary>確認・修正の一覧。「要確認だけ」に絞れる。</summary>
    public ObservableCollection<OcrReviewRow> ReviewItems { get; } = [];

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

    public string FormatDisplay
    {
        get => _formatDisplay;
        set
        {
            if (SetProperty(ref _formatDisplay, value))
            {
                OnPropertyChanged(nameof(IsCsv));
                OnSettingsChanged();
            }
        }
    }

    public bool IsCsv => string.Equals(_formatDisplay, FormatCsv, StringComparison.Ordinal);

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

    public bool LastRunSucceeded
    {
        get => _lastRunSucceeded;
        private set => SetProperty(ref _lastRunSucceeded, value);
    }

    public PdfReadPreview? Preview => _preview;

    public bool HasPreview => _preview is not null && !IsPreviewStale;

    public bool CanExecute => !IsBusy && HasPreview && _preview!.CanExecute;

    public bool CanRunOcr => !IsBusy
        && HasPreview
        && _preview!.Stage == PdfReadStage.NeedsOcr
        && _packStatus is { IsUsable: true };

    /// <summary>OCR が必要な PDF か(画面の確認欄を出すかどうか)。</summary>
    public bool NeedsOcr => _preview?.Stage == PdfReadStage.NeedsOcr;

    public bool HasReading => _reading is not null;

    public string ProgressText
    {
        get => _progressText;
        private set
        {
            if (SetProperty(ref _progressText, value))
            {
                OnPropertyChanged(nameof(HasProgress));
            }
        }
    }

    public bool HasProgress => ProgressText.Length > 0;

    private OcrReviewRow? _selectedItem;

    public OcrReviewRow? SelectedItem
    {
        get => _selectedItem;
        set
        {
            if (SetProperty(ref _selectedItem, value))
            {
                OnPropertyChanged(nameof(SelectedEngineText));
                OnPropertyChanged(nameof(HasSelectedItem));
                ConfirmSelectedCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public bool HasSelectedItem => SelectedItem is not null;

    public string SelectedEngineText => SelectedItem?.EngineText ?? "-";

    /// <summary>「要確認だけ」に絞る。120 ページを 1 件ずつ見せない。</summary>
    public bool ShowOnlyNeedsReview
    {
        get => _showOnlyNeedsReview;
        set
        {
            if (SetProperty(ref _showOnlyNeedsReview, value))
            {
                FillReviewItems();
            }
        }
    }

    public string ReviewSummaryText => _reading is null
        ? "-"
        : $"自動確定 {_reading.AutoAcceptedCount:N0} 件 / "
            + $"要確認 {_reading.NeedsReviewCount:N0} 件 / "
            + $"読取不能 {_reading.UnreadableCount:N0} 件 / "
            + $"確認済み {_reading.UserConfirmedCount:N0} 件"
            + (_reading.UserEditedCount > 0 ? $"(うち修正 {_reading.UserEditedCount:N0} 件)" : string.Empty);

    public string OcrPackText => _packStatus is null
        ? "-"
        : _packStatus.IsUsable
            ? "利用できます"
            : _packStatus.Message;

    public IReadOnlyList<MergeIssue> Issues => _preview is null
        ? []
        : [.. _preview.Issues.OrderByDescending(issue => issue.Severity)];

    public bool HasIssues => Issues.Count > 0;

    /// <summary>自動判定の結果(利用者に「これは表ですか」と聞かない)。</summary>
    public string KindText => _preview is null ? "-" : _preview.KindDisplay;

    public string SummaryText
    {
        get
        {
            if (_preview is null)
            {
                return "-";
            }

            if (_preview.Stage == PdfReadStage.NeedsOcr)
            {
                return $"{_preview.PageCount:N0} ページ / "
                    + $"OCR が必要 {_preview.OcrPageNumbers.Count:N0} ページ";
            }

            var rows = _preview.Kind == PdfDocumentKind.Table
                ? $"{Math.Max(_preview.TableRows.Count - 1, 0):N0} 行 × "
                    + $"{(_preview.TableRows.Count == 0 ? 0 : _preview.TableRows.Max(row => row.Length)):N0} 列"
                : $"{_preview.Lines.Count:N0} 行";

            return $"{_preview.PageCount:N0} ページ / 取り出し {rows}";
        }
    }

    public string OutputSummaryText => _preview is null || _preview.OutputFileName.Length == 0
        ? "-"
        : $"作成 {_preview.OutputFileName}"
            + (IsCsv ? $" / 文字コード {EncodingDisplay} / 引用符 {QuoteDisplay}" : string.Empty);

    public string IssueSummaryText => _preview is null
        ? "-"
        : $"注意 {_preview.WarningCount:N0} 件 / 実行できない問題 {_preview.BlockCount:N0} 件";

    /// <summary>プレビューの列見出し(文字 PDF と表 PDF で変える)。</summary>
    public string FirstColumnHeader => _preview?.Kind == PdfDocumentKind.Table ? "1 列目" : "ページ";

    public string SecondColumnHeader => _preview?.Kind == PdfDocumentKind.Table ? "2 列目" : "行";

    public string TextColumnHeader => _preview?.Kind == PdfDocumentKind.Table ? "以降の列" : "内容";

    private void SelectSource()
    {
        if (_pickPdfFile() is not { } path || path.Length == 0)
        {
            return;
        }

        SetSourceFile(path);
    }

    public void SetSourceFile(string path)
    {
        SourceFilePath = path;
        StatusText = "「PDF を解析」を押してください。";
        OnSettingsChanged();
    }

    /// <summary>PDF を解析して、種類の判定とプレビューを作る。</summary>
    public async Task AnalyzeAsync()
    {
        var request = BuildRequest();

        IsBusy = true;
        StatusText = "PDF を解析しています…(元の PDF は読み取りのみ)";
        try
        {
            _reading = null;
            ReviewItems.Clear();

            // OCR Pack は「スキャンページがあったときだけ」必要になる。
            // 無くてもここで失敗させない(文字情報のある PDF はそのまま読める)。
            _packStatus = await Task.Run(SafeInspectPack);
            var pack = _packStatus;

            _preview = await Task.Run(() => _planner.CreatePreview(request, pack));
            IsPreviewStale = false;
            FillPreviewRows();

            StatusText = _preview.Stage switch
            {
                PdfReadStage.NeedsOcr =>
                    $"スキャンされたページが {_preview.OcrPageNumbers.Count:N0} ページあります。"
                        + "「OCR で読み取る」を押してください。",
                PdfReadStage.Ready when _preview.CanExecute =>
                    "内容を確認して「ファイルを作成」を押してください。",
                _ => "このままでは作成できません。下の内容を確認してください。",
            };
        }
        catch (Exception ex)
        {
            _preview = null;
            StatusText = $"解析に失敗しました: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
            RaisePreviewProperties();
        }
    }

    /// <summary>
    /// OCR Pack の状態を見る。壊れていても例外を投げず、理由を持った状態を返す
    /// (native の読み込みで不可解な落ち方をさせない)。
    /// </summary>
    private OcrPackStatus SafeInspectPack()
    {
        try
        {
            return _inspectPack();
        }
        catch (Exception ex)
        {
            return OcrPackStatus.Broken(string.Empty, ex.Message);
        }
    }

    /// <summary>スキャンされたページを OCR で読み取る。数分かかるので中止できる。</summary>
    public async Task RunOcrAsync()
    {
        if (_preview is not { Stage: PdfReadStage.NeedsOcr } preview
            || _packStatus is not { IsUsable: true } pack)
        {
            return;
        }

        var pages = preview.OcrPageNumbers;
        var source = SourceFilePath;

        _ocrCancellation = new CancellationTokenSource();
        var token = _ocrCancellation.Token;

        IsBusy = true;
        ResultText = null;
        ProgressText = new OcrProgress(0, pages.Count, IsProbe: true).Text;
        StatusText = "スキャンされたページを読み取っています…(元の PDF は読み取りのみ)";

        var progress = new Progress<OcrProgress>(value => ProgressText = value.Text);

        try
        {
            var reading = await Task.Run(
                () =>
                {
                    using var engine = _loadEngine(pack);
                    return _scanReader.Read(engine, source, pages, progress, token);
                },
                token);

            _reading = reading;
            FillReviewItems();

            _preview = _planner.CompleteWithOcr(preview, reading);
            FillPreviewRows();

            StatusText = reading.Issues.Count > 0
                ? "このままでは作成できません。下の内容を確認してください。"
                : reading.IsComplete
                    ? "内容を確認して「ファイルを作成」を押してください。"
                    : $"読み取りました。要確認 {reading.NeedsReviewCount:N0} 件 / "
                        + $"読取不能 {reading.UnreadableCount:N0} 件を確認してください。";
        }
        catch (OperationCanceledException)
        {
            _reading = null;
            ReviewItems.Clear();
            StatusText = "読み取りを中止しました。ファイルは作成していません。";
        }
        catch (Exception ex)
        {
            _reading = null;
            ReviewItems.Clear();
            StatusText = $"読み取りに失敗しました: {ex.Message}";
        }
        finally
        {
            _ocrCancellation?.Dispose();
            _ocrCancellation = null;
            ProgressText = string.Empty;
            IsBusy = false;
            RaisePreviewProperties();
        }
    }

    private void CancelOcr() => _ocrCancellation?.Cancel();

    /// <summary>一覧に出す行を作る。既定は「要確認だけ」。</summary>
    private void FillReviewItems()
    {
        ReviewItems.Clear();
        if (_reading is null)
        {
            return;
        }

        foreach (var item in _reading.Items)
        {
            if (ShowOnlyNeedsReview && item.IsResolved)
            {
                continue;
            }

            ReviewItems.Add(new OcrReviewRow(item));
        }

        RaiseReviewProperties();
    }

    /// <summary>選んだ 1 件を「確認済み」にする。一覧を開いただけでは確認済みにしない。</summary>
    private void ConfirmSelected()
    {
        if (SelectedItem is null)
        {
            return;
        }

        SelectedItem.Confirm();
        AfterReviewChanged();
    }

    /// <summary>いま一覧に出ているものをまとめて確認済みにする。</summary>
    private void ConfirmAllShown()
    {
        foreach (var row in ReviewItems.ToList())
        {
            row.Confirm();
        }

        AfterReviewChanged();
    }

    private void AfterReviewChanged()
    {
        if (_preview is { OcrReading: not null } preview && _reading is not null)
        {
            _preview = _planner.CompleteWithOcr(preview, _reading);
            FillPreviewRows();
        }

        if (ShowOnlyNeedsReview)
        {
            FillReviewItems();
        }

        RaisePreviewProperties();
        RaiseReviewProperties();
    }

    private void FillPreviewRows()
    {
        PreviewRows.Clear();
        if (_preview is null)
        {
            return;
        }

        if (_preview.Kind == PdfDocumentKind.Table)
        {
            foreach (var row in _preview.TableRows.Take(PdfReadDefaults.PreviewRowLimit))
            {
                PreviewRows.Add(new PdfPreviewRow(
                    row.ElementAtOrDefault(0) ?? string.Empty,
                    row.ElementAtOrDefault(1) ?? string.Empty,
                    string.Join(" | ", row.Skip(2))));
            }

            return;
        }

        foreach (var line in _preview.Lines.Take(PdfReadDefaults.PreviewRowLimit))
        {
            PreviewRows.Add(new PdfPreviewRow(
                line.Page.ToString(System.Globalization.CultureInfo.InvariantCulture),
                line.Line.ToString(System.Globalization.CultureInfo.InvariantCulture),
                line.Text));
        }
    }

    private PdfReadRequest BuildRequest() => new()
    {
        SourceFilePath = SourceFilePath,
        OutputFormat = IsCsv ? PdfOutputFormat.Csv : PdfOutputFormat.Xlsx,
        CsvEncoding = EncodingDisplay switch
        {
            EncodingUtf8 => CsvOutputEncoding.Utf8,
            EncodingShiftJis => CsvOutputEncoding.ShiftJis,
            _ => CsvOutputEncoding.Utf8Bom,
        },
        CsvQuoteMode = string.Equals(QuoteDisplay, QuoteAll, StringComparison.Ordinal)
            ? CsvQuoteMode.All
            : CsvQuoteMode.Minimal,
        OutputSuffix = OutputSuffix,
    };

    private async Task ExecuteAsync()
    {
        if (_preview is null || !CanExecute)
        {
            return;
        }

        IsBusy = true;
        StatusText = "ファイルを作成しています…";
        try
        {
            var preview = _preview;
            var result = await Task.Run(() => _reader.Execute(preview));

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

    /// <summary>指定が変わったらプレビューを無効にする(古い内容のまま実行させない)。</summary>
    private void OnSettingsChanged()
    {
        ResultText = null;
        IsPreviewStale = true;

        // 設定が変われば読み取り直し。古い確認結果のまま作成させない。
        _reading = null;
        ReviewItems.Clear();
        SelectedItem = null;

        RaiseCommandStates();
        RaiseReviewProperties();
    }

    private void RaiseReviewProperties()
    {
        OnPropertyChanged(nameof(HasReading));
        OnPropertyChanged(nameof(ReviewSummaryText));
        ConfirmAllShownCommand.RaiseCanExecuteChanged();
        ConfirmSelectedCommand.RaiseCanExecuteChanged();
    }

    private void RaisePreviewProperties()
    {
        OnPropertyChanged(nameof(Preview));
        OnPropertyChanged(nameof(HasPreview));
        OnPropertyChanged(nameof(CanExecute));
        OnPropertyChanged(nameof(CanRunOcr));
        OnPropertyChanged(nameof(NeedsOcr));
        OnPropertyChanged(nameof(HasReading));
        OnPropertyChanged(nameof(ReviewSummaryText));
        OnPropertyChanged(nameof(OcrPackText));
        OnPropertyChanged(nameof(KindText));
        OnPropertyChanged(nameof(SummaryText));
        OnPropertyChanged(nameof(OutputSummaryText));
        OnPropertyChanged(nameof(IssueSummaryText));
        OnPropertyChanged(nameof(Issues));
        OnPropertyChanged(nameof(HasIssues));
        OnPropertyChanged(nameof(FirstColumnHeader));
        OnPropertyChanged(nameof(SecondColumnHeader));
        OnPropertyChanged(nameof(TextColumnHeader));
        RaiseCommandStates();
    }

    private void RaiseCommandStates()
    {
        SelectSourceCommand.RaiseCanExecuteChanged();
        AnalyzeCommand.RaiseCanExecuteChanged();
        RunOcrCommand.RaiseCanExecuteChanged();
        CancelOcrCommand.RaiseCanExecuteChanged();
        ConfirmSelectedCommand.RaiseCanExecuteChanged();
        ConfirmAllShownCommand.RaiseCanExecuteChanged();
        ExecuteCommand.RaiseCanExecuteChanged();
    }
}
