using System.Collections.ObjectModel;
using System.IO;
using ExcelBatchTool.Core.CsvTransform;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Ocr;
using ExcelBatchTool.Core.Recipes;
using ExcelBatchTool.Core.Pdf;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>プレビューに出す 1 行(文字 PDF は ページ/行/内容、表 PDF は列を連結)。</summary>
public sealed record PdfPreviewRow(string First, string Second, string Text);

/// <summary>「8. PDF を読み取る」(Phase 2F-A / 2F-B1)の ViewModel。</summary>
public sealed class PdfReadViewModel : ObservableObject, IRecipeHost
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

    /// <summary>OCR を始められる状態のプレビュー。読み取り直しのために残しておく。</summary>
    private PdfReadPreview? _scanPreview;
    private OcrReadMode _readMode = OcrReadMode.Auto;
    private FormTemplate? _template;
    private OcrDocumentReading? _reading;
    private OcrReviewSession? _session;
    private CancellationTokenSource? _ocrCancellation;
    private string _progressText = string.Empty;

    // 確認中は PDF を開いたままにする(ページ画像を出すため)。
    private IOcrEngine? _reviewEngine;
    private IOcrPageSource? _reviewSource;
    private OcrPageImageCache? _imageCache;
    private OcrPageImage? _pageImage;
    private (OcrPageImage Page, System.Windows.Media.Imaging.BitmapSource Source)? _pageImageSource;
    private int _currentPage;
    private double _zoom = 1;
    private double _viewportWidth = 640;
    private double _viewportHeight = 420;

    public PdfReadViewModel()
        : this(() => null)
    {
    }

    /// <summary>テスト用: ファイル選択と OCR Pack を差し替えられるようにする。</summary>
    internal PdfReadViewModel(
        Func<string?> pickPdfFile,
        Func<OcrPackStatus>? inspectPack = null,
        Func<OcrPackStatus, IOcrEngine>? loadEngine = null,
        RecipeStore? recipeStore = null,
        Func<string, bool>? confirm = null)
    {
        _pickPdfFile = pickPdfFile;
        Recipes = new RecipeAreaViewModel(
            this, recipeStore ?? new RecipeStore(), confirm ?? RecipeSaveGuard.AskInDialog);
        _inspectPack = inspectPack ?? (() => OcrPack.Inspect());
        _loadEngine = loadEngine ?? OcrPack.Load;

        SelectSourceCommand = new RelayCommand(SelectSource, () => !IsBusy);
        AnalyzeCommand = new RelayCommand(
            () => _ = AnalyzeAsync(), () => !IsBusy && SourceFilePath.Length > 0);
        RunOcrCommand = new RelayCommand(() => _ = RunOcrAsync(), () => CanRunOcr);
        CancelOcrCommand = new RelayCommand(CancelOcr, () => IsBusy && _ocrCancellation is not null);
        // 「まとめて確認済みにする」は置かない。要確認・読取不能は元のページを見ながら
        // 1 件ずつ確認する、というのがこの段階の安全の要なので、迂回路を作らない。
        ConfirmAndNextCommand = new RelayCommand(
            () => ConfirmAndAdvance(useEdit: true), () => CanConfirmSelected);
        ConfirmOriginalAndNextCommand = new RelayCommand(
            () => ConfirmAndAdvance(useEdit: false), () => CanConfirmSelected);
        CancelEditCommand = new RelayCommand(CancelEdit, () => SelectedItem is not null);
        NextReviewCommand = new RelayCommand(
            () => MoveReview(forward: true), () => HasReading && !IsBusy);
        PreviousReviewCommand = new RelayCommand(
            () => MoveReview(forward: false), () => HasReading && !IsBusy);
        PreviousPageCommand = new RelayCommand(
            () => ShowPage(CurrentPage - 1), () => HasReading && CurrentPage > 1);
        NextPageCommand = new RelayCommand(
            () => ShowPage(CurrentPage + 1), () => HasReading && CurrentPage < PageCount);
        ZoomInCommand = new RelayCommand(() => SetZoom(Zoom * 1.25), () => HasPageImage);
        ZoomOutCommand = new RelayCommand(() => SetZoom(Zoom / 1.25), () => HasPageImage);
        ActualSizeCommand = new RelayCommand(() => SetZoom(1), () => HasPageImage);
        FitCommand = new RelayCommand(FitToViewport, () => HasPageImage);
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

    public RelayCommand ConfirmAndNextCommand { get; }

    public RelayCommand ConfirmOriginalAndNextCommand { get; }

    public RelayCommand CancelEditCommand { get; }

    public RelayCommand NextReviewCommand { get; }

    public RelayCommand PreviousReviewCommand { get; }

    public RelayCommand PreviousPageCommand { get; }

    public RelayCommand NextPageCommand { get; }

    public RelayCommand ZoomInCommand { get; }

    public RelayCommand ZoomOutCommand { get; }

    public RelayCommand ActualSizeCommand { get; }

    public RelayCommand FitCommand { get; }

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

    /// <summary>
    /// 読み取りを始められるか。一度読み取ったあとでも押せる
    /// (帳票の項目を直してから読み取り直す、という流れがあるため)。
    /// </summary>
    public bool CanRunOcr => !IsBusy
        && !IsPreviewStale
        && _scanPreview is not null
        && _packStatus is { IsUsable: true };

    /// <summary>OCR が必要な PDF か(画面の確認欄を出すかどうか)。</summary>
    public bool NeedsOcr => _scanPreview is not null;

    public bool HasReading => _reading is not null;

    /// <summary>確認中でない(取り出した内容だけを見せる状態)。</summary>
    public bool HasNoReading => _reading is null;

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
            if (!SetProperty(ref _selectedItem, value))
            {
                return;
            }

            _session?.Select(value?.Item);

            // 選んだ項目のページへ自動で移動し、その位置を強調して見えるところへ寄せる。
            if (value is not null)
            {
                ShowPage(value.PageNumber);
                BringHighlightIntoView();
            }

            OnPropertyChanged(nameof(SelectedEngineText));
            OnPropertyChanged(nameof(HasSelectedItem));
            OnPropertyChanged(nameof(SelectedPositionText));
            OnPropertyChanged(nameof(SelectedReasonText));
            OnPropertyChanged(nameof(HasSelectedReason));
            RaiseHighlight();
            RaiseCommandStates();
        }
    }

    public bool HasSelectedItem => SelectedItem is not null;

    public bool CanConfirmSelected => SelectedItem is not null && !IsBusy;

    public string SelectedEngineText => SelectedItem?.EngineText ?? "-";

    public string SelectedPositionText => SelectedItem is null
        ? "-"
        : $"{SelectedItem.PageNumber} ページ {SelectedItem.LineNumber} 行目";

    /// <summary>
    /// なぜ確認が要るのか。自信が高くても形が怪しくて自動確定を見送ったときは、
    /// その理由をここに出す。理由が分からないまま「元のままで確認」を押されると、
    /// 確認の意味が無くなるため。
    /// </summary>
    public string SelectedReasonText => SelectedItem?.ReasonText ?? string.Empty;

    public bool HasSelectedReason => !string.IsNullOrWhiteSpace(SelectedReasonText);

    /// <summary>
    /// 自動確定も一覧に出すか。既定は出さない。
    /// 自動確定したからといって、元のページを見られなくはしない。
    /// </summary>
    public bool ShowAutoAccepted
    {
        get => _session?.ShowAutoAccepted ?? false;
        set
        {
            if (_session is null || _session.ShowAutoAccepted == value)
            {
                return;
            }

            _session.ShowAutoAccepted = value;
            OnPropertyChanged();
            FillReviewItems();
        }
    }

    // ── 元のページ画像 ─────────────────────────────────

    /// <summary>確認用に描くときの解像度。OCR(300dpi)より粗くてよい。</summary>
    public const int ViewDpi = 150;

    /// <summary>読み取り位置がこれより小さく映るなら、読めるところまで拡大する。</summary>
    public const double MinLegibleHeight = 24;

    /// <summary>拡大するときに目指す、読み取り位置の高さ。</summary>
    public const double PreferredHighlightHeight = 44;

    /// <summary>選んだ位置が見えるところへスクロールしてほしい、という合図。</summary>
    public event EventHandler<OcrDisplayRect>? ScrollToHighlightRequested;

    public int PageCount => _preview?.PageCount ?? 0;

    public int CurrentPage
    {
        get => _currentPage;
        private set
        {
            if (SetProperty(ref _currentPage, value))
            {
                OnPropertyChanged(nameof(PageText));
                RaiseCommandStates();
            }
        }
    }

    public string PageText => PageCount == 0 ? "-" : $"{CurrentPage} / {PageCount} ページ";

    public OcrPageImage? PageImage => _pageImage;

    public bool HasPageImage => _pageImage is not null;

    /// <summary>画面に出すページ画像。読み込んだら凍結して、UI スレッド以外からも触れるようにする。</summary>
    public System.Windows.Media.Imaging.BitmapSource? PageImageSource
    {
        get
        {
            if (_pageImage is not { } image)
            {
                return null;
            }

            if (_pageImageSource is { Source: var cachedSource, Page: var cachedPage }
                && ReferenceEquals(cachedPage, image))
            {
                return cachedSource;
            }

            var bitmap = new System.Windows.Media.Imaging.BitmapImage();
            bitmap.BeginInit();
            bitmap.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
            bitmap.StreamSource = new MemoryStream(image.Png, writable: false);
            bitmap.EndInit();
            bitmap.Freeze();

            _pageImageSource = (image, bitmap);
            return bitmap;
        }
    }

    /// <summary>ページ画像を表示する大きさ(拡大率を掛けたもの)。</summary>
    public double ImageDisplayWidth => (_pageImage?.Width ?? 0) * Zoom;

    public double ImageDisplayHeight => (_pageImage?.Height ?? 0) * Zoom;

    public double Zoom
    {
        get => _zoom;
        private set
        {
            if (SetProperty(ref _zoom, value))
            {
                OnPropertyChanged(nameof(ZoomText));
                OnPropertyChanged(nameof(ImageDisplayWidth));
                OnPropertyChanged(nameof(ImageDisplayHeight));
                RaiseHighlight();
            }
        }
    }

    public string ZoomText => $"{Zoom:P0}";

    /// <summary>選んでいる読み取り位置を、画像の上のどこへ描くか。</summary>
    public OcrDisplayRect Highlight => SelectedItem is { } row && _pageImage is { } image
        && row.PageNumber == CurrentPage
            ? OcrBoxMapper.ToDisplay(row.Item.BoundingBox, image, Zoom)
            : default;

    public bool HasHighlight => SelectedItem is { } row
        && _pageImage is not null
        && row.PageNumber == CurrentPage;

    public double HighlightLeft => Highlight.Left;

    public double HighlightTop => Highlight.Top;

    public double HighlightWidth => Highlight.Width;

    public double HighlightHeight => Highlight.Height;

    /// <summary>手元に置いているページ画像の枚数(メモリの確認用)。</summary>
    public int CachedPageCount => _imageCache?.Count ?? 0;

    public int PageRenderCount => _imageCache?.RenderCount ?? 0;

    /// <summary>画面の大きさが変わったら、fit の計算に使う値を更新する。</summary>
    public void SetViewport(double width, double height)
    {
        if (width <= 0 || height <= 0)
        {
            return;
        }

        _viewportWidth = width;
        _viewportHeight = height;
    }

    public void ShowPage(int page)
    {
        if (_imageCache is null || page < 1 || page > PageCount)
        {
            return;
        }

        if (_pageImage?.Page != page)
        {
            _pageImage = _imageCache.Get(page);
            _imageCache.Preload(page, PageCount);

            OnPropertyChanged(nameof(PageImage));
            OnPropertyChanged(nameof(PageImageSource));
            OnPropertyChanged(nameof(HasPageImage));
            OnPropertyChanged(nameof(ImageDisplayWidth));
            OnPropertyChanged(nameof(ImageDisplayHeight));
            OnPropertyChanged(nameof(CachedPageCount));
            OnPropertyChanged(nameof(PageRenderCount));
        }

        CurrentPage = page;
        RaiseHighlight();
    }

    private void SetZoom(double zoom) => Zoom = Math.Clamp(zoom, 0.1, 4);

    /// <summary>
    /// 選んだ位置を見えるようにする。
    ///
    /// ページ全体に合わせた倍率だと 1 行は数画素にしかならず、原文と見比べられない。
    /// 小さすぎるときだけ、読める大きさまで自動で拡大してからそこへ寄せる。
    /// 利用者が自分で拡大しているときは、その倍率を尊重して触らない。
    /// </summary>
    private void BringHighlightIntoView()
    {
        if (SelectedItem is not { } row || _pageImage is not { } image)
        {
            return;
        }

        var rect = OcrBoxMapper.ToDisplay(row.Item.BoundingBox, image, Zoom);

        if (rect.Height < MinLegibleHeight)
        {
            var boxHeight = row.Item.BoundingBox.Height * image.ScaleFromOcr;
            if (boxHeight > 0)
            {
                SetZoom(PreferredHighlightHeight / boxHeight);
                rect = OcrBoxMapper.ToDisplay(row.Item.BoundingBox, image, Zoom);
            }
        }

        ScrollToHighlightRequested?.Invoke(this, rect);
    }

    private void FitToViewport()
    {
        if (_pageImage is { } image)
        {
            SetZoom(OcrBoxMapper.FitZoom(image, _viewportWidth, _viewportHeight));
        }
    }

    private void RaiseHighlight()
    {
        OnPropertyChanged(nameof(Highlight));
        OnPropertyChanged(nameof(HasHighlight));
        OnPropertyChanged(nameof(HighlightLeft));
        OnPropertyChanged(nameof(HighlightTop));
        OnPropertyChanged(nameof(HighlightWidth));
        OnPropertyChanged(nameof(HighlightHeight));
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
            ResetReview();
            _scanPreview = null;

            // OCR Pack は「スキャンページがあったときだけ」必要になる。
            // 無くてもここで失敗させない(文字情報のある PDF はそのまま読める)。
            _packStatus = await Task.Run(SafeInspectPack);
            var pack = _packStatus;

            _preview = await Task.Run(() => _planner.CreatePreview(request, pack));
            _scanPreview = _preview.Stage == PdfReadStage.NeedsOcr ? _preview : null;
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
        if (_scanPreview is not { } preview || _packStatus is not { IsUsable: true } pack)
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
            // 読み取りが終わってからも、確認用にページ画像を出すため PDF は開いたままにする。
            var opened = await Task.Run(
                () =>
                {
                    var engine = _loadEngine(pack);
                    var pageSource = engine.Open(source);
                    var reading = _scanReader.Read(
                        pageSource, engine.Info, pages, BuildReadOptions(), progress, token);
                    return (engine, pageSource, reading);
                },
                token);

            CloseReviewSource();
            _reviewEngine = opened.engine;
            _reviewSource = opened.pageSource;
            var reading = opened.reading;

            _imageCache = new OcrPageImageCache(
                page => _reviewSource!.RenderPage(page, ViewDpi, CancellationToken.None),
                OcrPageImageCache.DefaultCapacity);

            _reading = reading;
            _session = new OcrReviewSession(reading);

            // 帳票として読む指定なのに項目がまだ決まっていなければ、
            // 1 ページ目の読み取りから候補を作る(利用者は名前を直すだけでよい)。
            if (IsFixedForm && TemplateFields.Count == 0)
            {
                SuggestTemplateFrom(reading);
            }

            // まずページ全体が入る大きさにしておく。このあと項目を選ぶと、
            // その読み取り位置が読める大きさまで自動で寄る。
            if (reading.Items.Count > 0)
            {
                ShowPage(reading.Items[0].PageNumber);
                FitToViewport();
            }

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
            ResetReview();
            StatusText = "読み取りを中止しました。ファイルは作成していません。";
        }
        catch (Exception ex)
        {
            ResetReview();
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

    // ── 同じ様式の帳票としてまとめて読む ────────────────

    private const string ModeAuto = "自動で判断する";
    private const string ModeLines = "文章として読む";
    private const string ModeTable = "表として読む";
    private const string ModeForm = "同じ様式の帳票として読む";

    public static IReadOnlyList<string> ReadModeOptions { get; }
        = [ModeAuto, ModeLines, ModeTable, ModeForm];

    public string ReadModeDisplay
    {
        get => _readMode switch
        {
            OcrReadMode.Lines => ModeLines,
            OcrReadMode.Table => ModeTable,
            OcrReadMode.FixedForm => ModeForm,
            _ => ModeAuto,
        };
        set
        {
            var mode = value switch
            {
                ModeLines => OcrReadMode.Lines,
                ModeTable => OcrReadMode.Table,
                ModeForm => OcrReadMode.FixedForm,
                _ => OcrReadMode.Auto,
            };

            if (_readMode == mode)
            {
                return;
            }

            _readMode = mode;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsFixedForm));
            OnSettingsChanged();
        }
    }

    public bool IsFixedForm => _readMode == OcrReadMode.FixedForm;

    /// <summary>帳票として読むときの項目一覧(画面で編集する)。</summary>
    public ObservableCollection<FormFieldRow> TemplateFields { get; } = [];

    public string TemplateSummaryText => TemplateFields.Count == 0
        ? "読み取る項目をまだ決めていません。"
        : $"{TemplateFields.Count:N0} 項目を 1 ページ 1 件として読み取ります。";

    /// <summary>
    /// 1 ページ目の読み取り結果から、項目の候補を作る。
    /// 「ラベル: 値」の形になっている行を拾って、値の場所を読み取り領域にする。
    /// 利用者は名前と要否を直すだけでよい。
    /// </summary>
    public void SuggestTemplateFrom(OcrDocumentReading reading)
    {
        TemplateFields.Clear();

        var firstPage = reading.Items
            .Where(item => item.PageNumber == reading.OcrPages.FirstOrDefault())
            .OrderBy(item => item.LineNumber)
            .ThenBy(item => item.IndexInLine)
            .ToList();

        foreach (var line in firstPage.GroupBy(item => item.LineNumber))
        {
            var parts = line.OrderBy(item => item.IndexInLine).ToList();

            // 「項目名」と「値」が別々に読み取れている場合。
            if (parts.Count >= 2)
            {
                var name = Label(parts[0].Text);
                if (name.Length > 0)
                {
                    TemplateFields.Add(new FormFieldRow(name, parts[^1].BoundingBox)
                    {
                        LabelArea = parts[0].BoundingBox,
                    });
                }

                continue;
            }

            // 「項目名: 値」が 1 つのまとまりとして読み取れている場合。
            // 読み取り位置を文字数の比で分け、右側を値の場所として扱う。
            var single = parts[0];
            var colon = single.Text.IndexOfAny([':', '：']);
            if (colon <= 0 || colon >= single.Text.Length - 1)
            {
                continue;
            }

            var label = Label(single.Text[..colon]);
            if (label.Length == 0)
            {
                continue;
            }

            var box = single.BoundingBox;
            var ratio = (colon + 1) / (double)single.Text.Length;
            var split = box.X + (box.Width * ratio);

            TemplateFields.Add(new FormFieldRow(
                label,
                new OcrBox(split, box.Y, box.Right - split, box.Height))
            {
                LabelArea = new OcrBox(box.X, box.Y, split - box.X, box.Height),
            });
        }

        OnPropertyChanged(nameof(TemplateSummaryText));
    }

    /// <summary>項目名として使える形に整える(末尾のコロンや空白を落とす)。</summary>
    private static string Label(string text) => text.Trim().TrimEnd(':', '：', ' ').Trim();

    private OcrReadOptions BuildReadOptions()
    {
        if (_readMode != OcrReadMode.FixedForm)
        {
            return new OcrReadOptions { Mode = _readMode };
        }

        _template = new FormTemplate
        {
            Name = SourceFileNameDisplay,
            Fields = [.. TemplateFields.Select(row => row.ToField())],

            // 項目名そのものを位置合わせの手がかりにする。
            // ページごとの多少のずれは、これで吸収できる。
            Anchors = [.. TemplateFields
                .Where(row => row.UseAsAnchor)
                .Select(row => new FormAnchor(row.Name, row.LabelArea))],
        };

        return new OcrReadOptions { Mode = OcrReadMode.FixedForm, Template = _template };
    }

    /// <summary>確認の状態を捨てて、開いていた PDF も閉じる。</summary>
    private void ResetReview()
    {
        _reading = null;
        _session = null;
        _pageImage = null;
        _pageImageSource = null;
        _currentPage = 0;
        ReviewItems.Clear();
        SelectedItem = null;
        _imageCache?.Clear();
        _imageCache = null;
        CloseReviewSource();

        OnPropertyChanged(nameof(PageImage));
        OnPropertyChanged(nameof(PageImageSource));
        OnPropertyChanged(nameof(HasPageImage));
        OnPropertyChanged(nameof(CachedPageCount));
        RaiseHighlight();
    }

    private void CloseReviewSource()
    {
        _reviewSource?.Dispose();
        _reviewSource = null;
        _reviewEngine?.Dispose();
        _reviewEngine = null;
    }

    /// <summary>
    /// 一覧に出す行を作る。既定は「人が見るべきもの」だけ。
    ///
    /// 絞り込みは最初の分類で行うので、確認済みにしても行は消えない
    /// (何を確認したのかが見えなくなると、取り消しもできなくなるため)。
    /// </summary>
    private void FillReviewItems()
    {
        var previous = SelectedItem?.Item;

        ReviewItems.Clear();
        if (_session is null)
        {
            SelectedItem = null;
            RaiseReviewProperties();
            return;
        }

        foreach (var item in _session.Visible)
        {
            ReviewItems.Add(new OcrReviewRow(item));
        }

        SelectedItem = ReviewItems.FirstOrDefault(row => ReferenceEquals(row.Item, previous))
            ?? ReviewItems.FirstOrDefault(row => !row.IsResolved)
            ?? ReviewItems.FirstOrDefault();

        RaiseReviewProperties();
    }

    /// <summary>
    /// 選んでいる 1 件を確認済みにして、次の未確認へ進む。
    /// <paramref name="useEdit"/> が false なら、元の読み取りのままで確認する。
    /// </summary>
    private void ConfirmAndAdvance(bool useEdit)
    {
        if (SelectedItem is not { } row || _session is null)
        {
            return;
        }

        _session.Select(row.Item);
        _session.ConfirmSelectedAndAdvance(useEdit ? row.EditedText : null);
        row.Refresh();

        AfterReviewChanged();
        SyncSelectionFromSession();
    }

    /// <summary>編集中の文字を元の読み取りへ戻す(確認済みにはしない)。</summary>
    private void CancelEdit()
    {
        if (SelectedItem is { } row)
        {
            row.ResetEdit();
        }
    }

    private void MoveReview(bool forward)
    {
        if (_session is null)
        {
            return;
        }

        _session.Select(SelectedItem?.Item);
        if (forward ? _session.MoveToNextUnresolved() : _session.MoveToPreviousUnresolved())
        {
            SyncSelectionFromSession();
        }
    }

    private void SyncSelectionFromSession()
    {
        if (_session?.Selected is not { } selected)
        {
            return;
        }

        SelectedItem = ReviewItems.FirstOrDefault(row => ReferenceEquals(row.Item, selected));
    }

    private void AfterReviewChanged()
    {
        if (_preview is { OcrReading: not null } preview && _reading is not null)
        {
            _preview = _planner.CompleteWithOcr(preview, _reading);
            FillPreviewRows();
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
        ResetReview();
        _scanPreview = null;
        OnPropertyChanged(nameof(TemplateSummaryText));

        RaiseCommandStates();
        RaiseReviewProperties();
    }

    private void RaiseReviewProperties()
    {
        OnPropertyChanged(nameof(HasReading));
        OnPropertyChanged(nameof(HasNoReading));
        OnPropertyChanged(nameof(ReviewSummaryText));
        OnPropertyChanged(nameof(ShowAutoAccepted));
        OnPropertyChanged(nameof(PageCount));
        OnPropertyChanged(nameof(PageText));
        RaiseCommandStates();
    }

    private void RaisePreviewProperties()
    {
        OnPropertyChanged(nameof(Preview));
        OnPropertyChanged(nameof(HasPreview));
        OnPropertyChanged(nameof(CanExecute));
        OnPropertyChanged(nameof(CanRunOcr));
        OnPropertyChanged(nameof(NeedsOcr));
        OnPropertyChanged(nameof(HasReading));
        OnPropertyChanged(nameof(HasNoReading));
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

    // ── 処理設定(レシピ) ─────────────────────────────
    //
    // 保存するのは「どう読むか」だけ。**元の PDF に関わるものは何も保存しない**
    // (ファイル名・保存場所・読み取った文字・ページの中身・個人情報の実値)。
    // 毎月同じ様式の帳票が届くとき、項目を作り直さずに済むようにするためのもの。

    public RecipeAreaViewModel Recipes { get; }

    RecipeType IRecipeHost.RecipeType => RecipeType.PdfRead;

    string? IRecipeHost.RecipeSaveBlockedReason
        => IsFixedForm && TemplateFields.Count == 0
            ? "読み取る項目がまだありません。1 度読み取って項目を作ってから保存してください。"
            : null;

    SavedRecipe IRecipeHost.CreateRecipe(string name) => new()
    {
        Name = name,
        Type = RecipeType.PdfRead,
        PdfRead = new PdfReadRecipe
        {
            ReadMode = ReadModeDisplay,
            OutputFormat = FormatDisplay,
            OutputSuffix = OutputSuffix,
            Encoding = EncodingDisplay switch
            {
                EncodingUtf8 => CsvOutputEncoding.Utf8,
                EncodingShiftJis => CsvOutputEncoding.ShiftJis,
                _ => CsvOutputEncoding.Utf8Bom,
            },
            QuoteMode = string.Equals(QuoteDisplay, QuoteAll, StringComparison.Ordinal)
                ? CsvQuoteMode.All
                : CsvQuoteMode.Minimal,
            Fields = [.. TemplateFields.Select(field => new PdfReadRecipeField
            {
                Name = field.Name,
                Kind = field.Kind,
                IsRequired = field.IsRequired,
                X = field.Area.X,
                Y = field.Area.Y,
                Width = field.Area.Width,
                Height = field.Area.Height,
            })],
        },
    };

    /// <summary>
    /// 設定を画面へ戻す。読み取りはやり直させる(前回の読み取り結果は引き継がない)。
    /// </summary>
    IReadOnlyList<string> IRecipeHost.ApplyRecipe(SavedRecipe recipe)
    {
        var payload = recipe.PdfRead!;
        var notes = new List<string>();

        if (ReadModeOptions.Contains(payload.ReadMode))
        {
            ReadModeDisplay = payload.ReadMode;
        }
        else
        {
            notes.Add($"「{payload.ReadMode}」という読み取り方は今のバージョンにありません。");
        }

        FormatDisplay = payload.OutputFormat;
        OutputSuffix = payload.OutputSuffix;
        EncodingDisplay = payload.Encoding switch
        {
            CsvOutputEncoding.Utf8 => EncodingUtf8,
            CsvOutputEncoding.ShiftJis => EncodingShiftJis,
            _ => EncodingUtf8Bom,
        };
        QuoteDisplay = payload.QuoteMode == CsvQuoteMode.All ? QuoteAll : QuoteMinimal;

        TemplateFields.Clear();
        foreach (var field in payload.Fields)
        {
            TemplateFields.Add(
                new FormFieldRow(
                    field.Name, new OcrBox(field.X, field.Y, field.Width, field.Height))
                {
                    Kind = FormFieldRow.KindOptions.Contains(field.Kind)
                        ? field.Kind
                        : FormFieldRow.KindOptions[0],
                    IsRequired = field.IsRequired,
                });
        }

        if (payload.Fields.Count > 0)
        {
            notes.Add(
                $"{payload.Fields.Count:N0} 項目を読み込みました。"
                    + "読み取る場所は前の PDF に合わせたものなので、"
                    + "今回の PDF で 1 度読み取って位置を確かめてください。");
        }

        OnPropertyChanged(nameof(TemplateSummaryText));
        RaiseCommandStates();
        return notes;
    }

    private void RaiseCommandStates()
    {
        SelectSourceCommand.RaiseCanExecuteChanged();
        AnalyzeCommand.RaiseCanExecuteChanged();
        RunOcrCommand.RaiseCanExecuteChanged();
        CancelOcrCommand.RaiseCanExecuteChanged();
        ConfirmAndNextCommand.RaiseCanExecuteChanged();
        ConfirmOriginalAndNextCommand.RaiseCanExecuteChanged();
        CancelEditCommand.RaiseCanExecuteChanged();
        NextReviewCommand.RaiseCanExecuteChanged();
        PreviousReviewCommand.RaiseCanExecuteChanged();
        PreviousPageCommand.RaiseCanExecuteChanged();
        NextPageCommand.RaiseCanExecuteChanged();
        ZoomInCommand.RaiseCanExecuteChanged();
        ZoomOutCommand.RaiseCanExecuteChanged();
        ActualSizeCommand.RaiseCanExecuteChanged();
        FitCommand.RaiseCanExecuteChanged();
        ExecuteCommand.RaiseCanExecuteChanged();
    }
}
