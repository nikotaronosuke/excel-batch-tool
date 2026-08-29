using System.Collections.ObjectModel;
using System.IO;
using ExcelBatchTool.Core.CsvTransform;
using ExcelBatchTool.Core.Merge;
using ExcelBatchTool.Core.Pdf;

namespace ExcelBatchTool.App.ViewModels;

/// <summary>プレビューに出す 1 行(文字 PDF は ページ/行/内容、表 PDF は列を連結)。</summary>
public sealed record PdfPreviewRow(string First, string Second, string Text);

/// <summary>「8. PDF を読み取る」(Phase 2F-A)の ViewModel。</summary>
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
    private readonly Func<string?> _pickPdfFile;

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

    public PdfReadViewModel()
        : this(() => null)
    {
    }

    /// <summary>テスト用: ファイル選択を差し替えられるようにする。</summary>
    internal PdfReadViewModel(Func<string?> pickPdfFile)
    {
        _pickPdfFile = pickPdfFile;

        SelectSourceCommand = new RelayCommand(SelectSource, () => !IsBusy);
        AnalyzeCommand = new RelayCommand(
            () => _ = AnalyzeAsync(), () => !IsBusy && SourceFilePath.Length > 0);
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
            _preview = await Task.Run(() => _planner.CreatePreview(request));
            IsPreviewStale = false;
            FillPreviewRows();

            StatusText = _preview.CanExecute
                ? "内容を確認して「ファイルを作成」を押してください。"
                : "このままでは作成できません。下の内容を確認してください。";
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
        RaiseCommandStates();
    }

    private void RaisePreviewProperties()
    {
        OnPropertyChanged(nameof(Preview));
        OnPropertyChanged(nameof(HasPreview));
        OnPropertyChanged(nameof(CanExecute));
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
        ExecuteCommand.RaiseCanExecuteChanged();
    }
}
