using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using ExcelBatchTool.App.ViewModels;
using ExcelBatchTool.Core.Ocr;
using Microsoft.Win32;

namespace ExcelBatchTool.App;

/// <summary>メイン画面。ファイルの追加(ドロップ / 選択)と解析結果の表示を行う。</summary>
public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();
        _viewModel = new MainViewModel(
            PickFiles, PickSavePath, PickSourceFile, recipeStore: null, pickPdfFile: PickPdfFile);

        // 差し替えられても追従できるよう、購読は DataContext に合わせる。
        DataContextChanged += MainWindow_DataContextChanged;
        DataContext = _viewModel;
    }

    private static string? PickSourceFile()
    {
        var dialog = new OpenFileDialog
        {
            Title = "転記元のファイルを選択",
            Filter = "Excel ブック / CSV (*.xlsx;*.csv)|*.xlsx;*.csv",
        };

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    private static string? PickPdfFile()
    {
        var dialog = new OpenFileDialog
        {
            Title = "読み取る PDF を選択",
            Filter = "PDF ファイル (*.pdf)|*.pdf",
        };

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    private static string[]? PickFiles()
    {
        var dialog = new OpenFileDialog
        {
            Title = "解析する Excel ファイルを選択",
            Filter = "Excel ブック (*.xlsx)|*.xlsx",
            Multiselect = true,
        };

        return dialog.ShowDialog() == true ? dialog.FileNames : null;
    }

    private static string? PickSavePath(string suggestedFileName)
    {
        var dialog = new SaveFileDialog
        {
            Title = "統合ファイルの保存先",
            Filter = "Excel ブック (*.xlsx)|*.xlsx",
            FileName = suggestedFileName,
            DefaultExt = ".xlsx",
            AddExtension = true,
            OverwritePrompt = true,
        };

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    private PdfReadViewModel? _subscribedPdfRead;

    private void MainWindow_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (_subscribedPdfRead is { } previous)
        {
            previous.ScrollToHighlightRequested -= PdfRead_ScrollToHighlightRequested;
            _subscribedPdfRead = null;
        }

        if (e.NewValue is MainViewModel main)
        {
            main.PdfRead.ScrollToHighlightRequested += PdfRead_ScrollToHighlightRequested;
            _subscribedPdfRead = main.PdfRead;
        }
    }

    /// <summary>
    /// 選んだ読み取り位置が見えるところへスクロールする。
    ///
    /// 確認の欄が出た直後や拡大率を変えた直後は、まだ配置が終わっていない。
    /// 自分でスクロール量を指示すると 0 に丸められてしまうので、
    /// 「この範囲を見せて」と頼んで配置のあとに動かしてもらう。
    /// </summary>
    private void PdfRead_ScrollToHighlightRequested(object? sender, OcrDisplayRect rect)
        => Dispatcher.BeginInvoke(
            DispatcherPriority.Background,
            () => PdfPageSurface.BringIntoView(
                new Rect(rect.Left, rect.Top, rect.Width, rect.Height)));

    /// <summary>
    /// ページ画像を出す枠の大きさを ViewModel へ伝える(「全体」の倍率計算に使う)。
    /// </summary>
    private void PdfPageScroll_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (sender is ScrollViewer scrollViewer && _subscribedPdfRead is { } pdf)
        {
            pdf.SetViewport(scrollViewer.ViewportWidth, scrollViewer.ViewportHeight);
        }
    }

    /// <summary>
    /// 大量の確認を現実的にするための操作。
    /// 誤って押しにくいものだけにし、押せるキーは画面にも書いてある。
    /// </summary>
    private void PdfReview_KeyDown(object sender, KeyEventArgs e)
    {
        if (_subscribedPdfRead is not { HasReading: true } pdf)
        {
            return;
        }

        var control = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;

        switch (e.Key)
        {
            case Key.Enter when control && pdf.ConfirmOriginalAndNextCommand.CanExecute(null):
                pdf.ConfirmOriginalAndNextCommand.Execute(null);
                e.Handled = true;
                break;

            case Key.Enter when pdf.ConfirmAndNextCommand.CanExecute(null):
                pdf.ConfirmAndNextCommand.Execute(null);
                e.Handled = true;
                break;

            case Key.Escape when pdf.CancelEditCommand.CanExecute(null):
                pdf.CancelEditCommand.Execute(null);
                e.Handled = true;
                break;

            case Key.F3 when pdf.NextReviewCommand.CanExecute(null):
                pdf.NextReviewCommand.Execute(null);
                e.Handled = true;
                break;
        }
    }

    private void Window_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        e.Handled = true;
    }

    private async void Window_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] paths)
        {
            await _viewModel.AddFilesAsync(paths);
        }
    }
}
