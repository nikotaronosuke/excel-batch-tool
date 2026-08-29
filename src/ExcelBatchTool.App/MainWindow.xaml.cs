using System.Windows;
using ExcelBatchTool.App.ViewModels;
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
