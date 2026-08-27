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
        _viewModel = new MainViewModel(PickFiles);
        DataContext = _viewModel;
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
