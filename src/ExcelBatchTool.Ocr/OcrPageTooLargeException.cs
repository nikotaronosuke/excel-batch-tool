namespace ExcelBatchTool.Ocr;

/// <summary>
/// 1 ページが大きすぎて読み取れないときに投げる。
/// 黙って落ちたり、メモリーを使い切ってアプリごと巻き込んだりしないため。
/// </summary>
public sealed class OcrPageTooLargeException(string message) : Exception(message);
