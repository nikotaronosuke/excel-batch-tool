# PdfFeasibility(Phase 2F-R: PDF 対応の実現性ベンチマーク)

製品本体とは**独立した検証専用コード**です。製品のソリューション
(`ExcelBatchTool.sln`)からは参照されず、製品の配布物にも含まれません。

- `PdfBench/` … PdfPig / Tabula / PDFium(PDFtoImage)/ PaddleOCR(PP-OCRv5)/
  SLANet / Tesseract / 罫線格子 + OCR の測定。テスト用 PDF の生成コードを含む
- `PdfBenchV2/` … Sdcb 2.x + Paddle Inference 2.x runtime に固定した第 2 スタック
  (日本語専用 rec: japan_PP-OCRv4)

## 安全上の決まり

- テスト用 PDF は**すべて架空データを生成コードから作る**。第三者の実データ・
  実案件のファイル・個人情報は使わない
- 生成した PDF・Ground Truth・測定結果(JSON)はリポジトリー外の
  作業フォルダーへ出力し、**commit しない**
- ベンチマークは PDF や結果を外部へ送信しない。ベンチマーク実行時のみ
  OCR モデル等をダウンロードすることがある(製品では同梱し、実行時 DL はしない)

使い方: `dotnet run --project PdfBench -- <work-dir> gen` で fixture を生成し、
`pdfpig` / `tabula` / `detect` / `paddle-*` / `grid-table` / `checkbox` / `tess`
の各コマンドで測定する(結果は `<work-dir>/out/*.json`)。
