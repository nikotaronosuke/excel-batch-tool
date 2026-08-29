# PdfFeasibility(PDF 対応の測定用コード)

製品本体とは**独立した検証専用コード**です。製品のソリューション
(`ExcelBatchTool.sln`)からは参照されず、製品の配布物にも含まれません。

- `PdfBench/` … PdfPig / Tabula / PDFium(PDFtoImage)/ PaddleOCR(PP-OCRv5)/
  SLANet / Tesseract / 罫線格子 + OCR の測定。テスト用 PDF の生成コードを含む
  (Phase 2F-R)
- `PdfBenchV2/` … Sdcb 2.x + Paddle Inference 2.x runtime に固定した第 2 スタック
  (日本語専用 rec: japan_PP-OCRv4)(Phase 2F-R)
- `PdfFusionBench/` … **二重読みの統合方式を Ground Truth の完全一致率で選ぶ**
  (Phase 2F-B1)。検出を 1 回だけ行い、同じ切り出し画像を 2 つの認識モデルへ通して
  結果を保存し、統合方式 11 種類 × 閾値 3 種類を後から機械照合する。
  製品が採用した規則はここで選んだもの

## 二重読みの統合方式を選んだ手順(Phase 2F-B1)

```
dotnet run --project PdfBench      -- <work-dir> gen       # 架空 fixture を作る
dotnet run --project PdfFusionBench -- <work-dir> capture  # 2 モデルで読んで保存
dotnet run --project PdfFusionBench -- <work-dir> fuse     # 統合方式を総当たりで比較
```

`capture` は時間がかかる(帳票 120 ページで数分)ので結果を JSON に残し、
`fuse` はそれを読むだけにしてある。統合方式を足して比べ直すのに再読み取りは要らない。

いちばん重視した指標は完全一致率ではなく、
**自動確定にしたのに間違っていた件数**(false AutoAccepted)。

## 傾き / スキャン表 / 定型帳票 / 印 を測った手順(Phase 2F-B2)

```
dotnet run --project PdfScanBench -- <packDir> <work-dir>          # 全部
dotnet run --project PdfScanBench -- <packDir> <work-dir> deskew   # 傾きだけ
dotnet run --project PdfScanBench -- <packDir> <work-dir> table    # 表だけ
dotnet run --project PdfScanBench -- <packDir> <work-dir> form     # 帳票だけ
dotnet run --project PdfScanBench -- <packDir> <work-dir> mark     # 印だけ
```

`<packDir>` は `tools/OcrPackBuilder` で組んだ Offline OCR Pack のフォルダー。
**製品と同じ経路そのもの**(`OcrPack.Load` → `PdfScanReader` → 確認)で測る。
ベンチ専用の近道を作ると測った数値が製品の挙動と一致しなくなるため。

帳票では完全一致率と並べて **指定した項目のうち結果に現れた割合**
(Expected field coverage)を出す。読めなかった項目が結果から消えていれば
ここが 100% を割るので、「消えた」ことを数字で見つけられる。

fixture の劣化・傾き・かすれは実物に寄せる。**印の線の太さを 2px(300dpi で
0.17mm)にしていたとき判定が 0/6 になった**が、これは fixture 側が細すぎた
だけで、実際のペン相当(6px)にすると 100% になった。
測定結果が悪いときは、まず fixture が現実的かを疑う。

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
