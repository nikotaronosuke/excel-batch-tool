# Phase 2F-R: PDF 対応の実現性ベンチマーク

日付: 2026-08-29
Public ベンチマークコード: `research/PdfFeasibility/PdfBench`(架空 fixture 生成コードのみ。実データ・結果データは commit しない)
作業ディレクトリ(ローカルのみ・未 commit): scratchpad の `pdf-bench/`(fixtures・GT・out の JSON)

目的: 「PDF → 構造化された field / table → 人間確認 → Excel / CSV」まで安全に持っていけるか。
文字が取れたかではなく、**納品できる完全一致率**で判定する。

## 1. 候補と統合方式

| 層 | 本命 | 比較 | 除外 |
|---|---|---|---|
| PC 生成 PDF の文字 | PdfPig | - | - |
| PC 生成 PDF の表 | PdfPig + Tabula(tabula-sharp) | 自前の位置ベース再構成 | - |
| ページ画像化 | PDFtoImage(PDFium) | - | MuPDF 系(AGPL/商用ライセンス懸念のため本命から除外) |
| スキャン OCR | PaddleOCR(PaddleSharp / Sdcb 経由のローカル推論) | Tesseract 5(baseline) | クラウド OCR(全面禁止) |
| スキャン表構造 | SLANet(PaddleOCR 表構造モデル、Sdcb 経由) | - | PP-StructureV3 full pipeline(Python sidecar が必要。§9) |
| チェックボックス | 固定位置の画素解析(自前) | - | - |

統合方式(§16)の比較:
- **A. 既存 .NET wrapper(PaddleSharp/Sdcb)で直接利用** — 今回の実測方式。追加プロセスなし、C# のみ。
- B. ONNX / OpenVINO で C# からローカル推論 — RapidOCR 系の ONNX モデル + onnxruntime。前後処理(CTC decode・dict)を自前実装する必要があり、実装量は A より大。A で精度・性能が足りたため今回は未実測。
- C. 公式 PaddleOCR(Python)を完全ローカル sidecar として同梱 — PP-StructureV3 や最新 PP-OCRv5/v6 server 系をそのまま使えるが、Python ランタイム同梱(数百 MB〜)と別プロセス管理が必要。§9 の表構造で SLANet が不足する場合の逃げ道として保留。

## 2. ライセンス(確認日 2026-08-29)

「無料」ではなく「Windows アプリに同梱して商用案件に使えるか」で確認。

| 部品 | 版 | ライセンス | 商用・再配布 | 根拠 |
|---|---|---|---|---|
| PdfPig | 0.1.16 | Apache-2.0 | 可(ライセンス文同梱) | github.com/UglyToad/PdfPig (LICENSE) |
| tabula-sharp (Tabula) | 1.0.1 | MIT | 可 | github.com/BobLd/tabula-sharp (LICENSE) |
| SkiaSharp | 4.151.1 | MIT(native skia は BSD-3) | 可 | github.com/mono/SkiaSharp |
| PDFtoImage | 5.4.0 | MIT | 可 | github.com/sungaila/PDFtoImage |
| PDFium(同梱 native) | - | Apache-2.0 / BSD-3 | 可(NOTICE 同梱) | pdfium.googlesource.com (LICENSE) |
| PaddleSharp(Sdcb.PaddleOCR / PaddleInference) | 3.3.1 | MIT | 可 | github.com/sdcb/PaddleSharp (LICENSE) |
| Paddle Inference native | 3.3.x | Apache-2.0 | 可 | github.com/PaddlePaddle/Paddle (LICENSE) |
| PaddleOCR モデル(PP-OCRv5 / SLANet / japan v4) | - | Apache-2.0 | 可(モデルも Apache) | github.com/PaddlePaddle/PaddleOCR (LICENSE) |
| oneDNN(mkldnn) | 3.6.2 | Apache-2.0 | 可 | github.com/uxlfoundation/oneDNN |
| Intel MKL(runtime.win64.mkl に含まれる場合) | - | Intel Simplified Software License | 再配布可の記載あり。**製品同梱前に最終確認**。回避策: Openblas ランタイム(BSD-3)へ差し替え可 | intel.com ISSL |
| OpenCvSharp4 (+runtime.win) | 4.13 | Apache-2.0(OpenCV 本体も Apache-2.0) | 可 | github.com/shimat/opencvsharp |
| Tesseract(.NET wrapper) | 5.2.0 | Apache-2.0 | 可。**VC++ ランタイム必要**(app-local 配布可) | github.com/charlesw/tesseract |
| tessdata(jpn) | fast | Apache-2.0 | 可 | github.com/tesseract-ocr/tessdata_fast |
| MuPDF / PyMuPDF / MuPDF.NET | - | AGPL / 商用 | **除外**(指示どおり) | mupdf.com |

クラウド OCR・外部 API・telemetry・実行時モデル DL は全面不使用(§19)。
ベンチマーク中の DL は tessdata(2.4MB)と japan v4 rec モデルのみ。製品では同梱する。

## 3. テストコーパス(すべて架空・生成コードは Public)

| 系統 | 内容 | 規模 |
|---|---|---|
| A born-digital text | 日本語文章 + 会社名/金額/電話(000 番台)/日付、2 段組 | 10 ページ |
| B born-digital table(罫線) | 商品コード/商品名/単価/在庫 | 10 ページ × 40 行 = 400 行 + ヘッダー |
| C 同(罫線なし) | 同上 | 同上 |
| D scan clean | 上記を 300dpi 画像化 → JPEG80 → 画像 PDF | text 5 / table 10 / form 120 |
| E scan degraded | 150dpi + ぼかし1.1 + 傾き±1〜3° + ごま塩ノイズ + かすれ + JPEG40 | text 5 / table 5 / form 30 |
| F fixed form | 店舗コード/担当者/日付/売上/Q1/Q2/備考(同一レイアウト) | 120 ページ |
| G checkbox | □はい/□いいえ/□未回答 + チェック・塗り・丸囲み・無印 | 120 ページ(F と同居) |
| H scanned table | B を 300dpi で画像 PDF 化 | 10 ページ |
| I handwriting 風 | UD デジタル教科書体 + 文字ごとの位置・角度の揺らぎ | 5 ページ |

Ground Truth は生成時に JSON で同時出力し、比較はすべて機械照合
(「画面でなんとなく読めた」は使っていない)。

## 4. 生成過程で得た実装知見(製品にも効く)

- **PDF の埋め込みテキストは互換コードポイントで返ることがある。**
  Skia + Yu Gothic の born-digital PDF で、月(U+6708)が康熙部首の ⽉(U+2F49)として
  抽出された。見た目は同一。製品の抽出パイプラインは **NFKC 正規化を必須**にする。
  本ベンチマークでは GT・抽出の両側へ等しく NFKC を適用。
- **tabula-sharp は連続する同一文字を重複除去で潰す**(188→18、11,236→1,236)。
  CJK フォントのグリフ幅の扱いに起因するとみられる。対策 =「構造は Tabula、
  セル内の文字は PdfPig の letter を bbox で詰め直す」hybrid で完全に回避できた。
- Tabula(spreadsheet)は表の外枠を「1 巨大セルの行」として返すため、その行の除去が必要。

## 5. 実測結果

exact = 空白除去 + NFKC 後の完全一致。すべて機械照合。

### 5.1 born-digital(OCR 不要の層)

| 測定 | 結果 | 速度 | 備考 |
|---|---|---|---|
| PdfPig 文章(10p × 4 フィールド) | **field exact 100%**・読み順 10/10 | 0.03 s/p | 会社名・金額・電話・日付とも完全 |
| Tabula 罫線表 + PdfPig 詰め直し hybrid(410 行 1,640 セル) | **cell exact 100%**・行数 10/10・ヘッダー 10/10 | 0.10 s/p | Tabula 素のままだと重複除去バグで 1.6% |
| Tabula 罫線なし(basic + 詰め直し) | cell exact 60% | 0.04 s/p | 列の矩形が合わない。単体では不採用 |
| 自前 header-guided 再構成(罫線あり) | **cell exact 100%** | 0.01 s/p | ヘッダー文字位置から列を決める方式 |
| 自前 header-guided 再構成(罫線なし) | **cell exact 100%** | 0.01 s/p | 罫線の有無に依存しない |

→ 目標 A(born-digital text ≥99.9%)・B(table cell ≥99%)を**達成**。

### 5.2 スキャン(PaddleOCR、CPU ローカル推論)

PP-OCRv5 ChineseV5 = 多言語(簡体・繁体・英・日)。japan_PP-OCRv4 = 日本語専用(2.x スタック)。

| 測定 | ChineseV5 | japan_PP-OCRv4 | Tesseract 5 (jpn fast) |
|---|---|---|---|
| 文章 clean 300dpi(20 フィールド) | 75%(落ちたのは会社名のみ) | **100%** | - |
| 文章 degraded 150dpi | 75% | 60% | - |
| 帳票 clean ×120p(813 フィールド) | **98.6%** | 84.9% | 80.3%(×30p) |
| ├ 店舗コード(英数字) | **100%** | 15% | - |
| ├ 日付 / 売上 / Q1 / Q2(数値) | **100% / 100% / 100% / 100%** | 100/96.7/95/94.2% | - |
| ├ 担当者(氏名) | 97.5% | 96.7% | - |
| └ 備考(かな漢字) | 91.4% | **100%** | - |
| 帳票 degraded ×30p | 87.7% | 73.4%(売上 10%) | 42.4% |
| 速度(clean 帳票) | 2.7 s/p | 1.4 s/p | 1.1 s/p |
| ピークメモリ | ~850MB | ~950MB | ~215MB |

**engine の性格が相補的**という結果:
- ChineseV5: 数値・英数字コード・日付・金額 = 全 120 ページで完全。ただし日本語の
  一部の字(支・一・月・カレーの「カレ」・促音ッ・長音ー・単→单)を
  **高 confidence のまま脱落・置換**する。
- japan v4: かな漢字は完全に近いが、英数字コード(S118-74)が 15% まで崩れる。
- Tesseract: 全面的に劣後(clean 80.3% / degraded 42.4%)。baseline として記録のみ。

→ 目標 C(clean scan 重要 field ≥98%)は ChineseV5 で**達成**(98.6%、数値系 100%)。
→ 目標 D(degraded ≥90%)は 87.7% で**わずかに未達**だが、§6 のとおり
  誤確定 0 のまま全 field を「自動確定 or 要確認」に分けられるため、条件付きで実用候補。

### 5.3 固定帳票の領域指定 OCR(§13 の比較)

| 方式 | clean ×120p | degraded ×30p | 速度 |
|---|---|---|---|
| A 全ページ OCR(v5) | **98.6%** | **87.7%** | **2.7 s/p** |
| B 領域切り出し + rec のみ(余白 21pt) | 87.3% | 54.2% | 3.3 s/p |
| B' 同(余白 16pt) | 59.5% | 37.9% | 3.6 s/p |
| B'' 領域切り出し + det+rec | 97.4%(6/7 フィールド 100%) | 44.8% | 6.4 s/p |

結論(予想と逆): **全ページ OCR が精度・速度とも最良**。
- rec 単体はクロップ余白 5pt の違いで 87%→60% と挙動が変わり、実用に耐えない。
- det+rec をクロップへ掛ければ clean では同等精度だが、呼び出し回数のぶん 2.4 倍遅い。
- 劣化版(±1〜3° の傾き)では固定座標そのものが外れ、領域方式は全滅する。
  領域指定を使うなら deskew(傾き補正)が前提。
- 製品は「全ページ OCR + ラベルの右側をアンカーにした field 対応付け」を本命とする。

### 5.4 スキャン表(300dpi・罫線あり・41 行 × 4 列 × 10p)

| 方式 | cell exact | 行数一致 | 速度 |
|---|---|---|---|
| SLANet(表構造モデル)+ OCR | **0%**(列格子が 1 列ずれ全滅) | 3/10 | 5.1 s/p |
| **罫線格子(OpenCV)+ OCR 割当** | **94.4%** | **10/10** | 11.5 s/p |

- SLANet は中身を読めているのに構造が崩れる(長い表で列を 1 本落とす)。不採用。
- 罫線をモルフォロジーで抽出して格子を作り、OCR 領域を中心座標で割り当てる
  古典的な方式は、行数 41/41 を全ページ再現。残り 5.6% の誤りは
  **すべて ChineseV5 の日本語弱点**(商品名列。数値・コード列はほぼ完全)。
  日本語 rec との二重読みで解消できる見込み。
- 罫線なしのスキャン表は未実測。**現時点の不得意**として明示する。

### 5.5 チェックボックス(画素解析・OCR 不使用)

| 条件 | exact | 速度 |
|---|---|---|
| clean 300dpi ×120p(チェック / 塗り / 丸囲み / 無印) | **100%** | 0.19 s/p |
| degraded(傾きあり) | 66.7% | 0.05 s/p |

→ 目標 F(≥99%)は clean で**達成**。傾きに弱いので deskew が前提(§5.3 と同じ)。

### 5.6 手書き風(characterization のみ・GO 条件外)

教科書体 + 文字ごとの揺らぎの合成データでは v5 / v4 とも field exact 100%。
これは**合成が機械認識に甘すぎる**ことを意味する(フォント由来の字形は保たれるため)。
実際の手書きへの一般化はできない。実案件の手書きは対象外のままとする。

## 6. Confidence(誤確定の実態)

ChineseV5・clean 帳票(813 フィールド)での閾値スイープ:

| 閾値 | 自動確定 | 要確認 | **誤確定(間違いを自動確定)** |
|---|---|---|---|
| 0.90 | 99.0% | 1.0% | 0.9% |
| 0.95 | 94.7% | 5.3% | 0.1% |
| **0.98** | **81.9%** | 18.1% | **0%** |

- 間違い 11 件の confidence 最大値は 0.966 → 0.98 閾値なら**誤確定ゼロ**で
  82% を自動確定 + 18% を人間確認に回し、100% のフィールドを確認可能状態にできる。
  劣化版でも @0.98 誤確定 0(自動確定 32%)。
- japan v4 は較正が甘く、@0.98 でも誤確定 3.4% 残る。**gate には v5 の confidence を使う**。
- Tesseract は meanConfidence が clean .83 / degraded .45 と判別力はあるが、素の精度が不足。

## 7. 自動 routing

PdfPig で 1 ページ目の「文字数・画像被覆率・細長い水平描画の本数」を見る判定:

- 8 ケース中 7 正解。スキャン 4 種は 4/4 で正しく scan 判定。
- 唯一の外れ: born-digital の帳票(記入欄の下線 = 罫線)を "table" と判定。
  埋め込みテキスト側へ routing される点は同じなので**影響は良性**。
- 「これはスキャン PDF ですか?」と利用者に聞かない設計は成立する。
  routing: 埋め込みテキストあり → PdfPig(+表なら罫線 hybrid)/ なし → render → OCR。

## 8. 性能・サイズ・依存

### 速度・メモリ(CPU のみ、Ryzen 級 1 台)

- born-digital: 0.01〜0.1 s/p・~130MB — 体感ゼロ
- OCR(v5): 2.7 s/p(22 p/min)・ピーク ~850MB。120 ページの帳票 = 5.4 分
- OCR(v4 2.x): 1.4 s/p(43 p/min)
- 罫線表格子 + OCR: 11.5 s/p(表は OCR 領域が 160+/p あるため)

### 配布サイズ(実測)

| 構成 | サイズ |
|---|---|
| 現行製品(self-contained) | 168.6 MB |
| + PDF 層のみ(PdfPig + Tabula + pdfium + Skia) | **+ 約 25 MB**(pdfium 6.9 / Skia 11.7 / 管理 DLL) |
| ベンチ全部入り self-contained 実測 | 610.9 MB |
| うち削れるもの(pdb 85 + ffmpeg videoio 27.3 + 重複モデル 20.8) | -133 MB |
| **OCR 込み現実推定(v5 スタック)** | **~478 MB** |
| Paddle runtime nupkg: 3.x = 359MB / **2.x = 154MB** | 2.x なら合計 ~460MB 級 |
| モデル単体(det 4.7 + cls 2.1 + japan rec 9.6 + SLANet 9.9) | ~26 MB |

主要内訳: paddle_inference_c 88.7 / mklml 88.4 / mkldnn 45.8 / phi 43.9 /
OpenCvSharpExtern 65.4(+ffmpeg 27.3 は削除可)/ v5 モデル 20.8。

→ 1GB は超えないが、**「本体(~195MB)+ Offline OCR Pack(~300MB 台)」の分割**が現実的。
   PDF 層(born-digital のみ)は本体同梱でよい軽さ。

### 依存・オフライン(§4 / §19)

- Python / Java / system PATH / 初回ネット DL: **すべて不要**(NuGet native 同梱で完結)
- モデルは同梱可能(Apache-2.0)。ベンチ中の DL は tessdata 2.4MB と japan v4 26MB のみ
- VC++ ランタイム: Paddle / OpenCV / Tesseract の native DLL は vcruntime 前提。
  Windows 11 は通常搭載済みだが、**app-local 同梱(再配布可)を製品要件にする**
- ネットワーク送信・cloud OCR・telemetry・常駐 server: なし(検証コードにも無い)

## 9. GO / NO-GO と推奨構成

### 判定: **GO(条件付き)** — 4 系統すべて実用候補

| 系統 | 判定 | 根拠 |
|---|---|---|
| 1. PC 生成 PDF | **実用候補** | field exact 100%・NFKC 正規化必須という条件も判明済み |
| 2. PC 生成の表 PDF | **実用候補** | cell exact 100%(hybrid / 自前再構成の両方) |
| 3. 印刷文字スキャン | **実用候補** | 帳票 98.6%・数値系 100%・誤確定 0 の閾値運用が成立 |
| 4. 同一レイアウト大量帳票 | **実用候補** | 120p を 5.4 分・自動確定 82% + 要確認 18% で全量確認可能 |
| 5. 複雑なスキャン表 | 条件付き | 罫線ありは 94.4%(格子方式)。罫線なしスキャン表は未実測 = 不得意 |
| 6. 手書き | 対象外 | characterization のみ。合成では測れないことも含めて記録 |

### GO の条件(これを外すと NO-GO 側に倒れる)

1. **日本語のかな漢字は v5 単独に依存しない。** v5 は「支・一・カレー」等を高 confidence の
   まま落とす。対策 = v5(数値・コード担当)+ japan rec(かな漢字担当)の二重読み照合、
   または confidence gate(v5 @0.98)で要確認へ回す運用。単一エンジン採用は不可。
2. **確定は 3 分類 UI(自動確定 / 要確認 / 読取不能)を前提にする。** 勝手に Excel へ
   確定しない(§20)。誤確定 0 の閾値が実測で存在することが根拠。
3. **傾き補正(deskew)を入れるまで、領域指定 OCR とチェックボックスは clean スキャン限定。**
4. 日本語専用 rec を使うには **Paddle 2.x runtime に固定**するか(SLANet も同様)、
   将来 PaddleSharp が新形式の japan モデルを載せるのを待つ。3.x runtime は旧形式を
   実行できない(実測: OneDnnContext does not have the input Filter。Openblas でも不可)。

### 推奨製品構成(実装するなら)

```
PDF 選択
 ↓ 自動判定(文字数・画像被覆・罫線)
 ├ 埋め込みテキストあり
 │   ├ 表: 罫線 hybrid(構造=Tabula/罫線、文字=PdfPig letter 詰め直し)
 │   │      罫線なしは header-guided 自前再構成      … cell 100%
 │   └ 文章/帳票: PdfPig + NFKC                      … field 100%
 └ スキャン
     ├ PDFium 300dpi render
     ├ PP-OCRv5(数値・コード)+ japan rec(かな漢字)の二重読み
     ├ v5 confidence @0.98 で 自動確定 / 要確認 / 読取不能
     ├ 表: 罫線格子 + OCR 割当(SLANet 不採用)
     └ チェックボックス: 固定座標の画素解析(deskew 前提)
配布: 本体 + Offline OCR Pack(分割)。完全オフライン。実行時 DL なし。
```

### 残る不得意(正直に列挙)

- 罫線なしのスキャン表(未実測)
- 手書き(対象外)
- 傾いたスキャンでの領域指定・チェックボックス(deskew 未実装)
- v5 と v4 の二重読み照合は設計提案であり、統合精度は未実測
- degraded で「自動確定率」は 32% まで落ちる(誤確定 0 は維持)

### 案件対応への見込み

- born-digital PDF → Excel/CSV: 即応可能水準(2E の CSV 変換と直結)
- 帳票スキャンの転記(同一レイアウト大量): 実用候補。100 ページ級を数分 + 2 割の目視確認
- スキャン表: 罫線ありに限り対応可能と案内できる
- 手書き帳票: 受けない(明示)

