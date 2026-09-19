# Owner Decision Log

Excel Batch Tool は、「Excel の機能をできるだけ多く自動化する」ことを目的にしていません。

実案件を調べながら、**大量・反復作業を減らしつつ、分からないものを勝手に補完せず、元ファイルを壊さないこと**を優先してきました。

この文書では機能一覧ではなく、プロジェクトオーナーとして「何を優先し、何を止め、何を採用しなかったか」が分かる判断だけをまとめます。

---

## 1. 最初に作ったのは編集機能ではなく、読み取り専用の安全解析

### 課題

Excel の一括処理で一番避けたかったのは、効率化より先に**元ファイルを壊すこと**でした。

数式・結合セル・図形・グラフ・画像・ピボット・外部参照・保護・入力規則などを持つ Workbook は、単純なセル書き換えでも意味を壊す可能性があります。

### 判断

Phase 0 は編集機能を一切持たない **read-only analyzer** から始めました。

入力ファイルは OS レベルで書き込めない `FileAccess.Read` で開き、解析前後に

- SHA-256
- ファイルサイズ
- 更新時刻

が一致することをテストしています。

「まず書けるようにして、危なければ後で制限する」ではなく、
**先に危険を見つける層を作ってから書き込み機能へ進む**順番を選びました。

**Evidence:** [Phase 0: read-only workbook safety analyzer](https://github.com/nikotaronosuke/excel-batch-tool/commit/da461404d10fdde32ec54c061bf50a8369107900)

---

## 2. Excel本体やCloudではなく、完全ローカル + Open XML を選んだ

### 課題

対象にしたかったのは、毎月同じ転記・統合・変換をしている事務作業や、クラウドソーシングの実作業です。

その用途では、

- Excel が入っていないPC
- 外部サービスへ業務データをアップロードできない環境
- マクロやスクリプトを保守できない利用者

も想定する必要がありました。

### 判断

製品の前提を、

- Windows desktop
- Excel本体不要
- 完全オフライン
- アプリ自身のネットワーク通信なし
- Microsoft Open XML SDK を中心に処理

としました。

ClosedXML のように Workbook 全体を再構築して保存する方式より、
「触る必要のないPartへ触れない」方針を取りやすい Open XML SDK を選んでいます。

機能数より **ファイル完全性とローカル完結** を中心価値にしました。

**Evidence:** [v1 Overview](spec/v1-overview.md) / [Design Decisions D-002〜D-004](decisions.md)

---

## 3. 対応できない要素は、黙って落とさず Block する

### 課題

Worksheet を別Workbookへ集約するとき、
値や書式だけコピーして「一応開けるファイル」を作ることはできます。

しかし元シートに、

- 数式
- Chart / Drawing / Image
- Table / PivotTable
- 条件付き書式
- 入力規則
- コメント
- ハイパーリンク
- 外部参照

などがあった場合、対応していない要素だけ消えても利用者は気づかないかもしれません。

### 判断

**保持できないものが1つでもあれば、理由を出して事前に止める**方針にしました。

対応できる分だけコピーして、未対応分を黙って捨てる経路は作りません。

後から入力規則や条件付き書式などを対応するときも、
「対応できるルールだけ部分コピー」は避け、未対応要素が混じればシート全体を Block する方針を維持しています。

**Evidence:** [Phase 1B.1: unsupported elements block instead of disappearing](https://github.com/nikotaronosuke/excel-batch-tool/commit/827f2e4c9ea8b0b0454408e30c7b2ee8fc0d6e95)

---

## 4. 元ファイルを直接変更せず、「copy → edit → validate → move」にした

### 課題

既存 Workbook を直接開いて変更すれば実装は簡単です。

ただし途中で例外・電源断・ファイルロックなどが起きると、
元ファイルそのものを壊す可能性があります。

### 判断

変更系でも Source は最後まで読み取り専用です。

```text
Source
  ↓ byte copy
temporary copy
  ↓ edit
re-open + validate
  ↓
final output
```

出力は常に**別名の新規ファイル**。

- 入力と同じパスは拒否
- 既存ファイルへの上書きも拒否
- 一時ファイルを再オープンして検証
- OpenXmlValidator を通す
- 成功後だけ最終パスへ移動

という順にしています。

rollback も「絶対に成功する」とは言わず、消せなかった可能性があればその事実を利用者へ伝えるようにしました。

**Evidence:** [Phase 1A: new-file-only output](https://github.com/nikotaronosuke/excel-batch-tool/commit/b846d359680cc3b5f8edf47118ff3ee0ee46ca24) / [docs/decisions.md D-023〜D-024](decisions.md)

---

## 5. 「安全な分だけ実行する」より、1件でも危なければバッチ全体を止める

### 課題

100個の変更のうち99個が安全で、1個だけ危険な場合、
99個だけ処理してしまう方が便利に見えます。

しかし利用者は「全部適用された」と思う可能性があり、
部分適用は結果の解釈を難しくします。

### 判断

複数セル一括変更では、**1セルでも guard に引っかかれば実行全体を Block** します。

同じセルへ2つの値が指定された場合も、

- 最初を採用
- 最後を採用

のような自動解決はしません。

貼り付け入力でも1行だけ不正なら、その行だけ無視せず入力全体を拒否します。

「できるところだけ進める」より、**何が実行されたかを利用者が確実に理解できること**を優先しました。

**Evidence:** [Phase 2B: one unsafe operation blocks the batch](https://github.com/nikotaronosuke/excel-batch-tool/commit/cba1b9ac0bd7a18fa039dcf738a560f0f5ab0e51)

---

## 6. キー照合で「賢い名寄せ」をしなかった

### 課題

転記・突合では、

- `00123` と `123`
- `ABC` と `abc`
- 前後空白
- 全角 / 半角

を自動で同一視すると便利そうに見えます。

しかし業務データでは、見た目が似ていても別IDであることがあります。

### 判断

キーは **文字列として完全一致** だけを採用しました。

- Trimしない
- case foldingしない
- 全角半角を揃えない
- Unicode正規化しない
- 数値へ変換しない
- 0埋めを補わない

一致しなければ、利用者へ「一致しない」と見せます。

0件一致・複数件一致も勝手に選ばず Block。

「人間ならたぶん同じだと思う」をコードに埋め込むより、**誤った行へ転記するリスクを避ける**判断です。

**Evidence:** [Phase 2C1: exact-string key matching](https://github.com/nikotaronosuke/excel-batch-tool/commit/265631084e781d7c02e4d5b55f76676c75d6606e) / [Phase 2C2](https://github.com/nikotaronosuke/excel-batch-tool/commit/72ffe33ab5181e7114af74d3a74257b020bb4687)

---

## 7. ロードマップは「作りたい機能」ではなく、実案件で順番を変えた

### 課題

最初から検索・置換や高度なExcel機能を増やしていくこともできました。

しかし市場調査を見ると、実際に繰り返し発注されていたのは、

- 複数セルへの定型入力
- 表から帳票への転記
- キーでの表突合
- システム間CSV変換
- 大量ファイルの反復処理

でした。

### 判断

**追加開発を想像で広げず、実案件で頻出する不足を先に実装**しました。

その結果、

- 複数セル Input Set
- Source Excel / CSV からの転記
- 表同士のキー突合
- CSV変換

を検索・置換より先へ移しています。

CSV変換を実装したときも、

> あるシステムから出した列構成を、別システムが要求する列名・順番へ変える

という実務ギャップを優先し、検索・置換を後ろへ回しました。

VBA編集も、案件があるからという理由だけで製品へ取り込まず、v1では明示的に対象外にしています。

**Evidence:** [market research summary](research/market-summary.md) / [Phase 2E: CSV transform before search/replace](https://github.com/nikotaronosuke/excel-batch-tool/commit/ce15518b09babd84589b9de92886f6f7592b1529)

---

## 8. 大規模データでは「同じ検証を全部にかける」より、目的ごとに検証範囲を変えた

### 課題

安全性を重視するなら、すべてのWorkbookへ完全な OpenXmlValidator をかけるのが一見正しく見えます。

しかし 10万行の転記元を実測すると、

- 全体 Validator: 約26秒 / 約1.4GB
- streaming scan: 約5.4秒 / 約1MB

でした。

転記元は**読むだけ**、転記先は**書き換える**ので、そもそも守る対象が違います。

### 判断

転記元は、値の意味に必要な

- sheet list
- shared strings
- styles

などを確認し、読めないセルはその場で Block。

転記先は従来どおり全体 Validator を維持しました。

「安全だから全部同じ検証」ではなく、**何から何を守る検証なのかを分け、実測でコストを見て決める**方針です。

**Evidence:** [Phase 2C1 performance and validation boundary](https://github.com/nikotaronosuke/excel-batch-tool/commit/265631084e781d7c02e4d5b55f76676c75d6606e)

---

## 9. PDF / OCR は「文字が取れた」ではなく、納品可能な完全一致率でGO/NO-GOした

### 課題

PDF案件が多いからという理由だけでOCRを製品へ入れると、
「読めるけれど納品には使えない」機能になる可能性があります。

### 判断

まず製品本体とは別に feasibility benchmark を作りました。

評価したのは、

- born-digital text
- born-digital table
- clean / degraded scan
- fixed form
- ruled / borderless table
- checkbox / mark
- handwriting-like characterization

など。

Ground Truthもfixture生成時に一緒に作り、**画面でなんとなく読めたかではなく exact match** で判定しました。

手書き風fixtureが100%読めても、
「合成フォントが簡単すぎるだけで実手書きへ一般化できない」と判断し、手書きは対象外のままにしました。

PDF/OCRを製品へ入れたのは、このベンチマークが条件付きGOになってからです。

**Evidence:** [Phase 2F-R: PDF feasibility benchmark](https://github.com/nikotaronosuke/excel-batch-tool/commit/3994ab4f7dbce92ac2f531bf886b4a3f361c5d99) / [benchmark report](research/pdf-feasibility-research.md)

---

## 10. OCRは「自動確定率」より false AutoAccepted 0 を優先した

### 課題

OCRでは confidence が高くても間違います。

実測では、

- `S001-24` → `SO01-24`
- `A0096` → `9600`

のように、2つのモデルが同じ間違いを高confidenceで返す例がありました。

単純にconfidence閾値を上げても、この種の誤りは消えません。

### 判断

結果を

- 自動確定
- 要確認
- 読取不能

へ分け、疑わしいものは**値を自動修正せず、人へ回す**ことにしました。

さらに、

- 項目種別と文字の形
- 同じ列での文字種パターン
- 上下反転で誤読しやすい数字
- 2行が1セルへ連結された形

などを auto-accept gate に使っています。

最終統合では、実案件相当の架空fixture群で **false auto-accept 0** を維持しました。

自動確定率を上げるために誤確定を許容するより、
**確認の手間が増えても誤った値を勝手に納品物へ入れない**方を選びました。

**Evidence:** [Phase 2F-B3: zero false auto-accepts](https://github.com/nikotaronosuke/excel-batch-tool/commit/21db92d13d5a75a5601bcc631524ebfa3b1dbd56)

---

## 11. 「新しいモデルだから採用」ではなく、同じ条件で測ってPP-OCRv5を却下した

### 課題

新しいOCRモデルを使えば品質が上がるように見えました。

### 実測

現在の製品経路で v4二重読み と v5 + japan v4 を比較すると、

- 帳票でほぼ同じ exact rate
- 480項目中の差は1項目
- v5構成は **2.4〜2.9倍遅い**
- Packも大きくなる

という結果でした。

### 判断

**PP-OCRv5は採用しませんでした。**

精度改善の大部分は新モデルではなく runtime 更新によって得られており、
同じ v4モデルでも大きく改善しました。

新しい技術へ置き換えること自体を成果にせず、
**利用者が得る改善量と処理時間の比**で選びました。

**Evidence:** [Phase 2F-B3: PP-OCRv5 measured and rejected](https://github.com/nikotaronosuke/excel-batch-tool/commit/21db92d13d5a75a5601bcc631524ebfa3b1dbd56)

---

## 12. 中間指標が正しくても、最終結果を確認するまで信用しない

### 課題

傾き補正では、角度推定自体はほぼ正しく動いていました。

そのため当初は「推定誤差が小さい = deskewも正しい」と考えていました。

しかし実際には、**補正方向が逆**でした。

```text
measured skew ≈ 1.98°
wrong rotation  → grid disappears
opposite rotate → grid restored
```

### 判断

角度推定の精度ではなく、**補正後の表が本当に戻ったか**まで検証するようにしました。

修正後、

- 傾いた罫線表: 17.5% → 88.9%、false auto-accept 0
- 傾いた帳票: 74.2% → 94.2%

まで改善しました。

「測定値が合っている」ことと「製品結果が正しい」ことを同一視しない、という判断です。

**Evidence:** [Phase 2F-B3: fix reversed deskew](https://github.com/nikotaronosuke/excel-batch-tool/commit/21db92d13d5a75a5601bcc631524ebfa3b1dbd56)

---

## このプロジェクトで優先したもの

Excel Batch Toolでは、対応範囲の広さより次を優先しています。

- 元ファイルを直接変更しない
- 分からない要素を黙って捨てない
- 部分成功を成功に見せない
- 自動名寄せや型変換で「たぶん同じ」を作らない
- 市場の実案件を見て実装順を変える
- 読み取りと書き換えで安全性の意味を分ける
- OCRは完全一致と誤確定で評価する
- 人間確認を失敗ではなく安全機構として扱う
- 新しいモデルでも実測で利点が無ければ採用しない
- 中間指標ではなく最終出力まで確認する

AIを使って実装していますが、
このリポジトリで重要なのはコード量ではなく、**どこまで自動化し、どこで意図的に止めたか**です。
