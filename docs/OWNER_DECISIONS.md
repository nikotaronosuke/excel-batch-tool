# Owner Decision Log

日本語 | [English](OWNER_DECISIONS.en.md)

Excel Batch Tool で判断記録として残すのは、実測や失敗で方針が変わった5件に絞ります。

## 1. まず編集機能ではなく read-only analyzer を作った

一括処理で最初に避けたかったのは「自動化できないこと」ではなく、元ファイルを壊すことでした。

Phase 0 は編集機能を持たず、`FileAccess.Read` で開く safety analyzer から開始。解析前後で SHA-256 / size / mtime が変わらないことをテストしました。

**Evidence:** [Phase 0](https://github.com/nikotaronosuke/excel-batch-tool/commit/da461404d10fdde32ec54c061bf50a8369107900)

## 2. 実案件調査を見てロードマップを入れ替えた

当初の機能候補をそのまま作るのではなく、CrowdWorks / Lancers の実案件を見て、

- 複数セル定型入力
- 表→帳票転記
- キー突合
- CSV変換

を優先しました。

その結果、CSV変換は検索・置換より先に実装。VBA編集は案件があっても製品へ統合しませんでした。

**Evidence:** [market research](research/market-summary.md) / [CSV transform](https://github.com/nikotaronosuke/excel-batch-tool/commit/ce15518b09babd84589b9de92886f6f7592b1529)

## 3. 10万行で測って validation 方針を変えた

転記元10万行で、

- full validator: **約26秒 / 約1.4GB**
- streaming scan: **約5.4秒 / 約1MB**

でした。

読むだけの source と、書き換える target に同じ検証をかけるのをやめ、source は意味解釈に必要な部分を streaming で確認、target は full validation を維持しました。

**Evidence:** [Phase 2C1](https://github.com/nikotaronosuke/excel-batch-tool/commit/265631084e781d7c02e4d5b55f76676c75d6606e)

## 4. PP-OCRv5 を同条件で測って採用しなかった

v4二重読みと v5 + japan v4 を同じ製品経路で比較すると、480項目の差はほぼ無い一方、v5構成は **2.4〜2.9倍遅い**結果でした。

新しいモデルという理由では採用せず、v4系を維持しました。

**Evidence:** [Phase 2F-B3](https://github.com/nikotaronosuke/excel-batch-tool/commit/21db92d13d5a75a5601bcc631524ebfa3b1dbd56)

## 5. deskew は角度推定が合っていたのに、補正方向が逆だった

傾き角の推定値はほぼ正しく見えたため、一度は処理も正しいと思いました。

実際の出力を見ると回転方向が逆で、罫線が崩れていました。最終画像まで確認するテストへ修正し、傾いた帳票 / 表の結果を再評価しました。

中間指標が正しいことと、製品結果が正しいことを同一視しない教訓になりました。

**Evidence:** [Phase 2F-B3 deskew fix](https://github.com/nikotaronosuke/excel-batch-tool/commit/21db92d13d5a75a5601bcc631524ebfa3b1dbd56)
