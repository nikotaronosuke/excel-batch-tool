# DECISIONS(append-only)

重要な設計判断を番号付きで記録する。過去の項目は書き換えず、変更する場合は
新しい番号で「D-xxx を更新」として追記する。

---

## D-001: Private spec / Public implementation の2リポジトリ構成 (2026-08-27)

- Private `excel-batch-tool-spec`: 仕様・調査・判断・現在地の正本。
- Public `excel-batch-tool`: source / tests / 公開 docs の正本。
- Public から Private の名前・URL・内部パスを参照しない。
  Private から Public commit SHA を参照するのは可。
- 理由: 市場調査・内部判断を公開せずに、実装のみをポートフォリオとして
  公開するため(既存 AI Problem-Solving Memory と同運用)。

## D-002: 技術スタックは .NET 8 + WPF + Microsoft Open XML SDK (2026-08-27)

- Windows 専用デスクトップ。C# / .NET 8 / WPF。
- Excel 読み書きは Microsoft Open XML SDK (`DocumentFormat.OpenXml`) を中心にする。
- Office Interop・Excel 本体依存・クラウド・Web API・外部 AI API・ログイン・
  アプリ自身のネットワーク通信は禁止。
- ClosedXML を主エンジンにしない。理由: ClosedXML は Workbook 全体を自前の
  オブジェクトモデルへ読み込み保存し直すため、未対応要素(Drawing / Chart /
  一部拡張)が保存時に欠落・破壊されるリスクがある。本製品の中心価値は
  「壊す可能性がある Workbook には安易に書き込まない」ことなので、
  パッケージ構造をそのまま扱える Open XML SDK を採用する。
- 詳細比較は docs/research/technology-summary.md。

## D-003: Phase 0 の読み取り専用保証は「物理的に書けない開き方」で実装 (2026-08-27)

- 対象ファイルは `FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)`
  で開き、そのストリームを `SpreadsheetDocument.Open(stream, isEditable: false)`
  に渡す。FileAccess.Read のため、コード上のバグがあっても OS レベルで
  書き込み不能。
- Phase 0 には保存 API・変更 API を一切実装しない(Core に Write 系メソッドを
  置かない)。
- 検証: 解析前後で SHA-256 / ファイルサイズ / LastWriteTimeUtc が一致することを
  自動テストで確認する(Public tests の FileImmutabilityTests)。

## D-004: Phase 0 の安全性分類ルール (2026-08-27)

Workbook 単位の分類は、検出された要素の最大深刻度で決める。

### ✖ 現在非対応(UnsupportedForNow)

| 要素 | 根拠 |
|---|---|
| マクロ / VBA(vbaProject パート、.xlsm 等) | v1 は VBA 編集対象外。書換時に VBA との整合を保証できない |
| 開けないファイル(壊れた zip / 暗号化・パスワード保護 / .xlsx 以外) | 解析自体が不可能。暗号化ファイルは OPC として開けない |

### ⚠ 注意が必要(NeedsAttention)

「将来の書換時に、単純なセル値書換だけでは壊れる・ずれる・意味が変わる
可能性がある要素」を持つ Workbook。書換不可という意味ではない。

| 要素 | 根拠 |
|---|---|
| 数式 | 書換後にキャッシュ値と数式の不整合が起きうる。参照先の移動にも追従が必要 |
| 結合セル | 行列の挿入・転記で範囲がずれる。値の書込位置が直感と異なる |
| Drawing / 図形 | アンカー(セル基準位置)が行列変更でずれる。未知の図形要素を壊しやすい |
| Chart | データ範囲参照を持つ。シート名変更・範囲変更で壊れる |
| 画像 | Drawing と同様にアンカーずれの可能性 |
| PivotTable | ソース範囲・キャッシュを持ち、単純書換で不整合になる |
| External Link(外部ブック参照) | 参照先が手元にない場合、値の意味が変わる |
| Sheet / Workbook 保護 | 書換がブロックされる、または保護を壊す可能性 |
| テーブル(ListObject) | 範囲定義を持ち、行追加・削除で定義とずれる |
| データ入力規則 | 適用範囲が行列変更でずれる |
| 条件付き書式 | 適用範囲・相対参照が行列変更でずれる |
| コメント(legacy / スレッド) | アンカーセルのずれ、パート間参照 |
| 定義名 | 範囲参照を持ち、シート・範囲変更で壊れる |
| ハイパーリンク | 適用範囲ずれ、外部参照 |
| OLE 埋め込み / ActiveX | 未対応バイナリを含み、書換時に壊しやすい |
| Custom XML パート | 他システム連携データの可能性があり、保持が必要 |

### ✅ 通常(Normal)

上記のいずれも検出されない Workbook(値・書式のみのデータ表など)。

### 補足

- 分類は「Phase 1 以降の書換系機能で追加の注意・処理が必要か」を
  ユーザーに事前提示するためのもの。Phase 0 では全ファイルが読み取り専用。
- 数式・結合セルを ⚠ に入れるため、実務ファイルの多くは ⚠ になる。
  これは仕様(「変更前に確認できる」が中心価値)であり、⚠ は危険という
  意味ではなく「書換時に確認が要る」の意味で UI 文言を設計する。

## D-005: Public リポジトリは MIT License (2026-08-27)

- 依存パッケージのライセンス: DocumentFormat.OpenXml = MIT、
  xunit = Apache-2.0、xunit.runner.visualstudio = Apache-2.0、
  Microsoft.NET.Test.Sdk = MIT。いずれも MIT での公開と両立する
  permissive license であることを確認(実際の解決版は docs/implementation/phase-log.md 参照)。
- よって Public repo は MIT License を採用。

## D-006: git author は GitHub noreply アドレス (2026-08-27)

- 両リポジトリとも repo-local 設定で
  `nikotaronosuke <242803730+nikotaronosuke@users.noreply.github.com>` を使う。
- 理由: Public の git history に個人メールアドレスを残さない
  (既存リポジトリと同じ慣例)。

## D-007: .NET 8 SDK はユーザー領域へ導入 (2026-08-27)

- この開発マシンには dotnet SDK が無かったため、Microsoft 公式
  dotnet-install.ps1 で `%LOCALAPPDATA%\Microsoft\dotnet` へ
  .NET 8 SDK をユーザー領域インストールした(システム全体へは非導入)。
- ビルド時は PATH 先頭に同ディレクトリを追加し `DOTNET_ROOT` を設定する。

## D-008: Phase 0 の使用範囲・行列数は SheetDimension を第一情報源にする (2026-08-27)

- 使用範囲は worksheet の `dimension`(SheetDimension)を読む。
  無い場合のみ行を走査して概算する。
- 行数・列数は使用範囲からの概算値として表示する(「概算」と明示)。
- 理由: 大量 Workbook 前提で、全セル走査を必須にすると解析が重くなる。
  正確な最終行判定は書換系 Phase で必要になった時に強化する。

## D-009: 解析の並列度と失敗隔離 (2026-08-27)

- 解析はバックグラウンドスレッドで実行し、UI スレッドを塞がない。
- ファイル単位で try/catch し、1 ファイルの失敗は AnalysisStatus.Failed として
  結果に記録、他ファイルの解析は継続する。
- 並列度は `Math.Max(1, Environment.ProcessorCount - 1)` を上限にする
  (UI 応答性を優先)。

## D-010: D-008 の補足 — SheetDimension の役割と Worksheet XML 走査の関係 (2026-08-28)

- D-008 の「SheetDimension を第一情報源にし、無い場合のみ行を走査する」は
  **「使用範囲・概算行列数」の算出に限った話**である。
- Safety Finding(数式・結合セル・シート保護・データ入力規則・条件付き書式・
  ハイパーリンク等)の検出のためには、SheetDimension の有無に関係なく
  **Worksheet XML 自体を OpenXmlReader でストリーミング走査する**(1 パス)。
- その際も DOM の全展開(Worksheet プロパティ経由のロード)は行わず、
  大量 Workbook・巨大シートでのメモリ使用を抑える。
- 行走査で得た最大行・spans 由来の最大列は、SheetDimension が無い場合の
  概算行列数のフォールバックとして同じ 1 パス内で流用する。
- 本 Decision は D-008 を書き換えず、その意味を補足するものである
  (DECISIONS は append-only)。

## D-011: 配布は self-contained win-x64 とし、Runtime 事前導入を要求しない (2026-08-28)

- 事象(2026-08-28 確認): framework-dependent な WPF exe を直接実行した際、
  システム側に .NET Desktop Runtime が無いため起動できず、
  Microsoft の Runtime ダウンロードページが自動で開いた。
- 判断: これはアプリの異常動作ではなく、framework-dependent build の
  実行要件(ホストが共有 Runtime を要求する挙動)によるもの。
- 製品方針: **配布版は self-contained Windows x64** とし、
  最終利用者に .NET Desktop Runtime の別インストールを要求しない。
  - RuntimeIdentifier: win-x64 / SelfContained: true
  - publish 設定は Public repo の
    `src/ExcelBatchTool.App/Properties/PublishProfiles/win-x64-self-contained.pubxml`
    に保持する(build/publish 成果物は commit しない)
- 単一ファイル化(PublishSingleFile)は必須にしない。フォルダー形式の
  self-contained publish を検証済みの基準構成とする(単一ファイル化の採否は
  配布方法を決める段階で再判断。採用時は本ファイルへ追記)。
- 開発環境メモ: この開発マシンは SDK がユーザー領域のみのため、
  framework-dependent exe の直接起動には DOTNET_ROOT が必要(D-007)。
  self-contained publish 成果物はこれに依存しないことを検証して配布判断の根拠とする。

## D-012: Phase 1A(表データの縦結合)の仕様と Safety Policy (2026-08-28)

Phase 1 は 1A(縦結合)と 1B(Sheet 集約)に分割する。本 Decision は 1A のみ。

### 1A のスコープ

- 複数 Workbook の Header 付きの表を、1 つの**新規** Workbook /
  1 Worksheet(既定名「統合結果」)へ縦結合する。1 Workbook につき 1 選択 Worksheet。
- 対象は .xlsx の Worksheet のみ(Chartsheet / MacroSheet / Dialogsheet は選択対象外)。
- 対象外(先回り実装しない): 既存 Workbook への追記、見た目のままの Sheet 複製、
  Workbook 内複数 Sheet の別 Sheet 集約、列名の手動 mapping UI、重複削除、
  表記揺れ修正、CSV、PDF、OCR、数式再計算、書式の完全コピー、
  Chart / Drawing / Image のコピー、Recipe 保存、永続的な処理履歴。

### Header ルール

- 表の 1 行目を Header 行として扱う(自動 Header 推定はしない)。
- 比較は **trim 後の文字列を Ordinal で厳密比較**。大文字小文字の同一視や
  Unicode 正規化などの「賢い推測」はしない(決定的であることを優先)。
- 空 Header → Block / 同一 Sheet 内の重複 Header → Block /
  不足 Header → Block / 余分な Header → Block。
  Header 名の集合が一致すれば、**列順が違っても** Header 名で対応付けて統合する。
- 出力列順は「最初に Header を読めた選択 Sheet」の Header 順。
- Header 列数は 1 列目から連続する非空セルで決める。その範囲内に空セルがあれば
  空 Header として Block。
- Header より右の列にデータがある場合は Warning(その範囲は統合しない)。

### 数式・結合セルの Block 方針

- 選択 Sheet に数式が 1 つでもあれば **Block**。理由: Open XML SDK は Excel の
  計算エンジンではなく、cached value が最新である保証がないため。
  将来 cached value 採用を検討する余地はあるが 1A では実装しない。
- 選択 Sheet の**表範囲**(1 行目〜最終データ行 × Header 列数)に結合セルが
  かかる場合は **Block**。単純な行データとして意味が曖昧になるため。
- どちらも **Sheet 単位**で判定する。非選択 Sheet に数式・結合セルがあっても、
  選択 Sheet が安全なら Block しない。

### Phase 0 Finding と 1A Safety Policy の関係

- Phase 0 の ⚠(NeedsAttention)は「既存 Workbook を書き換えるときの注意」であり、
  1A は入力を READ ONLY で読み Cell data だけを取り出して**完全に新しい Workbook を
  生成**する処理なので、**⚠ だから一律拒否はしない**。
- Block にするもの: 解析不能 / 対象外ファイル形式 / マクロ関連(Phase 0 ✖)/
  選択 Sheet の数式 / 選択 Sheet の表範囲の結合セル / Header 不一致 /
  空 Header / 重複 Header / 読み取りエラー / metadata 列名の衝突 /
  出力行数が 1 シート上限(1,048,576)超過 / 同一ファイルの重複選択。
- 原則 Block しないもの: Chart / Image / Drawing / 条件付き書式 / 入力規則 /
  コメント / ハイパーリンク / Sheet 保護 / Workbook 保護 / テーブル /
  PivotTable / Custom XML。**これらは出力へコピーしない**(データのみ統合)。
  入力側の保護解除なども一切行わない(入力は読むだけ)。

### metadata 列

- 「元ファイル」「元シート」を出力先頭に追加(初期 ON)。
- 追加する列名が元データの Header と衝突する場合は **Block** し、理由を表示する
  (黙って上書き・自動リネームはしない)。

### Cell value と日付

- 対応: Shared String / Inline String / 通常 String / Number / Boolean / Blank /
  Error(文字列として持ち越し)。Blank セルは出力しない。
- Header 以降の**完全空行は出力しない**。ただし途中の空行で打ち切らず、
  後続のデータ行は維持する(「途中の空行＝表の終了」と推測しない)。
- 日付: style(numFmt)を見て判定する。組み込み日付・時刻 ID と、
  ユーザー定義書式のトークン(y/d/e/g → 日付、h/s/AM/PM/[h] 等 → 時刻)で分類。
  workbook の date1904 を考慮し、serial を **1900 date system へ正規化**(+1462)して
  出力側に日付書式を設定する。
- **月(m)しか判定材料が無い書式は「分」と区別できないため Ambiguous とし、
  日付に誤変換せず数値として出力し Warning を出す**(黙って誤変換しない)。
- serial が範囲外のものも日付にせず数値のまま扱う。

### 出力 Workbook

- 完全新規作成。入力の Style / Drawing / Chart 等はコピーしない。
- 付けるのは Header 行の太字、先頭行固定、AutoFilter、日付表示形式のみ。

### 出力安全設計

- 入力ファイルは常に READ ONLY(FileAccess.Read)。書き込みは一切しない。
- 出力先が**既存ファイル**、または**入力ファイルと同一パス**の場合は Block。
  自動上書きはしない(利用者に別名を選ばせる)。
- 手順: 最終パスと**同じフォルダー**の一時ファイル `~ebt-merge-<guid>.xlsx` に生成
  → 読み取り専用で開き直し、シート構成・Header・行数を検証
  → `File.Move`(上書きなし)で最終パスへ確定。
  失敗・中止時は一時ファイルを削除し、壊れた中途半端な .xlsx を残さない。

### streaming 方針

- 入力: OpenXmlReader による 1 パスのストリーミング読み(DOM 全展開なし)。
  プレビューでは Header・行数・Block 要因などの metadata のみ保持し、
  実データを List で抱えない。
- 出力: OpenXmlWriter によるストリーミング書き込み。
  文字列は共有文字列表を作らず InlineString で書く(巨大な表でもメモリを抱えない)。
- 共有文字列表は Workbook 単位でのみ読み込む(一般的なトレードオフとして許容)。
- 過度な独自基盤は作らない。

## D-013: Phase 1A.1 — 出力列順の基準シートは明示的に選択する (2026-08-28)

D-012 の Header matching 仕様は変更しない。変更するのは
「**出力データ列の並び順をどの Sheet の Header 順から取るか**」を
利用者が明示できるようにする点だけ。

### 背景

D-012 では基準が「最初の選択 Sheet」だった。機能としては正しいが、
UI 上どれが基準か分からず、ファイル追加順で意図しない Sheet が基準になり得た。
利用者から見て「なぜこの列順になったのか」が説明できない状態だった。

### 確定仕様

- 統合対象が 1 件以上あるとき、そのうち**ちょうど 1 件**を「基準シート」として持つ。
  統合対象が 0 件なら基準なし(実行不可)。
- 基準シートの Header 順を、統合結果の**データ列**の並び順に使う。
  metadata 列(元ファイル / 元シート)は従来どおりデータ列の**前**に置く。
- **初期値**: 統合対象になった先頭の Source を初期の基準にする。
  ただし UI 上で「基準」であることを必ず明示する(暗黙のままにしない)。
- **UI**: 「統合するシートを選ぶ」一覧の各行に「基準」ラジオボタンを置き、
  一覧の上に「「基準」に選んだシートの見出し順で、統合結果の列を並べます。」と表示する。
  統合対象にできない行・対象から外した行は基準に選べない。
- **Include 解除時**: 基準にしていた Source を統合対象から外したら、
  残っている対象の**先頭**を新しい基準へ自動設定する(基準なし状態を放置しない)。
- **Sheet 変更時**: 基準 Source の選択 Worksheet を変えた場合は、
  同じ Source の新しい Worksheet をそのまま基準として扱う。
- **Preview stale**: 基準変更・Include 変更・Sheet 変更のいずれでも Preview を
  stale にし、「プレビューを更新」するまで「統合ファイルを作成」を実行できない。
- **Preview 表示**: 更新後の Preview に現在の基準(`ファイル名 / シート名`)を表示し、
  「出力する列」も基準シートの Header 順で表示する。

### Core 側 validation(UI だけを信用しない)

- `MergePlanner.CreatePreview` は基準を**明示的な引数**として受け取る。
  「最初の Source だから基準」という暗黙依存を Core から取り除いた。
- 次の場合は Planner 側で Block する:
  - 基準が未指定
  - 基準が統合対象に含まれていない(パス不一致 / Sheet 名不一致を含む)
  - 基準 Sheet 自体が Block されている(Header 問題・数式・結合セル等)
- UI 側で並べ替えて偶然先頭に置くような脆い実装はしない。
- `MergePreview.BaseSelection` / `BaseDisplay`、`MergeSourcePlan.IsBase` で
  解決された基準を返す。

### 変更しないもの(D-012 のまま)

1 行目 Header / trim 後 Ordinal 比較 / 空 Header Block / 重複 Header Block /
不足・余分 Header Block / Header 集合が一致すれば列順違いは統合可能。
Header 自動推定・manual mapping・表記揺れ補正は追加しない。

### 補足(テスト構成 / D-013)

基準の保持・自動再選択は ViewModel 側の状態機械なので、テストプロジェクトを
`net8.0-windows` + `UseWPF` にして App プロジェクトを参照し、ViewModel を直接検証する。
UI を起動する専用テスト基盤は作らない。
WPF 参照時は暗黙 using から `System.IO` が外れる(`System.Windows.Shapes.Path` との
衝突回避)ため、テストプロジェクトでは `Using Remove="System.Windows.Shapes"` +
`Using Include="System.IO"` を指定する。

## D-014: Phase 1B.1(基本要素を保持した Worksheet 集約)の仕様 (2026-08-28)

### Phase 1B を段階実装する理由

「任意の Excel Sheet の完全コピー」は、Chart / Drawing / PivotCache /
Conditional Formatting / Data Validation など Workbook をまたぐ依存を持つ
パートの移植が必要で、一度に作ると「一部を黙って落とす」実装になりやすい。
本製品の中心価値は安全性なので、**保持できる範囲を先に確定し、それ以外は
明示的に Block する**段階(1B.1)から始める。複雑な要素は 1B.2 以降で改めて設計する。

### 1B.1 のスコープ

- 複数 Workbook から選んだ Worksheet を、1 つの**新規** Workbook へ
  複数 Worksheet として並べる。Phase 1A と違い行データは混ぜない。
- 1 Workbook から複数 Worksheet を選択できる。対象は .xlsx の Worksheet のみ
  (Chartsheet / MacroSheet / Dialogsheet は選択対象外)。
- 入力 Workbook への書き込みは禁止(読み取り専用でしか開かない)。

### 保持する要素

セル値・セル型(Shared String / Inline String / String / Number / Boolean /
Blank / Error)、日付・時刻、セル書式(表示形式・フォント・塗りつぶし・罫線・配置)、
行の高さ / CustomHeight / 非表示 / アウトラインレベル / 行書式、
列の Min-Max / 幅 / CustomWidth / 非表示 / 列書式、結合セル、ウィンドウ枠の固定
(Pane + Selection)、シートの表示状態、SheetProtection、
SheetView のうち ShowGridLines / ShowRowColHeaders / RightToLeft / ZoomScale。

### 黙って落とさず Block する要素

- **Sheet 単位**: 数式 / グラフ / 図形 / 画像(シート背景を含む)/ 旧形式 VML /
  テーブル (ListObject) / ピボットテーブル / 条件付き書式 / データ入力規則 /
  コメント(旧・スレッド形式)/ ハイパーリンク / オートフィルター /
  OLE 埋め込み / ActiveX / リッチテキスト参照セル / 壊れた・重複する結合セル定義。
- **Workbook 単位**(シートを切り離すと意味が壊れるもの): マクロ (VBA) /
  他ブックへの外部参照(外部リンク・外部データ接続)/ 読み取り不能 / .xlsx 以外。
- 非選択 Sheet にだけ問題要素がある場合は、選択 Sheet を Block しない
  (Sheet 単位判定)。
- **Block せず、保持もしない要素**(データの意味を変えないレイアウト情報のため):
  印刷設定(pageMargins / pageSetup / printOptions / headerFooter / 改ページ)、
  sheetPr(タブ色等)、phoneticPr。この扱いは本 Decision で明示しておく。

### 数式

選択 Sheet に数式が 1 つでもあれば Block。別 Sheet 参照・定義名依存・外部ブック
参照・シート名変更による参照破壊の可能性があり、Open XML SDK は計算エンジンでも
ないため。**cached value だけを使って値へ変換することも行わない**(1B.2 で再設計)。

### 書式(style)の移植

- Workbook ごとに Stylesheet は独立しているので、StyleIndex をそのまま持ち込まない。
- Source ごとに Fonts / Fills / Borders を出力へ**追記**(offset 方式)し、
  ユーザー定義 NumberingFormat は 164 から採番し直す。CellFormat は clone して
  fontId / fillId / borderId / numFmtId を付け替え、
  「元の StyleIndex → 出力の StyleIndex」表を Source 単位で保持する。
- 過度な dedup はしない(安全性と追いやすさを優先)。
- 名前付きスタイル(cellStyleXfs / cellStyles)は取り込まず、出力の CellFormat は
  すべて既定の cellStyleXf(xfId=0)を指す。直接書式は cellXfs に入っているため
  通常のファイルでは見え方は変わらない。名前付きスタイル依存の書式は
  1B.2 の検討事項とする。
- テーマ(配色定義)は**最初の Source のものを出力へコピー**する。
  テーマが異なるブックが混在する場合は Warning を出す(Block はしない)。

### Shared String / リッチテキスト

- 複数 Workbook の SharedStringTable を index のまま結合しない。
  読み取り時に解釈し、出力では InlineString として書く(Phase 1A と同じ方針)。
- 文字単位の書式を持つ共有文字列(rPr を含む run)は、書式を落とさずに移せないため
  **黙って flatten せず Block** する。

### 日付

Phase 1A の判定基盤(`WorkbookReadContext`)を再利用する。重複実装しない。
Source が 1904 date system の場合は読み取り時に +1462 して 1900 系へ正規化し、
表示形式は元のものを remap して使う。1900 / 1904 が混在しても出力は同じ日付を表す。
曖昧な numFmt を誤変換しない方針も Phase 1A のまま。

### 結合セル

Phase 1A では Block だったが、1B.1 は Sheet 構造を保つ機能なので**対応対象**。
元 Sheet の MergeCell 範囲をそのまま出力 Sheet へ保持する。
範囲が重複する・解釈できない定義は Block。

### 出力シート名

- Excel 制約: 31 文字以内 / `: \ / ? * [ ]` 不可 / 先頭末尾のアポストロフィ不可 /
  予約名 "History" 不可 / Workbook 内で重複不可(大文字小文字は区別しない)。
- 重複時の初期提案は決定的に「名前」「名前 (2)」「名前 (3)」…とし、
  31 文字を超える場合は suffix 分だけ元の名前側を短くする。
- 利用者は出力シート名を編集できる。空 / 31 文字超過 / 使用不可文字 / 重複は
  **Block して理由を表示**し、勝手に別名へ置き換えて実行しない。
- 選択したシートがすべて非表示になる場合も Block(Excel が開けないため)。

### 出力シートの順序

Workbook の追加順 → 元 Workbook 内の Sheet 順。Planner は渡された選択順を
そのまま出力順として使う。Preview に最終出力順を表示する。
Drag & Drop 等の並べ替え UI は 1B.1 では作らない。

### 出力の安全設計と検証

- Phase 1A(D-012)の設計を再利用: 新規ファイルのみ / 既存ファイル上書き禁止 /
  入力と同一パス禁止 / 同一ディレクトリの一時ファイル → 再オープンして検証 →
  `File.Move`(上書きなし)で確定 / 失敗時は一時ファイル削除。
- 検証内容: SpreadsheetDocument として開ける / Workbook 存在 / シート数・名前・
  順序・表示状態の一致 / WorksheetPart の存在 / StyleIndex が範囲内 /
  **OpenXmlValidator で検証エラー 0 件**。
- 2026-08-28 時点で、許容が必要な既知の validation error は無い
  (テストでもエラー 0 件を確認済み)。将来許容が必要になった場合は
  勝手に無視せず本ファイルへ追記する。

### streaming

入力は OpenXmlReader による行単位の読み出し、出力は OpenXmlWriter による
逐次書き込み。セル全体を巨大な List に積まない。
Style mapping・列定義・結合範囲・SheetView など Workbook / Sheet 単位の
小さい metadata の保持は許容する。

### Phase 1B.2 で扱う候補(D-014 時点)

数式(参照の付け替えを含む)、Chart / Drawing / Image のコピー、Table (ListObject)、
条件付き書式、データ入力規則、コメント、ハイパーリンク、PivotTable / PivotCache、
名前付きスタイル依存、印刷設定・ページレイアウト、テーマの統合、
出力シートの並べ替え UI。

## D-015: Phase 1B.1.1 — 印刷設定の Block 化とシート表示状態の正確な保持 (2026-08-28)

Phase 1B.1 の GitHub レビューで見つかった 2 件の仕様ズレへの最小補正。
**D-014 のうち下記 2 点を更新する。それ以外の D-014 仕様は変更なし。**

### 1. 印刷・ページレイアウト情報は Block する(D-014 の該当記述を更新)

- D-014 では pageMargins / pageSetup / printOptions / headerFooter / 改ページ 等を
  「Block せず、保持もしない」としていた。これは Phase 1B.1 の元方針
  「保持しない要素を持つ選択 Sheet は、コピーせず続行ではなく Block する」からの
  逸脱だったため取り消す。
- **選択 Sheet に次の要素があれば Block する**(理由:
  「印刷設定・ページレイアウト情報を含むため、Phase 1B.1 では集約できません。」):
  `printOptions` / `pageMargins` / `pageSetup` / `headerFooter` /
  `rowBreaks` / `colBreaks` / `drawingHF`(ヘッダー・フッターの図)/
  `sheetPr` の `pageSetUpPr`(印刷の拡大縮小設定)。
- 判定は従来どおり **Sheet 単位**。非選択 Sheet にだけ印刷設定があっても、
  選択 Sheet を Block しない。
- 実務上の影響: Excel が保存した .xlsx は多くが pageMargins を持つため、
  1B.1 で集約できるシートは実質的に「印刷設定を持たないシート」に限られる。
  安全性を優先した仕様であり、印刷設定の移植は 1B.2 の課題とする。

### 2. sheetPr / phoneticPr の扱い(元指示で未定だったため、ここで確定)

D-014 では扱いを決めていなかったので、現状仕様を確認したうえで次に確定する。
Block を増やす範囲は 1. で挙げたものに限り、それ以外へは広げない。

- `sheetPr/pageSetUpPr`: 印刷設定なので **Block**(1. に含む)。
- `sheetPr/tabColor`(シート見出しの色): 保持しないが、セルの値・書式・構造を
  変えないため **Warning**(「シート見出しの色は引き継がれません。」)。
- `phoneticPr`(シートのふりがな設定)および共有文字列の `rPh`(セルのふりがな):
  保持しないため **Warning**(「ふりがな(phonetic)情報は引き継がれません。」)。
  日本語のファイルでは IME 入力により広く付くため、Block にすると実用性を
  大きく損なう。黙って落とさないという要件は Warning で満たす。
- `sheetPr` のその他の属性(codeName / filterMode 等)は、関連する機能
  (マクロ・オートフィルター)自体が既に Block 対象のため個別対応しない。
- 補足(2026-08-28): `pageSetUpPr` は D-016(Phase 1B.2A)で保持対象になった。
  `tabColor` / `phoneticPr` の Warning 方針は D-016 でも変更していない。
- 共有文字列の本文読み取りは `rPh` を含めないよう修正した
  (従来は InnerText を使っていたためふりがなが本文に混ざる可能性があった)。

### 3. シート表示状態を Visible / Hidden / VeryHidden で保持

- D-014 の実装では Hidden と VeryHidden を bool に潰し、出力ではどちらも
  Hidden として書いていた。「Sheet visibility を保持」を満たしていなかったため修正する。
- Core に `SheetVisibility { Visible, Hidden, VeryHidden }` を追加し、
  **scan → plan → output → 出力検証**まで 3 値のまま持ち回る。
  出力の `sheet/@state` は Visible なら属性なし、Hidden / VeryHidden は
  それぞれ対応する値を書く。出力検証でも 3 値で一致を確認する。
- Phase 0 の `SheetInfo` も `Visibility` を持つようにし、`IsHidden` は
  `Visibility != Visible` の派生プロパティにした(既存の呼び出しはそのまま動く)。
- UI 表示: 「非表示」「非常に非表示」を区別して表示する
  (解析タブのシート種類欄と、集約タブのシート一覧の両方)。UI の全面変更はしない。
- 「選択したシートがすべて非表示」の Block 判定は、Hidden と VeryHidden の
  どちらも非表示として扱う(従来どおり)。

## D-016: Phase 1B.2A — 印刷設定・ページレイアウトの安全な移植 (2026-08-28)

Phase 1B.2 の最初の段階。**D-015 の Block 方針のうち、安全に保持できるようになった
要素だけ Block を解除する。それ以外の Phase 1B.1 / 1B.1.1 仕様は変更なし。**

### 目的

D-015 で印刷・ページレイアウト情報を Block にしたが、Excel が保存した .xlsx は
ほとんどが pageMargins を持つため、1B.1 の実用範囲が極端に狭くなっていた。
1B.2A では「安全に意味を維持できる印刷設定」を出力へ移植し、
これらだけを理由に Block しない状態にする。

### 保持する要素(Block 解除)

- `printOptions`(horizontalCentered / verticalCentered / headings /
  gridLines / gridLinesSet など、要素の属性をそのまま)
- `pageMargins`(left / right / top / bottom / header / footer)
- `sheetPr/pageSetUpPr`(fitToPage / autoPageBreaks)。
  **sheetPr 全体は clone せず、pageSetUpPr だけを明示的に扱う**
- `pageSetup`(**r:id を持たない場合のみ**。paperSize / orientation / scale /
  fitToWidth / fitToHeight など)
- `headerFooter`(**文字列のみ**。oddHeader / oddFooter / evenHeader /
  evenFooter / firstHeader / firstFooter / differentOddEven / differentFirst /
  scaleWithDoc / alignWithMargins)
- `rowBreaks` / `colBreaks`(Break の id / min / max / man)。
  count / manualBreakCount は出力側で実際の内容に合わせて振り直す
- `_xlnm.Print_Area` / `_xlnm.Print_Titles`(下記)

### Block を継続する要素

- **`pageSetup` に r:id がある場合**: r:id はプリンター設定パート (devMode) への
  relationship。1B.2A では PrinterSettingsPart を移植しないので、
  r:id を黙って外すことも、Source の r:id を出力へ持ち込むこともせず Block する。
  文言:「プリンター固有の設定を含むため、現在のバージョンでは安全に集約できません。」
  PrinterSettingsPart 対応は 1B.2B 以降の候補。
- **ヘッダー・フッターの画像**: `drawingHF` / `legacyDrawingHF`、
  およびヘッダー文字列中の画像コード `&G`。画像を落として文字だけ残さない。
- **壊れた改ページ定義**: 位置が無い / 0 / 軸の上限超過 / min > max。
- 数式・グラフ・図形など Phase 1B.1 で Block だったものはそのまま。

### Print Area / Print Titles

Worksheet XML ではなく Workbook の Defined Name として存在するため、専用に扱う。

- 対象にする単純ケース: `localSheetId` が対象シートに一致し、参照先がその
  シート自身で、通常の A1 絶対参照であるもの。
  Print_Area は範囲(単一・複数)、Print_Titles は行範囲 / 列範囲 / その組合せ。
- **Block する**: 外部ブック参照(`[` を含む)/ 他シート参照 / 3D 参照 /
  `#REF!` / 数式的な定義 / 解釈できない形式 / 同一シートに同名定義が複数 /
  シート位置を特定できない場合。
- 文字列 replace では書き換えない。Print_Area / Print_Titles の限定文法だけを
  扱う小さな parser / formatter を用意する(Excel の数式パーサーは自作しない)。
- **シート名の quoting**: 解析時は引用符あり・なしの両方を受け付け、
  引用符内の `''` をエスケープとして解く。出力時は**常に引用符で囲み**、
  アポストロフィを `''` に置き換える(引用は常に付けてよいので、
  囲むかどうかの判定を実装しない)。
  例: `'大阪''支店'!$A$1:$F$100` を壊さない。
- **localSheetId**: Source の値をコピーせず、**出力 Workbook 内の実際の
  シート位置**へ振り直す。シートの一部選択や並び順の違いで位置が変わるため。

### その他の Defined Name

Print_Area / Print_Titles 以外は 1B.2A で一般対応しない。
選択シートに紐づく(localSheetId が一致する)それ以外の Defined Name を
検出したら、黙って落とさず Block する。既存の数式 Block 方針は維持。

### 出力とスキーマ順序

- Worksheet の子要素は CT_Worksheet の順で書く:
  `sheetPr` → `dimension` → `sheetViews` → `sheetFormatPr` → `cols` →
  `sheetData` → `sheetProtection` → `mergeCells` → `printOptions` →
  `pageMargins` → `pageSetup` → `headerFooter` → `rowBreaks` → `colBreaks`。
  末尾へ Append しない。
- Workbook では `definedNames` を `sheets` の後ろに置く。
- OpenXmlValidator のエラー 0 件を維持する(2026-08-28 時点で許容が必要な
  既知エラーは無い)。

### 出力後の検証

再オープン後、Preview で期待した内容と一致するかを確認する:
pageSetUpPr / printOptions / pageMargins / pageSetup / headerFooter の有無、
行・列の改ページ数、Print_Area / Print_Titles の参照文字列と localSheetId。

### UI

既存の「3. シートをまとめる」タブのまま。印刷設定を保持できるシートは通常どおり
選択・プレビュー・実行できる。Block する場合は理由を表示する。
Print Area / Titles の保持は自動で、設定項目を増やさない。

## D-017: Phase 1B.2B1 — ハイパーリンクの安全な移植 (2026-08-28)

**D-014 / D-015 のうちハイパーリンクの Block のみ解除する。他の仕様は変更なし。**

### Phase 1B.2B を段階化する理由

1B.2B の候補は依存関係の重さが大きく違うため、次の順に分ける:

1. **1B.2B1 Hyperlink** — Worksheet の `hyperlinks` 要素と、外部リンクの場合の
   WorksheetPart の HyperlinkRelationship を張り直せば済む、比較的独立した機能
2. 1B.2B2 Data Validation — `formula1` / `formula2` / `sqref` を持ち、
   数式・参照の解釈が必要
3. 1B.2B3 Conditional Formatting — Workbook Styles の DifferentialFormats (dxfs) と
   dxfId の remap、formula や x14 拡張まで絡む
4. 以降: Comments(Comments Part に加えて VML drawing の shape も絡む)、
   Table、Drawing / Image / Chart、Formula、PivotTable / PivotCache

今回扱うのは Hyperlink のみ。

### 外部リンク(Web / メール)

- 対応スキーム: `http` / `https` / `mailto`。
- **Source の r:id を再利用しない**。出力 WorksheetPart に
  `AddHyperlinkRelationship(uri, isExternal: true)` で新しい relationship を作り、
  その r:id を hyperlink 要素へ設定する。r:id の値が変わってよく、
  重要なのは target の意味が同じであること。
- 同じ URL が複数セルにあっても relationship の dedup はしない
  (安全で追いやすい方を優先)。
- Source Worksheet の .rels を丸ごとコピーしない。

### 外部リンクの Block

相対パス / `file://` / UNC / ローカルファイル / 別 Excel ブック / 上記以外のスキーム /
解釈できない URI / 参照先が見つからない r:id(dangling)/
relationship が外部でない場合。
理由: 出力ブックの保存先は元と異なり得るので、相対ファイルリンクは意味が変わる。
**勝手に絶対パスへ変換しない。**

### ブック内リンク

- `location` の限定文法だけを扱う(Excel の数式パーサーは自作しない):
  - シート名なし(`A10` / `A1:B5`)→ そのまま保持
  - シート名あり(`'売上'!A10` / `売上!A10`)→ 参照先シートの**出力シート名**へ書き直す
- 参照先が自分自身のシートでも、出力名へ書き直す必要があるので同じ経路で扱う。
- 別シート宛の解決には、既存の「元ファイル + 元シート → 出力シート名」対応表を使う
  (選択状況を知る必要があるため、走査ではなく Planner で解決する)。
- セル参照は `$` 付きも受け付け、`CellRangeParser` で解釈したうえで
  Excel の上限(16,384 列 / 1,048,576 行)を超えないことを確認する。

### ブック内リンクの Block

参照先シートが集約対象に選ばれていない / シートを特定できない /
名前定義への location / 数式的な location / 外部ブック参照 / 3D 参照 / `#REF!` /
セル位置を解釈できない。**黙ってリンクを削除して続行しない。**
参照先を追加選択して再プレビューすれば実行できるようにする。

### hyperlink の属性

`ref` / `location` / `tooltip` / `display` を保持し、`r:id` だけ出力側へ張り直す。
`ref` は `CellRangeParser` で有効な A1 セル / 範囲であることを確認する
(同じパーサーを重複実装しない)。
想定していない拡張属性・子要素を持つ hyperlink は、意味を保証できないので Block。

### シート名の quote / escape

D-016 で確定した方針と同じ。共通処理 `SheetReferenceSyntax` を新設し、
印刷範囲・印刷タイトルの parser とハイパーリンクの両方から使う
(同じロジックを複数箇所へコピーしない)。出力時は常に引用符で囲み、
アポストロフィは `''` へエスケープする。Print Area / Titles の動作は変えない。

### スタイル

ハイパーリンクのために青文字・下線などの独自スタイルを追加しない。
セルの見た目は Phase 1B.1 の style remap で元のまま引き継がれる。
ハイパーリンク機能とセル書式を混ぜない。

### セキュリティ / ネットワーク

リンクを保持するために URL へアクセスしない。DNS 参照・HTTP 要求・
リンク先ファイルの存在確認・リンクの自動オープンをしない。
読むのは文字列と relationship の metadata だけで、完全オフライン方針は維持する。

### 出力とスキーマ順序

`hyperlinks` は CT_Worksheet の順序どおり `mergeCells` より後、`printOptions` より前へ書く。
出力後の検証で、リンクの位置・location・tooltip・display の一致と、
外部リンクの relationship が解決できて target が一致することを確認する。
OpenXmlValidator のエラー 0 件を維持する。

### テストの補足(Public 安全監査の既知ヒット)

`mailto` スキームの動作確認のため、テストに `mailto:someone@example.invalid` を置く。
`.invalid` は RFC 2606 の予約 TLD で、実在しない架空アドレス。
個人のメールアドレスではないため許容する。安全監査で
メールアドレス検出にヒットするが、これは既知・許容の 1 件。

## D-018: Phase 1B.2B2A — 標準の入力規則(Data Validation)の安全な移植 (2026-08-28)

**D-014 / D-015 のうち入力規則の Block のみ解除する。他の仕様は変更なし。**

### Phase 1B.2B2 を段階化する理由

入力規則には、標準の `x:dataValidation` と Office 2010 以降の
`x14:dataValidation` があり、formula / sqref の構造も異なる。
一緒に扱うと「x14 を黙って落として標準だけ出力する」実装になりやすい。
そこで **2B2A = 標準形式の安全な部分集合**、**2B2B = x14 と複雑な formula** に分ける。

### 対応する type

`none` / `list` / `whole` / `decimal` / `date` / `time` / `textLength`。
`custom` は formula1 が任意の数式なので **Block**。未知の type も Block。

### list の対応範囲

- **直接指定**(`"未着手,進行中,完了"`): 意味を解釈して作り直さず、
  **元の formula 文字列をそのまま保持**する。引用符を勝手に足し引きしない。
  条件: 先頭と末尾が `"` / 内側に `"` を含まない / 全体が 255 文字以内
  (Excel のリスト直接指定の上限)。
- **同じシート内の完全絶対参照**(`$B$2:$B$10` / `$A$1`): シート名を含まないので
  出力でも位置が変わらず、そのまま保持できる。Excel の行列上限も確認する。
- **Block**: シート名付き参照 / 他シート / 名前定義 / INDIRECT・OFFSET 等の関数 /
  数式 / 外部ブック / 構造化参照・テーブル参照 / `#REF!` / 3D 参照 / 解釈不能。
  名前定義への対応は後段で設計する。

### whole / decimal / textLength

formula1 / formula2 は **数値定数のみ**(InvariantCulture で有限数として解釈できるもの)。
セル参照・名前定義・関数・式は Block。安全と確認できた場合は
**元の文字列表現をそのまま出力**する(再フォーマットしない)。

### date / time

- **date**: 数値 serial 定数のみ。Source が 1904 date system の場合、
  出力の 1900 系へ **+1462** して同じ日付を意味するようにする。
  変換は Phase 1A から使っている日付基盤に集約する
  (`MergeCellValue.NormalizeSerialTo1900` を新設し、`WorkbookReadContext` も
  これを使うようにした。同じ変換を別実装しない)。
- **time**: 数値定数のみ。**1462 補正はしない**(0.5 = 12:00 の意味を保つ)。
- セルの日付だけ正しく入力規則の日付条件がずれる、という状態を作らない。

### type / operator / formula の整合

- `list` / `none`: operator は意味を持たないため、存在したら Block
  (勝手に削除して正常化しない)。`list` は formula1 必須・formula2 不可。
  `none` は formula なし。
- `whole` / `decimal` / `date` / `time` / `textLength`:
  operator が `between` / `notBetween` なら formula1・formula2 とも必須。
  それ以外の operator では formula1 必須・formula2 があれば Block。
- 未知 operator は Block。OpenXmlValidator 任せにせず、意味上成立しない構造を
  事前に Block する。

### sqref と属性

- `sqref` は複数範囲(`A1:A10 C1:C10 E5`)に対応。各 token が有効な A1 参照で、
  Excel の上限内で、`#REF!` を含まないことを確認する。
  判定は共通の `A1RangeValidator`(`CellRangeParser` + 上限)に集約し、
  Phase 1B.2B1 のハイパーリンク側の独自上限チェックもこれに寄せた。
  Phase 0 / 1A が使う `CellRangeParser` 自体の意味は変えていない。
- 保持する属性: type / operator / allowBlank / showDropDown / showInputMessage /
  showErrorMessage / errorStyle / imeMode / promptTitle / prompt /
  errorTitle / error / sqref。**元の属性値をアプリ側の意味に読み替えない。**
- 未知の拡張属性・未知の子要素を持つ入力規則は Block。

### コンテナ属性と count

`dataValidations` の `disablePrompts` / `xWindow` / `yWindow` はそのまま保持。
`count` は **出力した件数から再計算**する(Source の count を鵜呑みにしない)。
コンテナに未知の拡張属性があれば Block。

### x14 入力規則

Worksheet の extLst にある `x14:dataValidations` / `x14:dataValidation` は
**Block**(「新しい形式の入力規則を含むため、現在のバージョンでは安全に集約できません。」)。
標準形式と x14 が同時に存在する場合も Block。x14 を黙って落として標準だけ出力しない。
検出は CLR 型ではなく名前空間 + 要素名で行う(型解決に依存しない)。
x14 の実データ移植は 2B2B 以降。

### 部分コピーの禁止

1 つのシートに複数の入力規則があり、1 件でも移植できないものがあれば、
その 1 件だけ落として続行せず **シート全体を Block** する。
理由には対象セル(sqref)と原因を含める。

### 出力方式とスキーマ順

- 意味を組み直さず、**検証済みの標準要素を clone** して必要最小限だけ変更する
  (変更するのは 1904 date の formula と、コンテナの count のみ)。
  既定属性を新しく付けない。list の直接指定を再 serialize して引用符を壊さない。
- `dataValidations` は CT_Worksheet の順で `mergeCells` の後・`hyperlinks` の前へ書く。
  条件付き書式(今も Block)は `dataValidations` の前に入る位置なので、
  将来追加しても順序を壊さない。
- OpenXmlValidator のエラー 0 件を維持する。

### Phase 0 への波及(x14 検出の穴埋め)

Phase 0 の解析でも x14 の入力規則を検出し、標準形式と同じ
「⚠ データ入力規則」Finding として見せるようにした。
新しい Finding 種別や UI 変更は増やしていない。Phase 0 の読み取り専用方針も維持。
条件付き書式など他の x14 拡張の検出は引き続き LATER。

## D-019: Phase 1B.2B2B — マスタシート参照の List 入力規則 (2026-08-28)

**D-018 のうち x14 リストと名前定義参照の Block だけ解除する。他の仕様は変更なし。**

### 主目的

「別シートに置いた候補一覧を参照するプルダウン」を、シート集約後も機能させる。
INDIRECT / OFFSET / 任意数式へは広げない。

### characterization(推測でパーサーを書かない)

実装前に Open XML SDK 3.5.1 で実ファイルを生成し、構造を確認した:

- 拡張 URI: `{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}`
- 名前空間: x14 = `.../spreadsheetml/2009/9/main`、
  xm = `.../excel/2006/main`(SDK は `xne` 接頭辞で出力するが名前空間は同じ)
- 構造: `x14:dataValidation > x14:formula1 > xm:f` と、
  `x14:dataValidation > xm:sqref`
- 標準の list が名前定義を参照するとき、`<formula1>` の中身は
  **先頭 `=` なしのそのままの名前**(例: `商品一覧`)
- `definedName` の本文も **先頭 `=` なし**で、シート名は引用符付き
  (例: `'商品マスタ'!$A$2:$A$50`)
- **`xr:uid` は SDK 3.5.1 の OpenXmlValidator(Office2010 / 2013)が
  「未宣言の属性」としてエラーにする**

### x14 の対応範囲

- **`type="list"` のみ**。whole / decimal / date / time / textLength / custom /
  none / 未知 type は Block(標準形式で対応済みの機能を理由なく x14 へ広げない)。
- 構造は `x14:formula1` 1 件・`formula2` なし・`xm:f` 1 件・`xm:sqref` 1 件に限定。
  未知の子要素・未知の属性(xr:uid を除く)があれば Block。
- 参照元として扱うのは、**同じ Source Workbook 内の通常 Worksheet 1 枚を指す
  完全絶対 A1 範囲**で、**1 セル / 1 列 / 1 行**のみ。
  2 次元範囲(`$A$1:$C$10`)は Block。直接指定("A,B,C")もそのまま保持する。
- 参照先シートが**集約対象に選ばれている場合のみ**対応。選ばれていなければ
  「参照先の『商品マスタ』シートが集約対象に含まれていないため…」と Block し、
  利用者が選び直して再プレビューすれば解除できる。
- シート名の書き換えは Planner の「元ファイル + 元シート → 出力シート名」表を使い、
  `SheetReferenceSyntax` で常に引用 + `''` エスケープする(D-016 / D-017 と同じ)。
- **x14 は x14 のまま**再構築する。標準形式へ変換しない(逆も行わない)。
  表現形式を変えて Excel の挙動が変わるのを避けるため。

### xr:uid

- 出力の x14 規則は **clone せず、対応が分かっているプロパティだけを写して
  新規に組み立てる**。clone すると元ブックの名前空間宣言(xr など)まで
  付いてくるため。結果として **xr:uid は出力に一切現れない**。
- 理由: (1) 複数ブックをまとめると Source の uid が衝突しうる、
  (2) 上記 characterization のとおり SDK の検証器が xr:uid をエラーにするため、
  新しい GUID を発行しても「検証エラー 0」を満たせない。
- xr:uid はリビジョン追跡用の内部識別子で利用者から見える意味を持たず、
  出力は履歴を持たない新規ブックなので、付けないことによる意味の欠落はない。
- テストでは出力の全 worksheet XML に xr 名前空間が現れないことを確認する。

### extLst の扱い

Source の extLst を**丸ごとコピーしない**。Office 2010 の入力規則拡張だけを
明示的に組み立て直す。それ以外の拡張(スパークライン等)が入っている場合は、
黙って捨てず **Block** する(「対応していない拡張情報を含むため…」)。

### 名前定義(Workbook スコープ)

- 対応するのは **ブック全体を対象とする単純な範囲名**のみ。
  参照先は同じ Source Workbook 内の通常 Worksheet 1 枚で、
  完全絶対の 1 行または 1 列の範囲。
- 名前の照合は Excel と同じく**大文字小文字を区別しない**。
- **Block**: シート固有の名前(localSheetId あり)/ 外部ブック / 3D 参照 /
  `#REF!` / INDIRECT・OFFSET 等の関数 / 数式 / 相対参照 / 2 次元範囲 /
  シート名なし / 参照先シート未選択 / 同名が複数。
- **属性**: hidden / function / vbProcedure / xlm / functionGroupId /
  shortcutKey / publishToServer / workbookParameter / customMenu /
  description / help / statusBar / comment のいずれかが付いている名前は Block。
  普通のユーザー作成の範囲名という狭い範囲に限定する。
- 出力では、参照先を出力シート名で組み立て直した
  `名前 → '出力シート名'!$A$2:$A$50` を **localSheetId なし**で作る。
  文字列 replace ではなく「限定 parser → 参照先解決 → formatter」で行う。
- 入力規則側の formula1 は名前のままなので書き換えない。
- 標準形式のリストからの名前参照も同時に Block 解除した(先頭 `=` は
  名前として引くときだけ許容する)。

### 名前の衝突

複数 Source から同名が来た場合、**勝手にリネームしない**。
出力後の参照先が完全に同一なら 1 つに統合し、少しでも違えば Block。
自動解決(`商品一覧_2` 等)は将来の候補。

### 部分コピーの禁止

安全な x14 リストと未対応の x14 規則が同じシートに混在する場合、
安全な分だけコピーせず**シート全体を Block**(D-018 と同じ方針)。

### 対象外(Block 継続)

INDIRECT / OFFSET / INDEX / FILTER / UNIQUE、動的名前定義、連動プルダウン、
Table の構造化参照、Custom Formula。数式評価と名前解決の範囲が急に広がるため、
今回は「固定範囲を参照する普通のマスタプルダウン」を完成させる。

### 検証

出力後に x14 の件数・type・`xm:sqref`・`xm:f`(書き換え後)・xr:uid 不在、
名前定義の名前・スコープ・参照先を照合する。OpenXmlValidator のエラー 0 件を維持。
ネットワークアクセスは一切なし(GUID 生成も行っていない)。

## D-020: Phase 1B.2B2B.1 — x14 属性の許容範囲を xr:uid だけに絞る (2026-08-28)

**D-019 の実装差異を直す最小補正。仕様そのものは D-019 のまま。**

### 見つかった差異

x14 入力規則の走査で、リビジョン名前空間
(`http://schemas.microsoft.com/office/spreadsheetml/2014/revision`)に属する属性を
**名前を問わず許容**していた。D-019 の方針は
「未知の属性は Block。ただし既知の xr:uid だけ入力上の例外として許容し、
出力には持ち込まない」なので、`xr:foo` のような未知属性まで通ってしまい、
そのあと出力要素の再構築時に**黙って消える**状態になっていた。

### 補正

入力で許容する拡張属性は、次の両方を満たすものだけにする:

- `NamespaceUri` == リビジョン名前空間
- `LocalName` == `"uid"`

これ以外(`xr:foo` などリビジョン名前空間の未知属性、他の名前空間の未知属性)は
理由付きで Block する。**未知属性を受け入れてから出力で落とす経路は作らない。**

### 変わらない点

- 入力に xr:uid があっても規則自体は保持し、出力へ uid を持ち込まないこと(D-019)。
- 出力要素を clone せず既知プロパティだけで組み立てること(D-019)。
- OpenXmlValidator のエラー 0 件、名前定義・マスタ参照の対応範囲(D-019)。

## D-021: Phase 1B.2B3A — 数式を使わない標準の条件付き書式だけを移植する (2026-08-28)

「条件付き書式があれば一律 Block」を解除する。ただし解除するのは、
**数式を評価せず、dxf を指すだけで意味が決まるルール**に限る。

### 対応する cfRule の type

`duplicateValues` / `uniqueValues` / `top10` / `aboveAverage` の 4 種類だけ。
どれもセル範囲の中で完結し、参照の書き換えが要らない。

### 保持するもの

- `sqref`: 元の文字列をそのまま使う。A1RangeValidator で Excel の上限内かを確認し、
  `#REF!` や解釈できない範囲は Block。
- `priority`: **元の値をそのまま使う。1,2,3 に振り直さない。**
  振り直すと `stopIfTrue` の評価順が変わり、見え方が変わりうるため。
  シート内で重複していれば Block(定義が壊れている)。
- `stopIfTrue`: あればそのまま、無ければ出力にも書かない。
- top10 の `rank` / `percent` / `bottom`、aboveAverage の
  `aboveAverage` / `equalAverage` / `stdDev`: あるものだけをそのまま写す。
  **元ファイルに無い属性を既定値として書き足さない。**

### Block する条件

数式を持つ type(`expression` / `cellIs`)、文字列・空白・エラー・日付条件
(`containsText` 系 / `containsBlanks` 系 / `containsErrors` 系 / `timePeriod`)、
カラースケール・データバー・アイコンセット、未知の type。
対応 type であっても、子要素(Formula など)を 1 つでも持つものは Block
(「対応 type だから Formula を無視してコピー」は禁止)。
`cfRule` / `conditionalFormatting` / `dxf` の未知拡張属性、`cfRule` 以外の子要素、
`pivot="true"`、ルール 0 件、priority 不正・重複も Block。

種類ごとの属性整合も見る: 対応 type に `operator` / `text` / `timePeriod` が付く、
top10 以外に rank 系が付く、aboveAverage 以外に平均系が付く、top10 に rank が無い、
rank が範囲外(個数 1〜1000 / 割合 0〜100)、`equalAverage` と `stdDev` の併用、
`stdDev` が 1 未満 — いずれも Block。

x14 の条件付き書式(extLst の `{78C0D931-6437-407d-A8EE-F0AAD7539E65}`)を含むシートも
Block。標準形式だけ写して x14 側を黙って落とすと、見た目が変わったことに気づけない。

### dxf(条件付き書式が使う書式)の対応付け

セル書式(CellFormats)とは**別の索引体系**なので、対応表も分けて持つ。

- `OutputStylesheetBuilder.AddDifferentialFormat(sourceKey, sourceIndex, dxf)` が
  「(Source, 元の dxf 位置) → 出力の dxf 位置」を覚えて出力位置を返す。
  同じ Source の同じ位置なら 1 つにまとめる。Source をまたぐ統合はしない(D-013 と同じ方針)。
- **参照されている dxf だけ**を出力へ写す。1 つも無ければ `dxfs` 要素自体を作らない。
- `count` は実際の件数で書き直す。
- 条件付き書式を足しても既存の CellStyle index はずれない(テストで固定)。

写せる dxf の子要素は font / numFmt / fill / alignment / border / protection。
alignment と protection は自己完結(索引参照を持たない)で、
Office2007 / 2010 / 2013 の OpenXmlValidator を 0 件で通ることを実測して許容した。
それ以外の子要素があれば Block。

**テーマ色は Block。** 出力ブックのテーマ次第で色が変わるため、RGB へ変換もしない。

### dxf の numFmt

実測で、dxf 内の `numFmt` は **ID によらず `formatCode` が必須**だった
(組み込み ID 14 でも `formatCode` 無しは検証エラー)。
そのため `formatCode` を持たない numFmt は Block。

`formatCode` を持っていても ID がブックの `numFmts` と食い違うと解釈が割れるので、
ユーザー定義 ID(164 以上)は出力側で採番し直し、同じ内容を `numFmts` にも登録して、
**どちらを見ても同じ書式になる**ようにする。採番はセル書式と同じカウンタを使う。

### 部分コピーの禁止

1 件でも移植できないルールがあれば、対応分だけコピーせず**シート全体を Block**
(D-018 / D-019 と同じ方針)。

### 出力位置と検証

CT_Worksheet の順序にしたがい `mergeCells` の後、`dataValidations` の前に書く。
出力後に、条件付き書式の件数・`sqref`・ルール数・type・priority・stopIfTrue・
種類ごとの属性・dxfId が指す書式の内容と表示形式・`dxfs` の count を照合する。
OpenXmlValidator のエラー 0 件を維持。

### 対象外(Phase 1B.2B3B 以降の候補)

カラースケール / データバー / アイコンセット(3B)、数式を使うルール(3C)。

## D-022: Phase 1 を v1 の安全 Baseline として CLOSE する (2026-08-28)

### 完成の定義を変える

Phase 1 の目的は **Excel Worksheet の完全クローンではない**。

> 安全に意味を維持できる Worksheet を集約し、
> 未対応要素は事前に検出して拒否する

これを v1 の完成定義とする。予定していた高度な要素をすべて実装してから
Phase 2 へ進む必要はない、と再評価した結果。

### v1 として CLOSE する範囲

- Phase 1A(表データの縦結合): COMPLETE(D-012)
- Phase 1A.1(基準シートの明示・選択): COMPLETE(D-013)
- Phase 1B.1 / 1B.1.1(基本要素を保持した Worksheet 集約): COMPLETE(D-014 / D-015)
- Phase 1B.2A(印刷設定・ページレイアウト): COMPLETE(D-016)
- Phase 1B.2B1(ハイパーリンク): COMPLETE(D-017)
- Phase 1B.2B2A(標準の入力規則): COMPLETE(D-018)
- Phase 1B.2B2B / .1(マスタ参照 List・x14 subset): COMPLETE(D-019 / D-020)
- Phase 1B.2B3A(数式を使わない条件付き書式): COMPLETE(D-021)

**Phase 1B.2B3A までを v1 の完成ラインとし、Phase 1 全体を CLOSED とする。**

### LATER へ移すもの(削除ではない)

Phase 1B.2B3B(カラースケール / データバー / アイコンセット)、
Phase 1B.2B3C(数式型の条件付き書式)、Comments、Table (ListObject)、
Drawing / Image / Chart、Formula、PivotTable / PivotCache、
PrinterSettings (devMode)、ヘッダー/フッター画像、
名前付きスタイル依存の書式、テーマ統合、出力シートの並べ替え UI。

これらは **「不要」と判断したのではない**。実使用で需要が確認できた時点で
LATER から戻す。roadmap にも項目として残す。

### なぜ今 Phase 2 へ移るか

未対応要素は黙って落とさず事前 Block できる状態が既にできている。
一方「複数ブックの同じセルをまとめて変える」ような一括変更は、
Worksheet fidelity をこれ以上上げるより実務価値が高いと判断した。
開発の優先順位をそちらへ移す。

### 既存 Decision を書き換えない

この判断は **D-014〜D-021 を一切書き換えない**。
各 Phase の技術判断はそのまま有効で、本項は「どこで区切るか」だけを決める。

## D-023: Phase 2A — 一括変更の安全 Baseline (2026-08-28)

既存 Workbook の内容を変える初めての機能。中心原則は 1 つ。

> **変更対象以外の Part には触れない。**

理解できない Part(グラフ・図・画像・テーブル・名前定義・条件付き書式など)は、
触らなければそのまま保持される。Phase 1B のように Workbook を新規再構築しない。

### Source を書き換えない

`SpreadsheetDocument.Open(..., isEditable: true)` を **Source に対して使わない**。

```
Source ─(byte copy)→ 一時ファイル ─(編集)→ 検証 ─(File.Move)→ 出力
```

Source は常に読み取り専用。実行前後で SHA-256 / サイズ / mtime が一致すること。

### AutoSave = false(characterization の結果)

コピーを編集するとき `OpenSettings.AutoSave` を **false** にし、
保存するのは対象 WorksheetPart だけにする。

実測: 既定(AutoSave = true)のままだと、シートを探すために読んだ
`xl/workbook.xml` まで書き戻される。AutoSave = false + 対象パートのみ保存なら、
変わる ZIP entry は `xl/worksheets/sheetN.xml` だけになる
(styles / sharedStrings / theme / drawing / image / table / rels /
[Content_Types].xml はいずれも展開後の内容が完全一致)。

SharedString のセルを InlineString へ変えても `xl/sharedStrings.xml` は不変。
これは「共有文字列表を増やさない」ための書き方でもある。

### 出力先

**元ファイルと同じディレクトリ**に `<元の名前>_変更済み.xlsx`(接尾辞は変更可)。
別フォルダー出力は導入しない。相対ハイパーリンクや sibling file 参照がある場合、
場所が変わるだけで参照の意味が変わりうるため。

既存ファイルがあれば **Block**。`_変更済み2` のような自動改名はしない。
控えファイル(後述)の名前が衝突する場合も Block。

### 扱う操作と値

操作は 1 種類だけ:「選択したすべてのシートの同じ 1 セルへ同じ定数値を書く」。
セル位置は A1 形式の**単一セル**のみ(`$B$2` も可)。範囲・シート名付きは不可。

値は **文字列 / 数値 / 空欄**の 3 つ。

- 文字列: `InlineString` で書く(sharedStrings を変更対象にしないため)
- 数値: InvariantCulture で有限数として検証。NaN / Infinity / 桁区切り付きは不可
- 空欄: セルと書式は残し、値の中身(formula / v / is)だけ消す

Boolean / Date / Time / 数式の書き込みは LATER。

### 書式は変えない

対象セルの `StyleIndex` は変更しない。新しい Style も作らず、Stylesheet も触らない。

### 存在しないセルは作らない

対象セルが物理的に無ければ **Block**。Style・行の書式・入力規則・テンプレートの
設計をどう継承するかを決める必要があるため、セルの新規作成は後段に回す。

### 数式(Phase 2A の意図的な制約)

**Workbook 内に数式が 1 件でもあれば Block。**

定数セルを変更すると、それを参照する数式の cached result が古いまま残る。
このツールは Excel の計算エンジンを持たず、依存関係の特定・calcChain・
calculation mode・fullCalcOnLoad を勝手に書き換えることもしない。
`CalculationChainPart` の存在も数式ありとみなす。

対象セル自体が数式の場合の guard も別に置く(数式を値へ置き換えない)。

### 対象セルの guard

結合セル(top-left でも)/ 保護シート / 入力規則(標準・x14 とも)の sqref に
含まれる / ハイパーリンクの ref に含まれる / リッチテキスト / セル値メタデータ
(`cm` `vm`。リンクされたデータ型など)を持つ — いずれも Block。

シートにピボットテーブルがある場合も Block(更新で値が入れ替わりうる)。
外部データ接続・外部ブック参照を持つ Workbook も Block。

**Chart / Drawing / Image / Table があるだけでは Block しない。**
それらの Part を一切変更しないため。Table 内の普通の値セルも、
他の Block 条件を満たさなければ対象にしてよい。

### 表示形式との組み合わせ

書式を変えずに値だけ差し替えるので、表示形式と値の種類が噛み合わないと
「気づけない意味の変化」が起きる(日付書式に 5 を入れると 1900-01-05、
`0%` に 50 を入れると 5000%)。そこで:

- 文字を書く: `numFmtId` が 0(General)または 49(`@`)のときだけ許可
- 数値を書く: 0 / 1 / 2 / 3 / 4(General と素の数値書式)のときだけ許可
- 空欄にする: 表示形式によらず許可(値が無ければ表示形式は効かない)

ユーザー定義(164 以上)を含め、上記以外は判断が曖昧なので **Block**。
通貨・パーセント・分数・指数・日付・時刻はすべてここに含まれる。

### Source は Validator クリーンなものだけ

元ファイルに OpenXmlValidator のエラーが 1 件でもあれば Block。
既に壊れているファイルを書き換えると原因の切り分けができなくなる。
出力も Validator エラー 0 を必須とする。既存エラーを baseline として
許容する方式は LATER。

### snapshot と stale guard

プレビュー時に各 Source の SHA-256 / サイズ / mtime を控え、実行直前に再確認する。
1 つでも変わっていれば **バッチ全体を中止**(一部だけ実行しない)。
画面側も、対象・セル位置・値の種類・値・接尾辞のいずれかが変わったら
プレビューを stale にして実行不可にする。

### バッチの原子性

1. 全 Source の snapshot 再確認
2. 全出力先(Workbook と控えファイル)が未使用か再確認
3. 各 Source を同じフォルダーの一時ファイルへコピーし、コピーだけ編集
4. 一時ファイルを開き直して検証(対象セルの型・値・StyleIndex、Validator 0)
5. 控えファイルも一時ファイルとして作る
6. **すべて成功してから**確定(File.Move)
7. 途中で 1 件でも失敗したら、一時ファイルを削除し、確定済みの出力も取り消す

**正直な限界**: 確定(rename)を 1 件ずつ行う区間で OS 障害や電源断が起きた場合、
一部だけ出力が残る可能性がある。ファイルシステムをまたぐ複数ファイルの
完全な原子性は保証できない。この区間はロールバックを試みるだけで、
「絶対に部分出力が残らない」とは言わない。

### 控えファイル(audit sidecar)

出力 Workbook の隣に `<出力名>.xlsx.audit.json` を作る。完全にローカル。

`schemaVersion` / `createdAt`(オフセット付き)/ `sourceFileName` /
`outputFileName` / `sourceSha256` / `outputSha256` / `operation` /
`changes[] { sheetName, cell, oldValue, oldType, newValue, newType }`。

**絶対パスは書かない**(ファイル名だけ)。ファイルの中身も保存しない
(対象セルの old / new だけ)。既に同名のファイルがあれば Block。

### No-op

現在値と新しい値が型を含めて同じセルは変更しない。すべてが No-op なら実行不可
(No-op のためだけに出力を作らない)。一部だけ No-op なら、変わるものだけ実行する。

### LATER

検索・置換、Date / Boolean / 数式の書き込み、存在しないセルの作成、
数式を考慮した変更(依存関係と再計算)、任意フォルダーへの出力、
既存エラーの baseline 許容、**in-place 上書き**。

## D-024: Phase 2A.1 — 取り消し(rollback)の結果を利用者へ正確に伝える (2026-08-28)

**D-023 の実装差異を直す最小補正。仕様そのものは D-023 のまま。**

### 見つかった差異

D-023 では「rollback は best effort。OS 障害等では一部出力が残る可能性があり、
"絶対に部分出力が残らない" とは言わない」と決めていた。

しかし `CellMutator.Execute` は、rollback のあと利用者へ

> 新しいファイルは作成していません

と**断定**していた。rollback の中身(`DeleteQuietly`)は `IOException` と
`UnauthorizedAccessException` を握りつぶす実装だったため、
確定済みの Workbook や控えファイルがロック等で削除できず**残っていても**、
「作成していません」と表示されうる状態だった。D-023 と安全方針に反する。

### 補正

rollback のあと、**本当に残っていないかを確認**してから文言を決める。

- 削除処理は「消せたか」を返す形にし、テストから失敗を再現できる
  最小の seam(`CellMutator.FileDeleter`)にした。大規模な DI 化はしない。
- 最終判断は削除処理の戻り値ではなく **`File.Exists`** で行う。
- 取り消し対象は「この実行が作ったもの」だけ(確定済みの出力と一時ファイル)。
  失敗の原因になった実行前から在るファイルには触れない。
- 一時ファイルは**作る前に**控える。作成途中で失敗しても取り消し対象から漏れない。

### 文言

すべて消えている場合:

> 一括変更に失敗しました: {理由}作成途中のファイルは取り消しました。
> 元のファイルは変更していません。

1 件でも残っている場合:

> 一括変更に失敗しました: {理由}取り消せなかったファイルが残っている可能性があります。
> 次のファイルを確認してください: ○○_変更済み.xlsx / ○○.audit.json。
> 元のファイルは変更していません。

- 「作成していません」とは言わない(一度作って取り消した可能性を反映する)。
- 残存ファイルは **`Path.GetFileName` のみ**。利用者向けメッセージに絶対パスを出さない。
- **中止(cancellation)時も同じ表示**を使う。

### 変わらない点

- **元ファイルの非変更保証は変わらない。** Source は読み取り専用でしか開かず、
  実行直前に snapshot(SHA-256 / サイズ / mtime)を再確認しているので、
  「元のファイルは変更していません」だけは断定してよい。
- rollback が best effort であること、rename 区間の完全な原子性を
  保証できないこと(D-023)も変わらない。今回変えたのは
  **その事実を利用者へどう伝えるか**だけ。

### 表示の一貫性

失敗の内容を成功と同じ見た目(緑の枠)で出すこと自体が断定になるため、
結果表示の色を実行結果に合わせるようにした。UI の構造は変えていない。

## D-025: Phase 2B — 複数セルへ異なる固定値をまとめて入力する Input Set (2026-08-28)

### なぜ検索・置換より先か

市場の再評価で、実務の一括作業は「どのセルへ、どの値を入れるか」という
複数項目の入力・転記が中心と判断した。B2=確認済み / D5=担当A / F8=1500 /
H10=空欄 のような入力セットを先に作る。これは Phase 2C
(Source Excel / CSV から値を供給する転記 Mapping)の安全基盤でもある。
検索・置換は Phase 2D へ。

### Operation model

`CellMutationRequest` を operation の list に変えた。

- `CellMutationOperationRequest` = CellReference + WriteKind + TextValue / NumberText
- `CellMutationRequest` = Targets + Operations + OutputSuffix
- 同じ入力セットを、選択したすべてのシートへ適用する
- Operation が 1 件だけなら Phase 2A と同じ結果になる(後方互換)

値の種類は Phase 2A と同じ **文字列 / 数値 / 空欄** のみ。operation ごとに指定できる。
Boolean / Date / Time / Formula は引き続き対象外。

### 重複 Block

セル参照は既存 `TargetCellAddress` で正規化して比べる($B$2 と B2 は同じセル)。
同じ入力セット内で重複したら **Block**。最初の値・最後の値を勝手に採用しない。

### 1 件でも不可なら全体 Block(部分適用禁止)

Phase 2A の guard(存在しないセル・数式セル・結合・保護・入力規則(標準/x14)・
ハイパーリンク・リッチテキスト・cm/vm・ピボット・表示形式の不適合)を
**operation × sheet ごと**に適用する。1 セルでも Block なら実行そのものを止める。
「安全な行だけ適用して続行」する経路は作らない。数式 Workbook の一律 Block、
Validator クリーン Source のみ、という Phase 2A の方針も変えない。

### Workbook / Sheet を N 回開き直さない

Operation が 500 件あっても、Scanner が Workbook を開くのは **1 ファイル 1 回**。
Workbook 単位の guard(数式・接続・外部参照・Validator)も 1 回だけ確認する。
シートの走査も **1 シート 1 回**で、その 1 パスで全対象セルの
存在・数式・メタデータ・結合/リンク/入力規則の被覆をまとめて集める。

実測(架空ブック 2 シート × 2,000 セル、Release):

| operations | Preview | Execute(2 シート適用) | Preview メモリ増分 |
|---|---|---|---|
| 100 | 約 320 ms | 約 290 ms(200 セル変更) | 約 5 MB |
| 500 | 約 314 ms | 約 373 ms(1,000 セル変更) | 約 6 MB |
| 2,000 | 約 348 ms | 約 1,042 ms(4,000 セル変更) | 約 8 MB |

Preview が件数によらずほぼ一定 = N 回 reopen になっていない証拠。
この結果から **Core に件数上限は設けない**(ファイル数上限 500 は保険として維持)。
UI の一覧(DataGrid)は仮想化されており、200 行の追加 9 ms /
プレビュー+描画 約 0.7 秒を確認した。

### Worksheet 単位でまとめて保存

Mutator は変更をシートごとに group し、対象セルを 1 回の走査でまとめて見つけ、
**Worksheet.Save() をシートにつき 1 回**にする。20 operations でも Save 20 回にしない。
AutoSave=false・対象 WorksheetPart のみ保存という D-023 の方式はそのまま。

2,000 operations を 2 シートへ適用した出力でも、変わる ZIP entry は
その 2 つの WorksheetPart だけであることを実測で確認(Package Part integrity)。

### No-op と控えファイル

No-op は operation ごとに判定する。一部が No-op なら変更分だけ書き、
**控えファイル(audit)には実際に変更したセルだけ**を記録する。
全 operation × 全シートが No-op なら実行不可(出力を作らない)。
changes[] は元々複数件対応の構造だったので **schemaVersion は 1 のまま**。
記録順は利用者が入力した operation 順(シート内)を保持する。

### UI

5 番目のタブは作らず、「4. セルをまとめて変更」の単一入力欄を
「変更するセル一覧」(セル / 種類 / 値の編集可能な一覧)へ発展させた。
行の追加・選択行の削除・**表からの貼り付け**(タブ区切り 3 列、
種類は 文字 / 数値 / 空欄 のみ、1 行でも不正なら全体拒否・部分追加なし)。
クリップボードの読み取りは UI 実行時のみで、解析ロジックはテスト可能な形に分離。
一覧のどこか(行の追加・削除・セル・種類・値・順序・接尾辞)が変わったら
プレビューを stale にする。

### 変わらない点

D-023 / D-024 の安全 baseline(Source read-only、コピー→temp→検証→確定、
same-directory 出力、上書き禁止、snapshot / preflight、rollback best effort と
残存の正確な通知、audit sidecar、Excel 非起動、ネットワークなし)はすべて維持。

### 次(このフェーズではやらない)

Phase 2C: 入力セットの Value を Source Excel / CSV から供給する転記 Mapping。
Phase 2D: 検索・置換。Date write / Formula-aware mutation / in-place は引き続き後段。

## D-026: Phase 2C1 — キー一致で 1 行を特定し、決まったセルへ転記する (2026-08-28)

### Phase 2C を 2 つに割る

市場の再調査で、Source → Target 転記には性質の違う 2 種類があると分かった。

- **2C1**: データ元の 1 行から複数項目を取り出し、帳票・月報・店舗別ファイルの
  「決まったセル」へ入れる。今回実装するのはこちら。
- **2C2**: データ元の表と転記先の表を、SKU / 商品コード / 社員番号などで
  **行同士**突合して列を更新する。後段。

検索・置換は Phase 2D へ。

### 変更エンジンを二重に作らない

Phase 2C 専用の Workbook 書き換えエンジンは作らない。D-023〜D-025 の安全基盤
(Source read-only、コピー → temp → 検証 → 確定、AutoSave=false、対象
WorksheetPart のみ Save、StyleIndex 維持、数式ブック Block、結合 / 保護 /
入力規則 / リンク / リッチテキスト guard、snapshot、preflight、rollback と
残存報告、audit、Package Part integrity、Validator 0)をそのまま使う。

そのために **`ResolvedCellMutation`**(FilePath / SheetName / Address / Value /
Provenance)という「解決済みの 1 変更」を挟む共通形を作った。

```
Phase 2B 手入力の入力セット ─┐
                            ├→ ResolvedCellMutation の並び
Phase 2C1 キー一致の行 ─────┘        ↓
                            MutationPlanBuilder(走査・guard・No-op・出力計画)
                                     ↓
                                 CellMutator(書き込み・検証・確定)
```

Phase 2C は「新しい値をどこから得るか」だけが違う。既存の Phase 2B API は壊していない。

### データ元

1 回の実行につき 1 ファイル。`.xlsx` と `.csv`。読み取り専用で、出力も変更もしない。
**データ元を転記先にすることは禁止**(読みながら書くことになるため)。

`.xlsx`: シートを 1 つ選び、**項目名の行を明示指定**する(自動推定しない)。
読み取りはストリーミングで、Worksheet 全体を DOM に載せない。

`.csv`: カンマ区切りのみ。引用符・引用符内のカンマ・二重化した引用符・
引用符内の改行・CRLF / LF を正しく扱うため、自前の `split(',')` ではなく
`Microsoft.VisualBasic.FileIO.TextFieldParser` を使う。
列数がヘッダーと違う行があれば Block(読み取り位置がずれるため)。

### 文字コード

BOM があれば UTF-8 → 無ければ厳密な UTF-8 で復号を試す → 不正バイトがあれば
CP932(`CodePagesEncodingProvider`)。UTF-16 の BOM は「扱えません」と Block し、
CP932 として化けさせない。BOM は `StreamReader` に取り除かせ、
先頭の項目名に混ぜない。判定結果は画面に表示する(手動選択は今回不要)。

### 項目名

前後の空白だけ Trim。**空の項目名・Trim 後の重複は Block**。
`商品コード_2` のような自動リネームはしない。

### キー

- データ元のキー列は**文字列のみ**。数値・日付・真偽・数式・リッチテキスト・
  空欄は Block。「00123」と「123」を表示形式から推測しないため。
- 転記先のキーセルは、存在する・素の文字列・空欄でない・数式でない・
  リッチテキストでない・結合でない・メタデータ参照が無いこと。
  読むだけなので、入力規則やハイパーリンクは問題にしない。
- **正規化しない。** Trim / 大文字小文字 / 全角半角 / Unicode 正規化 / 数値化 /
  0 詰め / ハイフン除去は一切しない。`"ABC"` と `"abc"` と `" ABC "` は別キー。
  一致しなければプレビューで利用者に見せる。名寄せは後段。
- 必要なキーは**ちょうど 1 行**に一致すること。0 件・2 件以上は Block。
  今回使わないキーの重複は Warning に留める。
- 複数の転記先シートが同じキーを持つのは許可(同じ行を両方へ適用)。
- **キーのセル自身を転記先にすることは Block。** 実行中に照合の基準を変えない。

### 対応付け

`SourceColumn` → `TargetCell` + 種類(**文字 / 数値**のみ)。
「空欄」はデータ元から供給する値ではないので入れない(空欄にしたいなら Phase 2B)。
転記先セルは `TargetCellAddress` で正規化して重複を Block(`D5` と `$D$5` は同じ)。
同じ項目を複数のセルへ入れるのは許可。

値の変換で推測はしない。空欄・数式・リッチテキスト・日付 / 通貨 / % 書式の数値は
すべて Block。数値マッピングでは、`.xlsx` は素の数値書式(`NumberFormatCompatibility`
を再利用)のみ、`.csv` は InvariantCulture の有限数(桁区切りは可)としてのみ読む。

### データ元と転記先で数式の扱いが違う

- **データ元**: 今回読むキー / 値のセルが数式なら Block。読まない列・シートの
  数式は転記結果に影響しないので Block しない。
- **転記先**: ブック内に数式が 1 件でもあれば Block(D-023 のまま)。

混同しないよう、理由も別に書く。

### データ元の検証範囲(実測にもとづく判断)

当初はデータ元も `OpenXmlValidator` でブック全体を検証していたが、
**10 万行のシートで約 26 秒・約 1.4 GB** かかることを実測した
(参考: 同じシートのストリーミング走査は約 5.4 秒・約 1 MB)。

転記元は読むだけで、防ぎたいのは「値を取り違えて読むこと」。
書き換え対象の検証(D-023: 壊れたファイルを書き換えない)とは目的が違う。
そこで、**値の意味を決めるパート(シート一覧・共有文字列・表示形式)だけ**を検証する。
セルそのものの想定外は読み取り側が「読み取れません」として Block するので、
黙って通ることはない。転記先は従来どおりブック全体の検証を維持する。

### 走査の回数

転記先は **1 ファイル 1 回**しか開かない。シートの走査も 1 回で、
対象セルの安全確認と**キーセルの読み取りを同じ 1 パス**で行う。

データ元は **1 回だけ** scan する。手順は
「転記先から必要なキーを集める → データ元を 1 pass → 必要な行だけ保持」。
転記先ごとに開き直さない。重複検出のためキーの一覧だけは通して見る。

### 件数の上限

Phase 2B の「上限なし」を機械的に流用せず、実測してから決めた。

架空データ(10 万行のデータ元、転記先 500 ファイル × 5 シート × 20 項目):

| 解決済み変更 | Preview (csv / xlsx) | Execute | メモリ |
|---|---|---|---|
| 1,000 | 2.4 s / 5.1 s | 0.8–1.0 s | 〜15 MB |
| 10,000 | 1.8 s / 5.2 s | 1.7–1.8 s | 〜18 MB |
| 50,000 | 3.5 s / 7.0 s | 8.6–8.9 s | 〜98 MB |

いずれも実用範囲だったため、**解決済み変更の件数に上限は設けない**。
ファイル数の上限 500 は保険として維持する。

### Package Part integrity

Phase 2B と同じ。変わってよいのは実際に変更した WorksheetPart だけ。
2,000 変更 × 2 シートでも変わる ZIP entry は対象 2 つだけであることを実測。
データ元(Excel / CSV)は完全に不変。

### 控えファイル(schemaVersion 2)

転記の出力だけ **schemaVersion 2 / operation `map-source-to-cells`** にする。
Phase 2A / 2B の手入力は **1 / `set-cell-value`** のまま(意味を変えない)。

追加するのは `dataSource`(fileName / sha256 / type / sheetName / headerRow /
keyColumn)と、各変更の `sourceColumn` / `key` / `sourceRowNumber`。
**絶対パスは書かない。** No-op は従来どおり記録しない。

行番号は `.xlsx` は実際のワークシート行番号、`.csv` は
**CSV レコード番号**(ヘッダーが 1)。引用符内の改行があるため物理行番号は使わない。

### snapshot

転記先に加えて**データ元(SHA-256 / サイズ / mtime)も控え**、実行直前に再確認する。
プレビュー後にデータ元が変わっていたらバッチ全体を中止する
(値そのものが変わってしまうため)。

### UI

5 番目のタブ「5. 表から転記」。既存 4 タブは維持。
データ元の選択 → シート / 項目名の行 → 「項目を読み込む」→
キー列(項目名から選ぶ)+ 転記先のキーセル → 対応付けの一覧 → 転記先シート → プレビュー。

利用者には `Column A` ではなく**項目名**を選ばせる。
プレビューは件数が増えるため、**最初から仮想化された DataGrid** を使う
(Phase 2B のプレビューは非仮想化のまま。今回は必須としない)。

出力の既定の接尾辞は `_転記済み`。同フォルダー出力・上書き禁止・
控えファイルの衝突も Block は Phase 2B と同じ。

### 次(このフェーズではやらない)

Phase 2C2(表同士の行キー突合・更新)、Phase 2D(検索・置換)。
Date / Boolean の書き込み、空欄の転記、行の追加、あいまい一致・名寄せ、
Formula-aware mutation、in-place 上書きも引き続き後段。

## D-027: Phase 2C2 — キーで行同士を突合し、既存行の指定列を更新する (2026-08-28)

### 2C1 との違い

2C1 は「データ元の 1 行 → 決まったセル(D5 / F8)」。
2C2 は「データ元の表 → 転記先の表」で、行番号ではなく**キーで行同士**を対応付けて
既存行の指定列を書き換える。商品マスタ・在庫表・台帳の更新が主用途。

### 既存行のみ・intersection のみ

更新するのは **Source と Target の両方に存在する unique key** の行だけ。

- Source にしか無いキー → **行を追加しない**(Warning で件数を知らせる)
- Target にしか無いキー → **行を削除・空欄化しない**(Warning で件数を知らせる)
- 一致が 0 件 → Block(「一致する行がありません。キー列や表記を確認してください。」)

実務では Source が差分・Target が履歴込みのマスタなど、片側だけの行が正常に
存在するため Block にしない。ただし黙って無視せず、プレビューの集計に
一致 / データ元のみ / 転記先のみ / 重複 / 空欄行 を必ず出す(TableMatchSummary)。

### キー

- Source / Target でキーの**列名は違ってよい**(SKU ↔ 商品コード)。
- 文字として完全一致のみ。正規化(Trim / 大小文字 / 全角半角 / 数値化 / 0 詰め)は
  一切しない。D-026 と同じ原則。
- 両側とも文字列のみ。**Target のキー列に非空欄の非文字列(数値・数式・
  リッチテキスト・cm/vm)が 1 つでもあればシートを Block**。「A001」と数値 1 を
  推測で比べない。
- キーが空欄の行は表の一部とみなさない。更新対象の列にも値が無ければ空行として
  読み飛ばし、値があれば Warning 付きで読み飛ばす。空欄同士を一致させない。

### 重複キー

- Source の同一キーが 2 行以上: そのキーが更新に**使われるなら Block**
  (どの行の値を使うか決めない)。使われないなら Warning。
- Target の同一シート内の同一キーが 2 行以上: 更新対象なら **Block**
  (最初の行・最後の行・両方のどれかを勝手に選ばない)。対象外なら Warning。
- **別シートに同じキーがあるのは正常**。同じ Source 行で両方を更新できる。

### 対応付け

SourceColumn → TargetColumn + 種類(文字 / 数値)。値の規則は D-026 の
`SourceValueConversion` をそのまま共有(空欄・数式・リッチテキスト・
日付 / 通貨 / % 書式は Block)。

- 同じ TargetColumn への重複は Block(定価と売価を同じ「価格」へ、は決めない)。
- 同じ SourceColumn を複数の TargetColumn へは可。
- **Target のキー列は更新先にできない**(照合の基準を実行中に変えない)。

### 転記先の表の読み方

- Target Header Row は全シート共通で明示指定(シートごとに違うのは LATER)。
- 項目名は Trim のみ。空・重複は Block(データ元と同じ `SourceHeaders.Validate` を
  件名だけ変えて共有)。ヘッダーの範囲は「その行に実在する最も右のセルまで」。
- **列は名前で解決**するので、シートごとに列の位置が違ってよい。余分な列があってよい。
  必要な列が 1 シートでも欠ければ全体 Block。
- UI では「項目を読み込む基準シート」を利用者が明示選択する(追加順の先頭を
  黙って基準にしない)。基準は候補表示に使うだけで、各シートはプレビューで個別検証。
- 「最初の空行で表終了」とはしない。実在する行を最後まで走査する(途中の空行は上記の扱い)。
- 行は昇順で並んでいる前提とし、項目名の行より前にデータ行が現れる XML は
  黙って読み違えず Block する(Excel が保存するファイルは昇順)。
- Excel Table(ListObject)が存在するだけでは Block しない。セル値だけを変え、
  Table のメタデータには触れない(calculated column 等は数式 guard で止まる)。

### エンジンを fork しない

新しく作ったのは「どの行のどのセルへ、どの値を書くかの解決」まで。

- 転記先の走査は新設 `TargetTableScanner` だが、セルの guard は既存
  `CellMutationScanner.ScanTargetCell` を internal 化して共有。Workbook 単位の
  guard(数式・接続・外部参照・Validator)も既存 `AddWorkbookBlocks` を共有。
- 解決後は `ResolvedCellMutation` → `MutationPlanBuilder` → `CellMutator` という
  2B / 2C1 と同一の経路。部分適用なし・snapshot・preflight・rollback と残存報告・
  Package Part integrity もそのまま。

### 走査の回数

- Source は **2 パスのストリーミング**。Pass A でキー列だけ(集合・重複・空欄数)、
  Target 走査で一致キーを確定してから、Pass B で**一致したキーの値だけ**を読む。
  10 万行 × 20 列を全部メモリに持たない。
- Target Workbook は 1 ファイル 1 回 open、各シート 1 パス。その 1 パスで
  ヘッダー解決・キー検査・一致行の確保・結合 / リンク / 入力規則の範囲収集まで行い、
  guard は一致行 × 対応付けのセルだけに掛ける。
- キーごとの再 open / 再 scan は無い。

### Validator と大規模 Target(実測)

D-023 の「壊れた Target を書き換えない」原則は軽くしない(preflight・出力とも
ブック全体の OpenXmlValidator を維持)。そのうえで実測した:

| 対象 | 時間 / メモリ |
|---|---|
| full Validator 単体(6 列) | 10k 行 1.5 秒 / 49 MB、50k 行 4.1 秒 / 213 MB、100k 行 7.4 秒 / 427 MB |
| target 10k・一致 1k・5 列 | Preview 4.3 秒 / Execute 1.4 秒 |
| target 50k・一致 10k・5 列(5 万セル更新) | Preview 8.1 秒 / Execute 6.4 秒 |
| target 100k・一致 50k・5 列(25 万セル更新) | Preview 13.0 秒 / Execute 12.4 秒 |
| target 100k・21 列(210 万セル) | Preview 37 秒 / Execute 41 秒、約 1.5 GB |

**上限**: 1 シートあたりのキー行数を、動作を確認した 100,000 行で打ち切り、
超えるものは「動作を確認した範囲を超えています」と Block する
(`TargetTableScanner.MaxKeyedRowsPerSheet`)。Validator を外すのではなく、
実測済みの範囲に対象を限定する、という §39 の初期案どおり。

### 併せて直した性能欠陥(全 Phase に効く)

初回実測で「50k セル更新の Execute が 244 秒」となり切り分けたところ、
`CellMutator` の**出力検証**が変更 1 件ごとに `Descendants<Cell>().FirstOrDefault`
で線形探索しており、大きな表で二次時間になっていた(書き込み側は 2B で
辞書化済みだったが、検証側が漏れていた)。シートごとに 1 パスで対象セルを
辞書に引く方式へ直し、244 秒 → **6.4 秒**。2A / 2B / 2C1 の検証も同じ経路なので
すべて速くなる。既存 636 テストに回帰なし。

### 控えファイル(schemaVersion 3)

2C2 の出力のみ **schemaVersion 3 / operation `map-source-table-to-target-table`**。
v1(手入力)/ v2(固定セル転記)の意味は変えない。

追加: `targetTable { headerRow, keyColumn }`、各 change の `targetColumn` と
`targetRowNumber`(v2 からの `sourceColumn` / `key` / `sourceRowNumber` に加えて)。
絶対パスなし。No-op は記録しない。

### 出力

既定の接尾辞は「_更新済み」(突合済みより日常語に近い)。同フォルダー・
上書き禁止・控えファイル衝突 Block は従来どおり。

### UI

6 番目のタブ「6. 表を突合して更新」。既存 5 タブは維持。プレビューは
2C1 と同じく最初から仮想化 DataGrid で、突合の集計(一致 / 片側のみ / 重複 / 空欄)を
常に表示する。対応付けの貼り付け(タブ区切り)は 2C2 本体を膨らませないため
今回は見送り(2B の operation 貼り付けは既存のまま)。

### 次(このフェーズではやらない)

Phase 2D(検索・置換)。行の追加・削除、複合キー、あいまい一致・名寄せ、
Date / Boolean / Formula の書き込み、空欄の転記、in-place 上書きは引き続き後段。

## D-028: Phase 2D — 処理設定(レシピ)の保存・再利用 v1 (2026-08-28)

### なぜ Search / Replace より先か

2C2 までで「毎月の更新作業」を安全に実行できるようになったが、毎回
キー列・項目名の行・対応付け・出力名を指定し直す必要がある。
機能を増やすより、すでに作れる処理を**繰り返し使えるようにする**方が
1 回あたりの時間を確実に減らす。Search / Replace は Phase 2E へ移す。

### 対象

タブ 4(セルをまとめて変更)/ 5(表から転記)/ 6(表を突合して更新)のみ。
Phase 1 系(表をまとめる・シートをまとめる)と解析タブは対象外。
必要が確認できた時点で拡張する。

### 「処理ルール」と「今回使うファイル」を分ける

レシピに入れるのは**ルールだけ**。

- 入れる: 名前・種類・項目名の行・キー・セル位置・対応付け・値の種類・
  固定値・出力の接尾辞・(.xlsx のとき)データ元のシート名
- 入れない: データ元 / 転記先の絶対パス・ファイル名・SHA・プレビュー結果・
  控えの内容・実行結果・一時パス・ユーザー名・PC 名

毎月ファイル名が変わっても同じレシピを使えることが目的なので、
ファイル名も保存しない(テストで JSON の中身を検査して固定)。
「基準シート」は今回選んだファイルに依存するため v1 では保存しない。

### 置き場所と形式

`%LOCALAPPDATA%\ExcelBatchTool\recipes.json` のみ。完全ローカル。
Git repo 内・publish フォルダー・Excel と同じフォルダーには作らない。

`schemaVersion` を必ず持ち(v1)、未知の版は勝手に解釈せず読み込みエラーにする。
種類は `cell-input-set` / `source-to-fixed-cells` / `source-table-to-target-table`、
値は `text` / `number` / `blank`、データ元は `xlsx` / `csv` という**固定文字列**で、
C# の名前と切り離す(後から C# 側を改名しても読める)。未知の文字列はエラー。

型ごとに payload を分けた(`object` / `dynamic` を使わない)。

### 読み込みの安全原則

レシピを読み込んでも:

- プレビューを自動実行しない
- Excel 処理を自動実行しない
- 出力を作らない

読み込み = 設定の反映 + **必ず stale** + 実行不可。
そこから利用者が今回のファイルを確認し、プレビューを更新し、
2A〜2C2 の安全チェックを毎回通してから実行する。
レシピは安全チェックを飛ばす機能ではない。autoRun / schedule は持たない。

画面に追加済みの Workbook や選択中のデータ元は消さない(選び直しの手間を
増やさない)が、レシピがそれを「確認済み」と扱うことはない。

### 保存できる条件

最後のプレビューが最新(fresh)で、実行できない問題(Block)が無いこと。
Warning は妨げない。

**例外を 1 つだけ置いた**: Block が「変更が必要なセルがありません」だけの場合は
保存できる。設定そのものは正しく、今月はたまたま値が同じというだけで、
翌月は変わり得るため。判定は文言比較ではなく
`CellMutationPreview.IsBlockedOnlyByHavingNothingToChange` に集約した。

なお「同じ名前の出力ファイルがすでにある」も設定としては正しいが、
これは Block のままにした(区別するには Block を「設定の問題」と
「今回のファイルの問題」に分類する設計が要り、2D の範囲を超える)。
実行直後にそのまま保存しようとすると止まる。TODO へ記録。

### 見つからない設定を推測しない

- 保存したシートが今回のファイルに無い → 先頭シートを黙って使わず、
  シート名を挙げて選び直してもらう
- 保存した項目が今回の項目名に無い → 似た名前を使わず、名前を挙げて空にする
- レシピが CSV 用なのに .xlsx を選んでいる → 互換扱いせず知らせる
  (CSV と .xlsx でシート・ヘッダーの意味が違うため、v1 は一致を求める)

この過程で既存の手動操作の欠陥も直した: 項目を読み込み直したとき、
選択済みのキー列が新しい項目名に無いと**黙って先頭の列に置き換わっていた**。
名前を挙げて空にするよう改めた(タブ 5・6 とも)。

### 保存の壊れにくさ

同じフォルダーへ一時ファイル → flush → `File.Replace` で置き換え、
1 つ前の内容は `recipes.json.bak` として残る。
既存を消してから書く方式は使わない。失敗時は元の `recipes.json` を維持する。

読めないファイルは**黙って空にしない・自動で上書きしない**。
「読み取れません」と伝え、Excel 処理そのものは通常どおり使える。
保存・更新・削除も、まず読み込みに成功しない限り書き込まない。

控えが正常でも**自動復旧はしない**(v1)。「壊れた本体を黙って控えで
上書きする」ことはせず、控えが残っている事実だけを伝える。

### 名前

前後の空白を落として 1〜60 文字、改行・制御文字は不可、
大文字小文字を区別せず同名は重複扱い。同名は黙って上書きしない。
「更新」「削除」は確認してから実行する。並び順は名前の昇順
(大文字小文字を区別しない)。「最近使った順」は使用履歴が要るので持たない。

### 実測(架空データ、Release)

| レシピ数 × 対応付け | 読み込み | 1 件保存 | ファイル |
|---|---|---|---|
| 10 × 5 | 0 ms | 3 ms | 14 KB |
| 100 × 20 | 5 ms | 11 ms | 344 KB |
| 1,000 × 20 | 30 ms | 39 ms | 3.4 MB |
| 1,000 × 100 | 162 ms | 215 ms | 15 MB |

保存は毎回ファイル全体を書き直す(壊さないための代償)が、現実的な規模では
数 ms、極端な 15 MB でも 0.2 秒。件数の上限は設けない。
画面側も 1,000 件の一覧で選択・表示 130 ms。

### レシピと控え(audit)は別物

レシピ = 次回も使いたい処理設定。控え = 今回何を変更したかの記録。
レシピに実行結果や控えを入れない。控えにレシピ情報を入れない。
audit の schemaVersion 1 / 2 / 3 は変更しない(テストで固定)。

### 実装の置き場所

`Core/Recipes`(モデル・検証・保存)と、画面側の共通領域
`RecipeAreaViewModel` + `IRecipeHost`。ViewModel が
`File.ReadAllText` / `JsonSerializer` を直接持たない。
DI framework は導入せず、置き場所(パス)の指定と確認ダイアログの
差し替えだけをコンストラクター引数にした。

Mutation Core をレシピ専用に fork していない。レシピは既存の
Request / ViewModel へ設定を流し込むだけで、D-023〜D-027 の安全性
(Source 不変・snapshot・Validator・数式 guard・結合・入力規則・リンク・
rollback・出力衝突・Package Part integrity・10 万行上限)は一切弱めていない。

### 次(このフェーズではやらない)

書き出し / 読み込み(他 PC への持ち出し)は v1 では行わない。固定値を含み得る
ため外部共有の設計と版の移行設計が要る。クラウド同期・自動実行・
Phase 1 系のレシピ化・レシピの暗号化の要否も LATER。

## D-029: Phase 2D.1 — 正常実行を根拠にしたレシピ保存 (2026-08-28)

D-028 は書き換えない。保存できる根拠を 1 つ足す補正。

### 直したこと

D-028 の保存条件は「最新プレビュー + 保存を妨げる Block なし」。
ところが正常に実行すると `<元名>_変更済み.xlsx` 等が作られ、
その後プレビューを取り直すと**今作った出力との同名衝突で Block** になる。

結果、利用者にとって最も自然な順序:

設定 → プレビュー → 実行 → 成功を確認 → 「次回も使いたい」 → 保存

だけができなかった。Excel に詳しくない人ほど「実際に成功してから保存したい」
と考えるので、そこを塞いだままにしない。

### 「保存のときだけ衝突を無視する」ことはしない

Block の意味は変えない。同名衝突は**再実行のときは引き続き Block**。
プレビュー・実行・Validator・snapshot・衝突確認・rollback のどれも迂回しない。

代わりに、保存だけに別の根拠を足す:

> 今の設定は、最後に正常終了した実行で使った設定とまったく同じ。
> つまりその実行で実際に確認・実行済みである。

### 保存できる条件(2D.1 以降)

A. 従来どおり: 最新プレビュー + 保存を妨げる Block なし
  (「今回は変わるところが無い」だけの Block は D-028 のとおり許す)
または
B. 今の設定が「最後に正常終了した実行の設定」と完全一致

`RecipeSaveGuard.ReasonFor(preview, isStale, matchesLastSuccessfulRun)`。
B が成立するときは A を見ない。

### 何を比較するか

レシピ JSON に入る**処理ルールだけ**を型付きで比較する
(`Core/Recipes/RecipeConfiguration`)。JSON 文字列の比較は
プロパティ順や将来の版に依存するので使わない。`object` / `dynamic` も使わない。

- タブ 4: operations(cell / kind / value・並び順)と outputSuffix
- タブ 5: sourceFileKind / sourceSheetName / headerRow / sourceKeyColumn /
  targetKeyCell / mappings / outputSuffix
- タブ 6: sourceFileKind / sourceSheetName / sourceHeaderRow / sourceKeyColumn /
  targetHeaderRow / targetKeyColumn / mappings / outputSuffix

文字は完全一致(前後の空白・大文字小文字も区別)。D-026 以来の原則と同じ。

### 何を比較しないか

レシピに入らないものは一致判定にも使わない: データ元 / 転記先の絶対パス・
ファイル名・SHA・転記先の選択・基準シート・プレビュー結果・控え・出力パス・実行日時。
名前・Id・作成日時・更新日時も除く(実行時点ではまだ名前が無いことがある)。

したがって:

- 翌月ぶんの別ファイルを選び直しても、ルールが同じなら一致のまま保存できる
- 転記先シートの選択を外しても一致のまま
- CSV 用の設定で .xlsx を選ぶと `sourceFileKind` が変わるので不一致

「一致するから実行してよい」ではない。**保存してよい**だけ。

### 控えの取り方

実行を始める直前にその実行で使う設定を控え、`Success == true` のときだけ
`_lastSuccessfulRecipe` に採用する。実行中に画面を触られても、根拠になるのは
「実際に流した設定」。Block・失敗・中止・rollback・残存はいずれも採用しない。

メモリ上だけで持ち、ファイルには書かない。アプリを閉じれば消える。
別のタブを見て戻ってきても、設定が変わっていなければ有効。

### 「一度成功したから何でも保存可能」にしない

成功後に値を変えれば不一致になり保存できない。元に戻せばまた保存できる。
判定は「変更イベントが起きたか」ではなく**今の内容が同じか**。

### バイパスしないもの

- `RecipeStore.Add` / `Update` の検証は従来どおり必ず通る
- 名前の決まり・同名禁止・壊れた recipes.json への書き込み拒否も従来どおり
- 自動保存はしない。名前を入れて「現在の設定を保存」を押したときだけ
- 既存レシピの自動更新もしない
- レシピ読み込み時の load → stale → preview → execute も従来どおり

### 画面

正常終了の直後に、処理設定の欄へ
「今回使った設定は、処理設定として保存できます。名前を入れて〜」と出すだけ。
checkpoint / snapshot / fingerprint のような言葉は画面に出さない。

### 3 タブ共通

タブ 4 / 5 / 6 で同じ意味で実装した。1 タブだけ特殊にしない。

### 残した課題

同名衝突以外の「今回のファイル由来の Block」(例: 転記先が別プロセスに
開かれている)も設定としては正しいが、B を満たさない限り保存できない。
Block を「設定の問題」と「今回のファイルの問題」に分類する設計は
今回も行わず、実測される必要が出た時点で検討する。

## D-030: Phase 2E — CSV 変換・整形 v1 (2026-08-28)

D-029 以前は書き換えない。

### なぜ Search / Replace より先か

実案件を調べ直したところ、次のような依頼が継続的に確認できた。

- システム出力 CSV → 別システム用 CSV
- 毎月の求人 CSV のフォーマット変換
- 商品 CSV の整形・変換
- 不要列の削除・列順の変更・固定値の追加

2C1 / 2C2 で CSV を**読む**ことはできるようになったが、
「指定されたフォーマットの CSV を**作る**」機能が無かった。
Search / Replace 単体より、こちらを先にする。Search / Replace は Phase 2F へ移す。

### v1 でできること

出力する列を 1 件ずつ決める。それだけに集中する。

- データ元の列を選ぶ / 出力する項目名を変える / 並び順を変える
- 出さない列を選ぶ(内部メモを外部へ出さない、など)
- 同じデータ元の列を複数の出力列へ使う
- 固定値の列・空欄の列を足す

### v1 でやらないこと

行の絞り込み・条件分岐・正規表現・検索置換・Trim の自動化・全角半角変換・
日付変換・数式・計算・文字列結合・列分割・重複削除・集計・あいまい一致・
CSV 同士の結合・Excel の書き換え・Web アップロード・in-place 上書き。

### データ元

1 回につき 1 ファイル。`.csv` / `.xlsx` とも 2C1 / 2C2 の reader をそのまま使う
(UTF-8 BOM / BOM なし / CP932、引用符内のカンマ・二重引用符・改行、CRLF / LF)。
`.xlsx` はシートと項目名の行を明示指定(自動推定なし)。読み取りのみ。
項目名は前後 Trim だけ。空・重複は Block(`SourceHeaders` を共有)。

### 出力する 1 列

`OutputName` / `ValueSourceKind`(SourceColumn / FixedText / Blank)/
`SourceColumn` / `FixedValue`。

- 出力する項目名は Trim 後 1 文字以上、改行・制御文字不可、**重複不可**。
  大文字小文字だけの違い(Price / price)も、取り違えを避けるため
  `OrdinalIgnoreCase` で重複として Block する。
- 存在しないデータ元の項目は Block。似た名前を代わりに使わない。
- 固定値は文字としてそのまま全行へ入れる。数値として解釈しない
  (CSV は文字の集まりとして扱い、Excel の型推測を持ち込まない)。

### .xlsx のセルをどう文字にするか(characterize して決めた)

通す: 素の文字列(共有文字列・インライン文字列)、素の数値、空欄。
止める: 数式、日付・時刻、通貨・パーセント・あいまいな表示形式、
リッチテキスト、**TRUE / FALSE**。

boolean は「許可を characterize する」対象だったので調べた。.xlsx の内部では
`t="b"` の 0 / 1 で、画面では TRUE / FALSE と見える。CSV に書くとき
「1」と書くか「TRUE」と書くかは受け取り側の都合であって、こちらで決められない。
2C1 / 2C2 と同じく **v1 では Block** とし、必要になったら選べる形で足す。

数値は `InvariantCulture` でそのまま(1200 → `1200`、0.5 → `0.5`)。

### 行

項目名の行より後を、同じ順で 1 行ずつ出す。追加・削除・並べ替え・
絞り込み・重複削除はしない。すべての項目が空欄の行だけ読み飛ばし、
件数を Warning で知らせる。一部空欄の行はそのまま出す。

### CSV の書き出し

`string.Join(",")` のような連結はしない。専用の `CsvWriter` を置き、
RFC 4180 相当:カンマ・引用符・CR・LF を含む項目は引用符で囲み、
中の引用符は 2 つに重ねる。「すべての項目を囲む」も選べる(既定は必要なときだけ)。
実案件に「全列ダブルクォート」の指定が実在するため。

文字コードは UTF-8 BOM(既定)/ UTF-8 / CP932。既定を BOM ありにしたのは、
Windows の Excel でそのまま開く使い方を想定しているため。
行の終わりは v1 では **CRLF 固定**(Windows 業務ツール向け)。

### 出力と確認

同じフォルダーへ `<元の名前>_変換済み.csv`。上書き禁止。
一時ファイルへ書く → 読み直して確認 → 移動、の順。

確認は「書きながら作った指紋」と「読み直して作った指紋」の突き合わせ。
項目の長さと中身を順に流し込んだ SHA-256 なので、**全項目を照合しつつ
どちらの内容もメモリに置かない**。項目名・行数・列数・BOM の有無も見る。
食い違えば作りかけを片付け、消せなかったものはファイル名で知らせる(D-024 と同じ)。

### 走査の回数

プレビューで 1 パス(行数・読み取れないセルの確認・先頭 20 行の見本)、
実行で 1 パス(書き出し)+ 読み直し 1 パス。
表を丸ごと DOM や DataTable に載せない。

プレビューで**全行を確かめる**ことにしたので、実行の途中で初めて
「読み取れないセル」が見つかることはない。

### 実測(架空データ、Release)

| データ元 | 行 × 列 | 出力列 | プレビュー | 実行 | メモリ |
|---|---|---|---|---|---|
| CSV 0.2 MB | 1,000 × 20 | 10 | 54 ms | 83 ms | 6 MB |
| CSV 2.1 MB | 10,000 × 20 | 10 | 216 ms | 483 ms | 2 MB |
| CSV 29.7 MB | 50,000 × 50 | 25 | 1.4 s | 2.6 s | 85 MB |
| CSV 59.9 MB | 100,000 × 50 | 25 | 3.0 s | 4.8 s | 165 MB |
| CSV 59.9 MB | 100,000 × 50 | 50 | 2.8 s | 6.7 s | 162 MB |
| .xlsx 6.5 MB | 100,000 × 20 | 10 | 5.0 s | 5.7 s | 8 MB |

Excel の書き換え(2C2 は 100k 行 × 21 列で約 1.5 GB)と違い、
OpenXML の DOM も Validator も通らないので桁違いに軽い。行数の上限は設けない。

### レシピ

Phase 2D / 2D.1 があるので最初から対応した。4 番目の種類 `csv-transform`。

保存: sourceFileKind / sourceSheetName / headerRow / outputColumns /
encoding / quoteMode / outputSuffix。
保存しない: データ元のパス・ファイル名・SHA・プレビュー・出力先。

**schemaVersion は 1 のまま**にした。既存の 3 種類の payload は変えず、
`csvTransform` を足しただけの追加変更で、2D 時点で書かれたファイルが
そのまま読める(テストで固定)。新しい enum(入れ方・文字コード・引用符)にも
C# 名と切り離した固定文字列を与えた。version 2 が要る破壊的変更ではない。

実行成功後の保存(D-029)も同じ形で効く。

### 控え

`<出力名>.csv.audit.json` に schemaVersion 1 / operation `csv-transform`。
Excel 書き換えの控え(v1 / v2 / v3)とは操作の意味が違うので、
同じ形式へ押し込まず別の形式にした。

source(ファイル名・SHA・種類・シート・項目名の行・文字コード)、
output(ファイル名・文字コード・引用符・行の終わり・行数・列数)、
columns(出力する項目名・入れ方・元の項目 / 固定値)。絶対パスは書かない。
固定値が控えに入るので、ローカルのファイルであることを README に明記した。

### 画面

7 番目のタブ「7. CSV を変換」。既存 6 タブは変更なし。
「元の項目をすべて追加」を置いて、20 列の CSV でも 1 クリックで並べられるようにした。
高さが足りないときは編集部分だけがスクロールする(最小サイズでも項目一覧へ届く)。

### 次(このフェーズではやらない)

Phase 2F(検索・置換)。行の絞り込み・並べ替え・重複削除・正規表現・
文字の正規化・計算・日付変換・CSV 同士の結合・Web アップロードは後段。

## D-031: Phase 2E.1 — 作業用ファイルの所有権 (2026-08-28)

D-030 は書き換えない。

### 見つかった差異

Phase 2E の CSV 変換は、作業用ファイルを固定名 `<出力名>.tmp` で作っていた。

1. **CSV の作業用ファイル**: `created.Add(tempPath)` を `FileStream` の
   作成より**前**に行っていた。実行前から同名のファイルがあると
   `FileMode.CreateNew` が失敗し、その catch から呼ばれる rollback が
   「この実行が作っていない既存ファイル」を削除し得た。
2. **控えの作業用ファイル**: `File.WriteAllText` を使っていたため、
   実行前から同名のファイルがあれば**上書き**した。さらに登録は書き込みの
   **後**だったので、書き込みが途中で失敗した場合に、変更されたファイルが
   rollback の管理外に残り得た。

いずれも D-023 / D-024 以来の
「取り消しの対象は、この実行が作ったものだけ」という原則に反する。

### 直し方

- 作業用ファイルの名前を、実行ごとに一意な `<最終名>.<guid>.tmp` にした。
- 作成は `FileMode.CreateNew`。名前が既に使われていたら**そのファイルには
  一切触れず**、別の名前で作り直す(8 回まで)。
- `CreateNew` がストリームを返した瞬間 = 所有が確定した時点で rollback の
  対象へ登録し、**そのあとで**中身を書き始める。所有しているのに未登録、
  あるいは登録済みなのに他人のもの、という窓を作らない。
- 控えも同じ経路にした。`CsvTransformAuditLog.Write` はパスではなく
  **呼び出し側が新規作成したストリーム**を受け取る形に変え、
  自分でファイルを開かない(既存ファイルを開く余地を無くす)。

### 変えていないこと

Phase 2E の仕様は変更していない。

- 最終的な出力名・控えの名前に一意な値は残らない(作業用ファイルにしか使わない)
- 出力・控えの上書き禁止はそのまま
- 取り消しで消せなかったものをファイル名で知らせるのもそのまま
- データ元は読み取りのみ

### なぜ固定名をやめたか

固定名のままでも「登録の順序を直す」だけで問題 1 は塞げるが、
それでも「前回の残りと同じ名前を使う」構造は残る。同じフォルダーで
同時に 2 つ動かした場合や、前回の残骸がある場合に、
どちらの持ち物か名前から判断できない。実行ごとに違う名前にすれば、
「自分が新規作成できたものだけが自分のもの」と一意に決まる。

## D-032: Phase 2F-R — PDF 対応の実現性ベンチマーク (2026-08-29)

D-031 以前は書き換えない。これは**実装決定ではなく、実装可否の調査の記録**。
詳細な数値・ライセンス・推奨構成は docs/research/pdf-feasibility-research.md。

### 何を測ったか

CrowdWorks / Lancers の PDF→Excel 案件を広く狙える前提で、
(1) PC 生成 PDF、(2) PDF 内の表、(3) スキャン PDF、(4) 同一レイアウト大量帳票、
(5) 複雑なスキャン表、(6) 手書き(characterization のみ)を、
すべて架空の生成コーパス + 機械照合の Ground Truth で実測した。
指標は「文字が取れたか」ではなく**納品できる完全一致率**(NFKC + 空白除去後)。

### 判定: GO(条件付き)。1〜4 の全系統が実用候補

- born-digital 文章: PdfPig で field exact **100%**
- born-digital 表: 罫線 hybrid(構造=Tabula、文字=PdfPig 詰め直し)と
  自前 header-guided 再構成の両方で cell exact **100%**(罫線なし含む)
- 印刷文字スキャン帳票(120 ページ): PP-OCRv5 で field exact **98.6%**、
  数値・コード・日付・金額は **100%**。2.7 s/ページ・CPU のみ
- confidence 閾値 0.98 で **誤確定 0** のまま 82% 自動確定 + 18% 要確認
  = 全フィールドを確認可能状態にできる(劣化スキャンでも誤確定 0 を維持)
- スキャン罫線表: SLANet は列格子が崩れ 0%。**罫線格子(OpenCV)+ OCR 割当**で
  94.4%・行数 41/41 を全ページ再現
- チェックボックス: 固定座標の画素解析で clean **100%**
- Tesseract は clean 80.3% / degraded 42.4% で全面劣後(baseline 記録のみ)

### GO に付けた条件

1. **単一 OCR エンジンに依存しない。** PP-OCRv5(多言語)は数値・英数字が完全な一方、
   日本語の一部の字(支・一・促音・長音)を高 confidence のまま落とす。
   日本語専用 japan_PP-OCRv4 はかな漢字がほぼ完全な一方、英数字コードが 15% まで崩れる。
   性格が相補的なので、二重読み照合または v5 confidence gate を前提にする。
2. 確定は「自動確定 / 要確認 / 読取不能」の 3 分類 UI 前提。勝手に Excel へ確定しない。
3. 領域指定 OCR・チェックボックスは deskew(傾き補正)実装まで clean スキャン限定。
   実測では**全ページ OCR + ラベルアンカーが領域指定より精度・速度とも優位**だった。
4. 日本語専用 rec と SLANet は Paddle 2.x runtime でしか動かない(3.x は旧形式モデルを
   実行不可)。採用時はランタイム版数の固定が必要。

### 事前予想と違った点(調査の価値があった箇所)

- 領域指定 OCR は全ページ OCR より精度も速度も悪い(クロップ余白 5pt で 87%→60%)
- SLANet(表構造モデル)より古典的な罫線検出のほうが圧倒的に正確
- 埋め込みテキストは康熙部首など互換コードポイントで返ることがある → NFKC 必須
- tabula-sharp は連続同一文字を潰す(188→18)→ PdfPig letter 詰め直しで回避
- 手書きは合成データでは測れない(100% になってしまう)ことも記録

### 配布・依存・ライセンス

- PDF 層のみなら現行 168.6MB **+約 25MB**。OCR 込み実測 611MB → 現実推定 ~478MB
  (2.x runtime なら ~460MB)。**「本体 + Offline OCR Pack」分割を推奨**
- Python / Java / 実行時 DL / cloud / telemetry: すべて不要・不使用
- 全依存が Apache-2.0 / MIT / BSD 系で商用再配布可。Intel MKL のみ製品同梱前に
  最終確認(Openblas 差し替えで回避可)。MuPDF 系は指示どおり除外(AGPL)
- VC++ ランタイムは app-local 同梱を製品要件にする

### 残る不得意

罫線なしスキャン表(未実測)、手書き(対象外)、傾きスキャンでの領域指定、
二重読み照合の統合精度(設計提案どまり)。

### このフェーズでやっていないこと

製品本体への組み込み(タブ・ViewModel・Recipe・README)はしていない。
GO 判定でも Phase 2F 本体実装には進まない(指示どおり)。

## D-033: Phase 2F-A — born-digital PDF の製品実装 (2026-08-29)

D-032 以前は書き換えない。

**Phase 2F-A COMPLETE ≠ Phase 2F COMPLETE。** 2F 全体は
「普通の PDF / 表 PDF / スキャン PDF / 大量定型帳票 / 要確認・修正 / Excel・CSV 出力」
まで含む。今回はそのうち **文字情報を持つ PDF(文章・表)だけ**を製品化した。

### 今回やったこと

タブ 8「PDF を読み取る」。PDF を 1 つ選ぶと種類を自動判定し、
文章 PDF は「ページ / 行 / 内容」、表 PDF は「そのままの行・列」で取り出して
新しい .xlsx / .csv を作る。元の PDF は読み取りのみ。完全オフライン。

### 今回やらなかったこと(2F-B 以降)

OCR・PaddleOCR・スキャン PDF の本処理・チェックボックス・deskew・
固定帳票設定・手書き・Offline OCR Pack・PP-Structure・Paddle 2.x runtime。
OCR 用の VC++ ランタイム同梱も無し(製品の依存は増えていない)。

### 自動判定(利用者に種類を聞かない)

ページごとに「文字数・画像の被覆率・横罫線・縦罫線・列のそろい」を測り、
文書全体を text / table / scan / mixed / unknown に分ける。

2F-R の判定式から 2 点変えた。どちらも研究時の閾値が実データで破綻したため。

1. **画像だけのページ**を「文字数 < 10」だけで決めていたが、
   1 行しか書いていない表紙のようなページを誤ってスキャン扱いした。
   「文字が 0」または「文字がごく少ない **かつ** ページの半分以上が画像」に変更。
2. **表の判定**を「横罫線 5 本以上」だけで決めていたが、
   2 行の小さな表(横 3 本)を落とし、記入欄の下線が並ぶ帳票を表と誤認した。
   **縦横の罫線が組み合わさって格子になっているか**(横 3 本以上 かつ 縦 3 本以上)、
   または**列の位置がそろった行が 6 割以上続くか**に変更。
   これで罫線なしの表(位置だけで列が分かれる表)も表として拾える。

### 表の復元(研究の知見をそのまま製品へ)

- 罫線あり: 構造は Tabula、**セルの文字は PdfPig の letter を bbox で詰め直す**。
  Tabula の戻り値は連続する同一文字を潰す(2F-R で実測)ので信用しない。
- 罫線なし: ヘッダー行の文字位置から列の左端を決める自前の再構成。
- Tabula は表の外枠を「1 行 1 セル」の格子としても返す。文字列の長さで
  切るのではなく、**行の高さが通常の 2 倍を超えるものを外枠として落とす**
  (構造で判断する)。
- 複数ページの表は 1 つにつなぐ。2 ページ目以降の先頭行が 1 ページ目の見出しと
  **完全一致**するときだけ見出しの繰り返しとして落とす。少しでも違えばデータ行。
  列数がページで違う場合は、いちばん広い列数に空欄で合わせたうえで Warning。

### NFKC

抽出した文字は必ず NFKC + 前後の空白落としだけを通す。
康熙部首(⽉ U+2F49)など、見た目が同じで別のコードポイントになる字が
混ざるため(2F-R で実測)。それ以外の意味を変える加工はしない
(大文字小文字・かな・表記ゆれの吸収はしない)。テストで固定した。

### 値の型

原則そのままの文字。**書き戻して同じ文字になる場合だけ**数値にする。
先頭 0(0123 / 007)、記号入り(000-1234-5678 / 1,200)、
長い桁は文字のまま。迷うものは文字。

### 部分的な成功を作らない

- スキャン PDF: 「OCR(次の段階で対応)が必要です」と理由を出して Block。
  赤いエラーではなく、対象外の理由として見せる。
- 混在 PDF(10 ページ中 8 ページが文字、2 ページが画像): **全体を Block**。
  どのページが画像かをページ番号で示す。8 ページ分のファイルは作らない。
- パスワード保護・破損・0 ページ・巨大(2000 ページ超)は個別の文言で Block。
  例外の型名やスタックトレースは画面に出さない。暗号化の判定は
  PdfPig の例外型が版で変わるため、型ではなく名前と文言で見る。

### 出力と安全性(既存の決まりを踏襲)

出力は新規 .xlsx(OpenXML で新規作成・数式なし)と .csv(**既存の CsvWriter を再利用**。
2E と同じ文字コード・引用符・CRLF)。
`<元の名前>_PDF抽出.xlsx / .csv`、上書き禁止、実行ごとに一意な作業用ファイル、
CreateNew で所有が確定してから取り消し対象へ登録(D-031)、
書いたものを読み直して行数・列数・各項目を照合してから確定、
失敗時は自分が作ったものだけ取り消して残存を名前で報告(D-024)。
snapshot は D-023 と同じく SHA-256・サイズ・更新日時で、実行直前に再確認。

.xlsx の検証は OpenXmlValidator + 読み直しの全項目照合、
.csv は既存と同じ読み直し照合。

### 控え

`<出力名>.audit.json`(schemaVersion 1 / operation `pdf-extract`)。
source(ファイル名・SHA-256・サイズ・ページ数・判定した種類・抽出方法・normalization)、
output(ファイル名・形式・文字コード・引用符・行数・列数)、warnings。
絶対パス禁止。**PDF の本文そのものは控えへ複製しない**。

### レシピ

2F-A では正式対応しない(OCR・固定帳票まで入ると保存すべき項目が変わるため)。
ただし将来レシピ化しやすいよう、設定は ViewModel 直書きではなく
`PdfReadRequest` という型付きの指定にまとめた。

### 依存(製品へ新規追加)

- PdfPig 0.1.16(Apache-2.0)
- Tabula 1.0.1(MIT)

いずれも 2F-R で調べた版と同じで、商用再配布可。ネイティブ依存も
実行時ダウンロードも無し。publish サイズは 168.6MB → **174.3MB(+5.7MB)**。
製品ランタイムのネットワーク参照は引き続き 0。

### 実測(架空データ、Release)

| 対象 | ページ | 行 | 解析 | 作成 | メモリ |
|---|---|---|---|---|---|
| 文章 | 10 | 400 | 1.2 s | 0.3 s | 18 MB |
| 文章 | 100 | 4,000 | 1.1 s | 0.7 s | 26 MB |
| 表 | 10 | 401 | 0.8 s | 0.1 s | 15 MB |
| 表 | 100 | 4,001 | 1.7 s | 0.3 s | 2 MB |

OCR ではないので待たせない範囲に収まっている。

### 精度

2F-R で 100% だった born-digital の文章・表(罫線あり / なし / 複数ページ)は、
製品実装でも exact 100% を維持していることをテストで固定した。

### 研究コードとの関係

`research/PdfFeasibility` は製品 solution から分離したまま。
製品側は必要なロジックだけを整理して書き直した(参照はしていない)。

## D-034: Phase 2F-B1 — スキャン PDF の OCR + 確認・修正 (2026-08-29)

D-033 以前は書き換えない。

**Phase 2F-B1 COMPLETE ≠ Phase 2F COMPLETE。** 2F 全体は
「文字 PDF / 表 PDF / スキャン PDF / 劣化スキャン / 混在 PDF / OCR 結果の確認・修正 /
スキャン表 / 傾き補正 / 同一レイアウト大量帳票 / checkbox / Excel・CSV 出力 /
完全オフライン配布」まで通ってから。今回はそのうち
**スキャン PDF の OCR と、確認・修正して出力するまで**を製品化した。
傾き補正・スキャン表・大量帳票・checkbox は 2F-B2、最終統合は 2F-B3。
手書きは 2F-R の判断どおり対象外のまま。

### 今回やったこと

2F-A では Block していた「画像だけのスキャン PDF」と「文字 + 画像の混在 PDF」を、
選ぶ → 自動判定 → 必要なページだけ OCR → 結果を確認・修正 → Excel / CSV へ出力、
まで通した。**確認していない読み取りが 1 件でも残っていれば出力できない。**

### 二重読みの構成は実測で決めた(§2 / §3)

まず「1 プロセスに 2 つのモデルを載せられるか」を再実測した。
日本語専用 rec(japan_PP-OCRv4)は旧形式で、**Paddle 3.x runtime では
device を問わず動かない**(Paddle status code 2 を再現)。一方 PP-OCRv5 は
3.x でしか動かない。つまり **v5 と japan は同じプロセスに同居できない**。

そこで 2.x runtime に載る組み合わせ(多言語 ch_PP-OCRv4 + japan_PP-OCRv4)で、
検出を 1 回だけ行い同じ切り出し画像を 2 つの認識モデルへ通し、
統合方式 11 種類 × 閾値 3 種類を Ground Truth と機械照合した
(research/PdfFeasibility/PdfFusionBench)。

帳票 120 ページ / 813 項目の実測:

| 項目 | 多言語のみ | 日本語のみ | 採用した規則 |
|---|---|---|---|
| 店舗コード(英数字) | **98.3%** | 30.8% | **98.3%** |
| 備考(かな漢字) | **0.0%** | 72.0% | **72.0%** |
| 担当者(氏名) | 49.2% | 79.2% | **79.2%** |
| 日付 | 100% | 100% | 100% |
| 全体 | 74.9% | 78.6% | **87.6%** |

**単一エンジンでは成立しない**ことが数字で出た。多言語 rec は
かな漢字を静かに落とし(備考 0%)、日本語 rec は英数字コードを崩す(30.8%)。
統合すると項目ごとに良いほうへ寄る。

### 採用した統合規則: agree-then-charclass-strict @0.98

- 2 つのモデルが**一致**し、自信が 0.98 以上 → 自動確定
- 一致したが自信が足りない → 要確認
- **割れたら必ず要確認**(どちらも自信満々でも自動確定にしない)
- 割れたときに画面へ出す文字は、**空白を除いてすべて ASCII なら多言語 rec、
  日本語が 1 文字でも混ざれば日本語 rec**

最後の閾値は測って決めた。「ASCII が 5 割以上なら多言語」にすると、
`金額:4,917,087円` のように ASCII が多数でも末尾に日本語が付く領域で
多言語 rec を選んでしまい、日本語の文章の完全一致が 100% → 75% に落ちた。

### 最重要指標は「自動確定にしたのに間違っていた件数」

@0.90 / @0.95 / @0.98 を測り、**0.98 だけが誤確定をほぼ 0(813 項目中 1 件)**に
抑えられた。自動確定率を上げるために誤確定を許容しない、という方針どおり。

### 認識器は NaN / Infinity を返す

空に近い切り出しで実測。そのまま閾値と比べると「自信 = 無限大」で
自動確定を通ってしまうため、**有限でない自信は必ず 0 に倒す**。テストで固定した。

### 切り出しは 180 度回ることがある

検出枠を切り出しただけでは上下が逆のまま認識され、電話番号が逆順に読まれた。
**向き判定(cls)を必ず通してから**認識へ渡す。

### 検出枠の score 絞り込みは使わない

Sdcb の `PaddleOcrDetector.GetScore` は、枠が画像の外へわずかにはみ出したときに
native 側で保護されていないメモリを読み、**プロセスごと落ちる**(120 ページの帳票で
再現。落ちるページは実行ごとに変わる)。.NET から捕まえられない種類の落ち方なので、
`BoxScoreThreahold = null` にして呼ばない形にした。

外して困らないことも実測した: 120 ページで枠は 2,169 → 2,191 件(+1%)しか増えず、
増えた枠は認識の自信が低いので要確認へ回る。確からしさの判断は、
認識モデル 2 つの一致と自信で行っている。

### 推論の実行方式も実測で決めた

同じモデル・同じページでの 1 ページあたり: **OpenBLAS 22.8 秒 / oneDNN 23.2 秒 /
ONNX 3.3 秒**。ONNX Runtime(MIT)+ Paddle2ONNX(Apache-2.0)を採用。
**Intel MKL は使わない**(再配布条件に曖昧さが残るため。§22 の指示どおり
少しでも不明なら条件の明確な構成へ)。行列演算の代替として OpenBLAS(BSD-3)を同梱。

### ページ単位の振り分け

document 全体で 1 種類と決めつけず、ページごとに
BornDigitalText / BornDigitalTable / Scan / Unknown へ分ける。
混在 PDF は**文字のページを OCR に通さない**(元から正しい文字を
画像にし直して精度を落とさない)。

ただし**表のページとスキャンのページが混ざる PDF は Block**。
行と列の意味を安全に揃えられないため、無理に 1 つの表へまとめない。

### 傾き・スキャン表は認識の前に止める

認識は 1 ページ数秒かかるので、**先に画像だけを見る安い確認を全ページに通す**
(100dpi で傾きと罫線格子を測る)。傾きが 1.5 度を超えるページ、
縦横の罫線が格子になっているページが見つかったら、**認識を始める前に**
理由を出して止める。無理に文章として出力しない。次の段階(2F-B2)へ送る状態を持つ。

### 確認・修正

- 状態は 自動確定 / 要確認 / 読取不能 / 確認済み
- 「修正して確認」と「元の読み取りのままで正しいと確認」の両方ができる。
  **一覧を開いただけでは確認済みにしない**
- 元の OCR 結果(モデルごとの読みと自信)は残したまま、修正値を別に持つ。
  画面には両モデルの読みを並べて出すので、どちらが違うのかを見て直せる
- 「要確認だけ表示」で絞れる。120 ページを 1 件ずつ見せない
- 未確認が残っている間は Block を付けたままにする(黙って飛ばさない)

### Offline OCR Pack

製品本体は OCR ランタイムを一切参照しない。本体が知っているのは
`ExcelBatchTool.Core` の `IOcrEngine` だけで、実体(`ExcelBatchTool.Ocr`)は
`ocr` フォルダーから実行時に読み込む。おかげで

- **Pack が無くても本体は普通に起動し、文字情報のある PDF はそのまま読める**
- 本体の配布サイズに OCR ランタイムが乗らない(174.3MB のまま)

Pack には目録(`pack.json`: ファイルごとのサイズと SHA-256)を入れ、
読み込む前に検査する。欠損・サイズ違い・中身違い・未知の版は、
**native の読み込みに入る前に**利用者向けの文言で Block する。

モデル・辞書・ランタイムはすべて Pack の中のファイルから読む。
**実行時のダウンロードは無い**(モデルの取得は Pack を組み立てるときだけ)。
モデル置き場を退避して動くことを実測で確認した。

### VC++ ランタイム

PE の import を全ファイル調べた。必要とするのは ONNX 系の DLL
(onnxruntime / onnxruntime_providers_shared / paddle2onnx)だけで、
paddle_inference_c / openblas / OpenCvSharpExtern / pdfium / libSkiaSharp は
ランタイムを静的に取り込んでいる。**製品本体は Windows 標準の Universal CRT だけ**
(VCRUNTIME140 系への依存 0)。

Pack には VCRUNTIME140 / VCRUNTIME140_1 / MSVCP140 を app-local で同梱し、
利用者に別途インストールを求めない。

### 出力は 2F-A のものをそのまま再利用

OCR 専用の書き出しは作らない。`<元の名前>_PDF抽出.*`、上書き禁止、
一意な作業用ファイル(D-031)、読み直しての全項目照合、失敗時は自分が作ったものだけ
取り消し(D-024)、SHA-256 / サイズ / 更新日時での不変確認(D-023)。
文章型スキャンは 2F-A と同じ「ページ / 行 / 内容」へ合流する。

### 控え

OCR を通した場合だけ **schemaVersion 2**(`ocr` の節が増える)。
文字情報だけから作った場合は 1 のまま(「版 1 が意味すること」を後から変えない)。
`ocr` にはモデル名・実行方式・閾値と、自動確定 / 要確認 / 読取不能 / 確認済み /
人が修正した件数だけを書く。**読み取った文字そのものは控えへ複製しない。**

### 中止

OCR は数分かかるのでキャンセルを必須にした。中止すると元の PDF は変わらず、
出力も控えも作らず、作業用ファイルも残らない。

### 実測(架空データ、Release、12 コア)

| 対象 | ページ | 項目 | 完全一致 | 自動確定 | **誤確定** | 要確認が捕捉 | s/page |
|---|---|---|---|---|---|---|---|
| 文章 clean | 5 | 20 | **100.0%** | 0.0% | **0.00%** | 100.0% | 4.72 |
| 帳票 clean | 120 | 813 | **87.9%** | 24.4% | **0.12%**(1 件) | 99.0% | 2.20 |
| 文章 degraded ※ | 5 | 20 | 10.0% | 0.0% | 0.00% | 100.0% | 3.72 |
| 帳票 degraded ※ | 30 | 203 | 47.8% | 3.9% | 0.00% | 100.0% | 3.39 |

※ 製品では傾きの判定で止まる。次の段階が相手にする精度を知るための測定。

混在 PDF(文字 50 ページ + 画像 50 ページ): **7.11 秒/ページ**(OCR にかけた 50 ページ換算)、
メモリ増分 35MB、出力 2,000 行。文字のページは OCR を通さず PdfPig で読んでいる。
1 ページあたりの行が長い(35 文字超 × 20 行)ため帳票より遅い。

読み取り項目の内訳(帳票 clean 120 ページ): 2,191 件 =
自動確定 302 / 要確認 1,720 / 読取不能 169。ピークメモリ増分 37MB。
2F-R の PP-OCRv5 単体 2.7 秒/ページに対し、**2 モデルを通して 2.20 秒/ページ**。

**誤確定は 813 項目中 1 件(0.12%)、劣化版では 0 件。**
間違いのうち 99〜100% は要確認へ回っており、人が見れば気づける状態になっている。
自動確定率は 24.4% と低いが、これは方針どおり
(自動確定率を上げるために誤確定を許容しない)。

### 2F-R(v5・3.x)との差について

2F-R の帳票 clean は PP-OCRv5 単体で 98.6% だった。今回は 87.9% に下がっている。
**v5 と日本語専用 rec は同じプロセスに載せられない**ため、二重読みを成立させるには
2.x に載る多言語 rec(v4)を使うしかない、というのがこの差の理由。

ただし目的は完全一致率の最大化ではなく**誤確定を 0 に寄せること**で、そちらは
v5 単体(@0.98 で誤確定 0・自動確定 82%)と同じく達成できている。
v5 を使う道(別プロセス化など)は 2F-B3 の統合時に再評価する。

## D-035: Phase 2F-B1.1 — OCR 確認の安全性補正 (2026-08-29)

D-034 以前は書き換えない。

**Phase 2F-B1.1 COMPLETE 後も Phase 2F は進行中。** 傾き補正・スキャン表・
大量定型帳票・checkbox は 2F-B2、最終統合は 2F-B3。

### なぜ必要だったか

2F-B1 の確認画面には OCR 結果・自信・状態・修正欄・2 モデルの読みはあったが、
**元の PDF のページ画像が無かった**。読み取りが正しいかは元の文書と見比べないと
決められないので、これが無いと「人が確認する」が成立していない。
2 つのモデルが**両方とも間違う**こともあるので、なおさら原文が要る。

もう 1 つ、「表示中をすべて確認済みにする」があった。原文と見比べずに
まとめて確認済みにできてしまい、安全設計を迂回できる状態だった。

### 元のページ画像

- 選んだ項目のページを 150dpi で描いて出す(OCR は 300dpi、画面はもっと粗くてよい)
- 読み取り位置を赤い枠で示す。座標は「OCR の画素 → 画像の画素 → 画面の画素」を
  掛けて出す純粋な計算にしてあり、画面なしでテストできる
- ページ送り / 拡大 / 縮小 / 100% / 全体表示
- **項目を選ぶとそのページへ自動で移動し、枠のところへ寄る。**
  ページ全体を表示した倍率(A4 で約 24%)だと 1 行は 10 画素ほどにしかならず
  原文と見比べられないので、**枠が 24 画素より小さくなるときだけ**
  読める大きさ(枠の高さ 44 画素)まで自動で拡大する。
  利用者が自分で拡大しているときはその倍率を尊重して触らない
- 同じページの中で選び直したときは、画像を描き直さず枠だけ動かす

### 画像の持ち方

120 ページを 300dpi のまま抱えない。**今のページと前後 1 ページ = 3 枚**だけを
手元に置き、古いものから捨てる。実測(120 ページの帳票を 400 回行き来):
管理メモリ +0.6MB、ワーキングセットは 100 回目以降ほぼ一定(195MB 前後)、
手元は常に 3 枚、1 回の切り替え 194ms。**増え続けない。**

一時ファイルは作らない。ページ画像はメモリの中だけで、確認用に PNG を
ディスクへ残すことはしない(テストで固定)。

### 「まとめて確認済みにする」を削除

指示どおり案 A(削除)を採った。代わりに飛ばせない形で速くする:

- 次の要確認へ / 前の要確認へ
- **修正して確認 → 次へ** / **元のままで確認 → 次へ**
- Enter(修正して確認して次へ)/ Ctrl+Enter(元のまま確認して次へ)/
  Esc(修正を取り消す)/ F3(次の要確認)。押せるキーは画面に出す

自動で進むのは**次の未確認**だけで、未確認を飛び越えない。
「まとめて確認」に相当するメソッドが増えていないことをテストで固定した。

### 一覧の絞り込みを直した

2F-B1 は「未解決だけ」を出していたので、確認した行がその場で消えていた。
何を確認したのか見えず、取り消すこともできない。
**最初の分類で絞る**ように変えたので、確認済みにしても行は残る。
「自動確定も表示」で自動確定した内容も原文と見比べられる(§10)。

### 読取不能

読取不能も同じように原文とその位置を見られる。人が読めるなら
そのまま入力して確認済みにできる。読めなければ未解決のまま残り、出力は Block。

### 誤自動確定の再測定 — 結論: 0 件

2F-B1 は 813 項目中 1 件(0.12%)と報告していた。**これは測り方の誤りだった。**

その 1 件を具体的に見ると:

```
p81 売上 GT="2,636,155" / 自動確定した領域="2026/02/10" 自信 0.9998
```

自動確定された領域は**日付の項目を正しく読んだもの**で、売上の値はどの領域も
読んでいなかった(検出漏れ)。それを「いちばん近い領域」として売上に結び付けたため、
検出漏れが誤確定として数えられていた。

**「その項目を読もうとした領域」だけを対応付ける**ように直した(文字正解率 0.5 未満の
領域は、その項目について何も主張していない = 検出漏れとして別に数える)。
結果:

| 閾値 | 完全一致 | 自動確定 | **誤確定** | 要確認が捕捉 |
|---|---|---|---|---|
| 0.98 | 87.6% | 24% | **0 件** | 100% |
| 0.985 | 87.6% | 23% | **0 件** | 100% |
| 0.99 | 87.6% | 22% | **0 件** | 100% |
| 0.995 | 87.6% | 20% | **0 件** | 100% |
| 0.999 | 87.6% | 14% | **0 件** | 100% |

劣化版・文章版も全閾値で 0 件。対応付けの下限を 0.2〜0.7 で振っても
どこでも 0 件で、都合のよい値を選んだわけではないことも確かめた。

**閾値は 0.98 のまま。** 上げても安全は増えず(どれも 0)、自動確定率だけが
24% → 14% に落ちるため。追加の gate も入れない
(効果が 0 の安全策のために自動確定率を下げるのは筋が悪い)。

誤確定を 0 にしているのは閾値ではなく**「2 つのモデルが一致し、かつ自信が足りている
ときだけ自動確定する」という条件**そのもの。これを不変条件としてテストで固定した
(文字列 7 種 × 自信 8 段階の総当たりで、自動確定になったものは必ず
「両モデルが同じ・空でない・min 自信 ≥ 0.98」であることを確認)。

検出漏れは別の話として残る(帳票 clean で 813 項目中 89 件)。これは OCR の
拾い漏れで、確認の安全とは別問題。2F-B2 以降の課題。

### Pack の版違いを利用者向けの文言にした

作業中、アプリだけを更新して Pack を古いままにしたところ
「Method 'RenderPage' ... does not have an implementation」という .NET の
内部的な文言がそのまま画面に出た。目録の検査は Pack の中身しか見ないので
ここは通ってしまう。型や実装が噛み合わない失敗をまとめて捕まえ、
**「Pack の版がアプリと合っていません。同じ配布物のものへそろえてください」**
と言い換えるようにした。

### 画面と ViewModel の結び付け方を直した

スクロール要求を MainWindow の生成時に購読していたため、
DataContext を差し替えると購読が外れた状態になっていた
(UI ハーネスで画像が動かないことから判明)。**DataContext に合わせて
購読し直す**形へ変更。差し替えても追従する。

### 実測(架空データ、Release)

- 長時間の確認: 120 ページを 400 回行き来して 管理メモリ +0.6MB /
  ワーキングセット +8.7MB(100 回目以降ほぼ一定)/ 手元 3 枚 / 194ms per 切替
- OCR 自体の速度は変えていない(画像を出すのは確認のときだけ)

## D-036: Phase 2F-B2 — 傾き補正 / スキャン表 / 定型帳票 / 印 (2026-08-29)

D-035 以前は書き換えない。

**Phase 2F-B2 COMPLETE ≠ Phase 2F COMPLETE。** このあと必ず
Phase 2F-B3(全方式統合 / Pack 最終配布 / 実案件相当の総合 fixture /
Phase 2F 最終 close 判定)を行う。手書きは 2F-R の判断どおり対象外のまま。

### いちばん大事だった課題: 項目が消えないこと

B1.1 で分かった「帳票 813 項目中 89 件が、どの領域にも読まれず結果から消える」を
無くした。**指定した項目の数と、確認に出る件数を必ず一致させる。**

読み取り領域が 1 つも見つからなかった項目は、消さずに
**状態「見つからない」**として 1 件残す。位置は指定した領域を使うので、
確認画面で元のページのその場所を見て手で入力できる。
未確認のまま残れば出力は Block。

型の上でも守れるようにした。`FormFieldExtractor.Read` は
「指定した項目 → 読み取り結果」を 1 対 1 で返す(見つからなくても返す)ので、
途中で件数が減る余地が無い。

### 領域指定は「切り出して読む」ではない

2F-R で、領域を切り出して読む方式は全ページ OCR より悪かった
(clean 98.6% → 87.3%、劣化 87.7% → 54.2%)。この知見を無視しない。

**全ページを OCR し、その結果のうちどれがこの項目に属するかを領域で選ぶ。**
領域は「そこだけ読む」ためではなく「読んだものを割り当てる」ために使う。

### 位置合わせ

項目名そのもの(「店舗コード」など)を手がかりにして、ページごとのずれを
平行移動として求める。手がかりが指定の位置から離れすぎている場合は使わない
(同じ文字が別の場所にあるときに大きく動かして壊さない)。
複数の手がかりの中央値を取るので、1 つ外れても効かない。

### 傾き補正

- 0.3 度未満は直さない(わずかな傾きのために回すと補間で文字がぼやける)
- 6 度を超えるものは直さず、取り込み直しを促して Block
- 角度を測れなかったページは**回さずそのまま読む。止めない**
  (測れないのは行らしい塊が少ないページで、多くは実際には傾いていない。
  測れないというだけで止めると、そういうページが軒並み扱えなくなる)
- 読み取りは直した画像で行い、**確認に使う位置は必ず元のページの座標へ戻す**。
  戻す変換(`DeskewTransform`)を明示的に持ち、往復して元に戻ることをテストで固定した

### 実装中に見つけた、想定と違った点

1. **傾きの推定が罫線に引きずられる。** 罫線表では、ページを横切る罫線が
   文字の行より強い塊になり、傾き 0 度の表が「6 度超」と判定されて止まった。
   **罫線を先に取り除いてから**文字の行だけで角度を測るようにした。
2. **「角度を測れない = 止める」は厳しすぎた。** 上の修正前は、罫線表が軒並み
   Block されていた。測れないページは回さずに読む、へ変更。
3. **細い外枠の罫線は消える。** 線が画素の境目に来ると濃さが半分ずつに割れ、
   二値化で落ちる。実測では 4 列の表が 2 列になり、いちばん左と右の列が
   **まるごと落ちた**。罫線の外側に文字があれば、そこにも区切りがあったものとして
   足すようにした(判断には文字の**中心**を使う。端で判断すると、枠のすぐ内側にある
   見出しで区切りが増えて行が 1 つずれた)。
4. **向きの判定を弱めると読み取りが落ちる。** 短い数字の切り出しが上下逆に
   読まれている(「1,037」→「E0'」)と考えて閾値を上げたところ、
   罫線表 61.3% → 38.7%、罫線なし 87.9% → 64.3% と悪化した。
   検出した枠は上下逆で返ることが多く、この判定がそれを直している。既定へ戻した。

### 傾いた罫線表は止める

まっすぐな罫線表はセル一致 61.3% なのに、2 度傾いた表は 4.8% まで落ちた
(誤確定も 8 件出た)。罫線の位置がわずかな残り傾き(推定誤差 平均 0.33 度)で
ずれ、行と列が崩れるため。**崩れた表をそれらしく出すより理由を示して止める。**
2F-B3 で扱う。

### 印(チェック / 塗りつぶし / 丸 / ばつ)

文字として読ませず、箱の中とラベルの周り(丸囲みの線が通るところ)の
黒画素の割合で判定する。1 位と 2 位の差が小さいとき、薄いときは
**自動で決めず人へ回す**。自動確定にするのは差がはっきりついたときだけ。

### 一致していても自動確定にしない場合(誤確定 0 のために足した)

二重読みの「両方が一致した」という根拠は、**2 つのモデルが同じ間違いをしない**
ことを前提にしている。ところが採用した 2 つはどちらも PP-OCR v4 系で、
字形の取り違えは共通して起きる。実測した誤確定 6 件は、すべてこの形だった。

| 出どころ | 読み | 正 | 自信 |
| --- | --- | --- | --- |
| 帳票 p1 / p5 / 傾き p1 / 傾き p2 店舗コード | SO01-24 など | S001-24 など | 98.4〜98.8% |
| 罫線あり p5 r16c0 | 9600 | A0096 | 99.8% |
| 罫線なし p4 r9c0 | 6900 | A0069 | 99.7% |

前者は 0 と O の取り違え、後者は切り出しが上下逆のまま両モデルへ入ったもの。
**自信でも一致でも捕まえられない**(閾値を上げても消えず、正しい読みが人へ回るだけ)。
そこで「形」で判断する 2 つを足した。

1. **項目の種類で見る**(`FieldAutoAcceptPolicy`)。コードの項目に
   取り違えやすい字(0/O・1/l/I・5/S・8/B・2/Z・6/G・9/q)があれば自動確定しない。
   数量・金額の項目に数字と区切り以外があれば自動確定しない
2. **表の列で見る**(`ColumnShapeGuard`)。同じ列の他の行と文字の種類
   (数字 / 英字 / 日本語)が食い違うセルは自動確定しない。
   記号は数えない(桁区切りを数えると金額の「999」と「1,234」が別扱いになり、
   正しいセルまで人へ回してしまう。実際にこれを踏んでから直した)。
   いちばん多い形が 4 割に満たない列は判断しない

どちらも**自動確定を見送って人へ回すだけ**で、読み取った内容は書き換えない。
外したときの損は確認の手間が増えることだけなので、厳しすぎるほうが害が大きい。

実測の結果、**帳票 4 種・表 3 種・印 3 種のすべてで誤確定 0**。
代償は帳票の自動確定 46% → 32% など。**自動確定の割合より誤確定 0 を採る**
(指示の「自動確定率を上げるために誤確定を許容しない」に従う)。

なぜ確認が要るのかは、確認の画面で理由として出す。
理由が分からないまま「元のままで確認」を押されると確認の意味が無くなるため。

### 出力

- 表: 行と列のまま。2 ページ目以降の見出しの繰り返しは落とす
- 帳票: **1 ページ 1 件**。列は指定した項目の順で、読めなかった項目も列として残る
- 印: 選んだ選択肢の文字を列へ

既存の書き出し・上書き禁止・作業用ファイル・取り消し・読み直し照合・snapshot は
そのまま再利用。確認の画面も B1.1 のものをそのまま使う(別の危ない確認画面を作らない)。

### 控え

`ocr` の節へ mode / 傾きを直したページ数 / 表として読んだページ数 /
帳票として読んだページ数 / 指定した項目の総数 / 見つからなかった件数 を追加。
**読み取った文字そのものは引き続き複製しない。**

## D-037: Phase 2F-B3 — PDF 対応の最終統合と close 判定 (2026-08-29)

Phase 2F-A / 2F-B1 / 2F-B1.1 / 2F-B2 で作ったものを 1 つの機能として束ね、
実案件相当の総合試験まで通した。D-036 以前は変更しない。

### いちばん大きかった発見: 傾きを直す向きが逆だった

2F-B2 から、**傾きを直すつもりで倍にしていた。**

罫線表の傾き 2 度のページで、傾きを 1.98 度と正しく測れているのに、
直したあとの画像から罫線が 1 本も拾えなかった。試しに逆向きへ回すと
22 行 3 列の格子がそのまま出てきた。

```
rotate  -1.98: rows=0  cols=0   ← 製品がやっていた向き
rotate  +1.98: rows=22 cols=3   ← 正しい向き
```

2F-B2 で「傾いた罫線表は実用にならない(4.8%)」「帳票 傾き 50.8%」
「印 傾き 53.3%」と報告した数値は、すべてこれが原因だった。
向きを直した結果:

| 対象 | 直す前 | 直した後 |
| --- | --- | --- |
| 罫線表 傾き 2° | 17.5%(誤確定 101) | **88.9%**(誤確定 0) |
| 帳票 傾き ±2° | 74.2% | **94.2%** |

**測った角度が合っていても、直した結果を確かめていなければ意味がない。**
角度の推定誤差(平均 0.33 度)だけを見て良しとしていたのが抜けだった。

### OCR 構成: PP-OCRv5 は採らない(実測で決めた)

2F-B1 では「japan_PP-OCRv4 は旧形式なので Paddle 3.x では動かない」と実測し、
2.x に留めていた。**3.3.1 で測り直すと、PP-OCRv5 と japan_PP-OCRv4 が
同じプロセスで両方読み込めた**(3.0.x では再現する。3.3.1 では起きない)。
別プロセスも IPC も要らない。

そのうえで 3 つの構成を、同じ製品経路・同じ fixture で比べた。

| | v4 二重読み | v5 + japan v4 | 差 |
| --- | --- | --- | --- |
| 帳票 120p | 99.0% | 99.0% | なし |
| 帳票 ずれ | 94.2% | 94.2% | なし |
| 帳票 傾き | 74.2% | 74.2% | なし |
| 帳票 拡大 | 99.2% | 100.0% | +1 項目 |
| 1 ページ | **1.6〜1.9 秒** | 3.7〜4.6 秒 | **2.4〜2.9 倍** |
| Pack | **249.1MB** | 254.8MB | +5.7MB |

**採らない。** 完全一致は 480 項目中 1 件しか変わらず、時間が 2.4 倍以上になる。
v5 は Paddle2ONNX が変換できず native 実行へ落ちるのが原因
(`Paddle2ONNX do't support convert the Model, fall back to using Paddle Inference`)。

2F-R が v5 に期待していた差(98.6% 対 87.9%)は、**runtime を 2.6.1 から 3.3.1 へ
上げただけで埋まった**。同じ v4 モデルのまま帳票 80.2% → 99.0%、検出漏れ 4 件 → 0 件。

### 検出漏れ(2F-B1.1 の 89/813)

| 時点 | 検出漏れ |
| --- | --- |
| 2F-B1.1 | 89 / 813 |
| 2F-B2(Missing として表面化) | 0 件消失(状態として残す) |
| 2F-B3(runtime 3.3.1) | **0 / 480**(そもそも起きない) |

消えない仕組みは残したまま、起きる回数そのものが 0 になった。

### 誤確定 0 のために足したもの(2F-B2 から継続)

runtime を上げて読み取りが良くなると、**新しい種類の誤確定**が出た。

1. **罫線が落ちて 2 行が 1 区画に入る。**「A0017」と「A0018」が
   「A0017A0018」という 1 セルになり、自信 99.6% で自動確定していた。
   → 区画の高さが他より明らかに高く、中の文字が縦に離れた塊へ分かれるときだけ
   区切りを足す(`SplitTallBands`)。折り返しは高さが増えないので割らない。
   解消しなかった連結は自動確定しない(`IsMerged`)。
   劣化した罫線表が 50.0% → **86.1%** になった
2. **上下逆でも数として読める値。**「90」を「06」と読み、自信 99.1〜99.8% で
   一致していた。0・1・8 は上下逆でも同じ形、6 と 9 は入れ替わる。
   その字だけでできている値は自動確定しない。
   自動確定の割合は 1〜3 ポイントしか下がらなかった

### コードの項目の自動確定を戻した

2F-B2 の「取り違えやすい字が 1 つでもあれば自動確定しない」は安全だが粗く、
数字を含むコードがほぼ全部人へ回っていた(帳票の自動確定 46% → 32%)。

**項目 × 全ページで形を学ぶ**ようにした(`FieldShapePattern`)。
店舗コードが全ページ `A999-99` の形なら、その形どおりに読めたページは
自動確定してよい。形が違うページ(`SO01-24` は英字が 1 つ増える)だけ人へ回す。

**値は決して書き換えない。** 形から推測して O を 0 に直すことはしない。
判断にだけ使う。学べるのは 8 ページ以上あって 6 割以上が同じ形のときだけで、
学べない項目は 2F-B2 の粗い決まりへ落とす。

帳票の自動確定が 32% → **75%** へ戻り、誤確定は 0 のまま。

### 表の空欄を黙って空欄にしない

罫線から作った格子のうち、文字が 1 つも割り当たらなかった区画は、
これまで確認にも出力にも現れなかった。**元の表に文字があるのに検出されなかった
セルが、黙って空欄として出る**状態だった。

その区画に黒い画素があるかを測って分ける。文字があるのに読めなかった区画だけ
「読取不能」として人へ回し、もともと空の区画はそのまま空欄として出す。
空欄の多い表で確認の手間が現実的でなくなるのを避けつつ、取りこぼしを無くす。

### 傾いた表は止めずに通す。ただし壊れていたら止める

2F-B2 は傾いた表を一律で止めていた。向きの誤りを直した結果 88.9% まで上がったので、
**傾きを直してから行と列へ戻すところまで通す**。
戻せたかどうかは出来上がった表の形で判断し、中身のあるセルが半分に満たなければ
「表の行と列を戻せないページ」として理由を示して止める(崩れた表は出さない)。

### 大きすぎる PDF

- OCR に回せるのは **1,000 ページまで**(文字 PDF の上限 2,000 とは別)。
  1 ページ 1.6〜3.8 秒かかるので、無制限に走らせない
- 200 ページを超えるときは、かかる時間の見込みを先に出す(止めはしない)
- 1 ページの画素数の上限も設ける(A4 300dpi の約 9 倍)。
  図面のような巨大なページでメモリーを使い切って落ちるより、理由を出して止める

### PDF の処理設定(レシピ)

タブ 8 も既存の RecipeStore へ載せた。保存するのは
**読み取り方 / 出力形式 / CSV の設定 / 帳票の項目(名前・種類・読む場所)** だけ。

**元の PDF に関わるものは保存しない** — ファイル名も、保存場所も、
読み取った文字も、ページの中身も、個人情報の実値も入れない。
入れる場所そのものを型に作らず、保存したファイルの中身にも出てこないことを
テストで固定した。

読み込んでも自動では実行しない。読む場所は前の PDF に合わせたものなので、
「今回の PDF で 1 度読み取って位置を確かめてください」と伝える。

## D-038: D-001 を更新 — spec リポジトリを廃止し、本リポジトリへ統合 (2026-09-19)

- D-001 で決めた「Private spec / Public implementation の 2 リポジトリ構成」をやめる。
- 仕様・ロードマップ・設計判断・調査結果・実装記録は、すべて本リポジトリの
  `docs/` に置く。本リポジトリが唯一の正本。
- 移してきたのは文書の内容だけで、spec 側の Git 履歴は引き継がない。
- 移さなかったもの: spec 側の README(2 リポジトリ構成の説明)、作業手順書、
  作業セッションログ、内部 TODO の作業待ち行列、対応 commit SHA の一覧。
  いずれも 2 リポジトリ運用のための記録で、統合後は意味を持たない。
- D-006(git author は GitHub noreply アドレス)は引き続き有効。
