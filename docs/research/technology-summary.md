# 技術調査サマリー — Excel 処理エンジンの選定

作成: 2026-08-27
結論: **Microsoft Open XML SDK(DocumentFormat.OpenXml)を中心にする**(DECISIONS D-002)

## 比較

### Microsoft Open XML SDK(採用・中心)

- Microsoft 公式。MIT License。.NET 8 対応。
- .xlsx(OPC パッケージ)をパート単位でそのまま扱う低レベル API。
- 読み取り専用で開ける。触っていないパート(Drawing / Chart / 拡張要素)は
  バイト列として保持され、破壊しない。
- ストリーミング読み(OpenXmlReader)で大容量にも対応可能。
- 欠点: 高レベル API がなく実装コストが高い(セル値の型解決、共有文字列、
  A1 参照計算などを自前実装)。数式再計算エンジンは無い。
- 判断: 本製品の中心価値は「壊さない」こと。パート構造をそのまま
  保持できることが最優先のため、実装コストを払っても採用する。

### ClosedXML(不採用: 主エンジンにしない)

- MIT License。高レベル API で生産性が高い。
- ただし Workbook 全体を自前オブジェクトモデルに読み込み、保存時に
  再構築する設計。未対応・部分対応の要素(Drawing、Chart、一部拡張、
  条件付き書式の一部など)が保存で欠落・変形するリスクがある。
- 既存の実業務 Workbook を開いて保存する用途では「開いて保存しただけで
  壊れる」可能性を排除できないため、主エンジンにしない。
- 将来、新規ファイル生成(統合結果の新規 Workbook 出力)に限定して
  補助利用する余地はある(採用時は DECISIONS に追記)。

### openpyxl(不採用)

- Python 製。本製品は C#/.NET のため言語が不一致。
- Python ランタイム同梱は self-contained Windows 配布を複雑にする。
- openpyxl 自体も Chart / 図形の保存で欠落があることが知られており、
  「壊さない」要件に対する優位性もない。

### Excel アドイン型製品(不採用)

- Excel 本体が必要になり、「Excel 本体不要」の中心価値と矛盾。
- 大量 Workbook を開かずに処理する、という価値も実現できない。

### Excel MCP 系 / AI 連携ツール(不採用)

- AI・外部サービス前提の構成は「完全ローカル・外部送信なし・AI 内蔵対象外」
  という v1 方針と矛盾。
- 非技術者ユーザーに MCP ホスト環境を要求できない。

## Open XML SDK 採用時の設計上の注意(Phase 0 で実装済みの範囲)

- セル値は CellValue + DataType + SharedStringTablePart を解決して読む。
- 使用範囲は SheetDimension を第一情報源にする(D-008)。
- 読み取りは FileAccess.Read の FileStream 経由で物理的に書込不能にする(D-003)。
- 書換系(Phase 1 以降)では、変更しないパートには触れない実装を徹底する。
