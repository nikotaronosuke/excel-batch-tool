# Excel Batch Tool

> **Excel Batch Tool は開発コードネームです。正式な製品名は未定です。**

Windows 上で、複数の Excel Workbook を Microsoft Excel を開くことなく、
**完全ローカル**で安全に一括処理することを目指しているデスクトップアプリです。

**現在は初期開発段階です。** 完成品・安定版ではありません。

## 現在できること(Phase 0: 読み取り専用の Workbook 解析)

現在の Phase では、Excel ファイルを**書き換える前に安全に把握する**ための
読み取り専用解析のみを提供します。**対象ファイルは一切変更しません。**

- `.xlsx` ファイルの Drag & Drop / ファイル選択(複数同時)
- Workbook ごとの情報表示
  - シート数・シート名・使用範囲・概算行数/列数・ファイルサイズ
- 書き換え時に注意が必要な要素の検出
  - 数式 / 結合セル / 図形 / グラフ / 画像 / ピボットテーブル / 外部参照 /
    シート・ブック保護 / テーブル / データ入力規則 / 条件付き書式 /
    コメント / 定義名 / ハイパーリンク / OLE・ActiveX / マクロ関連 など
- 解析結果の分類表示
  - ✅ 通常 / ⚠ 注意が必要 / ✖ 現在非対応

## 方針

- **Windows 専用**(.NET 8 / WPF)
- **Microsoft Excel 本体不要**(Office Interop 不使用。
  [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) で直接読み取り)
- **完全ローカル**(クラウドサービス・外部 API・ログイン不使用。
  アプリ自身はネットワーク通信機能を持ちません)
- 対象は `.xlsx`(`.xls` は対象外)

## ビルド

.NET 8 SDK が必要です。

```
dotnet build
dotnet test
dotnet run --project src/ExcelBatchTool.App
```

## 構成

```
src/ExcelBatchTool.Core   … 解析エンジン(読み取り専用)
src/ExcelBatchTool.App    … WPF デスクトップアプリ
tests/                    … テスト(架空データで生成した Workbook を使用)
```

テストでは、解析前後のファイルの SHA-256 ハッシュ・サイズ・最終更新日時を
比較し、**解析が元ファイルを一切変更しないこと**を自動検証しています。

## License

MIT License([LICENSE](LICENSE) を参照)

依存パッケージ:

- [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml)(MIT)
