# ドキュメント

このリポジトリの仕様・設計・調査・実装記録をまとめた場所です。
アプリの使い方と機能一覧は、リポジトリ直下の [README.md](../README.md) にあります。

| 文書 | 内容 |
| --- | --- |
| [spec/v1-overview.md](spec/v1-overview.md) | v1 の目的・対象ユーザー・中心価値・対象外・安全性方針 |
| [implementation/roadmap.md](implementation/roadmap.md) | Phase 0〜5 の計画と、どこまで完了したか |
| [implementation/phase-log.md](implementation/phase-log.md) | Phase ごとの実装内容・実測値・検証結果 |
| [implementation/open-issues.md](implementation/open-issues.md) | 未対応と分かっている課題・将来対応 |
| [decisions.md](decisions.md) | 設計判断の記録(D-001〜、追記のみ) |
| [research/market-summary.md](research/market-summary.md) | 企画背景にした市場調査の要約 |
| [research/technology-summary.md](research/technology-summary.md) | Excel 処理エンジンの比較と選定理由 |
| [research/pdf-feasibility-research.md](research/pdf-feasibility-research.md) | PDF 対応の実現性ベンチマークと GO / NO-GO 判定 |

ベンチマークを実行するコードそのものは
[../research/PdfFeasibility/](../research/PdfFeasibility/) にあります。

## この文書群の出自

これらの文書は、もともと実装とは別の非公開リポジトリで管理していたものです。
2026-09-19 に、実装とドキュメントを 1 つのリポジトリにまとめる方針へ変更し
(decisions.md の D-038)、文書の内容だけをここへ移しました。

そのため 2026-08 以前の記録には、当時の 2 リポジトリ運用を前提にした
「Public 側 / Private 側」という言い回しが残っています。現在はこのリポジトリが
唯一の正本です。
