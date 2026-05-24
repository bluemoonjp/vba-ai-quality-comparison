# vba-ai-quality-comparison エージェントルール

## 適用範囲

このルールは、この独立リポジトリ全体に適用します。親 `_Workspace` のルールも、矛盾しない範囲で適用します。

## 公開データ境界

- 架空で公開可能な合成データだけを使う。
- 実在の顧客、社員、支店、業務、クレーム内容、実パス、workbookメタデータ、スクリーンショット、secret、private source materialを追加しない。
- `.xlsm` は、メタデータと隠し内容を確認するまで高リスク成果物として扱う。
- workbook、prompt log、生成物、スクリーンショット、exported VBAを追加する前に `docs/public-boundary.md` を確認する。

## 実験の一貫性

- 比較対象は、AIの使い方による成果物差であり、人手による隠れた改善ではない。
- 全条件で同じ課題仕様、サンプルデータ、期待結果、評価基準を使う。
- 各条件で確認質問の回数は制限しない。修正依頼は1回までにする。
- prompt、確認質問、回答、修正依頼、モデル/モード/日付、出力、検証メモを記録する。
- 先行条件の生成物やレビュー結果を後続条件へ見せない。見えてしまった場合は、実験汚染として記録する。

## エージェント役割

- 計画と進行管理の主担当は Orchestrator とする。
- Orchestrator は実ファイル作成、VBA生成、評価本文作成を直接担当しない。依頼、受領、検収、レビュー依頼、Issue整理を担当する。
- 実作業は Worker / Explorer が担当し、重要成果物は Reviewer high の確認を受ける。
- subagent 出力は提案として扱い、Orchestrator が公開境界と実験統制を確認して統合判断を行う。

## VBAとExcelの安全性

- レビュー対象のVBAは `.bas` / `.cls` などのテキスト形式で保存する。
- モデル成果物として `.xlsm` を保存しない。マクロ付きブックはローカル検証用scratchとして扱う。
- サンプル workbook は公開境界チェックを通過したものだけ `experiments/**/samples/workbooks/` に保存してよい。
- 明確なタスクと対象ファイルがない状態で、マクロ実行や既存workbookへの書き込みをしない。
- VBAの動作検証はWindows版Excelを前提にする。
- 静的レビューだけでは最終検証完了としない。

## GitとGitHub

- このディレクトリは独立Gitリポジトリ。親 `_Workspace` にこのリポジトリのファイルをcommitしない。
- commitは、意味のあるIssueまたはフェーズ単位に分ける。
- GitHub Issuesをこのリポジトリの作業トラッカーとして使う。
- scaffoldや公開安全点検は、該当Issueのコメントまたは `experiments/<id>/verification/` 配下の監査記録へ残す。
