# Issue #4 AI向け課題仕様監査

監査日: 2026-05-23

## 対象範囲

Issue #4 では、全AI条件へ渡す課題仕様と共通promptを固定しました。AI条件の実行、VBA生成、サンプルworkbook編集、`.xlsm` 作成は行っていません。

## 固定した判断

- prompt言語: 日本語。
- 実験上のやり取り上限: 仕様確認1回、修正依頼1回。
- 期待するコード: 標準モジュール用VBA。読みやすい英語識別子を推奨。
- 説明言語: 日本語。
- 探索起点: `ThisWorkbook.Path\daily\YYMM\`。
- 対象月の判定元: 月次workbook名 `月次クレーム集計YYMM.xlsm`。
- 日次日付の判定元: 日次workbook名 `クレーム集計YYMMDD.xlsx`。
- 転記キー: 支店コード + 業務コード。
- 行種別: `DETAIL` は転記対象、`SUBTOTAL` は入力キーとして扱わない。
- 再実行: 先に日別転記セルをクリアし、結果を再構築する。
- 参照設定: AIが判断してよいが、追加参照を使う場合は理由と代替案を説明する。

## 更新成果物

- `docs/task-spec.md`: AI向けの詳細課題仕様。
- `prompts/common-task-prompt.md`: 全AI条件へ渡す共通prompt。
- `prompts/clarification-answer-bank.md`: 1回だけ許可する確認質問への標準回答。

## 引き渡し

- Issue #5 では `prompts/common-task-prompt.md` を全AI条件の初回promptとして使う。
- AIが確認質問をした場合は、`prompts/clarification-answer-bank.md` から回答する。
- promptと回答bankの外から追加ヒントを出さない。
- 評価前に生成物を人手修正しない。

## 後続作業

- 実際のAI条件実行は Issue #5 に延期。
- Windows版Excelでの検証は Issue #6 に延期。
- マクロ付き `.xlsm` 成果物は、VBA実装を選定または検証する段階まで延期。
