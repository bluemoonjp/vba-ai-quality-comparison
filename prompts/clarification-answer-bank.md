# Clarification Answer Bank

Use these standard answers when an AI condition asks its one allowed clarification question. Keep the answer concise and do not add extra hints beyond the relevant item.

## Folder Location

日次ファイル探索起点は `ThisWorkbook.Path\daily\YYMM\` です。`YYMM` は月次ブック名 `月次クレーム集計YYMM.xlsm` から判定します。配下フォルダは再帰探索してください。

## Target Month

対象月は月次ブック名から判定します。サンプル対象は `2602`, `2604`, `2605` です。日次ファイル名 `クレーム集計YYMMDD.xlsx` から日付を判定してください。

## Matching Key

転記キーは `支店コード + 業務コード` です。支店名と業務名は表示・確認用で、主キーではありません。

## Rerun Behavior

再実行時は `DETAIL` 行の日別転記範囲 `H:AL` をクリアしてから、日次ファイルを読み直して再集計してください。二重計上は不可です。

## Subtotal Behavior

`SUBTOTAL` 行は入力キーではありません。小計は同じ `小計グループ` の `DETAIL` 行を集計してください。大支店は支店単独、小支店は地域単位の小計です。

## Error Handling

未知の支店、未知の業務、空欄件数、非数値件数、重複ファイル、対象外ファイル、`日次集計` シートなしは転記せず、ログまたは処理結果として分かるようにしてください。

## Reference Settings

追加参照設定を使うかはAIの判断に任せます。使う場合は参照設定名、理由、設定手順、late binding または標準VBAだけの代替案を説明してください。

## Output Format

標準モジュール用の完全なVBAコード、入口プロシージャ名、配置場所、実行方法、参照設定の有無、エラー処理、前提、テスト観点を出してください。
