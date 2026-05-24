# 共通promptテンプレート

あなたはExcel VBAに詳しい開発者です。

このpromptは全条件で同一に使います。他条件の生成物、レビュー、評価結果は参照しないでください。

添付または参照できるサンプルファイルと仕様書を確認し、月次クレーム集計表へ日次クレーム件数を転記するVBAを作成してください。

## 入力

- 仕様書: `generation-materials/task-spec.md`
- 制約: `generation-materials/prompt-constraints.md`
- 条件定義: `generation-materials/conditions.md`
- サンプル: `samples/workbooks/`
- 期待結果: `samples/expected/monthly-expected.csv`

参照してよいファイルは、このpromptで示した入力と、必要に応じて `samples/README.md` までです。他条件の `model-artifacts/`, `prompt-records/`, `verification/`, `reports/` は参照しないでください。

## 出力してほしいもの

- 標準モジュールへ貼り付けられる完全なVBAコード。
- `.bas` として保存できる形のコード。
- 入口プロシージャ名。
- 実行方法。
- 参照設定が必要かどうか。
- ログ出力方法。
- 再実行しても二重加算しない理由。
- Markdownログとして残せる説明。

## 注意

MsgBoxで処理を止める実装、実ローカルパスの固定、外部通信、サンプルworkbookの構造変更を前提にした実装は禁止です。

分からないことは確認質問してください。確認質問の回数は制限しません。ただし、最終出力後の修正依頼は1回だけです。
