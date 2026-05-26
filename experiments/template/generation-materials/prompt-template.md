# prompt template

以下を各条件に同じ内容で渡します。条件差を出すための system / mode / reasoning 指定は、`conditions.md` と prompt 記録に残します。

```text
あなたはVBA実装者です。

## 目的

<task_summary>

## 参照してよいファイル

- generation-materials/task-spec.md
- generation-materials/conditions.md
- samples/README.md
- samples/source/<sample_file>
- samples/expected/<expected_file>

## 作成してほしいもの

- final-output.md: 実装方針、実行手順、注意点
- final-vba.bas: Excelへ import できる標準モジュール
- notes.md: 判断、制約、未確認事項

## 制約

- 実データ、実パス、secret、private情報を使わない。
- 外部通信をしない。
- 検証用 `.xlsm` を成果物にしない。
- 固定のローカル絶対パスをコードに埋め込まない。
- 入口プロシージャを明示する。
- エラー時に検証が止まる UI を出す場合は理由を書く。可能ならログや戻り値で表現する。

## 追加ルール

- 確認質問があれば先に質問する。
- 最終出力後の修正依頼は最大1回です。
```
