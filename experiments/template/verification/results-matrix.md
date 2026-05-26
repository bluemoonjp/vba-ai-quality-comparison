# results matrix

判定値は `pass`, `fail`, `blocked`, `n/a` のいずれかを使います。理由は短く書き、詳細は `excel-run-log.md` へリンクします。

| condition_id | import | run | transfer | task_specific_check | rerun_safe | no_msgbox | major_defect | note |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `codex-chatgpt-medium` |  |  |  |  |  |  |  |  |
| `codex-chatgpt-xhigh` |  |  |  |  |  |  |  |  |
| `codex-plan-medium` |  |  |  |  |  |  |  |  |
| `codex-plan-xhigh` |  |  |  |  |  |  |  |  |
| `codex-multi-review` |  |  |  |  |  |  |  |  |

任意条件へ差し替えた場合は、`generation-materials/conditions.md` の条件IDと行を揃えます。

## major_defect の目安

- 実行不能。
- 主要な集計結果が期待値と一致しない。
- 再実行で二重加算または破壊的変更が起きる。
- `MsgBox` / `InputBox` などで自動検証が停止する。
- 実ローカルパス、実データ、外部通信、追加参照などを前提にしている。
