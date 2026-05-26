# 初回比較レポート

対象: `001-claims-monthly-simple`

このレポートは、同じExcel VBA課題を5つのCodex利用条件で生成し、正常系の実Excel評価と生成物記録から、日本人向けに初回の比較結果を整理したものです。評価結果の正本は `verification/results-matrix.md`、実Excel実行の詳細は `verification/excel-run-log.md` です。

## 結論

今回の正常系実Excel評価では、5条件すべてが合格でした。すべての条件でVBE import、入口プロシージャ実行、期待値との一致、月合計、再実行時の二重加算防止、MsgBoxで停止しないことを満たし、重大欠陥は確認されませんでした。

そのため、今回の差は「正確性」ではなく、生成過程、構造、レビュー容易性、説明の厚み、次回複雑化時の伸びしろに出ました。とくに、コードが短く読みやすいか、対象月の扱いが固定か自動抽出か、ログや例外処理の説明がどれだけ明確か、レビュー役が入ったことで確認記録が厚くなったかが比較ポイントです。

保守性は `verification/results-matrix.md` では未評価のままです。本レポートでも点数化せず、VBA本文と説明文を読んだ定性的観察に留めます。

また、全条件で修正依頼0回だったため、修正後改善度は比較不能です。全条件で確認質問0回または標準回答採用だったため、今回の仕様では質問力の差も測れませんでした。

## 実験概要と評価方法

課題は、日次ファイル `クレーム集計YYMMDD.xlsx` の `日次集計` シートから、月次ファイル `月次クレーム集計YYMM.xlsm` の `月次集計` シートへ、支店コード + 業務コード単位でクレーム件数を転記するVBAを作ることです。対象月は2026年5月の合成データです。

必須ルールは、転記前に日別列と月合計をクリアすること、同じ日付・同じ支店+業務が複数行ある場合は合算すること、転記後に月合計を計算すること、同じブックで2回実行しても二重加算しないこと、MsgBoxで処理を止めずログに残すことです。

比較した条件は次の5つです。

| condition_id | 実行形態 |
| --- | --- |
| `codex-chatgpt-medium` | ChatGPT代替としてのprompt応答のみ、reasoning medium |
| `codex-chatgpt-xhigh` | ChatGPT代替としてのprompt応答のみ、reasoning xhigh |
| `codex-plan-medium` | 短い計画から実装まで、ファイル直接操作あり、reasoning medium |
| `codex-plan-xhigh` | 短い計画から実装まで、ファイル直接操作あり、reasoning xhigh |
| `codex-multi-review` | Orchestrator、Explorer、Worker、Reviewerによるマルチエージェント構成 |

評価では、各条件の `final-vba.bas` を評価前に修正せず、Excel VBEへimportして実行しました。期待値照合では `branch_code + business_code` をキーにし、`day_01`、`day_02`、`day_15`、`month_total` と、その他日別列が0のままであることを確認しました。同じworkbookセッションで再実行し、二重加算が起きないことも確認しています。

## 比較表

| condition_id | 実Excel評価 | 入口プロシージャ | 構造・説明の観察 | 定性的な見立て | 主な根拠 |
| --- | --- | --- | --- | --- | --- |
| `codex-chatgpt-medium` | 合格 | `ImportDailyClaimsToMonthly` | 標準的で十分に分割され、必要事項は揃っている | 初回正常系には十分。過不足の少ない回答。 | `verification/results-matrix.md`, `model-artifacts/codex-chatgpt-medium/final-output.md`, `model-artifacts/codex-chatgpt-medium/final-vba.bas` |
| `codex-chatgpt-xhigh` | 合格 | `ImportDailyClaimsToMonthly` | mediumより説明とログ観点が厚く、対象外キーや数値不正の説明が明確 | 異常系追加時に観察しやすそう。 | `verification/results-matrix.md`, `model-artifacts/codex-chatgpt-xhigh/final-output.md`, `model-artifacts/codex-chatgpt-xhigh/final-vba.bas` |
| `codex-plan-medium` | 合格 | `RunClaimsMonthlyTransfer` | 5条件中もっとも短めで把握しやすく、計画と実装方針が簡潔 | レビューしやすい一方、複雑化時の拡張余地は追加確認が必要。 | `verification/results-matrix.md`, `prompt-records/codex-plan-medium.md`, `model-artifacts/codex-plan-medium/final-vba.bas` |
| `codex-plan-xhigh` | 合格 | `ImportDailyClaimsToMonthly` | 最も大きく、防御的な補助関数、ログ、エラー処理の説明が厚い | 複雑化への伸びしろは大きいが、初回課題にはやや重い。 | `verification/results-matrix.md`, `prompt-records/codex-plan-xhigh.md`, `model-artifacts/codex-plan-xhigh/final-vba.bas` |
| `codex-multi-review` | 合格 | `TransferDailyClaimsToMonthly` | 中程度の大きさで、Explorer、Worker、Reviewerの記録が残る | 成果物単体より、レビュー可能性と監査記録の厚さが強み。 | `verification/results-matrix.md`, `model-artifacts/codex-multi-review/worker-log.md`, `model-artifacts/codex-multi-review/reviewer-log.md`, `model-artifacts/codex-multi-review/final-vba.bas` |

全条件で、`verification/results-matrix.md` 上の `import`、`run`、`transfer`、`monthly_total`、`rerun_safe`、`no_msgbox` はpassです。`maintainability` は全条件で未評価です。

## 条件別の強み・弱み

`codex-chatgpt-medium` は、prompt応答のみでも正常系に必要な要件を満たしました。入口、配置場所、実行方法、参照設定不要、ログ、再実行安全性が説明されており、初回の「貼って動かす」用途には十分です。一方で、生成過程の記録や判断の分解はplan系・multi-review系より薄く、後から「なぜその設計にしたか」を追う材料は少なめです。

根拠: `prompt-records/codex-chatgpt-medium.md`, `model-artifacts/codex-chatgpt-medium/final-output.md`, `model-artifacts/codex-chatgpt-medium/final-vba.bas`。

`codex-chatgpt-xhigh` は、同じprompt応答のみでも、ログ項目やエラー時の説明がより丁寧でした。固定パスを避ける説明、外部通信なし、構造変更前提なしといった公開境界にも言及があり、読む側に安心材料があります。弱みとしては、正常系評価ではmediumとの正確性差が出なかったため、推論設定の高さが今回の合否には反映されませんでした。

根拠: `prompt-records/codex-chatgpt-xhigh.md`, `model-artifacts/codex-chatgpt-xhigh/final-output.md`, `model-artifacts/codex-chatgpt-xhigh/final-vba.bas`。

`codex-plan-medium` は、コード量が比較的少なく、構造を追いやすい点が強みです。短い計画から成果物を直接作っており、実装意図も簡潔に残っています。一方、対象月を `2605` として定数化しているため、今回の仕様には合いますが、次回に複数月や月次ブック名からの自動判定を求める場合は修正が必要になりそうです。

根拠: `prompt-records/codex-plan-medium.md`, `model-artifacts/codex-plan-medium/final-output.md`, `model-artifacts/codex-plan-medium/final-vba.bas`。

`codex-plan-xhigh` は、最も説明と防御的な実装が厚い条件でした。月次ブック名から対象月を抽出し、フォルダ存在、ファイル名、数値、キー正規化などの補助関数を多く持っています。複雑化時には伸びしろがありそうですが、初回正常系だけを見るとコードが大きく、非プログラマーが読むには負荷が高い可能性があります。

根拠: `prompt-records/codex-plan-xhigh.md`, `model-artifacts/codex-plan-xhigh/final-output.md`, `model-artifacts/codex-plan-xhigh/final-vba.bas`。

`codex-multi-review` は、VBA単体の合否に加えて、Explorerの仕様整理、Workerの判断、Reviewerの指摘なし判定がログとして残った点が特徴です。成果物の正確性だけでなく、公開境界や汚染防止を後から確認しやすい構成でした。一方で、今回はReviewer指摘なし、修正依頼0回だったため、レビューが実際に品質を引き上げたかまでは測れていません。

根拠: `prompt-records/codex-multi-review.md`, `model-artifacts/codex-multi-review/explorer-log.md`, `model-artifacts/codex-multi-review/worker-log.md`, `model-artifacts/codex-multi-review/reviewer-log.md`, `model-artifacts/codex-multi-review/final-vba.bas`。

## 失敗しなかった点 / 差が出た点

失敗しなかった点は明確です。5条件すべてで、日次から月次への転記、月合計計算、再実行安全性、MsgBoxなし、固定実パスなし、外部通信なし、追加参照強制なしを満たしました。実Excel評価でも、期待CSVとの照合はすべて一致しました。

また、全条件が `Scripting.Dictionary` をlate bindingで使い、支店コード + 業務コードをキーにする方針を採用しました。転記前に日別列と月合計をクリアしてから再集計する設計も共通しており、今回の最重要要件である二重加算防止に対して、全条件が同じ方向の解を出しています。

差が出たのは、主に次の点です。

| 観点 | 差の内容 |
| --- | --- |
| 対象月の扱い | `codex-plan-medium` と `codex-multi-review` は対象月を定数化。`codex-chatgpt-medium`、`codex-chatgpt-xhigh`、`codex-plan-xhigh` は月次ブック名などから判断する方向。 |
| コード量 | `codex-plan-medium` が短め、`codex-plan-xhigh` が最も大きい。短さはレビュー容易性、大きさは防御的処理の厚さにつながる。 |
| ログ説明 | xhigh系とmulti-reviewは、警告、エラー、処理サマリーの説明が比較的厚い。 |
| 生成過程 | prompt応答のみ条件は生成過程が薄く、plan系は計画が残り、multi-reviewは役割別ログが残る。 |
| 監査しやすさ | `codex-multi-review` はReviewerログで公開リスク、VBA本文一致、参照範囲を確認しており、後から追いやすい。 |

## 確認質問と修正依頼の観察

確認質問は、全条件で0回または標準回答採用でした。標準回答として、日次フォルダは月次ファイルと同じフォルダ配下の `daily/2605/`、対象月は `2605`、ログはログシート、追加参照は避ける方針が使われています。

今回の仕様は、必要な前提が `generation-materials/clarification-answer-bank.md` に整理されており、モデルが追加質問をしなくても進められる状態でした。そのため、「曖昧な仕様をどれだけ質問で詰められるか」という質問力の差は測れていません。

修正依頼も全条件で0回でした。すべてfirst outputをそのままfinal outputとして扱っているため、修正依頼後にどの条件がどれだけ改善するかは比較不能です。今回見えているのは、初回生成物として正常系を通せたか、そしてその生成物がどれだけ読みやすく説明されているかです。

## 限界

この実験は初回検証であり、正常系中心です。未知キー、非数値、重複ファイル、欠損日次ファイル、複数月、月末以外の31日列、シート欠落、日次ファイルの拡張子違いなどの異常系は、十分には評価していません。

保守性は `results-matrix.md` で未評価です。本レポートの保守性に関する記述は、コード量、関数分割、説明の厚み、ログ記録の読みやすさからの定性的観察であり、採点結果ではありません。

実Excel評価では、VBE importのためにscratch内でCP932変換コピーを使っています。追跡対象の `.bas` は変更していませんが、VBE import時の文字コードは今後も検証手順上の注意点です。また、cloud-synced folder配下で `ThisWorkbook.Path` がURLになる問題を避けるため、一時ドライブ経由で実行しています。この回避は評価環境上の手当であり、各モデル成果物の修正ではありません。

さらに、今回は5条件すべてが合格したため、正確性の優劣はつけられません。差を読むには、より複雑な仕様や、あえて曖昧さを残した仕様での追加検証が必要です。

## 次回複雑化案

次回は、正常系では差が出にくかったため、仕様を少し複雑にして比較するのがよさそうです。

1. 未知キー、非数値、空欄、日次シート欠落などの軽い異常系を追加する。
2. 同一日・同一支店+業務の重複行を増やし、合算とログの妥当性を見る。
3. 複数月または対象月を可変にして、対象月固定実装と自動判定実装の差を見る。
4. 欠損日次ファイルや対象外ファイルを混ぜ、処理継続と警告ログの質を見る。
5. 月次表に小計や保護列を追加し、セル範囲固定の実装がどこまで耐えられるかを見る。
6. 修正依頼を1回必ず入れる設計にし、修正後改善度を比較できるようにする。
7. あえて仕様に曖昧さを残し、確認質問の質を測れる課題にする。

初回の結論としては、「正しく動くVBAを出す」だけなら5条件すべてが到達しました。次回は、例外処理、変更耐性、質問力、レビューでの改善効果を測れる条件にすると、Codexの使い方による差がより見えやすくなります。
