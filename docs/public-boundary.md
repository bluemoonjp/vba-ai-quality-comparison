# 公開境界

このリポジトリはPublic公開を前提にしています。commitする成果物は、すべて公開してよい内容である必要があります。

## 追加してよい成果物

- 合成workbook仕様、合成マスタデータ、期待結果。
- `.bas` や `.cls` などのexported VBA。
- 公開用に整理したprompt、生成ログ、レビュー記録。
- 公開前チェックを通過したサンプルworkbook。
- 架空データだけを使ったdocs、評価基準、検証ケース、比較サマリ。

## 追加しない成果物

- 実運用データ、実クレーム内容、実支店名、顧客名、社員名、会社固有の識別子。
- 実ローカルパス、ユーザー名、端末名、OneDriveパス、ネットワークパス、社内ファイル名。
- secret、token、APIレスポンス、private connector output、private source material。
- private screenshot、実workbook由来のスクリーンショット。
- 未確認のOfficeバイナリ。対象は `.xls`, `.xlsx`, `.xlsm`, `.xlsb` を含む。
- 隠しシート、コメント、名前定義、リンク、ドキュメントプロパティなどに非公開情報を含む可能性があるworkbook。

## Workbook公開前チェック

workbookを `experiments/<id>/samples/workbooks/` へ移してcommitする前に、次を確認します。

- workbookが架空データだけを使っている。
- ドキュメントプロパティと作成者メタデータが公開可能である。
- hidden / very hidden sheetを確認済みである。
- コメント、メモ、threaded comments、セルメタデータを確認済みである。
- 名前定義、数式、外部リンク、pivot cache、connection、queryに実パスや非公開情報がない。
- VBAモジュール、フォーム、参照設定、定数に実パス、非公開名、secretがない。
- Windows版Excelで開き、想定するマクロ/セキュリティ挙動を説明できる。
- ファイル名自体が公開可能である。

## Git境界

このディレクトリは、親workspace配下にある独立Gitリポジトリです。commit、remote、tag、GitHub Issueは、明示がない限りこのリポジトリに属します。

このリポジトリのファイルを親workspaceリポジトリへcommitしません。親workspaceのprivateな文脈をこの公開リポジトリへコピーしません。

## 既定の保存ルール

未確認の生成workbookは `experiments/<id>/samples/generated/` や `experiments/<id>/model-artifacts/<condition>/scratch/` などのignore済みscratch場所に置きます。

公開前チェック済みのサンプルworkbookだけを `experiments/<id>/samples/workbooks/` へ置けます。モデル成果物としての `.xlsm` は保存せず、`.bas` とMarkdownログを正本にします。

初回検証 `001-claims-monthly-simple` では、サンプル `.xlsx` を保存し、各モデルが作成する成果物は `.bas` などのテキスト形式だけ保存します。
