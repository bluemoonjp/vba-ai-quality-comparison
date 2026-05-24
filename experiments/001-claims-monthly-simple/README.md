# 001-claims-monthly-simple

初回検証です。日次のクレーム集計ファイルから、月次集計表へ支店+業務別の件数を転記するVBAを、Codexの使い方別に作成させて比較します。

## 検証範囲

- 対象月: 2026年5月の合成データ。
- 日次ファイル名: `クレーム集計YYMMDD.xlsx`。
- 月次ファイル名: `月次クレーム集計YYMM.xlsm`。
- 主キー: 支店コード + 業務コード。
- 初回は正常系中心。
- 必須確認: 転記、月合計、再実行時の二重加算防止。

## 対象外

- 実務データ。
- 複雑な小計。
- 重複ファイル、未知キー、非数値などの網羅的な異常系。
- モデル成果物としての `.xlsm` 保存。

## 条件

条件定義は `generation-materials/conditions.md` を正本にします。

- `codex-chatgpt-medium`
- `codex-chatgpt-xhigh`
- `codex-plan-medium`
- `codex-plan-xhigh`
- `codex-multi-review`

## 成果物

- 生成用素材: `generation-materials/`
- サンプル: `samples/`
- prompt記録: `prompt-records/`
- モデル成果物: `model-artifacts/`
- 評価: `verification/`
- 人間向けレポート: `reports/`
