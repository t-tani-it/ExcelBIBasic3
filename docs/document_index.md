# 文書索引（ExcelBIBasic3）

- 版: RevA
- 作成日: 2026-10-22
- 目的: 現行仕様と履歴文書の優先順位を明示する
- 対象: v0.4 方式A、および別セッション・別モデル向けデザインバリアント

## 1. 優先順位

同一事項の記載が衝突する場合、次の順で上位を優先する。

1. `AGENTS.md`
2. `docs/decision_log.md` の最新決定（D-22、D-24、D-25、D-26）
3. `docs/v0.4_method_a_addendum.md`
4. `docs/variant_requirements.md`（別セッション・別モデル用）
5. `README.md`
6. `docs/progress.md`、`docs/handover.md`、`docs/test_spec.md` 第17節、第18節
7. RevA 文書（履歴）

- 原本は不変とし、改訂は複写による Rev 進行とする。
- 追補は RevA 文書の該当節を置換する。追補に記載のない事項は RevA を参照する。
- `docs/variant_requirements.md` は別成果物の要求であり、v0.4 の仕様を変更しない。
- 数値、セル番地、シート名は、特に断らない限り `prototype_v04.xlsx` の実測値を正とする。

## 2. 現行文書

| 文書 | 役割 | 状態 | 備考 |
|---|---|---|---|
| `AGENTS.md` | プロジェクト永続ルール | 現行 | v0.4 の不変条件とデザインバリアントの方針を含む |
| `docs/decision_log.md` | 設計上の決定と根拠 | 現行 | D-22、D-24、D-25、D-26 を優先する |
| `docs/v0.4_method_a_addendum.md` | v0.4 方式Aの追補仕様 | 現行 | RevA の PQ・データモデル前提を置換する |
| `docs/variant_requirements.md` | 別セッション・別モデル向け要求 | 現行 | v0.4 とは別ファイルの成果物に適用する |
| `README.md` | 利用者向けの操作と再ビルド | 現行 | v0.4 の操作説明を含む |
| `docs/progress.md` | タスクと状態 | 現行 | v0.4 の結果とバリアントのタスクを含む |
| `docs/handover.md` | フェーズ引き継ぎ | 現行 | v0.4 とバリアントの引き継ぎを含む |
| `docs/test_spec.md` 第17節 | v0.4 の受入結果 | 現行 | 173 / 173 合格の記録 |
| `docs/test_spec.md` 第18節 | バリアントの受入 | 現行 | 別モデル成果物の検証項目 |

## 3. 履歴文書

次の文書は仕様策定時点のRevAであり、方式Aの現行仕様ではない。参照する場合は、必ず本索引と追補を先に読む。

| 文書 | 履歴としての内容 | 現行との主な相違 |
|---|---|---|
| `docs/requirements.md` | 当初の要件定義 | PQ、データモデル、DAX、CUBE、ゲージ前提 |
| `docs/design_spec.md` | 当初の設計仕様 | データモデル、CUBE、PT-01からPT-06、ゲージ前提 |
| `docs/implementation_plan.md` | 当初の実装計画 | P1のPQ・モデル、CUBE、Node.js工程前提 |
| `docs/calculation_spec.md` | 計算仕様 | 第1から第4節、第6から第10節は有効。第5章は旧方式 |
| `docs/compatibility_matrix.md` | 当初の対応版マトリクス | データモデル、DAX、CUBE、GETPIVOTDATAの注記が旧方式 |
| `docs/test_spec.md` 第4から第13節 | 当初のテスト仕様 | データモデル、ゲージ、チェックボックス前提 |
| `opencode_prompt.md` | 当初の依頼プロンプト | 原本。変更しない |

## 4. 別セッション・別モデルの読み順

1. `AGENTS.md`
2. `docs/decision_log.md` の D-22、D-24、D-25、D-26
3. `docs/v0.4_method_a_addendum.md`
4. `docs/variant_requirements.md`
5. `docs/test_spec.md` 第17節
6. `prototype_v04.xlsx` と `verify_v04_windows.py`
7. 必要に応じて `sales_fixture.json`、`windows_v04_results.json`

- 別モデルは RevA の `requirements.md`、`design_spec.md`、`implementation_plan.md`、`calculation_spec.md`、`compatibility_matrix.md` を現行仕様として扱わない。
- 別モデルは v0.4 の売上分析の見た目を模倣しない。`docs/variant_requirements.md` のデザイン可変要件に従う。
- 別モデルは `prototype_v04.xlsx` を上書きしない。成果物は別ファイルとする。

## 5. 更新ルール

- フェーズ完了時は `docs/handover.md` と `docs/progress.md` を更新する。
- 設計判断を変更した場合は `docs/decision_log.md` へ追記する。
- 現行仕様を変更した場合は、RevA 原本を書き換えず、追補または新しい Rev を作成する。
- 文書の追加・改名・廃止は本索引へ反映する。
- 公開禁止情報、ローカルパス、APIキー、個人ID、認証情報を含めない。

## 6. 変更履歴

| 日付 | 内容 |
|---|---|
| 2026-10-22 | 初版作成。v0.4 追補、バリアント要求、履歴文書の優先順位を記録 |
