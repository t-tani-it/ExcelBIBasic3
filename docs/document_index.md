# 文書索引（ExcelBIBasic_DeepSeek_V4.1）

- 版: RevC
- 作成日: 2026-10-22
- 目的: 現行仕様と履歴文書の優先順位を明示する
- 対象: v0.4 方式A、v0.5 デザインバリアント B、v0.6 デザイン C 設計、および別セッション・別モデル向け要求
- 旧称: `ExcelBIBasic3`（2026-10-22 に改名）。旧名を含む文書は歴史記録として残す。改名の詳細は `docs/project_rename.md` を参照する。

## 1. 優先順位

同一事項の記載が衝突する場合、次の順で上位を優先する。

1. `AGENTS.md`
2. `docs/decision_log.md` の最新決定（D-22、D-24、D-25、D-26、D-27、D-28、D-29、D-30、D-31、D-32）
3. `docs/v0.4_method_a_addendum.md`
4. `docs/variant_requirements.md`（別セッション・別モデル用）
5. `docs/variant_design_b/design_spec.md`（v0.5 実装済みの設計）
6. `docs/variant_design_c/design_spec.md`（v0.6 実装済み RevC）
7. `docs/next_model/START_HERE.md`（v0.6 次モデル入口）
8. `README.md`
9. `docs/progress.md`、`docs/handover.md`、`docs/test_spec.md` 第17節、第18節
10. RevA 文書（履歴）

- 原本は不変とし、改訂は複写による Rev 進行とする。
- 追補は RevA 文書の該当節を置換する。追補に記載のない事項は RevA を参照する。
- `docs/variant_requirements.md` は別成果物の要求であり、v0.4 の仕様を変更しない。
- 数値、セル番地、シート名は、v0.4 は `prototype_v04.xlsx`、v0.5 は `prototype_v05_design_b.xlsx` の実測値を正とする。

## 2. 現行文書

| 文書 | 役割 | 状態 | 備考 |
|---|---|---|---|
| `AGENTS.md` | プロジェクト永続ルール | 現行 | v0.4 の不変条件とデザインバリアントの方針を含む |
| `docs/decision_log.md` | 設計上の決定と根拠 | 現行 | D-22、D-24〜D-32 を優先する |
| `docs/v0.4_method_a_addendum.md` | v0.4 方式Aの追補仕様 | 現行 | RevA の PQ・データモデル前提を置換する |
| `docs/variant_requirements.md` | 別セッション・別モデル向け要求 | 現行 | v0.4 とは別ファイルの成果物に適用する |
| `docs/variant_design_b/design_spec.md` | v0.5 デザインバリアント B の設計 | 現行 | 配色、KPI、グラフ、スライサー、ヒートマップ、採用 UI |
| `docs/variant_design_b/test_results.md` | v0.5 の検証結果 | 現行 | 233 / 233 合格、3 状態の個別検査 |
| `docs/variant_design_b/known_issues.md` | v0.5 の既知の問題 | 現行 | 未検証範囲、制約、設計判断 |
| `docs/variant_design_b/change_history.md` | v0.5 の変更履歴 | 現行 | v0.4 からの変更、固定要件の維持確認 |
| `docs/variant_design_b/handover.md` | v0.5 の引き継ぎ | 現行 | コンテキスト 0% からの再開手順 |
| `docs/project_rename.md` | プロジェクト改名メモ | 現行 | 旧名、変更対象、対象外、テスト影響、GitHub、作成者方針 |
| `docs/next_model/START_HERE.md` | v0.6 次モデル入口 | 現行 RevC | 成果物、読み順、D-30〜D-32、専用fixture、禁止事項 |
| `docs/next_model/START_HERE_RevB.md` | v0.6 次モデル入口 | 履歴 | データ・UI再修正前の RevB |
| `docs/next_model/design_brief.md` | v0.6 デザインブリーフ | 現行 | v0.5 評価、Warm Slate Editorial、外部 UI 方針、点数表 |
| `docs/next_model/design_tokens.json` | v0.6 デザイントークン | 現行 RevB | 色、文字、間隔、グラフ、ヒートマップ、印刷、アクセシビリティ |
| `docs/next_model/build_cookbook.md` | v0.6 ビルド再利用集 | 現行 RevC | 専用fixture、候補ビルド、年別検証、コメント・グラフ検査 |
| `docs/next_model/build_cookbook_RevB.md` | v0.6 ビルド再利用集 | 履歴 | 専用データ導入前の RevB |
| `docs/next_model/acceptance_matrix.md` | v0.6 受入マトリクス | 現行 RevC | 固定要件、259 チェック、年別データ、コメント、グラフ境界 |
| `docs/next_model/acceptance_matrix_RevA.md` | v0.6 初版受入マトリクス | 履歴 | 変更禁止の RevA |
| `docs/next_model/acceptance_matrix_RevB.md` | v0.6 初回実装受入マトリクス | 履歴 | 専用データ refresh 前の RevB |
| `docs/next_model/ui_reference_log.md` | 外部 UI 参照ログ | 現行 RevB | v0.6 の採用・不採用、URL、ライセンス、Excel 実装 |
| `docs/next_model/reference_v05.png` | v0.5 画面参照画像 | 現行 | タイトルバーを含まない売上分析の画面範囲 |
| `docs/variant_design_c/design_spec.md` | v0.6 デザイン C の設計・実装結果 | 実装済み RevC | 専用fixture、年別タイムライン、コメント矩形、Q列内グラフ |
| `docs/variant_design_c/design_spec_RevB.md` | v0.6 初回実装設計 | 履歴 | データ更新前の RevB |
| `docs/variant_design_c/design_spec_RevA.md` | v0.6 デザイン C の初版設計 | 履歴 | 実装前の RevA。変更禁止 |
| `build_v06_design_c.py` | v0.6 ビルドスクリプト | 現行 | v0.4 から別成果物を作成 |
| `verify_v06_windows.py` | v0.6 実機検証スクリプト | 現行 | 既存 233 件以上の検証 |
| `prototype_v06_design_c.xlsx` | v0.6 デザイン C 成果物 | 現行 | v0.4 / v0.5 とは別ファイル |
| `generate_sales_fixture_v06.py` | v0.6 架空データ生成 | 現行 | 決定的生成、2026年売上を約8%増 |
| `sales_fixture_v06.json` | v0.6 専用fixture | 現行 | 576 行、2025 / 2026 年の月別・商品別値 |
| `docs/variant_design_c/test_results.md` | v0.6 検証結果 | 現行 RevC | 259 / 259、自動・年別タイムライン・コメント・Q列グラフ |
| `docs/variant_design_c/test_results_RevA.md` | 初回 v0.6 検証結果 | 履歴 | 専用fixture導入前の RevA |
| `docs/variant_design_c/test_results_RevB.md` | v0.6 初回実装検証結果 | 履歴 | 専用データ refresh 前の RevB |
| `docs/variant_design_c/known_issues.md` | v0.6 既知の問題 | 現行 RevC | 手動 UI、実印刷、環境差、未検証範囲 |
| `docs/variant_design_c/known_issues_RevA.md` | 初回 v0.6 既知の問題 | 履歴 | 専用fixture導入前の RevA |
| `docs/variant_design_c/known_issues_RevB.md` | v0.6 初回実装既知の問題 | 履歴 | 専用データ refresh 前の RevB |
| `docs/variant_design_c/preview.png` | v0.6 画面プレビュー | 現行 | タイトルバーを含まない画面画像 |
| `windows_v06_results.json` | v0.6 自動検証結果 | 現行 | 259 / 259 合格 |
| `README.md` | 利用者向けの操作と再ビルド | 現行 | v0.4、v0.5、v0.6 設計の操作説明を含む |
| `docs/progress.md` | タスクと状態 | 現行 | v0.4〜v0.6 の結果と残タスクを含む |
| `docs/handover.md` | フェーズ引き継ぎ | 現行 | v0.4〜v0.6 の計画、結果、残タスクを含む |
| `docs/test_spec.md` 第17節 | v0.4 の受入結果 | 現行 | 173 / 173 合格の記録 |
| `docs/test_spec.md` 第18節 | バリアントの受入 | 現行 | v0.5 の 233 / 233 合格と未検証範囲 |

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
2. `docs/decision_log.md` の D-22、D-24、D-25、D-26、D-27
3. `docs/v0.4_method_a_addendum.md`
4. `docs/variant_requirements.md`
5. `docs/test_spec.md` 第17節
6. `prototype_v04.xlsx` と `verify_v04_windows.py`
7. 必要に応じて `sales_fixture.json`、`windows_v04_results.json`

- 別モデルは RevA の `requirements.md`、`design_spec.md`、`implementation_plan.md`、`calculation_spec.md`、`compatibility_matrix.md` を現行仕様として扱わない。
- 別モデルは v0.4 の売上分析の見た目を模倣しない。`docs/variant_requirements.md` のデザイン可変要件に従う。
- 別モデルは `prototype_v04.xlsx` を上書きしない。成果物は別ファイルとする。

### 4.1 v0.5 実装後の読み順

1. `AGENTS.md`
2. 本索引
3. `docs/decision_log.md` の D-22、D-24、D-25、D-26、D-27、D-28、D-29
4. `docs/variant_design_b/design_spec.md`
5. `docs/variant_design_b/test_results.md`
6. `docs/variant_design_b/known_issues.md`
7. `docs/variant_design_b/handover.md`
8. `docs/variant_design_b/change_history.md`
9. `prototype_v05_design_b.xlsx`、`verify_v05_windows.py`、`windows_v05_results.json`

- v0.5 は v0.4 と別ファイルである。v0.4 のブックと仕様文書を上書きしない。
- v0.5 の手動 UI、実印刷、性能は未検証である。`known_issues.md` を先に確認する。

### 4.2 v0.6 デザイン C の読み順

1. `docs/next_model/START_HERE.md`
2. `docs/next_model/design_brief.md`
3. `docs/next_model/design_tokens.json`
4. `docs/next_model/acceptance_matrix.md`
5. `docs/variant_design_c/design_spec.md`
6. `docs/next_model/ui_reference_log.md`
7. ビルド時に `docs/next_model/build_cookbook.md`
8. 必要に応じて `docs/variant_requirements.md`、`docs/v0.4_method_a_addendum.md`、`docs/variant_design_b/known_issues.md`

v0.6 の実装後に参照する文書は、`docs/variant_design_c/test_results.md` と `docs/variant_design_c/known_issues.md` である。

- v0.6 は `prototype_v04.xlsx` を出発点とし、`prototype_v06_design_c.xlsx` を別ファイルとして作成する。
- v0.5 は視覚参照であり、ビルド入力にしない。
- 外部 UI は 3〜5 件まで参照し、採用元、URL、ライセンス、採用要素、改変内容を `ui_reference_log.md` に記録する。
- 参考画像そのものはコミットしない。ブック内に外部リンク、外部画像、外部フォントを埋め込まない。

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
| 2026-10-22 | v0.5 デザインバリアント B の設計、検証、既知の問題、変更履歴、引き継ぎを追加 |
| 2026-10-22 | プロジェクト名を `ExcelBIBasic_DeepSeek_V4.1` へ変更。改名メモを追加（D-29） |
| 2026-10-22 | v0.6 デザイン C の設計、次モデル入口、デザインブリーフ、トークン、受入マトリクス、ビルド集、外部 UI 参照ログを追加（D-28） |
| 2026-10-22 | v0.6 初回実装・検証資料を追加し、RevA / RevB を保存（D-30） |
| 2026-10-22 | v0.6 年次データ、年別タイムライン、コメント図形、グラフ境界を更新し、259 / 259 を記録（D-31、D-32） |
| 2026-10-22 | v0.6 受入・検証・既知の問題の RevC と RevB 履歴を追加 |
