# フェーズ引き継ぎ（ExcelBIBasic3）

- 更新日: 2026-10-21
- 対象フェーズ: P0 仕様策定（完了）
- 次フェーズ: P1 データ基盤（未着手）

## 1. 方針（計画）
- ExcelBIBasic2 の試作（SUMIFS 版）を土台に、見積・原価・利益シミュレーターを Power Query、データモデル、ピボットテーブル、ピボットグラフ、スライサー、タイムライン構成へ再設計する。
- 売上分析は作り直す。見積・原価・利益は関数による即時再計算を維持する。
- 購入者向けはマクロなしの `.xlsx` とする。VBA と外部リンクは使用しない。
- 文書 10 点（`AGENTS.md` と `docs/` の 9 文書）を作成し、実装と検証の基準とする。

## 2. 実装（実行）
1. 資料 14 点の存在確認と `prototype.xlsx` の構造調査を実施した。
2. 実測値を確定した。
   - シート 11 枚（採算シミュレーション、見積書、見積入力、原価(木工所)、原価(農家)、設定、サンプル・入力例、計算、使い方・検証、売上分析、売上データ）
   - 売上データ 576 行（2025-01 から 2026-12、8 商品 × 3 件）
   - メモ 74 件、`INDIRECT` 24 件、グラフ 5 個、ピボットとスライサーとタイムラインは 0
   - 計算 B4 から B12、B29:B41、D16:J25、見積入力 H16:I28、原価 L4 と M5 から M7 の式
   - 売上分析の合計（47,492,700 円、18,852,426 円、差額率 0.3969541845378342）
3. 仕様方針を合意し、文書 10 点を作成した。
   - `AGENTS.md`
   - `docs/requirements.md`、`docs/design_spec.md`、`docs/calculation_spec.md`、`docs/test_spec.md`、`docs/implementation_plan.md`、`docs/compatibility_matrix.md`、`docs/decision_log.md`
   - `docs/progress.md`、`docs/handover.md`
4. 参照画像 3 点を確認し、構成と配色のみを仕様へ反映した（保存なし）。

## 3. 結果
| 項目 | 結果 |
|---|---|
| 仕様書 | 10 文書の作成完了 |
| 要件 | FR-001 から FR-025、NFR-001 から NFR-007 の確定 |
| 決定 | D-01 から D-21 の記録 |
| 未確定 | U-01 から U-07（実装フェーズで判定） |
| 既存資産 | `prototype.xlsx`、再ビルド工程、検証スクリプトは変更なし |
| Git | `git init` と分割コミットを実施 |
| 成果物ブック | 未生成（P1 以降で生成） |

## 4. 残タスク
| フェーズ | 内容 | 完了条件 |
|---|---|---|
| P1 | PQ-01 から PQ-03、リレーション、DAX メジャー | 合計値が期待値と一致 |
| P2 | PT-01 から PT-06、PG-01 から PG-04、スライサー、タイムライン | 連動の合格 |
| P3 | KPI 4 表示、ヒートマップ、スパークライン | 表示形式と連動の合格 |
| P4 | ゲージ、目標テーブル、チェックボックス | 非連動と条件切替の合格 |
| P5 | 見積入力 B12:F12、採算 D12 | 状態表示の一元化 |
| P6 | テスト仕様の全項目 | 合格または保留の判定 |
| P7 | 再生成、文書更新、コミット | 引き継ぎ可能な状態 |

## 5. 実行コマンド
| 目的 | コマンド |
|---|---|
| venv の作成 | `python -m venv .venv` |
| PowerShell の有効化 | `.\.venv\Scripts\Activate.ps1` |
| cmd の有効化 | `.\.venv\Scripts\activate.bat` |
| pywin32 の導入 | `python -m pip install pywin32` |
| 骨格生成 | `node build.mjs`（出力先は `README.md` を参照） |
| 強化工程 | `node enhance.mjs` |
| 後処理 | `python finalize_ooxml.py` |
| 実機検証 | `python verify_excel_windows.py` |
| 旧版検証 | `python verify_v02_windows.py` |
| 状態確認 | `git status`、`git log --oneline` |

## 6. 再開手順（コンテキスト 0% から）
1. 本ファイルと `docs/progress.md` を読む。
2. `docs/decision_log.md` で確定事項を確認する。
3. `docs/implementation_plan.md` の第 3 節に従い、P1 から着手する。
4. 実装中に方式を判定した場合は `docs/decision_log.md` の未判定事項を更新する。

## 7. 参照ファイル
| ファイル | 用途 |
|---|---|
| `opencode_prompt.md` | 要件の原本 |
| `README.md` | 再ビルド手順 |
| `prototype.xlsx` | 調査対象（試作） |
| `package_checks.json` | 構造検査の期待値 |
| `notes_manifest.json` | メモ 74 件の位置と本文 |
| `engine_tests.json` | 既存 40 テストの内容 |
| `sales_fixture.json` | 576 行の期待値 |
| `test_report.md` | v0.2 の検証結果 |
