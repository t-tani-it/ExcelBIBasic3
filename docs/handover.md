# フェーズ引き継ぎ（ExcelBIBasic3）

- 更新日: 2026-10-22
- 対象フェーズ: v0.4 方式 A 実装・検証（完了）
- 次フェーズ: 別セッション・別モデルによるデザインバリアント実装、手動 UI 確認、P4 ゲージ・目標管理の要否判断

## 1. 方針（計画）
- 当初は ExcelBIBasic2 の試作（SUMIFS 版）を土台に、Power Query、データモデル、ピボットテーブル、ピボットグラフ、スライサー、タイムライン構成へ再設計する計画であった。
- D-22 によりデータモデル方式が実行不能となったため、v0.4 は D-24 の方式 A（通常ピボット、共有キャッシュ、スライサー、タイムライン、ピボットグラフ）で実装する。
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
本節は仕様策定時点の記録である。v0.4 の実装結果は第 9 節を参照する。

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
| 手動 UI | スライサー、タイムライン、印刷、再起動後の状態確認 | 操作結果の記録 |
| P4 | ゲージ、目標テーブル、チェックボックス | 要否判断と非連動の合格 |
| データモデル | 対話セッションでの PQ モデル接続確認 | D-22 の再開条件を満たす環境 |
| 運用 | 商品追加時の売上データ登録と PT 更新手順 | 手順書の確認 |

## 5. 実行コマンド
| 目的 | コマンド |
|---|---|
| v0.4 ビルド | `.\.venv\Scripts\python.exe build_v04_a.py prototype_v03.xlsx prototype_v04.xlsx` |
| v0.4 検証 | `.\.venv\Scripts\python.exe verify_v04_windows.py prototype_v04.xlsx` |
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
2. `docs/decision_log.md` の D-22 と D-24 で確定事項を確認する。
3. `prototype_v04.xlsx` を開き、`verify_v04_windows.py` で 173 / 173 合格を再確認する。
4. 方式 A の再ビルドは `build_v04_a.py` を使用する。データモデル方式の再開は D-22 の再開条件を満たす環境でのみ行う。
5. 方式を変更した場合は `docs/decision_log.md` へ追記する。

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
| `prototype_v04.xlsx` | v0.4 成果物（方式 A） |
| `build_v04_a.py` | v0.4 ビルドスクリプト |
| `verify_v04_windows.py` | v0.4 実機検証スクリプト |
| `windows_v04_results.json` | v0.4 検証結果 |

## 8. v0.3 実装フェーズ（2026-10-21）

### 方針（計画）

P1 から P4 を COM で実装し、P5 で見積入力の備考欄を復帰する計画であった。P1 は Power Query 3 本をデータモデルへ読み込み、リレーションと DAX メジャーを作成する工程である。

### 実装（実行）

P5 を `build_v03_p5.py` で実装した。あわせて `.venv` を作成し、pywin32 を導入した。

P1 は次の方式を順に試したが、いずれもモデル接続の作成で停止した。

1. `Workbook.Queries.Add` と `Connections.Add2(CreateModelConnection=True)`
2. 一時マクロ（`.xlsm`）内での同操作
3. `Application.OnTime` によるアイドル実行
4. 可視インスタンスでの実行

クエリ定義と接続定義（`CreateModelConnection=False`）は成功する。`#table` リテラルのモデル接続も成功する。ワークシートテーブルを参照するクエリのみが停止する。ダイアログウィンドウは存在せず、Mashup エンジンの停止と判断した。

### 結果

| 項目 | 結果 |
|---|---|
| P5 | 完了。`verify_v03_windows.py` 30 / 30 合格 |
| 計算エンジン | `verify_excel_windows.py` 13 / 13 合格 |
| P1 から P4 | 未実装。D-22 に記録 |
| 成果物 | `prototype_v03.xlsx`（P5 適用済み、absPath と作成者情報を除去） |

### 残タスク

1. 対話セッションで Excel を起動し、Power Query を手動でデータモデルへ読み込めるか確認する。
2. 読み込める場合、P1 の工程を再作成し、P1 から P4 を実装する。
3. 旧仕様の D12:F12 検査 3 件を `verify_v02_windows.py` から除外し、v0.3 の検査へ置き換える。

### 実行コマンド

| 目的 | コマンド |
|---|---|
| P5 の適用 | `.\.venv\Scripts\python.exe build_v03_p5.py prototype.xlsx prototype_v03.xlsx` |
| v0.3 差分検証 | `.\.venv\Scripts\python.exe verify_v03_windows.py prototype_v03.xlsx` |
| v0.2 回帰検証 | `.\.venv\Scripts\python.exe verify_v02_windows.py prototype_v03.xlsx` |
| 計算エンジン検証 | `.\.venv\Scripts\python.exe verify_excel_windows.py prototype_v03.xlsx` |

## 9. v0.4 実装フェーズ（2026-10-22）

### 方針

D-22 によりデータモデル方式が実行不能であるため、D-24 の方式 A を採用した。通常のピボットテーブルを共有キャッシュから作成し、業種・商品スライサーと日付タイムラインで絞り込む。前年比と前年比較は専用 PT で実装し、前年データを参照するためタイムライン非連動とする。

### 実装

- `build_v04_a.py` を追加した。`prototype_v03.xlsx` を入力に `prototype_v04.xlsx` を生成する。
- `SalesHistory` に `年月`、`年`、`月` の補助列を追加した。
- `商品マスタ` シートと `ProductMaster` テーブルを追加した。方式 A では PT 未接続である。
- 非表示の `制御` シートに共有キャッシュ 1 個と PT-01 から PT-08 を作成した。
- 業種スライサー、商品スライサー、日付タイムラインを作成した。タイムラインは PT-01 から PT-06 へ接続した。
- `売上分析` に KPI 4 表示、PG-01 から PG-04、商品×月ヒートマップ、スパークライン 8 本を配置した。
- OOXML 後処理で absPath と作成者情報を除去した。

### 結果

| 項目 | 結果 |
|---|---|
| 実機検証 | 173 / 173 合格 |
| 共有キャッシュ | 1 個 |
| ピボットテーブル | 8 個（PT-01 から PT-08） |
| ピボットグラフ | 4 個（PG-01 から PG-04） |
| 業種スライサー | 2 項目、全 PT 接続 |
| 商品スライサー | 8 項目、全 PT 接続 |
| タイムライン | 6 PT 接続、PT-07/08 は非連動 |
| KPI 売上 | 47,492,700 円（2026 年・全業種） |
| KPI 差額 | 18,852,426 円 |
| KPI 差額率 | 0.3969541845378342 |
| KPI 前年比 | -0.011131169147547193 |
| ヒートマップ | 8 商品 × 24 か月、条件付き書式 2 件 |
| スパークライン | 8 本 |
| メモ / INDIRECT | 74 件 / 24 件 |
| 見積・原価・利益 | 維持（利益 72,000 円、税込見積 176,000 円） |
| 公開禁止情報 | absPath 0、ローカルパス 0、作成者固定 |

### 制約

- Power Query、データモデル、DAX、リレーション、CUBE 関数は未使用である。
- 商品マスタは表示順と有効フラグの管理用であり、商品スライサーは売上データの商品列を使用する。
- タイムライン初期値は 2026 年である。2025 年のヒートマップ列は、タイムラインを全期間へ広げるまで空欄となる。
- PT-07 と PT-08 は前年データを参照するためタイムライン非連動である。業種・商品スライサーには連動する。
- 売上データを編集した場合は「すべて更新」で PT を更新する。
- ゲージ、目標テーブル、チェックボックスは未実装である。
- ヒートマップの性能と手動 UI 操作は未測定である。

### 残タスク

1. 手動 UI 確認（スライサー、タイムライン、印刷、再起動後の状態）を実施する。
2. P4 ゲージ・目標管理の要否を判断する。
3. 商品追加時の売上データ登録と PT 更新手順を確認する。

### 実行コマンド

| 目的 | コマンド |
|---|---|
| v0.4 ビルド | `.\.venv\Scripts\python.exe build_v04_a.py prototype_v03.xlsx prototype_v04.xlsx` |
| v0.4 検証 | `.\.venv\Scripts\python.exe verify_v04_windows.py prototype_v04.xlsx` |
| ビルド込み後処理 | `build_v04_a.py` が `finalize_v03_ooxml.sanitize_package` を呼び出す |

## 10. 別セッション向け要求整理（2026-10-22）

### 方針

- 別セッション・別モデルは `prototype_v04.xlsx` を複製し、既存の見積・原価・利益シートを維持する。
- `売上分析` シートは削除し、同名のシートを 0 から再作成する。
- 優れた UI があれば採用してよい。固定要件を壊さず、採用元、採用理由、ライセンスを設計書へ記録する。
- デザイン差は、配色、KPI カードの構成、グラフ構成のうち 2 つ以上とする。
- v0.4 のブックと文書は上書きしない。成果物は別ファイル・別文書とする。

### 成果物

| 文書 | 内容 |
|---|---|
| `docs/document_index.md` | 現行文書と履歴文書の優先順位 |
| `docs/v0.4_method_a_addendum.md` | v0.4 方式Aの追補仕様 |
| `docs/variant_requirements.md` | 別セッション・別モデル向けの固定要件とデザイン可変要件 |
| `docs/decision_log.md` D-25 | デザインバリアントの要求分離 |
| `docs/decision_log.md` D-26 | v0.4 ヒートマップの実装相違 |

### 注意

- RevA の `requirements.md`、`design_spec.md`、`implementation_plan.md`、`calculation_spec.md`、`compatibility_matrix.md` は履歴として扱う。
- v0.4 のヒートマップは `IFERROR` で空欄化しており、0 円、データなし、選択範囲外の 3 状態を完全には区別していない。バリアントでは 3 状態の区別を必須とする。
- バリアントの受入は `docs/test_spec.md` 第18節と `docs/variant_requirements.md` に従う。
- バリアントの手動 UI、印刷、性能は、実施していない場合は「未検証」と記録する。

### 実行コマンド

| 目的 | コマンド |
|---|---|
| バリアントの要求確認 | `docs/variant_requirements.md` を読む |
| バリアントの設計開始 | `prototype_v04.xlsx` を複製し、`売上分析` シートを削除して再作成する |
| v0.4 の再確認 | `.\.venv\Scripts\python.exe verify_v04_windows.py prototype_v04.xlsx` |

