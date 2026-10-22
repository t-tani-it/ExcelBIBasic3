# フェーズ引き継ぎ（ExcelBIBasic3）

- 更新日: 2026-10-22
- 対象フェーズ: v0.6 デザインバリアント C 実装・自動検証（完了）
- 次フェーズ: v0.6 手動 UI・実印刷・追加性能確認、v0.5 / v0.4 の手動確認、P4 判断

> 2026-10-22: プロジェクト名を `ExcelBIBasic_DeepSeek_V4.1` に変更した。入口文書のみ新名とし、本文の `ExcelBIBasic3` は歴史記録として残す。v0.6 デザイン C を担当する次モデルは、本ファイルではなく `docs/next_model/START_HERE.md` から開始する。改名後の再検証は v0.4 173 / 173、v0.5 233 / 233 合格である。

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
| v0.5 手動 UI | スライサー、タイムライン、実印刷、性能、再起動後の状態確認 | 操作結果の記録 |
| v0.4 手動 UI | スライサー、タイムライン、印刷、再起動後の状態確認 | 操作結果の記録 |
| P4 | ゲージ、目標テーブル、チェックボックス | 要否判断と非連動の合格 |
| データモデル | 対話セッションでの PQ モデル接続確認 | D-22 の再開条件を満たす環境 |
| 運用 | 商品追加時の売上データ登録と PT 更新手順 | 手順書の確認 |

## 5. 実行コマンド
| 目的 | コマンド |
|---|---|
| v0.4 ビルド | `.\.venv\Scripts\python.exe build_v04_a.py prototype_v03.xlsx prototype_v04.xlsx` |
| v0.4 検証 | `.\.venv\Scripts\python.exe verify_v04_windows.py prototype_v04.xlsx` |
| v0.5 ビルド | `.\.venv\Scripts\python.exe build_v05_design_b.py prototype_v04.xlsx prototype_v05_design_b.xlsx` |
| v0.5 検証 | `.\.venv\Scripts\python.exe verify_v05_windows.py prototype_v05_design_b.xlsx windows_v05_results.json` |
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
2. `docs/decision_log.md` の D-22、D-24、D-25、D-26、D-27 で確定事項を確認する。
3. `prototype_v04.xlsx` を開き、`verify_v04_windows.py` で 173 / 173 合格を再確認する。
4. `prototype_v05_design_b.xlsx` を開き、`verify_v05_windows.py` で 233 / 233 合格を再確認する。
5. v0.5 の設計と未検証範囲は `docs/variant_design_b/design_spec.md` と `docs/variant_design_b/known_issues.md` を確認する。
6. 方式 A の再ビルドは `build_v04_a.py`、バリアント B の再ビルドは `build_v05_design_b.py` を使用する。データモデル方式の再開は D-22 の再開条件を満たす環境でのみ行う。
7. 方式を変更した場合は `docs/decision_log.md` へ追記する。

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
| `prototype_v05_design_b.xlsx` | v0.5 デザインバリアント B の成果物 |
| `build_v05_design_b.py` | v0.5 ビルドスクリプト |
| `verify_v05_windows.py` | v0.5 実機検証スクリプト |
| `windows_v05_results.json` | v0.5 検証結果 |
| `docs/variant_design_b/` | v0.5 の設計、検証、既知の問題、変更履歴、引き継ぎ |

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

## 11. v0.5 デザインバリアント B 実装・検証（2026-10-22）

### 方針

v0.4 の複製から `売上分析` シートだけを削除し、同名のシートを 0 から再作成した。既存の見積・原価・利益、売上データ、商品マスタ、制御シートの PT-01 から PT-08 と共有キャッシュ 1 個は維持する。方式 A の集計構造を変えず、配色、KPI カード、グラフ構成、スライサー配置、ヒートマップの状態表現を変更した。

### 実装

- `build_v05_design_b.py` を追加した。`prototype_v04.xlsx` を入力に `prototype_v05_design_b.xlsx` を生成する。
- 旧 `売上分析` を削除する前に、旧スライサーキャッシュ `SL_業種`、`SL_商品`、`TL_日付` を削除した。
- 新しいスライサーキャッシュ `SL_業種_B`、`SL_商品_B`、`TL_日付_B` を作成した。
- 新しい `売上分析` をシート索引 10、`商品マスタ` の直前へ作成した。
- KPI を 2 × 2 のカード配置とした。値セルは `E6`、`I6`、`E11`、`I11` である。
- グラフを `PG-B01` から `PG-B04` の 4 個作成した。種類はエリア + 折れ線、折れ線 + マーカー、集合縦棒、円である。
- ヒートマップを `B51:Y58` の 192 セルで作成した。数式は `ISERROR` と `COUNTIFS` を使用し、`IFERROR` を使用しない。
- 条件付き書式を 3 件作成した。空白、テキスト、3 色スケールである。
- 列スパークラインを `Z51:Z58` の 8 本作成した。
- 印刷範囲を `A1:Z62`、横向き、1 ページに設定した。
- 保存後に OOXML 後処理を実行し、絶対パスと作成者情報を除去した。

### デザイン差

| 項目 | v0.4 | v0.5 デザインバリアント B |
|---|---|---|
| 配色 | 深緑、ティール、黄緑、淡緑 | スレート、コーラル、アンバー、クリーム |
| KPI カード | 横 4 枚 | 2 × 2 の 4 枚 |
| グラフ構成 | 折れ線、集合縦棒、横棒、ドーナツ | エリア + 折れ線、折れ線 + マーカー、縦棒、円 |
| スライサー | 上部の横並び | 左側のフィルタ帯 |
| ヒートマップ | 緑系 3 色、条件付き書式 2 件 | 暖色系 3 色、条件付き書式 3 件 |
| スパークライン | 商品行ごとに 8 本 | 列スパークライン 8 本 |

デザイン差は、配色、KPI カード、グラフ構成の 3 項目である。要件の最低条件である 2 項目以上を満たす。外部 UI の採用はなく、レイアウトは自作である。採用元、理由、ライセンスは `docs/variant_design_b/design_spec.md` に記録した。

### 結果

| 項目 | 結果 |
|---|---|
| 実機検証 | 233 / 233 合格 |
| シート数 | 13 |
| `売上分析` の索引 | 10 |
| 共有キャッシュ | 1 個 |
| ピボットテーブル | 8 個（PT-01 から PT-08） |
| ピボットグラフ | 4 個（PG-B01 から PG-B04） |
| 業種スライサー | 2 項目、全 PT 接続 |
| 商品スライサー | 8 項目、全 PT 接続 |
| タイムライン | 6 PT 接続、PT-07/08 は非連動 |
| KPI 売上 | 47,492,700 円（2026 年・全業種） |
| KPI 差額 | 18,852,426 円 |
| KPI 差額率 | 0.3969541845378342 |
| KPI 前年比 | -0.011131169147547193 |
| ヒートマップ | 8 商品 × 24 か月、条件付き書式 3 件 |
| スパークライン | 8 本、列種別 |
| メモ / `INDIRECT` | 74 件 / 24 件 |
| 見積・原価・利益 | 維持（利益 72,000 円、税込見積 176,000 円） |
| 公開禁止情報 | absPath 0、ローカルパス 0、作成者固定 |

### ヒートマップの 3 状態

| 状態 | 表示 | 塗り | 確認方法 |
|---|---|---|---|
| 売上 0 円 | `0万円` | 3 色スケールの最小色 | 一時コピーで `テーブル` の 2026-01 を 0 円へ変更 |
| データなし | `－` | サンド | 一時コピーで同じ判定式を実行 |
| 選択範囲外 | 空欄 | 塗りなし | 実際の `B51` と 2025-01 の検証式 |
| 売上あり | `170万円` など | 3 色スケール | 実際の `N51` など |

### 制約

- タイムラインの `TimelineStyleLight2` 設定は COM エラーとなり、`TimeSlicerStyleLight1` へフォールバックする。表示と接続は検証済みである。
- 手動 UI 操作、実印刷、性能、Excel 再起動後の状態は未検証である。
- 現行データに 0 円とデータなしの実セルがないため、3 状態のうち 2 状態は一時コピーで確認している。
- 2025 年のヒートマップ列は、タイムラインを全期間へ広げるまで空欄である。
- 商品マスタは方式 A のため PT とスライサーに未接続である。商品追加時は `商品マスタ` と `SalesHistory` の両方へ登録する。
- 対応版は Microsoft 365 のみ実機確認済みである。

### 成果物

| ファイル | 内容 |
|---|---|
| `prototype_v05_design_b.xlsx` | v0.5 デザインバリアント B のブック |
| `build_v05_design_b.py` | v0.5 ビルドスクリプト |
| `verify_v05_windows.py` | v0.5 実機検証スクリプト |
| `windows_v05_results.json` | v0.5 検証結果 |
| `docs/variant_design_b/design_spec.md` | 設計書 |
| `docs/variant_design_b/test_results.md` | 検証結果 |
| `docs/variant_design_b/known_issues.md` | 既知の問題 |
| `docs/variant_design_b/change_history.md` | 変更履歴 |
| `docs/variant_design_b/handover.md` | 引き継ぎ |

### 残タスク

1. スライサーのクリック、複数選択、クリアを確認する。
2. タイムラインのドラッグ、期間変更、全期間拡張を確認する。
3. Excel を再起動し、選択状態、表示位置、印刷設定を確認する。
4. 横向き 1 ページの実印刷を確認する。
5. 低スペック端末で操作時間を計測する。
6. 商品追加の操作手順を確認する。
7. 必要に応じて Excel 2016、2019、2021、2024 で確認する。
8. 色覚多様性の観点でコントラスト比を確認する。

### 実行コマンド

| 目的 | コマンド |
|---|---|
| v0.5 ビルド | `.\.venv\Scripts\python.exe build_v05_design_b.py prototype_v04.xlsx prototype_v05_design_b.xlsx` |
| v0.5 検証 | `.\.venv\Scripts\python.exe verify_v05_windows.py prototype_v05_design_b.xlsx windows_v05_results.json` |
| v0.4 の再確認 | `.\.venv\Scripts\python.exe verify_v04_windows.py prototype_v04.xlsx` |

## 12. v0.6 デザインバリアント C 実装・検証（2026-10-22）

### 方針（計画）

- 作業ルートを Luna コピーに固定し、親・兄弟プロジェクトを参照しない。
- v0.5 を一時出力先で 233 / 233 再確認してから、`prototype_v04.xlsx` を読み取り専用の出発点として v0.6 を作成する。
- 方式 A と既存計算を維持し、Warm Slate Editorial、4 グラフ、3 状態ヒートマップ、証拠資料を作成する。

### 実装（実行）

- AGENTS.md の先頭へ Luna 固有ルールを追加し、作業ルートと Git 操作の優先順位を明記した。
- Python 3.11 ランチャーがなかったため、指定フォールバックで Python 3.9.7 の `.venv` を作成した。pywin32 312 を導入し、COM 読込を確認した。
- v0.5 の一時回帰は 233 / 233 合格した。`prototype_v05_design_b.xlsx` と v0.5 専用文書は変更していない。
- `build_v06_design_c.py`、`verify_v06_windows.py`、`prototype_v06_design_c.xlsx`、`windows_v06_results.json` を作成した。
- 設計書 RevB、UI 参照ログ、検証結果、既知の問題、プレビュー画像、進捗・索引を更新した。D-30 に補助範囲とコントラスト調整を記録した。
- Git は未初期化のままであり、git init、commit、push は実行していない。

### 結果

| 項目 | 結果 |
|---|---|
| 自動検証 | 240 / 240 合格（既存 233 件を維持） |
| デザイン評価 | 20 / 24。重大項目に 0 点なし |
| PDF 確認 | A4 横向き、1 ページ。実印刷は未検証 |
| 性能 | 3 回の中央値。v0.5 開く 0.284 秒 / 再計算 0.035 秒、v0.6 開く 0.230 秒 / 再計算 0.034 秒 |
| 公開禁止情報 | 自動走査 0 件。作成者は `ExcelBIBasic_DeepSeek_V4.1` |
| 成果物 | `prototype_v06_design_c.xlsx`、ビルド・検証スクリプト、結果 JSON、設計・検証文書、`preview.png` |

### 残タスク

- スライサー操作、タイムライン操作、Excel 再起動後の状態を実機確認する。
- 実印刷、100% / 125% 表示、色覚多様性を確認する。
- スライサー操作性能と低スペック環境の性能を測定する。
- Python 3.11 系と Excel の完全な更新番号で再確認する。
- v0.5 / v0.4 の手動 UI 確認、P4 ゲージ・目標管理の要否判断を継続する。

### 実行コマンド

| 目的 | コマンド |
|---|---|
| v0.5 事前回帰 | `.\.venv\Scripts\python.exe verify_v05_windows.py prototype_v05_design_b.xlsx $env:TEMP\windows_v05_results_check.json` |
| v0.6 ビルド | `.\.venv\Scripts\python.exe build_v06_design_c.py prototype_v04.xlsx prototype_v06_design_c.xlsx` |
| v0.6 検証 | `.\.venv\Scripts\python.exe verify_v06_windows.py prototype_v06_design_c.xlsx windows_v06_results.json` |
## 13. v0.6 データ差分・UI修正（2026-10-22）

### 方針（計画）

- 共通 fixture、v0.4 / v0.5 を変更せず、v0.6 専用の決定的な架空データを作る。
- 2025年の基準データを保ち、2026年は税抜売上総額を約8%増とし、月別・商品別構成を変える。
- 年別タイムライン選択、メモ図形、グラフの列境界を自動検証する。

### 実装（実行）

- `generate_sales_fixture_v06.py` から `sales_fixture_v06.json` を生成した。乱数は使わず、共通 `sales_fixture.json` は変更していない。
- v0.6 ビルドはコピー先の `SalesHistory` だけを fixture の値へ差し替え、v0.4 入力は読み取り専用で維持した。
- 2025年は2025-01〜2025-12、2026年は2026-01〜2026-12 の月別・商品別配分を変更した。
- メモ74件を矩形・自動サイズへ設定した。4グラフは最終列幅の後でQ列内へ収めた。
- タイムラインを2025年のみ、2026年のみ、両年に切り替え、ヒートマップ追従を一時コピーで検証した。
- 正式ブックを置き換える前にユーザーが開いていたため、閉じてもらってから合格候補を反映した。

### 結果

| 項目 | 結果 |
|---|---|
| 2025年売上 / 原価 / 差額 | 48,027,300 円 / 28,999,224 円 / 19,028,076 円 |
| 2026年売上 / 原価 / 差額 | 51,869,100 円 / 31,729,100 円 / 20,140,000 円 |
| 2026年の売上成長 | 2025年比 7.9992% |
| 自動検証 | 259 / 259 合格 |
| コメント | 74 / 74 矩形、自動サイズ、本文一致 |
| グラフ | 4 / 4 Q列内 |
| タイムライン | 2025のみ、2026のみ、両年の3ケース合格 |
| v0.4 / v0.5 | ブックおよび専用文書は変更なし |
| Git | 未初期化。git init、commit、push は未実施 |

### 残タスク

- マウスによるスライサー・タイムライン操作、Excel 再起動後の状態を確認する。
- 実印刷、100% / 125% 表示、色覚多様性シミュレーションを確認する。
- スライサー操作時間と低スペック環境での性能を測定する。

### 実行コマンド

| 目的 | コマンド |
|---|---|
| v0.6 fixture 再生成 | `.\.venv\Scripts\python.exe generate_sales_fixture_v06.py` |
| v0.6 候補ビルド | `.\.venv\Scripts\python.exe build_v06_design_c.py prototype_v04.xlsx prototype_v06_design_c_candidate.xlsx` |
| 候補検証 | `.\.venv\Scripts\python.exe verify_v06_windows.py prototype_v06_design_c_candidate.xlsx windows_v06_candidate_results.json` |
