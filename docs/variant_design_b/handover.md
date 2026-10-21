# デザインバリアント B 引き継ぎ（ExcelBIBasic3）

- 更新日: 2026-10-22
- 対象: `prototype_v05_design_b.xlsx`
- 状態: 実装・自動検証・文書化 完了、手動 UI・実印刷・性能 未検証
- 次の担当: 実操作の確認、印刷と性能の確認、必要に応じた微調整

## 1. 最初に読む順序

1. `AGENTS.md`
2. `docs/document_index.md`
3. `docs/decision_log.md` の D-22、D-24、D-25、D-26、D-27
4. `docs/v0.4_method_a_addendum.md`
5. `docs/variant_requirements.md`
6. `docs/test_spec.md` 第18節
7. `docs/variant_design_b/design_spec.md`
8. `docs/variant_design_b/test_results.md`
9. `docs/variant_design_b/known_issues.md`
10. `docs/variant_design_b/change_history.md`

## 2. 現在の状態

| 項目 | 状態 |
|---|---|
| v0.5 ブック | 生成済み |
| 自動検証 | 233 / 233 合格 |
| 設計書 | 作成済み |
| 検証結果 | 作成済み |
| 既知の問題 | 作成済み |
| 変更履歴 | 作成済み |
| 手動 UI 操作 | 未検証 |
| 実印刷 | 未検証 |
| 性能 | 未検証 |
| v0.4 のブック | 未変更 |

## 3. 成果物

| ファイル | 内容 |
|---|---|
| `prototype_v05_design_b.xlsx` | v0.5 デザインバリアント B のブック |
| `build_v05_design_b.py` | v0.4 から v0.5 を生成するビルドスクリプト |
| `verify_v05_windows.py` | v0.5 の実機検証スクリプト |
| `windows_v05_results.json` | 233 件の検証結果 |
| `docs/variant_design_b/design_spec.md` | 設計書 |
| `docs/variant_design_b/test_results.md` | 検証結果 |
| `docs/variant_design_b/known_issues.md` | 既知の問題 |
| `docs/variant_design_b/change_history.md` | 変更履歴 |
| `docs/variant_design_b/handover.md` | 本引き継ぎ |

## 4. 実行コマンド

PowerShell とコマンドプロンプト（cmd）のどちらでも、プロジェクト直下で実行する。仮想環境の有効化は次のとおりである。

| シェル | 有効化 |
|---|---|
| PowerShell | `.\.venv\Scripts\Activate.ps1` |
| コマンドプロンプト（cmd） | `.\.venv\Scripts\activate.bat` |

| 目的 | コマンド |
|---|---|
| v0.5 ビルド | `.\.venv\Scripts\python.exe build_v05_design_b.py prototype_v04.xlsx prototype_v05_design_b.xlsx` |
| v0.5 検証 | `.\.venv\Scripts\python.exe verify_v05_windows.py prototype_v05_design_b.xlsx windows_v05_results.json` |
| v0.4 の再確認 | `.\.venv\Scripts\python.exe verify_v04_windows.py prototype_v04.xlsx` |
| 状態確認 | `git status --short` |
| 差分確認 | `git diff --stat` |

## 5. ビルドの前提

- Windows 版 Excel と pywin32 が必要である。
- 入力は `prototype_v04.xlsx` である。v0.5 を入力にしない。
- 出力先の `prototype_v05_design_b.xlsx` が Excel で開かれている場合は閉じる。
- `prototype_v04.xlsx` が Excel で開かれている場合は、読み取り専用で開くか、複製した入力を使用する。
- ビルドは旧スライサーキャッシュ `SL_業種`、`SL_商品`、`TL_日付` を削除する。
- ビルドは旧 `売上分析` を削除し、新しいシートを索引 10 に作成する。
- PT-01 から PT-08 と共有キャッシュ 1 個は再利用する。
- 保存後に `finalize_v03_ooxml.sanitize_package` を実行し、絶対パスと作成者情報を除去する。

## 6. 検証の前提

- 検証対象は読み取り専用で開く。
- 0 円とデータなしの検査は一時コピーで行い、保存しない。
- 検証結果は `windows_v05_results.json` に保存する。
- 233 件すべて合格した場合だけ終了コード 0 となる。
- 不合格がある場合は、検証名、期待値、実測値を確認する。

## 7. 検証結果の要点

| 項目 | 結果 |
|---|---|
| シート数 | 13 |
| `売上分析` の索引 | 10 |
| 共有キャッシュ | 1 個 |
| ピボットテーブル | 8 個 |
| 業種スライサー | 2 項目、8 PT 接続 |
| 商品スライサー | 8 項目、8 PT 接続 |
| タイムライン | 6 PT 接続、PT-07/08 は非接続 |
| KPI 売上 | 47,492,700 円 |
| KPI 差額 | 18,852,426 円 |
| KPI 差額率 | 0.3969541845378342 |
| KPI 前年比 | -0.011131169147547193 |
| グラフ | PG-B01 から PG-B04 の 4 個 |
| ヒートマップ | 8 商品 × 24 か月、192 セル |
| 3 状態 | 0 円は `0万円`、データなしは `－`、範囲外は空欄 |
| スパークライン | 8 本、列種別 |
| 既存計算 | 利益 72,000 円、税込見積 176,000 円 |
| メモ / `INDIRECT` | 74 件 / 24 件 |
| 配布制約 | VBA、外部リンク、保護、固定枠、クエリ、データモデルなし |

## 8. 3 状態の確認方法

| 状態 | 確認方法 | 結果 |
|---|---|---|
| 0 円 | 一時コピーで `テーブル` の 2026-01 の売上を 0 にし、PT を更新 | `N51` が `0万円` |
| データなし | 一時コピーの `AA3` に同じ判定式を置く | `－` |
| 選択範囲外 | 実際の `B51` と、`テーブル` 2025-01 の検証式 | 空欄 |
| 既存データ | `テーブル` 2026-01 の検証式 | 170 万円 |

検証式は `GETPIVOTDATA`、`ISERROR`、`COUNTIFS` を使用する。`IFERROR` は使用しない。

## 9. 既知の問題

- タイムラインの `TimelineStyleLight2` 設定は COM エラーとなり、`TimeSlicerStyleLight1` へフォールバックする。表示と接続は検証済みである。
- 手動 UI、実印刷、性能、Excel 再起動後の状態は未検証である。
- 現行データに 0 円とデータなしの実セルがないため、3 状態のうち 2 状態は一時コピーで確認している。
- 2025 年のヒートマップ列は、タイムラインを全期間へ広げるまで空欄である。
- 商品マスタは方式 A のため PT とスライサーに未接続である。商品追加時は `商品マスタ` と `SalesHistory` の両方へ登録する。
- 対応版は Microsoft 365 のみ実機確認済みである。

詳細は `known_issues.md` を参照する。

## 10. 残タスク

1. スライサーのクリック、複数選択、クリアを確認する。
2. タイムラインのドラッグ、期間変更、全期間拡張を確認する。
3. Excel を再起動し、選択状態、表示位置、印刷設定を確認する。
4. 横向き 1 ページの実印刷を確認する。
5. 低スペック端末で操作時間を計測する。
6. 商品追加の操作手順を確認する。
7. 必要に応じて Excel 2016、2019、2021、2024 で確認する。
8. 色覚多様性の観点でコントラスト比を確認する。
9. 結果を `test_results.md` と `known_issues.md` へ追記する。

## 11. 変更禁止箇所

- `prototype_v04.xlsx` を上書きしない。
- 既存の見積・原価・利益の計算式を変更しない。
- メモ 74 件と `INDIRECT` 24 件を削除しない。
- ヒートマップへ `IFERROR` を再導入しない。
- `ProductMaster` を PT またはスライサーへ接続しない。
- シート保護、セル保護、ウィンドウ枠固定を追加しない。
- VBA、外部リンク、Power Query、データモデル、DAX、リレーション、CUBE 関数を追加しない。
- 旧スライサーキャッシュ名を新しいブックへ戻さない。

## 12. コンテキスト 0% からの再開手順

1. リポジトリを取得し、プロジェクト直下へ移動する。
2. `.venv` を作成し、pywin32 を導入する。
3. `prototype_v04.xlsx` と `prototype_v05_design_b.xlsx` の存在を確認する。
4. `verify_v05_windows.py` を実行し、233 / 233 合格を確認する。
5. 不合格がある場合は、`windows_v05_results.json` の `checks` で最初の不合格を確認する。
6. 変更が必要な場合は、`build_v05_design_b.py` を修正し、v0.4 から再ビルドする。
7. 再ビルド後は必ず `verify_v05_windows.py` を再実行する。
8. 文書は `docs/variant_design_b/` を更新し、共有の作業管理文書へ結果を追記する。
9. コミットは、実装、検証、文書の論理的な区切りで分割する。

## 13. トラブルシューティング

| 症状 | 確認 | 対応 |
|---|---|---|
| ビルドが `name 'XL_TOP' is not defined` で停止 | COM 定数の定義 | `XL_TOP = -4160` を確認 |
| タイムラインのスタイル設定エラー | ビルドログ | `TimeSlicerStyleLight1` のフォールバックを確認 |
| スライサーのスタイル比較が不合格 | COM の `Style` がオブジェクト | `str()` で比較 |
| データなしの塗りが 0 円と同じ | 条件付き書式のテキストルール | サンド塗りを確認 |
| 2025 年の列が空欄 | タイムライン初期値 | 仕様どおり。全期間へ広げる |
| 商品マスタの追加が反映されない | PT 接続 | 方式 A の制約。売上データにも登録する |
| 出力ファイルが開けない | Excel のロック | 出力先を閉じる |
| 検証が途中で停止 | 非表示の Excel プロセス | 自前の非表示プロセスだけ終了し、再実行する |

## 14. 完了条件

- `verify_v05_windows.py` が 233 / 233 合格する。
- `docs/variant_design_b/` の 5 文書が存在する。
- `docs/test_spec.md` 第18節が合格へ更新されている。
- `docs/progress.md` の T-21 が完了である。
- `docs/handover.md` に v0.5 の節がある。
- `README.md` に v0.5 の使い方とビルド・検証がある。
- 公開禁止情報が 0 件である。
- v0.4 のブックと仕様文書が変更されていない。
