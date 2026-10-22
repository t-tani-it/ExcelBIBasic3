# v0.6 ビルド・検証クックブック（ExcelBIBasic_DeepSeek_V4.1）

- 版: RevC
- 作成日: 2026-10-22
- 対象: v0.6 デザイン C のビルド担当
- 目的: v0.5 のビルド・検証資産を再利用し、COM の罠と再作業を減らす
- 改訂内容: v0.6 専用データ生成、259 件検証、年別ヒートマップ、コメント・グラフ修正を追記

## 1. 環境

| 項目 | 前提 |
|---|---|
| OS | Windows |
| Excel | Microsoft 365（16.0.20430.20182、x64）で実機確認済み |
| Python | 3.11.15 |
| 仮想環境 | プロジェクト直下の `.venv` |
| COM | pywin32 312 |
| シェル | PowerShell を基本とする |

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -c "import win32com.client; print('pywin32 OK')"
```

`requirements.txt` は pywin32 のバージョンを固定する。Excel COM を使うため、Excel がインストールされ、対話セッションで起動できる必要がある。

## 2. 最初に実行する回帰確認

v0.6 を触る前に v0.5 の 233 / 233 を確認する。出力は一時ファイルへ出し、追跡済みの `windows_v05_results.json` を不要に上書きしない。

```powershell
.\.venv\Scripts\python.exe verify_v05_windows.py prototype_v05_design_b.xlsx $env:TEMP\windows_v05_results_check.json
```

不合格なら、Excel の残存プロセス、ブックのロック、pywin32 の導入を確認する。

## 3. v0.6 の成果物名

| 用途 | 名前 |
|---|---|
| ビルドスクリプト | `build_v06_design_c.py` |
| 検証スクリプト | `verify_v06_windows.py` |
| データ生成スクリプト | `generate_sales_fixture_v06.py` |
| データ fixture | `sales_fixture_v06.json` |
| ブック | `prototype_v06_design_c.xlsx` |
| 検証結果 | `windows_v06_results.json` |
| 設計書 | `docs/variant_design_c/design_spec.md` |
| 画面画像 | `docs/variant_design_c/preview.png` |

ビルドの入力は `prototype_v04.xlsx` とする。v0.5 は視覚参照であり、入力にしない。

### 3.1 v0.6 の実行コマンド

PowerShell で Luna 作業ルートから実行する。

```powershell
.\.venv\Scripts\python.exe generate_sales_fixture_v06.py
.\.venv\Scripts\python.exe build_v06_design_c.py prototype_v04.xlsx prototype_v06_design_c_candidate.xlsx
.\.venv\Scripts\python.exe verify_v06_windows.py prototype_v06_design_c_candidate.xlsx windows_v06_candidate_results.json
```

- `build_v06_design_c.py` は既存の出力ブックを上書きせず停止する。候補を検証した後に正式名へ反映する。
- v0.5 の事前回帰は、追跡済み JSON ではなく `%TEMP%` の一時出力先へ保存する。
- v0.6 検証の合格は 233 件以上であり、現在の実装は 259 / 259 である。既存 233 件の検査を削除してはならない。
- 共通 `sales_fixture.json` は変更しない。v0.6 専用 fixture は 576 行、2026 年税抜売上は 2025 年比約 8% 増である。

## 4. ビルドスクリプトの再利用マップ

`build_v05_design_b.py` は約 1,500 行である。全部を読み直さず、次の範囲を関数単位で再利用する。

| 行範囲の目安 | 内容 | v0.6 での扱い |
|---|---|---|
| 129〜243 | `log`、`rgb`、`column_letter`、`count_notes`、`count_indirect`、`close_to` | 再利用 |
| 244〜390 | `validate_input`、`validate_v04`、`delete_old_dashboard` | 再利用 |
| 391〜566 | `style_dashboard_shell` | デザイン差し替え |
| 567〜764 | `create_kpi_card`、`create_kpi_cards`、`create_insight_panel` | デザイン差し替え |
| 765〜902 | `connect_slicer`、`create_slicers` | 再利用。位置とサイズだけ変更 |
| 903〜1164 | `create_chart`、`style_chart`、`create_charts` | デザイン差し替え |
| 1165〜1310 | `create_heatmap` | 数式ロジックは再利用。書式と配置は差し替え |
| 1311〜1504 | `finalize_state`、`execute_logic`、`main` | 再利用。作成者名を上書き |

### 4.1 触ってよい関数

- `style_dashboard_shell`
- `create_kpi_card`
- `create_kpi_cards`
- `create_insight_panel`
- `create_chart`
- `style_chart`
- `create_charts`
- `create_heatmap` の書式部分

### 4.2 触ると壊れやすい関数

- `validate_v04` の不変条件
- `delete_old_dashboard` の旧スライサーキャッシュ削除
- `connect_slicer` の接続対象
- `create_slicers` のタイムライン接続
- `create_heatmap` の `ISERROR` / `COUNTIFS` ロジック
- `finalize_state` の保存・後処理
- `main` の引数と終了コード

## 5. 検証スクリプトの再利用マップ

`verify_v05_windows.py` は約 870 行である。共通検査を再利用し、v0.6 固有の期待値だけ差し替える。

| 行範囲の目安 | 内容 | v0.6 での扱い |
|---|---|---|
| 106〜360 | `rgb`、`check`、`check_close`、`check_true`、カウント、検索、ヒートマップ位置 | 再利用 |
| 361〜399 | `scan_package` | 再利用 |
| 400〜507 | `test_zero_state`、`test_no_data_state` | 再利用 |
| 508〜800 | `run_workbook_checks` | v0.6 用に期待値を差し替え |
| 801〜870 | `json_default`、`main` | 再利用 |

v0.5 の検証をそのまま v0.6 に実行しない。色、セル範囲、グラフ名、KPI の期待値は v0.6 用に更新する。

## 6. COM の罠

| 症状 | 原因 | 対応 |
|---|---|---|
| `name 'XL_TOP' is not defined` | COM 定数の未定義 | `XL_TOP = -4160` を定義する |
| タイムラインのスタイル設定で例外 | `TimelineStyleLight2` が COM エラー | `TimeSlicerStyleLight1` へフォールバックする |
| スライサーのスタイル比較が不合格 | `Style` がオブジェクト | `str()` で比較する |
| 旧スライサーが残る | キャッシュ削除の順序 | ダッシュボード削除前に旧キャッシュを消す |
| グラフにフィールドボタンが出る | 既定値 | `ShowAllFieldButtons = False` |
| 出力ファイルが開けない | Excel のロック | 出力先を閉じる。残存プロセスを確認する |
| ビルドが途中で止まる | 非表示 Excel の残存 | 自前の非表示プロセスだけ終了して再実行する |
| 色が保存されない | ピボットグラフの再描画 | 保存前に再適用し、検証で確認する |

### 6.1 基本の COM 手順

1. `pythoncom.CoInitialize()` を呼ぶ。
2. `win32com.client.DispatchEx("Excel.Application")` で専用インスタンスを作る。
3. `Visible = False`、`DisplayAlerts = False` を設定する。
4. ビルド中は `Calculation = xlCalculationManual` とし、保存前に自動へ戻して再計算する。
5. `try/finally` でブックを閉じ、Excel を終了する。
6. 例外時も自前の Excel プロセスだけを終了する。

### 6.2 避ける操作

- `Select` と `Activate` の多用。
- クリップボード経由の大量コピー。
- `UsedRange` 全体への一括書式設定。
- 複数 Excel インスタンスの同時ビルド。
- 保護ビューでのブックオープン。
- 画面更新を有効にしたままの大量セル書き込み。

## 7. OOXML 後処理

v0.6 でも `finalize_v03_ooxml.sanitize_package` を再利用する。

- `x15ac:absPath` を除去する。
- 作成者メタデータを `ExcelBIBasic_DeepSeek_V4.1` に上書きする。
- 既存の `finalize_v03_ooxml.CONFIG` を直接書き換えない。v0.6 ビルド側でコピーして上書きする。
- v0.6 の検証に `docProps/core.xml` の作成者確認を追加する。
- ローカルパス、API キー、個人 ID が 0 件であることを確認する。

## 8. 基準値

| 項目 | 値 |
|---|---|
| シート数 | 13 |
| `売上分析` の索引 | 10 |
| 共有キャッシュ | 1 |
| ピボットテーブル | 8 |
| 業種スライサー | 2 項目、8 PT 接続 |
| 商品スライサー | 8 項目、8 PT 接続 |
| タイムライン | 6 PT 接続、PT-07 / PT-08 は非接続 |
| 2025 年売上 / 原価 | 48,027,300 円 / 28,999,224 円 |
| 2026 年売上 / 原価 | 51,869,100 円 / 31,729,100 円 |
| KPI 売上 | 51,869,100 円 |
| KPI 差額 | 20,140,000 円 |
| KPI 差額率 | 0.38828512544077304 |
| KPI 前年比 | 0.07999200454741362 |
| 2026 年売上成長 | 2025 年比 7.9992% |
| ヒートマップ | 8 商品 × 24 か月、192 セル |
| 年別ヒートマップ | 2025のみ / 2026のみ / 両年の選択で追従 |
| コメント図形 | 74 件、四角形、自動サイズ |
| グラフ範囲 | 4 個すべて Q 列内 |
| スパークライン | 8 本、列種別 |
| 既存計算 | 利益 72,000 円、税込見積 176,000 円 |
| メモ | 74 件 |
| `INDIRECT` | 24 件 |
| v0.5 自動検証 | 233 / 233 |

## 9. トラブルシューティング

| 症状 | 確認 | 対応 |
|---|---|---|
| v0.5 の 233 件が落ちる | Excel の版、pywin32、ブックのロック | 環境を直してから v0.6 へ進む |
| スライサーが 8 PT に接続されない | `connect_slicer` の `skip` 引数 | PT-07 / PT-08 の扱いを確認する |
| タイムラインが前年比較に連動する | 接続対象 | PT-07 / PT-08 を除外する |
| ヒートマップの 3 状態が崩れる | 数式と条件付き書式の優先度 | `ISERROR`、`COUNTIFS`、`StopIfTrue` を確認する |
| 0 円が空欄になる | 数値 0 と文字列の扱い | `0"万円"` と条件付き書式を確認する |
| データなしが空欄になる | `COUNTIFS` の列名 | `SalesHistory[商品]`、`SalesHistory[年月]` を確認する |
| 2025 年のヒートマップが空欄 | 初期タイムラインは2026年 | タイムラインで2025年または全期間を選択し、年別テストを実行する |
| コメント図形で文字切れ | 図形種別、自動サイズ | `normalize_comment_shapes`、矩形設定、本文一致検査を確認する |
| グラフがQ列を超える | ヒートマップが列幅を変えた後の位置 | `fit_charts_to_grid` を実行し、4形状の右端チェックを確認する |
| 印刷で文字が切れる | 用紙、余白、倍率 | 横向き 1 ページ、改ページプレビューで確認する |
| ファイルが重い | 画像、グラフ、条件付き書式の範囲 | 画像を減らし、適用範囲を 192 セルに限定する |

## 10. 公開情報チェック

push 前に次を確認する。

```powershell
git status --short
git diff --cached
```

- ローカルパス、API キー、個人 ID、トークン、認証情報がない。
- 参考画像そのものをコミットしていない。
- 外部リンクをブック内に埋め込んでいない。
- ライセンス表記が必要な素材は、設計書またはブック内のテキスト注記にある。
- 問題がなければ「公開禁止情報は含まれていません」と表示する。

## 11. コミット分割

1. ビルドスクリプトとブック
2. 検証スクリプトと結果 JSON
3. 設計書、既知の問題、変更履歴、引き継ぎ
4. 共有文書の更新

1 コミットにまとめず、`git log --oneline` で履歴が読める粒度にする。
