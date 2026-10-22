# 次モデル入口（ExcelBIBasic_DeepSeek_V4.1）

- 版: RevA
- 作成日: 2026-10-22
- 対象: v0.6 デザイン C を担当する別セッション・別モデル
- 旧名: `ExcelBIBasic3`（過去文書には旧名が残る）

## 1. 最初に読むもの

このファイルだけで作業の全体像が分かる。詳細は必要なときだけ参照する。

1. 本ファイル
2. `docs/next_model/design_brief.md`
3. `docs/next_model/design_tokens.json`
4. `docs/next_model/acceptance_matrix.md`
5. ビルド時に `docs/next_model/build_cookbook.md`
6. デザイン方向の詳細として `docs/variant_design_c/design_spec.md`

次の文書は、必要になった場合だけ読む。

- `docs/variant_requirements.md`（固定要件の原典）
- `docs/v0.4_method_a_addendum.md`（方式 A の現行仕様）
- `docs/variant_design_b/design_spec.md`（v0.5 の設計）
- `docs/variant_design_b/known_issues.md`（v0.5 の未検証範囲）

`docs/handover.md`、`docs/progress.md`、`docs/decision_log.md` の全履歴は読まない。必要な決定は D-22、D-24、D-25、D-26、D-27、D-28、D-29 だけである。

## 2. 目的

`prototype_v04.xlsx` を出発点に、`売上分析` シートだけを 0 から再作成し、v0.5 と同等以上のモダンデザインを持つ v0.6 デザイン C を作る。

- 成果物: `prototype_v06_design_c.xlsx`
- ビルド: `build_v06_design_c.py`
- 検証: `verify_v06_windows.py`
- 検証結果: `windows_v06_results.json`
- 設計・検証文書: `docs/variant_design_c/`
- 画面画像: `docs/variant_design_c/preview.png`

v0.4 と v0.5 のブック・文書は上書きしない。

## 3. 出発点と参照

| 用途 | ファイル |
|---|---|
| 出発点 | `prototype_v04.xlsx` |
| 視覚参照 | `prototype_v05_design_b.xlsx`、`docs/next_model/reference_v05.png` |
| ビルドのベース | `build_v05_design_b.py` |
| 検証のベース | `verify_v05_windows.py` |
| 回帰基準 | `windows_v05_results.json` |
| データ理解 | `sales_fixture.json` |

次は歴史資産であり、v0.6 のベースにしない。

- `build.mjs`、`enhance.mjs`、`finalize_ooxml.py`
- `build_v03_p5.py`、`build_v04_a.py`
- `prototype.xlsx`、`prototype_v03.xlsx`
- `verify_v02_windows.py`、`verify_v03_windows.py`、`verify_v04_windows.py`
- `opencode_prompt.md`

## 4. 最初に確認するコマンド

PowerShell でプロジェクト直下から実行する。

```powershell
.\.venv\Scripts\python.exe --version
.\.venv\Scripts\python.exe -c "import win32com.client; print('pywin32 OK')"
.\.venv\Scripts\python.exe verify_v05_windows.py prototype_v05_design_b.xlsx $env:TEMP\windows_v05_results_check.json
```

v0.5 が 233 / 233 合格することを確認してから v0.6 を触る。不合格なら環境かブックが壊れているため、作業を止めて `docs/next_model/build_cookbook.md` のトラブルシューティングを見る。

## 5. 変更してよいもの

- `売上分析` シートのレイアウト、配色、フォントサイズ、余白
- KPI カードの形、並び、注記
- グラフの種類、配置、 series の見せ方
- スライサー、タイムラインの配置とスタイル
- ヒートマップの色、罫線、凡例
- スパークラインの見せ方
- 注記、操作説明の見せ方
- 外部 UI の良い要素を Excel で再現した装飾

## 6. 変更してはいけないもの

- 方式 A（通常 PT、共有キャッシュ 1 個、スライサー、タイムライン、ピボットグラフ）
- Power Query、データモデル、DAX、リレーション、CUBE 関数
- VBA、ActiveX、マクロ、外部リンク、外部接続
- シート保護、セル保護、ウィンドウ枠固定
- 既存の見積・原価・利益の計算式
- メモ 74 件、`INDIRECT` 24 件
- `ProductMaster` の PT / スライサー接続
- ヒートマップの `IFERROR` 再導入
- `TimelineStyleLight2` の再試行
- `prototype_v04.xlsx` と `prototype_v05_design_b.xlsx` の上書き
- 旧スライサーキャッシュ名 `SL_業種`、`SL_商品`、`TL_日付` の再利用

## 7. デザイン目標

- v0.5 の強み（左フィルタ帯、2 × 2 KPI、3 状態ヒートマップ、データ注記）を維持する。
- v0.5 の弱み（文字階層が小さい、2 分類の円グラフ、エリア + 折れ線の混雑、印刷・性能の証拠不足）を改善する。
- 外部 UI は 3〜5 件まで参照する。良い要素は Excel のセル、図形、グラフ、スライサー、条件付き書式で再現する。
- 参考画像の貼り付け、外部リンク、外部フォント、外部アイコンフォントの埋め込みは行わない。
- デザイン点数は 20 / 24 以上、重大項目に 0 がないこと。

## 8. 時間予算と停止条件

- 参考探し: 20 分
- デザイン決定: 30 分
- ビルド: 45 分
- 検証・目視: 25 分
- 合計: 約 2 時間
- 再ビルドは最大 2 回
- 233 / 233 とデザイン 20 / 24 に達したら、それ以上の微調整をしない
- VBA なしで難しい動的フィルタ要約は追わない

## 9. 完了条件

- `verify_v06_windows.py` が 233 / 233 合格する。
- デザイン点数が 20 / 24 以上で、重大項目に 0 がない。
- 手動確認の未実施範囲を `known_issues.md` に明記する。
- 画面画像 `docs/variant_design_c/preview.png` がある。
- 外部 UI の採用記録とライセンスが `ui_reference_log.md` と設計書にある。
- ブック内に外部リンク、ローカルパス、API キー、個人 ID がない。
- v0.4 / v0.5 のブックと文書を変更していない。
- 共有文書（`docs/progress.md`、`docs/handover.md`、`docs/document_index.md`、`docs/decision_log.md`）を更新している。

## 10. 10 問セルフチェック

作業開始前に、次の 10 問へ答えられること。

1. どのファイルを出発点にするか。
2. どのファイルを視覚参照にするか。
3. どのシートを削除して再作成するか。
4. 8 個の PT のうち、タイムライン非連動はどれか。
5. ヒートマップの 3 状態は何か。
6. 変更してはいけない既存計算は何か。
7. 233 / 233 の検証コマンドは何か。
8. デザイン合格点は何点か。
9. 外部 UI を使う場合に記録する項目は何か。
10. v0.4 / v0.5 を上書きしない理由は何か。

答えられない項目があれば、該当する文書だけを読む。全部を最初から読み直さない。
