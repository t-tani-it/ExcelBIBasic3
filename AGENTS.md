# ExcelBIBasic_DeepSeek_V4.1 - プロジェクト永続ルール（AGENTS.md）

- 旧称: `ExcelBIBasic3`（2026-10-22 に改名）
- 改名の詳細は `docs/project_rename.md` を参照する。過去の RevA 文書と v0.4 / v0.5 の旧名は歴史記録として残す。

## 概要
- 本プロジェクトは、ExcelBIBasic2 の試作（SUMIFS 版）を土台に、見積・原価・利益シミュレーターを再設計するものである。
- v0.4 の売上分析は D-24 の方式 A（通常ピボット、共有キャッシュ、スライサー、タイムライン、ピボットグラフ）で実装する。当初の Power Query・データモデル方式は D-22 により中止する。
- 判断と作業の基準は `docs/` 配下の仕様書とする。仕様変更は `docs/decision_log.md` への追記と `docs/progress.md` の更新を伴う。
- 文書の原本は不変とし、改訂は複写による Rev 進行とする。

## 対象と成果物
- 対象ブック: `prototype.xlsx`（試作、読み取り専用の調査対象）
- 生成物: 購入者向け `.xlsx`（マクロなし、外部リンクなし）
- 現行仕様: `docs/document_index.md`、`docs/v0.4_method_a_addendum.md`、`docs/decision_log.md`
- 別成果物の要求: `docs/variant_requirements.md`。v0.4 の仕様を変更しない
- v0.5 成果物: `prototype_v05_design_b.xlsx`。設計と検証は `docs/variant_design_b/` を参照する
- v0.6 設計: `docs/variant_design_c/design_spec.md`。次モデル向け入口は `docs/next_model/START_HERE.md` である
- 履歴仕様: `docs/requirements.md`、`docs/design_spec.md`、`docs/calculation_spec.md`、`docs/implementation_plan.md`、`docs/compatibility_matrix.md`。RevA の旧方式であり、現行仕様として扱わない
- テスト: `docs/test_spec.md`。第17節を v0.4 の正とし、第18節をバリアントの受入とする
- 作業管理: `docs/progress.md`（タスクと状態）、`docs/handover.md`（フェーズ引き継ぎ）
- v0.4 再ビルド工程: `build_v04_a.py`、`finalize_v03_ooxml.py`
- v0.5 再ビルド工程: `build_v05_design_b.py`
- v0.6 再ビルド工程（予定）: `build_v06_design_c.py`、`verify_v06_windows.py`
- v0.2 以前の再ビルド工程（履歴）: `build.mjs`、`enhance.mjs`、`finalize_ooxml.py`
- 検証: `verify_v05_windows.py`、`verify_v04_windows.py`、`verify_v03_windows.py`、`verify_v02_windows.py`、`verify_excel_windows.py`

## 開発環境
- OS は Windows を前提とする。
- v0.4 は Windows 版 Excel の COM（pywin32）で生成し、Python で OOXML 後処理と実機検証を行う。v0.2 以前の Node.js 工程は履歴として残す。
- 仮想環境は venv を標準とする。他案件の conda env を流用しない。
- venv はプロジェクト直下の `.venv` に統一する（`python -m venv .venv`）。
- 実機検証には pywin32 が必要である。導入先は本プロジェクトの `.venv` とする。
- `.venv/`、`venv/`、`__pycache__/`、`~$*` は `.gitignore` で Git 除外とする。
- 参照環境: Microsoft 365（16.0.20430.20182、x64）。対応版の範囲は `docs/v0.4_method_a_addendum.md` 第12節を参照する。旧 `docs/compatibility_matrix.md` は履歴である。

## 不変条件（変更禁止）
- メモは既存 74 件を維持する。削除禁止。追加は新規要件に限り許容する。
- `INDIRECT` は 24 件を維持する。内訳は計算シート B5・B6・B9・B11 と D16:D25・E16:E25 である。
- シート保護、セル保護、ウィンドウ枠固定は行わない。
- VBA、ActiveX、外部リンク、マクロは使用しない。
- 保存値は円とする。万円は表示変換に限る。
- 見積・原価・利益の計算セル（計算、採算シミュレーション、見積入力、原価(木工所)、原価(農家)）は計算仕様に従い維持する。

## 仕様上の確定事項
- 売上分析は D-24 の方式 A で作り直す。Power Query、データモデル、DAX、リレーション、CUBE 関数は使用しない。SUMIFS 版は新ブックへ引き継がない。
- 商品マスタは別テーブル化する（追加を見越す）。方式 A ではピボット・スライサーに未接続とする。
- 業種・商品スライサーは全ピボットテーブルと全ピボットグラフへ接続する。日付タイムラインは PT-01 から PT-06 へ接続し、前年参照の PT-07 と PT-08 は非連動とする。ゲージ用ピボットのみ独立とする。
- KPI は円、グラフとヒートマップは万円（例 `84万円`）で表示する。実装は `ROUND(円/10000,0)` と書式 `0"万円"` である。
- 状態表示は採算シミュレーション D12 のみとする。見積入力 D12:F12 の状態表示は削除し、B12:F12 を備考入力へ復帰させる。
- チェックボックスは P4 で実装する場合、フォームコントロールとする（セル内チェックボックスは Microsoft 365 限定のため不採用）。v0.4 では未実装である。
- ゲージは P4 で実装する場合、データバー方式とする。評価期間は別入力とし、スライサー非連動である旨を画面へ明記する。v0.4 では未実装である。
- 返品、複数税率、年度締めは対象外とする。
- v0.4 のヒートマップは `IFERROR` で空欄化しており、0 円、データなし、選択範囲外の 3 状態を完全には区別していない。バリアントでは 3 状態の区別を必須とする。

## デザインバリアント（別セッション・別モデル）

- 別セッション・別モデルは `docs/variant_requirements.md` に従う。
- `prototype_v04.xlsx` を複製し、既存の見積・原価・利益シートは維持する。
- `売上分析` シートは削除し、同名のシートを 0 から再作成する。旧セル、旧書式、旧グラフ、旧スパークライン、旧条件付き書式を継承しない。
- 優れた UI の採用を許容する。固定要件を壊さず、採用元、採用理由、ライセンスを設計書へ記録する。
- v0.6 以降は外部のデザイン、テンプレート、配色、アイコン、グラフ配置を積極的に参照する。参照元は 3〜5 件とし、`docs/next_model/ui_reference_log.md` へ URL、ライセンス、採用要素、改変内容を記録する（D-28）。
- 外部 UI の良い要素は、Excel のセル、図形、グラフ、スライサー、条件付き書式で再現する。ブック内に外部リンク、外部画像、外部フォント、外部アイコンフォントを埋め込まない。
- ライセンス不明の素材、有料テンプレート、ロゴは使用しない。参考画像そのものを public リポジトリへ保存しない。
- デザイン差は、配色、KPI カードの構成、グラフ構成のうち 2 つ以上とする。
- v0.4 のブックと文書を上書きしない。成果物は別ファイル・別文書とする。
- 文書の優先順位は `docs/document_index.md` に従う。
- バリアントの受入は `docs/test_spec.md` 第18節に従う。
- v0.5 デザインバリアント B は 233 / 233 合格である。手動 UI、実印刷、性能は未検証である。
- v0.5 のヒートマップは `IFERROR` を使用しない。`ISERROR` と `COUNTIFS` で 3 状態を区別する。
- v0.5 のタイムラインは `TimeSlicerStyleLight1` を使用する。`TimelineStyleLight2` の設定は COM エラーとなる。
- v0.6 は `prototype_v04.xlsx` を出発点とし、`prototype_v06_design_c.xlsx` を別ファイルとして作成する。v0.4 / v0.5 のブックと文書を上書きしない。
- v0.6 のデザイン方向は Warm Slate Editorial とし、文字階層、余白、グラフ適合性、印刷・性能の証拠を v0.5 より改善する。
- v0.6 のデザイン合格は 20 / 24 以上、重大項目に 0 がないこととする。機能の合格は 233 / 233 である。

## コーディング規約（スクリプト）
- 各ファイルの冒頭に、目的、仕様、前提、入出力、注意点、拡張ポイント、変更履歴をコメントで記載する。
- 設定値はコード冒頭にまとめ、`config` 相当の定数または辞書として定義する。意味をコメントで説明する。
- 関数は単一責務とする。`main()`、`validate_input()`、`execute_logic()`、`format_output()` を基本構成とする。
- 型ヒントを必須とする。docstring には要約、引数、戻り値、例外、副作用、使用例を記載する。
- 出力は決定的とする。日時や乱数に依存する値を成果物へ埋め込まない。
- 既存構造を尊重し、生成物のシート名・セル番地・書式は仕様書の記載と一致させる。

## セキュリティ・公開禁止情報
- ローカルパス、API キー、個人 ID、トークン、認証情報をコード、設定、ログ、成果物へ含めない。
- GitHub へ push する前に公開禁止情報の有無を確認する。問題がなければ「公開禁止情報は含まれていません」と表示する。

## GitHub 運用ルール
- リモートは HTTPS 方式、認証は gh CLI とする。
- リポジトリは public とする。作成例は `gh repo create <name> --public --source . --remote origin` である。
- ブランチは `main` とし、リモートを追跡する。
- コミットは論理的な区切りで分割する（骨格、仕様書、作業管理など）。
- push 前に `git status` と `git diff --cached` で禁止情報を確認する。

## ビルド・実行・検証
- 作業ディレクトリは `ExcelBIBasic_DeepSeek_V4.1` とする。
- v0.4 の再ビルドは次のコマンドで行う。詳細な引数は `README.md` を参照する。
  - `.\.venv\Scripts\python.exe build_v04_a.py prototype_v03.xlsx prototype_v04.xlsx`
- v0.4 の実機検証は次のコマンドで行う。
  - `.\.venv\Scripts\python.exe verify_v04_windows.py prototype_v04.xlsx`
- v0.5 の再ビルドは次のコマンドで行う。
  - `.\.venv\Scripts\python.exe build_v05_design_b.py prototype_v04.xlsx prototype_v05_design_b.xlsx`
- v0.5 の実機検証は次のコマンドで行う。
  - `.\.venv\Scripts\python.exe verify_v05_windows.py prototype_v05_design_b.xlsx windows_v05_results.json`
- v0.6 の再ビルドと検証は、次モデルが作成する `build_v06_design_c.py` と `verify_v06_windows.py` で行う。コマンドは `docs/next_model/build_cookbook.md` に追記する。
- v0.2 以前の Node 工程は履歴であり、v0.4 と v0.5 の再生成には使用しない。
- 検証は `package_checks.json`、`engine_tests.json`、`notes_manifest.json`、`sales_fixture.json` の期待値と照合する。v0.4 は `windows_v04_results.json`、v0.5 は `windows_v05_results.json` を正とする。
- 実機検証が未実施の範囲は「実機保証なし」と明示する。

## README記載ルール
- README の「セットアップ」と「起動」には、PowerShell とコマンドプロンプト（cmd）の両方の仮想環境有効化手順を記載する。
- 各コマンドの対象シェルを明記する。
  - PowerShell: `.\.venv\Scripts\Activate.ps1`
  - cmd: `.\.venv\Scripts\activate.bat`

## ドキュメント規約
- 文体は「である調」とする。表内のセルは述語なしとする。
- 丸付き数字は使用しない。通常の数字で記載する。
- 単位は万円表示と円保存を区別して記載する。
- 事実、推論、未確認を区別し、未確認事項は「確認が必要」と明示する。

## フェーズ引き継ぎ
- フェーズ完了時に `docs/handover.md` を更新する。
- 記載内容は、方針（計画）、実装（実行）、結果、残タスク、実行コマンドとする。
- `docs/progress.md` のタスク状態を同時に更新する。

## 変更履歴
- 2026-10-21: 初版作成（仕様策定フェーズの確定事項を反映）。
- 2026-10-22: D-24 の方式 A、PT-07/08 のタイムライン非連動、v0.4 のビルド・検証コマンドを反映。
- 2026-10-22: D-25 のデザインバリアント要求、D-26 のヒートマップ相違、文書索引と追補を反映。
- 2026-10-22: D-27 の v0.5 デザインバリアント B、233 / 233 合格、ビルド・検証コマンドを反映。
- 2026-10-22: プロジェクト名を `ExcelBIBasic_DeepSeek_V4.1` へ変更。入口文書のみ新名とし、旧名は歴史記録として残す（D-29）。
- 2026-10-22: D-28 の外部 UI 参照方針、v0.6 デザイン C の成果物名、233 / 233 とデザイン 20 / 24 の受入を反映。
