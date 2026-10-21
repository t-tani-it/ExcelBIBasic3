# ExcelBIBasic3 - プロジェクト永続ルール（AGENTS.md）

## 概要
- 本プロジェクトは、ExcelBIBasic2 の試作（SUMIFS 版）を土台に、見積・原価・利益シミュレーターを Power Query／データモデル／ピボットテーブル／ピボットグラフ／スライサー／タイムライン構成へ再設計するものである。
- 判断と作業の基準は `docs/` 配下の仕様書とする。仕様変更は `docs/decision_log.md` への追記と `docs/progress.md` の更新を伴う。
- 文書の原本は不変とし、改訂は複写による Rev 進行とする。

## 対象と成果物
- 対象ブック: `prototype.xlsx`（試作、読み取り専用の調査対象）
- 生成物: 購入者向け `.xlsx`（マクロなし、外部リンクなし）
- 仕様書: `docs/requirements.md`、`docs/design_spec.md`、`docs/calculation_spec.md`、`docs/test_spec.md`、`docs/implementation_plan.md`、`docs/compatibility_matrix.md`、`docs/decision_log.md`
- 作業管理: `docs/progress.md`（タスクと状態）、`docs/handover.md`（フェーズ引き継ぎ）
- 再ビルド工程: `build.mjs`、`enhance.mjs`、`finalize_ooxml.py`
- 検証: `verify_excel_windows.py`、`verify_v02_windows.py`

## 開発環境
- OS は Windows を前提とする。
- Node.js で OOXML を直接生成し、Python で後処理と実機検証を行う。
- 仮想環境は venv を標準とする。他案件の conda env を流用しない。
- venv はプロジェクト直下の `.venv` に統一する（`python -m venv .venv`）。
- 実機検証には pywin32 が必要である。導入先は本プロジェクトの `.venv` とする。
- `.venv/`、`venv/`、`__pycache__/`、`~$*` は `.gitignore` で Git 除外とする。
- 参照環境: Microsoft 365（16.0.20430.20182、x64）。対応版の範囲は `docs/compatibility_matrix.md` を参照する。

## 不変条件（変更禁止）
- メモは既存 74 件を維持する。削除禁止。追加は新規要件に限り許容する。
- `INDIRECT` は 24 件を維持する。内訳は計算シート B5・B6・B9・B11 と D16:D25・E16:E25 である。
- シート保護、セル保護、ウィンドウ枠固定は行わない。
- VBA、ActiveX、外部リンク、マクロは使用しない。
- 保存値は円とする。万円は表示変換に限る。
- 見積・原価・利益の計算セル（計算、採算シミュレーション、見積入力、原価(木工所)、原価(農家)）は計算仕様に従い維持する。

## 仕様上の確定事項
- 売上分析は Power Query とデータモデルで作り直す。SUMIFS 版は新ブックへ引き継がない。
- 商品マスタは別テーブル化する（追加を見越す）。
- スライサー（業種・商品）とタイムライン（日付）を全ピボットテーブルと全ピボットグラフへ接続する。ゲージ用ピボットのみ独立とする。
- 表示単位は万円（例 `84万円`）とする。実装は `ROUND(円/10000,0)` と書式 `0"万円"` である。
- 状態表示は採算シミュレーション D12 のみとする。見積入力 D12:F12 の状態表示は削除し、B12:F12 を備考入力へ復帰させる。
- チェックボックスはフォームコントロールとする（セル内チェックボックスは Microsoft 365 限定のため不採用）。
- ゲージはデータバーで実装する。評価期間は別入力とし、スライサー非連動である旨を画面へ明記する。
- 返品、複数税率、年度締めは対象外とする。

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
- 作業ディレクトリは `ExcelBIBasic3` とする。
- 再ビルドは次の順で行う。詳細な引数は `README.md` を参照する。
  - `node build.mjs`
  - `node enhance.mjs`
  - `python finalize_ooxml.py`
- 実機検証は次の順で行う。
  - `python verify_excel_windows.py`
  - `python verify_v02_windows.py`
- 検証は `package_checks.json`、`engine_tests.json`、`notes_manifest.json`、`sales_fixture.json` の期待値と照合する。
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
