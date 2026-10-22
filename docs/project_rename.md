# プロジェクト改名メモ（ExcelBIBasic_DeepSeek_V4.1）

- 版: RevA
- 作成日: 2026-10-22
- 旧名: `ExcelBIBasic3`
- 新名: `ExcelBIBasic_DeepSeek_V4.1`
- 状態: 改名実施済み

## 1. 目的

プロジェクト名を `ExcelBIBasic3` から `ExcelBIBasic_DeepSeek_V4.1` へ変更する。改名の影響を入口文書と今後の成果物に限定し、検証済みの v0.4 / v0.5 資産と過去の設計記録を壊さないことを目的とする。

## 2. 変更対象

| 対象 | 内容 |
|---|---|
| ローカルフォルダ | `ExcelBIBasic_DeepSeek_V4.1` へ変更 |
| GitHub リポジトリ | `t-tani-it/ExcelBIBasic_DeepSeek_V4.1` へ変更 |
| `AGENTS.md` | タイトルと作業ディレクトリを新名へ変更。外部 UI 参照方針を追記 |
| `README.md` | 新名、リポジトリ URL、次モデル入口への導線を追記 |
| `docs/document_index.md` | 新名へ変更。改名メモと次モデル文書を索引へ追加 |
| `docs/next_model/START_HERE.md` | 新名で新規作成 |
| `docs/project_rename.md` | 本メモを新規作成 |
| 新規 v0.6 文書 | 新名で作成 |

## 3. 変更しない対象

次の対象は歴史記録または検証済み資産として旧名のまま残す。一括置換は行わない。

- 過去の RevA 文書（`docs/requirements.md`、`docs/design_spec.md`、`docs/implementation_plan.md`、`docs/calculation_spec.md`、`docs/compatibility_matrix.md`）
- `docs/variant_design_b/` の v0.5 設計・検証・引き継ぎ文書
- `prototype_v04.xlsx` と `prototype_v05_design_b.xlsx` の作成者メタデータ
- `verify_v03_windows.py` と `verify_v04_windows.py` の作成者期待値 `ExcelBIBasic3`
- `finalize_v03_ooxml.py` の作成者設定
- `build_v03_p5.py` などの過去スクリプトのコメント
- Git のコミット履歴

## 4. 理由

- 一括置換は「原本は不変」の文書規約に反する。
- v0.3 / v0.4 の検証は作成者 `ExcelBIBasic3` を期待しており、変更すると過去資産の再検証が壊れる。
- 過去の文書に残る旧名は、プロジェクトの履歴として意味がある。
- Git 履歴の書き換えは、コミットハッシュと公開リンクを壊すため行わない。

## 5. テストへの影響

テストファイルの中身は変更しない。テストは `prototype_v04.xlsx` などの相対パスを参照しており、プロジェクトフォルダ名に依存しない。

| 項目 | 影響 | 対応 |
|---|---|---|
| `verify_v03_windows.py` | なし | 作成者を変更しないため |
| `verify_v04_windows.py` | なし | 作成者を変更しないため |
| `verify_v05_windows.py` | なし | 作成者メタデータを検査していない |
| `.venv` | あり | フォルダ改名後に再作成する |
| OpenCode セッション | あり | 新ディレクトリへ移動する |
| Git remote | あり | 新 URL へ更新する |

改名後に次の検証を実行し、回帰がないことを確認する。

```powershell
.\.venv\Scripts\python.exe verify_v04_windows.py prototype_v04.xlsx
.\.venv\Scripts\python.exe verify_v05_windows.py prototype_v05_design_b.xlsx <一時出力先>\windows_v05_results_rename_check.json
```

v0.5 の検証結果は一時出力とし、追跡済みの `windows_v05_results.json` を不要に上書きしない。

## 6. 仮想環境

`.venv` には絶対パスが含まれるため、フォルダ改名後に再作成する。Git 管理外であり、リポジトリの内容には影響しない。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -c "import win32com.client; print('OK')"
```

## 7. GitHub

- 旧 URL: `https://github.com/t-tani-it/ExcelBIBasic3`
- 新 URL: `https://github.com/t-tani-it/ExcelBIBasic_DeepSeek_V4.1`
- GitHub の旧 URL は新 URL へリダイレクトされる。
- 旧名で新しいリポジトリを作成するとリダイレクトが切れるため、作成しない。
- ローカル remote は新 URL へ更新する。

```powershell
git remote set-url origin https://github.com/t-tani-it/ExcelBIBasic_DeepSeek_V4.1.git
git remote -v
```

## 8. ブック作成者メタデータ

| ブック | 作成者 | 扱い |
|---|---|---|
| `prototype_v04.xlsx` | `ExcelBIBasic3` | 歴史資産として維持 |
| `prototype_v05_design_b.xlsx` | `ExcelBIBasic3` | 歴史資産として維持 |
| `prototype_v06_design_c.xlsx` | `ExcelBIBasic_DeepSeek_V4.1` | 新規作成時に設定 |

v0.6 のビルドでは、`finalize_v03_ooxml.py` の定数を直接書き換えず、ビルド側で作成者名を上書きする。v0.6 の検証には作成者メタデータの確認を追加する。

## 9. 完了条件

- ローカルフォルダと GitHub リポジトリが新名である。
- `git remote -v` が新 URL を指す。
- `.venv` が再作成され、pywin32 が導入されている。
- `verify_v04_windows.py` と `verify_v05_windows.py` が合格する。
- 入口文書が新名であり、歴史文書に旧名が残っている。
- `docs/project_rename.md` が存在する。
- 公開禁止情報が含まれていない。

## 10. 実施結果（2026-10-22）

| 項目 | 結果 |
|---|---|
| ローカルフォルダ | `ExcelBIBasic_DeepSeek_V4.1` へ変更済み |
| GitHub リポジトリ | `t-tani-it/ExcelBIBasic_DeepSeek_V4.1` へ変更済み |
| Git remote | 新 URL へ更新済み |
| `.venv` | Python 3.11.15 で再作成、pywin32 312 導入済み |
| v0.4 検証 | 173 / 173 合格 |
| v0.5 検証 | 233 / 233 合格 |
| 新規文書の公開情報 | ローカルパス、API キー、個人 ID、トークンなし |

v0.5 の再検証は一時出力へ保存し、追跡済みの `windows_v05_results.json` を上書きしていない。v0.3 / v0.4 の検証コードと作成者期待値は変更していない。

## 11. 変更履歴

| 日付 | 内容 |
|---|---|
| 2026-10-22 | 初版作成。改名スコープ、対象外、テスト影響、GitHub、作成者方針を記録 |
| 2026-10-22 | 実施結果を追記。ローカル、GitHub、remote、`.venv`、v0.4 173 / 173、v0.5 233 / 233 を記録 |
