# 実装計画（ExcelBIBasic3）

- 版: RevA
- 作成日: 2026-10-21
- 対象: 仕様書確定後の実装フェーズ

## 1. 前提
- 仕様書一式（`docs/requirements.md` から `docs/decision_log.md`）が確定している。
- 実機検証環境は Windows と Excel デスクトップアプリである。
- 生成は OOXML 直接生成（Node.js）と Excel COM（pywin32）の併用とする。
- VBA は成果物へ含めない。COM 操作はビルド工程のスクリプトに限定する。

## 2. フェーズ構成
| フェーズ | 内容 | 成果物 | 完了条件 |
|---|---|---|---|
| P0 | 仕様策定 | 仕様書 10 文書 | 全文書の確定 |
| P1 | 基盤データ | PQ クエリ、データモデル、リレーション、DAX メジャー | モデル合計が期待値と一致 |
| P2 | ピボットとフィルタ | PT-01 から PT-06、PG-01 から PG-04、スライサー、タイムライン | 連動テストの合格 |
| P3 | KPI と可視化 | KPI 4 表示、ヒートマップ、スパークライン | 連動と表示形式の合格 |
| P4 | ゲージと条件切替 | ゲージ、目標テーブル、チェックボックス | 非連動と条件切替の合格 |
| P5 | 既存改訂 | 見積入力 B12:F12、採算 D12 | 状態表示の一元化 |
| P6 | 検証 | テスト仕様の全項目 | 合格または保留の判定 |
| P7 | 文書とコミット | progress、handover、分割コミット | 引き継ぎ可能な状態 |

## 3. 実装順序と依存
1. P1 を先行する。モデルが未確定の段階で PT を作成しない。
2. P2 は P1 の合計一致後に着手する。PT は `制御` シートへ作成する。
3. P3 は P2 の連動確認後に着手する。CUBE 式の連動可否を先に判定する。
4. P4 は P3 と独立に着手可能とする。ゲージは PT-06 の作成後に着手する。
5. P5 は P1 から P4 の値に影響しないため、任意の時点で着手可能とする。
6. P6 は全フェーズ完了後に実施する。

## 4. COM 操作の要点
| 操作 | API の候補 | 注意 |
|---|---|---|
| PQ クエリの追加 | `Workbook.Queries.Add` | M 式は文字列で投入。外部ソースなし |
| モデルへのロード | クエリのロード先設定 | ロード先は「データモデルのみ」 |
| DAX メジャー | `Model.ModelMeasures.Add` | 名称は日本語。書式指定は任意 |
| リレーション | `Model.ModelRelationships` | 多対一、日付は日付マスタへ |
| PT の作成 | `PivotCaches.Create`（`xlExternal`、接続 `ThisWorkbookDataModel`） | バージョン引数の確認が必要 |
| スライサー | `SlicerCaches.Add2` | 種類の指定（スライサー） |
| タイムライン | `SlicerCaches.Add2`（タイムライン種別） | API 名と引数の確認が必要 |
| スライサーの接続 | `SlicerCache.SlicerPivotTables.AddPivotTable` | ゲージ用 PT を除外する |
| ピボットグラフ | PT から `PivotChart` を作成し、シートへ移動 | 配色は実装時に設定 |
| 日付テーブルのマーク | UI またはモデル API | API の有無は確認が必要 |
| 条件付き書式 | `Range.FormatConditions.Add` | 色スケール、データバー、数式 |
| フォームコントロール | `Shapes.AddFormControl` | リンクセルは `制御` シート |

- API 名と引数は候補であり、実装時に実機で確認する。確認結果は `decision_log.md` へ記録する。

## 5. ビルドと検証の手順
1. `node build.mjs` で骨格を生成する。出力先は環境変数 `PROTOTYPE_OUTPUT_DIR` に従う（詳細は `README.md`）。
2. `node enhance.mjs` で強化工程を適用する。
3. `python finalize_ooxml.py` で後処理（INDIRECT 復元、保護解除、VML 統合など）を実行する。
4. COM 工程（PQ、モデル、PT、スライサー、条件付き書式、フォームコントロール）を実行する。
5. `python verify_excel_windows.py` と `python verify_v02_windows.py` で検証する。
6. `package_checks.json`、`notes_manifest.json`、`engine_tests.json`、`sales_fixture.json` と照合する。
7. 実機で手動確認（連動、性能、印刷）を行う。

- 実行前に `.venv` を作成し、pywin32 を導入する。
- 各工程は冪等とする。同一入力から同一成果物を生成する。

## 6. リスクと対策
| リスク | 影響 | 対策 |
|---|---|---|
| CUBE 式のスライサー連動不成立 | KPI とヒートマップの前提崩れ | KPI は GETPIVOTDATA へ切替。ヒートマップは PT 方式へ切替 |
| タイムラインの CUBE 式への非反映 | KPI の期間不一致 | PT-05 の GETPIVOTDATA へ切替 |
| 日付テーブルのマーク不可 | 前年売上の不成立 | `CALCULATE` と `FILTER` 方式へ切替。不可時は前年比を非表示 |
| ヒートマップの再計算遅延 | 操作性の低下 | セル数を削減し、PT 方式へ切替。性能を計測して記録 |
| PT-06 のレイアウト変化 | ゲージの値ズレ | 監視セルで検出。行数の固定を確認 |
| ピボットグラフの配色保存 | 見た目の不一致 | テーマ配色と個別設定の優先順を確認 |
| 商品追加時の行不足 | ヒートマップの欠落 | 20 行を予約し、追加手順を文書化 |
| Excel の版差 | 機能の不成立 | `compatibility_matrix.md` の範囲で判定し、実機未検証を明示 |

## 7. 完了条件
- `docs/test_spec.md` の新規項目が合格または保留の判定となる。
- 計算仕様の変更禁止セルが維持される（`INDIRECT` 24 件、メモ 74 件以上）。
- 保護、固定枠、VBA、外部リンクが存在しない。
- `docs/handover.md` と `docs/progress.md` が更新される。
- 公開禁止情報がないことを確認する。

## 8. コミット計画
| 順序 | 内容 | 例 |
|---|---|---|
| 1 | 骨格と仕様書 | `docs: add specification set for ExcelBIBasic3` |
| 2 | 作業管理 | `docs: add progress and handover` |
| 3 | データ基盤 | `feat: add power query and data model` |
| 4 | ピボットとフィルタ | `feat: add pivot tables, charts, slicers, timeline` |
| 5 | KPI と可視化 | `feat: add kpi, heatmap, sparklines` |
| 6 | ゲージと切替 | `feat: add gauge and condition checkboxes` |
| 7 | 既存改訂 | `fix: restore remarks input and unify status` |
| 8 | 検証 | `test: add verification results` |

## 9. 未確定事項
| 項目 | 判断時期 |
|---|---|
| タイムライン作成 API の確定 | P2 の着手時 |
| 日付テーブルのマーク可否 | P1 の完了時 |
| KPI の方式確定 | P3 の着手時 |
| ヒートマップの性能判定 | P3 の中間時点 |
| ピボットグラフ配色の設定方法 | P2 の完了時 |
