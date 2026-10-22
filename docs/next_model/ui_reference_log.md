# 外部 UI 参照ログ（ExcelBIBasic_DeepSeek_V4.1）

- 版: RevA
- 作成日: 2026-10-22
- 対象: v0.6 デザイン C
- 目的: 外部のデザイン、テンプレート、配色、アイコン、グラフ配置の参照元とライセンスを記録する

## 1. 方針

- 参考元は 3〜5 件まで。
- 参考画像そのものは public リポジトリへ保存しない。URL と採用メモだけを残す。
- 良い要素は Excel のセル、図形、グラフ、スライサー、条件付き書式で再現する。
- 外部リンク、外部画像、外部フォント、外部アイコンフォントをブック内に埋め込まない。
- ライセンスが不明な素材、有料テンプレート、ロゴは使用しない。
- 採用した要素は、設計書 `docs/variant_design_c/design_spec.md` にも記録する。

## 2. 参照元

| ID | 参考元 | URL | ライセンス | 採用する要素 | Excel での実装 | 改変内容 | 状態 |
|---|---|---|---|---|---|---|---|
| R-01 | IBM Carbon Design System | https://carbondesignsystem.com/ | Apache-2.0 | 8 px 基準の間隔、12 カラム、文字階層、余白の考え方 | 列幅・行高・セル結合・カード余白で近似 | 色とフォントは本プロジェクト用に変更 | 採用 |
| R-02 | Financial Times Visual Vocabulary | https://github.com/Financial-Times/chart-doctor/tree/main/visual-vocabulary | MIT | 目的別のグラフ選択。推移は線、比較は棒、構成は積み上げ棒 | Excel の折れ線、横棒、100% 積み上げ横棒 | 2 分類の円グラフを廃止 | 採用 |
| R-03 | Microsoft Fluent UI System Icons | https://github.com/microsoft/fluentui-system-icons | MIT | 16〜20 px の線アイコン、モノクロ、最小限の使用 | PNG 画像または Excel 図形。最大 6 個 | 色は Navy または Gray に統一 | 候補 |
| R-04 | Lucide | https://lucide.dev/ | ISC | 線アイコンのスタイル | 原則使用しない。Fluent と混在させない | 採用しない | 不採用 |
| R-05 | Material Symbols | https://github.com/google/material-design-icons | Apache-2.0 | アイコンの線幅、余白の考え方 | 原則使用しない。Fluent と混在させない | 採用しない | 不採用 |

### 2.1 ライセンス URL

| ID | ライセンス | URL |
|---|---|---|
| R-01 | Apache-2.0 | https://github.com/carbon-design-system/carbon/blob/main/LICENSE |
| R-02 | MIT | https://github.com/Financial-Times/chart-doctor/blob/main/LICENSE |
| R-03 | MIT | https://github.com/microsoft/fluentui-system-icons/blob/main/LICENSE |
| R-04 | ISC | https://github.com/lucide-icons/lucide/blob/main/LICENSE |
| R-05 | Apache-2.0 | https://github.com/google/material-design-icons/blob/master/LICENSE |

## 3. 採用基準

- 情報の読み取りやすさが上がる。
- Excel の標準機能で再現できる。
- 日本語の文字が崩れない。
- コントラストが確保できる。
- ライセンスが明確で、再配布条件を満たせる。
- 装飾の追加ではなく、階層・比較・余白の改善につながる。

## 4. 不採用の基準

- 参考画像の貼り付けが必要。
- 外部フォントやアイコンフォントの同梱が必要。
- 有料テンプレートの再配布が必要。
- ロゴや商標の使用が必要。
- Excel で再現できないアニメーションや特殊グラフ。
- 画面を詰めすぎる、または情報密度を下げすぎる。

## 5. ライセンス記録

| 項目 | 記録内容 |
|---|---|
| 参考元の名称 |  |
| URL |  |
| ライセンス名 |  |
| ライセンス URL |  |
| 商用利用 | 可 / 不可 / 条件付き |
| 再配布 | 可 / 不可 / 条件付き |
| 帰属表示 | 必要 / 不要 |
| 採用要素 |  |
| 改変内容 |  |
| 判断 | 採用 / 不採用 |
| 判断日 |  |

## 6. 帰属表示

- MIT、ISC、Apache-2.0 は、原則として著作権表示とライセンス文の保持が必要である。
- 配布物単体で帰属表示が必要な場合は、ブック内の注記にテキストで記載する。ハイパーリンクは入れない。
- リポジトリ内の設計書にも出典とライセンスを記載する。
- 参考画像そのものはコミットしない。

## 7. v0.6 での想定

- R-01 はレイアウトと余白の考え方だけを採用する。Carbon の画像やフォントは使わない。
- R-02 はグラフ選択の考え方だけを採用する。FT の画像やレイアウトをコピーしない。
- R-03 を使う場合も、アイコンは 4〜6 個まで。KPI ラベルやセクション見出しの補助に限定する。
- R-04 と R-05 は不採用とし、アイコンスタイルの混在を避ける。
- 色は `design_tokens.json` の v0.5 系パレットを基本とする。外部パレットを採用する場合は、このログに追加する。

## 8. 変更履歴

| 日付 | 内容 |
|---|---|
| 2026-10-22 | 初版作成。Carbon、FT Visual Vocabulary、Fluent、Lucide、Material Symbols を記録 |
