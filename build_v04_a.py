"""v0.4 方式 A（通常ピボット・スライサー・タイムライン）の売上分析ビルド。

目的:
    既存 v0.3 の SUMIFS 版売上分析を、通常のピボットテーブル（共有キャッシュ）、
    業種・商品スライサー、日付タイムライン、ピボットグラフ、ヒートマップへ
    置き換える。方式 A のため Power Query・データモデル・DAX は使用しない。

仕様:
    - `SalesHistory` に年月・年・月の補助列を追加する。
    - `商品マスタ` シートと `ProductMaster` テーブルを追加する（方式 A では未接続）。
    - 非表示の `制御` シートに共有キャッシュと PT-01 から PT-08 を作成する。
    - 業種・商品スライサーと日付タイムラインを作成する。
      PT-07（前年比）と PT-08（前年比較）は前年データ参照のためタイムライン非連動。
    - `売上分析` を KPI 4 表示、ピボットグラフ 4 個、商品×月ヒートマップへ再構築する。
    - 円は KPI と保存値、万円はグラフとヒートマップの表示に使用する。
    - 保存後に OOXML を後処理し、絶対パスと作成者情報を除去する。

前提:
    - 入力は v0.3 成果物 `prototype_v03.xlsx` である。
    - Windows と Microsoft Excel デスクトップアプリ、pywin32 が利用できる。
    - 元ブックのメモ 74 件、INDIRECT 24 件、保護・固定枠なしを維持する。

入出力:
    - 入力: 引数 1（v0.3 xlsx のパス）。
    - 出力: 引数 2（v0.4 xlsx のパス）。

注意点:
    - Excel のタイムライン期間設定は文字列日付で行う（datetime では反映されない）。
    - 計算フィールドはピボットキャッシュ共有のため、最初の PT でのみ追加する。
    - 既存の計算セル（見積・原価・利益）とメモ・INDIRECT には触れない。
    - 保存のたびに `finalize_v03_ooxml.sanitize_package` を実行する。

拡張ポイント:
    - ゲージ・目標管理・チェックボックスは次フェーズで追加する。
    - 商品マスタをデータモデルへ接続する場合は方式 B へ移行する。

変更履歴:
    2026-10-22: 初版作成（方式 A の本実装）。
"""

from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

from finalize_v03_ooxml import CONFIG as OOXML_CONFIG
from finalize_v03_ooxml import sanitize_package

# ---------------------------------------------------------------------------
# 設定値
# ---------------------------------------------------------------------------
CONFIG: dict[str, Any] = {
    # シート名
    "data_sheet": "売上データ",
    "dashboard": "売上分析",
    "control": "制御",
    "product_master": "商品マスタ",
    # テーブル名
    "table": "SalesHistory",
    "product_table": "ProductMaster",
    # 売上データへ追加する補助列（構造化参照）
    "helper_columns": (
        ("年月", '=TEXT([@日付],"yyyy-mm")'),
        ("年", "=YEAR([@日付])"),
        ("月", "=MONTH([@日付])"),
    ),
    # ヒートマップの月（2025-01 から 2026-12）
    "months": tuple(
        f"{year}-{month:02d}" for year in (2025, 2026) for month in range(1, 13)
    ),
    # 商品マスタ（表示順はヒートマップの行順）
    "products": (
        ("P-01", "テーブル", "木工所", 1, True),
        ("P-02", "椅子", "木工所", 2, True),
        ("P-03", "収納棚", "木工所", 3, True),
        ("P-04", "木製小物", "木工所", 4, True),
        ("P-05", "野菜5kg箱", "農家", 5, True),
        ("P-06", "トマト箱", "農家", 6, True),
        ("P-07", "米5kg", "農家", 7, True),
        ("P-08", "季節の果物", "農家", 8, True),
    ),
    # ピボットテーブルの配置（非表示シート。重ならない位置）
    "pt_anchors": {
        "PT-01": "A2",
        "PT-02": "A12",
        "PT-03": "A26",
        "PT-04": "A56",
        "PT-05": "H2",
        "PT-06": "H14",
        "PT-07": "H20",
        "PT-08": "H28",
    },
    # タイムライン初期範囲と全期間
    "timeline_start": "2026-01-01",
    "timeline_end": "2026-12-31",
    "full_start": "2025-01-01",
    "full_end": "2026-12-31",
    # 期待値（2026 年・全業種）
    "expected_sales": 47492700,
    "expected_diff": 18852426,
    "expected_rate": 0.3969541845378342,
    "expected_yoy": -0.011131169147547304,
    # 不変条件
    "expected_notes": 74,
    "expected_indirect": 24,
    # 配色（Excel COM 用 RGB）
    "colors": {
        "dark_green": (31, 92, 74),
        "teal": (46, 158, 143),
        "yellow_green": (140, 198, 63),
        "light_green": (221, 239, 227),
        "mid_green": (127, 199, 178),
        "dark_teal": (31, 122, 102),
        "white": (255, 255, 255),
        "text": (64, 64, 64),
        "gray": (117, 117, 117),
        "border": (217, 217, 217),
        "paper": (250, 250, 247),
    },
}

# COM 定数（可読性のため名前を付ける）
XL_SUM = -4157
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2
XL_DATA_FIELD = 4
XL_TABULAR_ROW = 1
XL_CENTER = -4108
XL_LEFT = -4131
XL_BOTTOM = -4107
XL_LANDSCAPE = 2
XL_SHEET_HIDDEN = 0
XL_LINE = 4
XL_COLUMN_CLUSTERED = 51
XL_BAR_CLUSTERED = 57
XL_DOUGHNUT = -4120
XL_NONE = -4142


def log(*args: object) -> None:
    """標準出力へ即時出力する。

    Args:
        *args: 出力対象。

    Returns:
        なし。

    Side Effects:
        標準出力へ書き込む。
    """
    print(*args, flush=True)


def rgb(color: tuple[int, int, int]) -> int:
    """RGB タプルを Excel COM の色値へ変換する。

    Args:
        color: (赤, 緑, 青) の 0 から 255。

    Returns:
        Excel COM の色値。

    Side Effects:
        なし。

    Examples:
        >>> rgb((46, 158, 143))
    """
    red, green, blue = color
    return red + (green << 8) + (blue << 16)


def column_letter(index: int) -> str:
    """1 始まりの列番号を列文字へ変換する。

    Args:
        index: 列番号（1 は A）。

    Returns:
        列文字。

    Side Effects:
        なし。

    Examples:
        >>> column_letter(2)
        'B'
    """
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def count_notes(book: object) -> int:
    """ブック全体のメモ件数を数える。

    Args:
        book: 対象ブック。

    Returns:
        メモ件数。

    Side Effects:
        なし。
    """
    return sum(sheet.Comments.Count for sheet in book.Worksheets)


def count_indirect(book: object) -> int:
    """INDIRECT を含む数式の件数を数える。

    Args:
        book: 対象ブック。

    Returns:
        INDIRECT を含む数式のセル数。

    Side Effects:
        なし。
    """
    total = 0
    for sheet in book.Worksheets:
        for cell in sheet.UsedRange:
            try:
                if cell.HasFormula and "INDIRECT" in cell.Formula.upper():
                    total += 1
            except Exception:
                continue
    return total


def validate_input(path: Path) -> None:
    """入力ファイルの存在と拡張子を検証する。

    Args:
        path: 入力 xlsx。

    Returns:
        なし。

    Raises:
        FileNotFoundError: 入力が存在しない場合。
        ValueError: 拡張子が .xlsx でない場合。

    Side Effects:
        なし。
    """
    if not path.is_file():
        raise FileNotFoundError(f"入力ファイルがありません: {path}")
    if path.suffix.lower() != ".xlsx":
        raise ValueError(f"入力は .xlsx です: {path}")


def validate_workbook(book: object, app: object, cfg: dict[str, Any]) -> None:
    """v0.3 入力ブックの前提を検証する。

    Args:
        book: 対象ブック。
        app: Excel Application。
        cfg: 設定値。

    Returns:
        なし。

    Raises:
        ValueError: 前提が崩れている場合。

    Side Effects:
        なし。
    """
    names = [book.Worksheets(i).Name for i in range(1, book.Worksheets.Count + 1)]
    required = (cfg["data_sheet"], cfg["dashboard"], "計算", "見積入力", "設定")
    for name in required:
        if name not in names:
            raise ValueError(f"必須シートがありません: {name}")
    for name in (cfg["control"], cfg["product_master"]):
        if name in names:
            raise ValueError(f"既に追加済みのシートがあります: {name}")
    if book.PivotCaches().Count != 0:
        raise ValueError("入力に既存のピボットキャッシュがあります")
    if book.SlicerCaches.Count != 0:
        raise ValueError("入力に既存のスライサーがあります")
    if book.HasVBProject:
        raise ValueError("入力に VBA が含まれています")
    if book.LinkSources():
        raise ValueError("入力に外部リンクがあります")
    if book.FileFormat != 51:
        raise ValueError(f"入力が .xlsx ではありません: {book.FileFormat}")
    if book.Queries.Count != 0:
        raise ValueError("入力に Power Query があります")

    data = book.Worksheets(cfg["data_sheet"])
    if cfg["table"] not in [lo.Name for lo in data.ListObjects]:
        raise ValueError(f"テーブルがありません: {cfg['table']}")
    table = data.ListObjects(cfg["table"])
    if table.Range.Address != "$A$4:$I$580":
        raise ValueError(f"SalesHistory の範囲が想定外です: {table.Range.Address}")
    if app.WorksheetFunction.CountA(data.Range("J4:L580")) != 0:
        raise ValueError("売上データの J4:L580 に既存値があります")

    notes = count_notes(book)
    indirect = count_indirect(book)
    if notes != cfg["expected_notes"]:
        raise ValueError(f"メモ件数が想定外です: {notes}")
    if indirect != cfg["expected_indirect"]:
        raise ValueError(f"INDIRECT 件数が想定外です: {indirect}")
    if any(sheet.ProtectContents for sheet in book.Worksheets):
        raise ValueError("入力にシート保護があります")

def add_helper_columns(book: object, cfg: dict[str, Any]) -> None:
    """`SalesHistory` に年月・年・月の補助列を追加する。

    Args:
        book: 対象ブック。
        cfg: 設定値。

    Returns:
        なし。

    Raises:
        ValueError: テーブル範囲または補助列の値が想定外の場合。

    Side Effects:
        `売上データ` のテーブルを A4:L580 へ拡張する。
    """
    data = book.Worksheets(cfg["data_sheet"])
    table = data.ListObjects(cfg["table"])
    for name, formula in cfg["helper_columns"]:
        column = table.ListColumns.Add()
        column.Name = name
        column.DataBodyRange.Formula = formula
    address = table.Range.Address
    if address != "$A$4:$L$580":
        raise ValueError(f"補助列追加後の範囲が想定外です: {address}")
    expected = (("J5", "2025-01"), ("K5", 2025), ("L5", 1))
    for cell, want in expected:
        got = data.Range(cell).Value
        if got != want:
            raise ValueError(f"{cell} の値が想定外です: {got!r}")
    data.Range("J:J").ColumnWidth = 12
    data.Range("K:K").ColumnWidth = 8
    data.Range("L:L").ColumnWidth = 8
    data.Range("J4:L580").Locked = False
    data.Range("J4:L580").FormulaHidden = False
    for cell in ("J4", "K4", "L4"):
        data.Range(cell).HorizontalAlignment = XL_CENTER


def create_product_master(book: object, cfg: dict[str, Any]) -> None:
    """商品マスタシートと `ProductMaster` テーブルを作成する。

    Args:
        book: 対象ブック。
        cfg: 設定値。

    Returns:
        なし。

    Raises:
        ValueError: テーブル範囲が想定外の場合。

    Side Effects:
        `売上分析` の直前に可視シートを追加する。
    """
    # 売上データの直前に追加する（売上分析の直後）
    book.Worksheets(cfg["data_sheet"]).Activate()
    sheet = book.Worksheets.Add()
    sheet.Name = cfg["product_master"]

    colors = cfg["colors"]
    sheet.Range("A2").Value = "商品マスタ"
    sheet.Range("A2").Font.Size = 18
    sheet.Range("A2").Font.Bold = True
    sheet.Range("A2").Font.Color = rgb(colors["dark_green"])
    sheet.Range("A3").Value = (
        "商品の表示順と有効フラグを管理します。"
        "方式 A ではピボット・スライサーに未接続です。"
    )
    sheet.Range("A3").Font.Size = 10
    sheet.Range("A3").Font.Color = rgb(colors["gray"])
    sheet.Range("A3").WrapText = True

    headers = ("商品 ID", "商品名", "業種", "表示順", "有効")
    for col, value in enumerate(headers, start=1):
        sheet.Cells(4, col).Value = value
    for row, product in enumerate(cfg["products"], start=5):
        for col, value in enumerate(product, start=1):
            sheet.Cells(row, col).Value = value

    table = sheet.ListObjects.Add(1, sheet.Range("A4:E12"), False, 1)
    table.Name = cfg["product_table"]
    try:
        table.TableStyle = "TableStyleMedium4"
    except Exception as exc:
        log(f"商品マスタのテーブルスタイル設定を省略: {exc}")
    if table.Range.Address != "$A$4:$E$12":
        raise ValueError(f"ProductMaster の範囲が想定外です: {table.Range.Address}")

    # ヘッダーは濃緑で統一する
    header = sheet.Range("A4:E4")
    header.Interior.Color = rgb(colors["dark_green"])
    header.Font.Color = rgb(colors["white"])
    header.Font.Bold = True
    header.HorizontalAlignment = XL_CENTER
    sheet.Range("D5:E12").HorizontalAlignment = XL_CENTER

    sheet.Range("A:A").ColumnWidth = 10
    sheet.Range("B:B").ColumnWidth = 18
    sheet.Range("C:C").ColumnWidth = 12
    sheet.Range("D:D").ColumnWidth = 10
    sheet.Range("E:E").ColumnWidth = 8
    sheet.Rows(2).RowHeight = 28
    sheet.Rows(3).RowHeight = 30

    note = sheet.Range("A14:E16")
    note.Merge()
    note.Value = (
        "スライサーは売上データの商品列を使用します。商品を追加する場合は"
        "売上データにも同じ商品名を登録してください。表示順はヒートマップの行順です。"
    )
    note.WrapText = True
    note.VerticalAlignment = XL_CENTER
    note.HorizontalAlignment = XL_LEFT
    note.Font.Size = 9
    note.Font.Color = rgb(colors["gray"])
    note.Interior.Color = rgb(colors["light_green"])
    sheet.Range("A1:E20").Locked = False
    sheet.Range("A1:E20").FormulaHidden = False
    sheet.Tab.Color = rgb(colors["light_green"])


def add_calc_data_field(pt: object, field_name: str, caption: str) -> None:
    """計算フィールドをデータ領域へ追加し、表示名と書式を設定する。

    Args:
        pt: ピボットテーブル。
        field_name: 計算フィールド名。
        caption: データフィールドの表示名。

    Returns:
        なし。

    Side Effects:
        ピボットテーブルのデータ領域を変更する。
    """
    pt.PivotFields(field_name).Orientation = XL_DATA_FIELD
    data_field = pt.DataFields(pt.DataFields.Count)
    data_field.Name = caption
    data_field.NumberFormat = '0"万円"'


def make_pivot(
    cache: object,
    sheet: object,
    name: str,
    anchor: str,
    rows: tuple[str, ...] = (),
    columns: tuple[str, ...] = (),
    values: tuple[tuple[str, str], ...] = (),
    row_grand: bool = False,
    column_grand: bool = False,
    tabular: bool = True,
) -> object:
    """1 個のピボットテーブルを作成して配置する。

    Args:
        cache: 共有ピボットキャッシュ。
        sheet: 配置先シート。
        name: ピボットテーブル名。
        anchor: 左上セル。
        rows: 行フィールド名。
        columns: 列フィールド名。
        values: (フィールド名, 表示名) のデータフィールド。
        row_grand: 行集計の表示有無。
        column_grand: 列集計の表示有無。
        tabular: 表形式レイアウトの有無。

    Returns:
        作成したピボットテーブル。

    Side Effects:
        シートへピボットテーブルを追加する。
    """
    pt = cache.CreatePivotTable(sheet.Range(anchor), name)
    for field in rows:
        pt.PivotFields(field).Orientation = XL_ROW_FIELD
    for field in columns:
        pt.PivotFields(field).Orientation = XL_COLUMN_FIELD
    for field_name, caption in values:
        if field_name in ("売上万円", "差額万円"):
            add_calc_data_field(pt, field_name, caption)
        else:
            pt.AddDataField(pt.PivotFields(field_name), caption, XL_SUM)
            pt.DataFields(pt.DataFields.Count).NumberFormat = "#,##0"
    pt.RowGrand = row_grand
    pt.ColumnGrand = column_grand
    if tabular and rows:
        pt.RowAxisLayout(XL_TABULAR_ROW)
    pt.RefreshTable()
    return pt


def create_control_sheet(
    book: object, cfg: dict[str, Any]
) -> tuple[dict[str, object], object]:
    """非表示予定の制御シートに共有キャッシュと PT-01 から PT-08 を作成する。

    Args:
        book: 対象ブック。
        cfg: 設定値。

    Returns:
        (PT 名から COM オブジェクトへの辞書, 共有キャッシュ)。

    Raises:
        ValueError: 作成後の範囲が想定外の場合。

    Side Effects:
        `制御` シートを追加し、ピボットテーブルを作成する。
    """
    book.Worksheets(cfg["data_sheet"]).Activate()
    sheet = book.Worksheets.Add()
    sheet.Name = cfg["control"]
    sheet.Range("A1").Value = "制御シート（ピボット・スライサー用。非表示）"
    sheet.Range("A1").Font.Size = 10
    sheet.Range("A1").Font.Color = rgb(cfg["colors"]["gray"])

    cache = book.PivotCaches().Create(1, cfg["table"])
    anchors = cfg["pt_anchors"]
    pts: dict[str, object] = {}

    # PT-01 業種別（計算フィールドは共有キャッシュへ最初に追加する）
    pt = cache.CreatePivotTable(sheet.Range(anchors["PT-01"]), "PT-01")
    pt.CalculatedFields().Add("売上万円", "=税抜売上/10000")
    pt.CalculatedFields().Add("差額万円", "=計上原価控除後差額/10000")
    pt.PivotFields("業種").Orientation = XL_ROW_FIELD
    add_calc_data_field(pt, "売上万円", "売上（万円）")
    pt.RowGrand = False
    pt.ColumnGrand = False
    pt.RowAxisLayout(XL_TABULAR_ROW)
    pt.RefreshTable()
    pts["PT-01"] = pt

    # PT-02 商品別（売上降順）
    pt = make_pivot(
        cache,
        sheet,
        "PT-02",
        anchors["PT-02"],
        rows=("商品",),
        values=(("売上万円", "売上（万円）"),),
    )
    try:
        pt.PivotFields("商品").AutoSort(2, "売上（万円）")
        pt.RefreshTable()
    except Exception as exc:
        log(f"PT-02 の並べ替えを省略: {exc}")
    pts["PT-02"] = pt

    # PT-03 年月別（売上と差額）
    pts["PT-03"] = make_pivot(
        cache,
        sheet,
        "PT-03",
        anchors["PT-03"],
        rows=("年月",),
        values=(("売上万円", "売上（万円）"), ("差額万円", "差額（万円）")),
    )

    # PT-04 業種×商品
    pt = make_pivot(
        cache,
        sheet,
        "PT-04",
        anchors["PT-04"],
        rows=("業種", "商品"),
        values=(("売上万円", "売上（万円）"),),
    )
    try:
        pt.RepeatAllLabels(2)
    except Exception as exc:
        log(f"PT-04 のラベル繰り返しを省略: {exc}")
    pts["PT-04"] = pt

    # PT-05 商品×年月（ヒートマップ参照元）
    pts["PT-05"] = make_pivot(
        cache,
        sheet,
        "PT-05",
        anchors["PT-05"],
        rows=("商品",),
        columns=("年月",),
        values=(("売上万円", "売上（万円）"),),
    )

    # PT-06 KPI（円の売上合計と差額合計）
    pts["PT-06"] = make_pivot(
        cache,
        sheet,
        "PT-06",
        anchors["PT-06"],
        values=(("税抜売上", "売上合計"), ("計上原価控除後差額", "差額合計")),
        row_grand=True,
        column_grand=True,
        tabular=False,
    )

    # PT-07 前年比（年別。タイムライン非連動）
    pts["PT-07"] = make_pivot(
        cache,
        sheet,
        "PT-07",
        anchors["PT-07"],
        rows=("年",),
        values=(("売上万円", "売上（万円）"),),
    )

    # PT-08 前年比較（月×年。タイムライン非連動）
    pts["PT-08"] = make_pivot(
        cache,
        sheet,
        "PT-08",
        anchors["PT-08"],
        rows=("月",),
        columns=("年",),
        values=(("売上万円", "売上（万円）"),),
    )

    sheet.Range("A:A").ColumnWidth = 22
    sheet.Range("B:G").ColumnWidth = 14
    sheet.Range("H:AG").ColumnWidth = 9
    sheet.Range("A1:AG80").Locked = False
    sheet.Range("A1:AG80").FormulaHidden = False
    sheet.Tab.Color = rgb(cfg["colors"]["gray"])
    return pts, cache

def rebuild_dashboard_cells(dashboard: object, cfg: dict[str, Any]) -> None:
    """売上分析のセルを初期化し、タイトル・KPI・注記を配置する。

    Args:
        dashboard: 売上分析シート。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        既存のセル内容・書式・入力規則を削除し、新しいレイアウトを設定する。
    """
    colors = cfg["colors"]
    # 既存の圖形は呼び出し側で削除済み。セルと入力規則を初期化する。
    dashboard.Range("A1:Z100").UnMerge()
    dashboard.Range("A1:Z100").Clear()
    try:
        dashboard.Range("A1:Z100").Validation.Delete()
    except Exception:
        pass

    dashboard.Range("A1:Z100").Font.Name = "Yu Gothic UI"
    dashboard.Range("A1:Z100").Font.Size = 10
    dashboard.Range("A1:Z100").Font.Color = rgb(colors["text"])
    dashboard.Range("A1:Z100").Interior.Color = rgb(colors["paper"])
    dashboard.Range("A1:Z100").Locked = False
    dashboard.Range("A1:Z100").FormulaHidden = False

    widths = {
        "A": 18,
        "B": 12,
        "C": 12,
        "D": 18,
        "E": 12,
        "F": 12,
        "G": 14,
        "H": 14,
        "I": 13,
        "J": 14,
        "K": 13,
        "L": 13,
        "M": 11,
    }
    for column, width in widths.items():
        dashboard.Range(f"{column}:{column}").ColumnWidth = width
    dashboard.Range("N:Y").ColumnWidth = 11
    dashboard.Range("Z:Z").ColumnWidth = 18

    heights = {2: 30, 3: 18, 5: 150, 6: 20, 8: 18, 9: 32, 11: 34, 44: 22, 45: 20}
    for row, height in heights.items():
        dashboard.Rows(row).RowHeight = height
    dashboard.Range("46:53").RowHeight = 20
    dashboard.Rows(55).RowHeight = 32
    dashboard.Rows(57).RowHeight = 18
    dashboard.Rows(58).RowHeight = 18

    # タイトル
    dashboard.Range("A2").Value = "売上分析"
    dashboard.Range("A2").Font.Size = 20
    dashboard.Range("A2").Font.Bold = True
    dashboard.Range("A2").Font.Color = rgb(colors["dark_green"])
    dashboard.Range("B2:F2").Merge()
    dashboard.Range("B2").Value = "商品と季節の動きをつかむ"
    dashboard.Range("B2").Font.Size = 11
    dashboard.Range("B2").Font.Color = rgb(colors["gray"])
    dashboard.Range("B2").VerticalAlignment = XL_CENTER

    dashboard.Range("A3:L3").Merge()
    dashboard.Range("A3").Value = (
        "すべて架空の販売履歴です。現在の見積・予定案件利益には加算されません。"
    )
    dashboard.Range("A3").Font.Size = 9
    dashboard.Range("A3").Font.Color = rgb(colors["gray"])

    # スライサー帯
    dashboard.Range("A5:L5").Interior.Color = rgb(colors["light_green"])
    dashboard.Range("A6:L6").Merge()
    dashboard.Range("A6").Value = (
        "業種・商品スライサーとタイムラインに連動"
        "（前年比・前年比較は前年データを参照）"
    )
    dashboard.Range("A6").Font.Size = 9
    dashboard.Range("A6").Font.Color = rgb(colors["gray"])
    dashboard.Range("A6").VerticalAlignment = XL_CENTER

    # KPI カード
    cards = (
        ("A8:C8", "A9:C9", "A8", "税抜売上（円）", "A9"),
        ("D8:F8", "D9:F9", "D8", "計上原価控除後差額（円）", "D9"),
        ("G8:I8", "G9:I9", "G8", "差額率", "G9"),
        ("J8:L8", "J9:L9", "J8", "前年比", "J9"),
    )
    for head_range, value_range, head_cell, title, value_cell in cards:
        dashboard.Range(head_range).Merge()
        dashboard.Range(value_range).Merge()
        dashboard.Range(head_cell).Value = title
        dashboard.Range(head_range).Interior.Color = rgb(colors["light_green"])
        dashboard.Range(head_cell).Font.Size = 9
        dashboard.Range(head_cell).Font.Bold = True
        dashboard.Range(head_cell).Font.Color = rgb(colors["dark_green"])
        dashboard.Range(head_cell).HorizontalAlignment = XL_CENTER
        dashboard.Range(head_cell).VerticalAlignment = XL_CENTER
        dashboard.Range(value_range).Interior.Color = rgb(colors["white"])
        dashboard.Range(value_cell).Font.Size = 20
        dashboard.Range(value_cell).Font.Bold = True
        dashboard.Range(value_cell).Font.Color = rgb(colors["dark_green"])
        dashboard.Range(value_cell).HorizontalAlignment = XL_CENTER
        dashboard.Range(value_cell).VerticalAlignment = XL_CENTER
        card_range = f"{head_range.split(':')[0]}:{value_range.split(':')[1]}"
        dashboard.Range(card_range).Borders.LineStyle = 1
        dashboard.Range(card_range).Borders.Color = rgb(colors["border"])
        dashboard.Range(value_range).Borders.LineStyle = 1
        dashboard.Range(value_range).Borders.Color = rgb(colors["border"])

    dashboard.Range("A9").Formula = '=ROUND(GETPIVOTDATA("売上合計",制御!$H$14),0)'
    dashboard.Range("D9").Formula = '=ROUND(GETPIVOTDATA("差額合計",制御!$H$14),0)'
    dashboard.Range("G9").Formula = '=IF(A9=0,"算出不可",D9/A9)'
    dashboard.Range("J9").Formula = (
        '=IFERROR(GETPIVOTDATA("売上（万円）",制御!$H$20,"年",2026)'
        '/GETPIVOTDATA("売上（万円）",制御!$H$20,"年",2025)-1,"算出不可")'
    )
    dashboard.Range("A9").NumberFormat = "#,##0"
    dashboard.Range("D9").NumberFormat = "#,##0"
    dashboard.Range("G9").NumberFormat = "0.0%"
    dashboard.Range("J9").NumberFormat = "0.0%"

    dashboard.Range("A11:L11").Merge()
    dashboard.Range("A11").Value = (
        "差額は売上履歴の計上原価のみを控除。"
        "会計上の粗利・営業利益や現在の見積利益とは区別します。"
    )
    dashboard.Range("A11").Font.Size = 9
    dashboard.Range("A11").Font.Color = rgb(colors["gray"])
    dashboard.Range("A11").WrapText = True
    dashboard.Range("A11").VerticalAlignment = XL_CENTER

    try:
        dashboard.PageSetup.PrintArea = "$A$1:$Z$58"
        dashboard.PageSetup.Orientation = XL_LANDSCAPE
        dashboard.PageSetup.Zoom = False
        dashboard.PageSetup.FitToPagesWide = 1
        dashboard.PageSetup.FitToPagesTall = 1
    except Exception as exc:
        log(f"印刷設定を省略: {exc}")
    dashboard.Tab.Color = rgb(colors["teal"])


def connect_slicer(slicer_cache: object, pts: dict[str, object], skip: str) -> None:
    """スライサーキャッシュへ全ピボットテーブルを接続する。

    Args:
        slicer_cache: スライサーキャッシュ。
        pts: ピボットテーブル辞書。
        skip: 既に接続済みの PT 名。

    Returns:
        なし。

    Raises:
        RuntimeError: 接続に失敗した場合。

    Side Effects:
        ピボットテーブルの絞り込み対象を増やす。
    """
    for name, pt in pts.items():
        if name == skip:
            continue
        try:
            slicer_cache.PivotTables.AddPivotTable(pt)
        except Exception as exc:
            raise RuntimeError(
                f"スライサー接続に失敗: {slicer_cache.Name} -> {name}: {exc}"
            ) from exc


def create_slicers(
    book: object, dashboard: object, pts: dict[str, object], cfg: dict[str, Any]
) -> dict[str, object]:
    """業種・商品スライサーと日付タイムラインを作成する。

    Args:
        book: 対象ブック。
        dashboard: 売上分析シート。
        pts: ピボットテーブル辞書。
        cfg: 設定値。

    Returns:
        スライサーキャッシュの辞書。

    Raises:
        RuntimeError: 接続に失敗した場合。

    Side Effects:
        売上分析シートへスライサーとタイムラインを配置する。
    """
    colors = cfg["colors"]
    top = dashboard.Range("A5").Top + 8
    industry_left = dashboard.Range("A5").Left + 8
    product_left = dashboard.Range("D5").Left + 8
    timeline_left = dashboard.Range("J5").Left + 8

    industry = book.SlicerCaches.Add2(pts["PT-01"], "業種", "SL_業種")
    industry_shape = industry.Slicers.Add(
        dashboard, pythoncom.Empty, "SL_業種_Shape", "業種", top, industry_left, 120, 80
    )
    connect_slicer(industry, pts, "PT-01")

    product = book.SlicerCaches.Add2(pts["PT-02"], "商品", "SL_商品")
    product_shape = product.Slicers.Add(
        dashboard, pythoncom.Empty, "SL_商品_Shape", "商品", top, product_left, 180, 140
    )
    connect_slicer(product, pts, "PT-02")

    timeline = book.SlicerCaches.Add2(pts["PT-03"], "日付", "TL_日付", 2)
    timeline_shape = timeline.Slicers.Add(
        dashboard,
        pythoncom.Empty,
        "TL_日付_Shape",
        "日付",
        top,
        timeline_left,
        330,
        100,
    )
    # 前年比・前年比較は前年データを参照するためタイムライン非連動
    for name in ("PT-01", "PT-02", "PT-04", "PT-05", "PT-06"):
        try:
            timeline.PivotTables.AddPivotTable(pts[name])
        except Exception as exc:
            raise RuntimeError(
                f"タイムライン接続に失敗: {name}: {exc}"
            ) from exc

    for shape in (industry_shape, product_shape, timeline_shape):
        try:
            shape.Placement = 3
        except Exception:
            pass
    try:
        product_shape.NumberOfColumns = 2
    except Exception as exc:
        log(f"商品スライサーの列数設定を省略: {exc}")

    timeline.TimelineState.SetFilterDateRange(
        cfg["timeline_start"], cfg["timeline_end"]
    )
    for name in ("PT-01", "PT-02", "PT-03", "PT-04", "PT-05", "PT-06"):
        pts[name].RefreshTable()
    log(
        "slicers:",
        industry.SlicerItems.Count,
        product.SlicerItems.Count,
        "timeline PTs:",
        timeline.PivotTables.Count,
    )
    return {"industry": industry, "product": product, "timeline": timeline}

def create_chart(
    dashboard: object,
    pt: object,
    name: str,
    chart_type: int,
    anchor: str,
    title: str,
    width: int = 350,
    height: int = 280,
) -> object:
    """ピボットテーブルからピボットグラフを作成する。

    Args:
        dashboard: 売上分析シート。
        pt: 元ピボットテーブル。
        name: 圖形名（PG-01 など）。
        chart_type: Excel のグラフ種類。
        anchor: 左上セル。
        title: グラフタイトル。
        width: 幅（ポイント）。
        height: 高さ（ポイント）。

    Returns:
        作成した Chart オブジェクト。

    Side Effects:
        売上分析シートへピボットグラフを追加する。
    """
    left = dashboard.Range(anchor).Left + 2
    top = dashboard.Range(anchor).Top + 2
    shape = dashboard.Shapes.AddChart2(-1, chart_type, left, top, width, height)
    shape.Name = name
    try:
        shape.Placement = 3
    except Exception:
        pass
    chart = shape.Chart
    chart.SetSourceData(pt.TableRange1)
    chart.HasTitle = True
    chart.ChartTitle.Text = title
    chart.ChartTitle.Font.Size = 11
    chart.ChartTitle.Font.Bold = True
    chart.ChartTitle.Font.Name = "Yu Gothic UI"
    chart.ChartTitle.Font.Color = rgb(CONFIG["colors"]["dark_green"])
    chart.ChartArea.Font.Name = "Yu Gothic UI"
    chart.ChartArea.Font.Size = 9
    chart.ChartArea.Font.Color = rgb(CONFIG["colors"]["text"])
    try:
        chart.ChartArea.Format.Fill.ForeColor.RGB = rgb(CONFIG["colors"]["white"])
        chart.PlotArea.Format.Fill.ForeColor.RGB = rgb(CONFIG["colors"]["white"])
        chart.ChartArea.Format.Line.Visible = 0
    except Exception:
        pass
    return chart


def style_chart(
    chart: object, chart_type: int, legend: bool, colors: dict[str, tuple[int, int, int]]
) -> None:
    """グラフの軸・凡例・系列色を設定する。

    Args:
        chart: 対象グラフ。
        chart_type: グラフ種類。
        legend: 凡例の表示有無。
        colors: 配色辞書。

    Returns:
        なし。

    Side Effects:
        グラフの書式を変更する。
    """
    chart.HasLegend = legend
    if legend:
        try:
            chart.Legend.Position = XL_BOTTOM
            chart.Legend.Font.Size = 8
            chart.Legend.Font.Name = "Yu Gothic UI"
        except Exception:
            pass

    if chart_type != XL_DOUGHNUT:
        for axis_index in (1, 2):
            try:
                axis = chart.Axes(axis_index)
                axis.HasMajorGridlines = False
                axis.TickLabels.Font.Size = 8
                axis.TickLabels.Font.Name = "Yu Gothic UI"
                axis.Format.Line.ForeColor.RGB = rgb(colors["border"])
            except Exception:
                pass
    try:
        chart.ChartGroups(1).GapWidth = 60
    except Exception:
        pass
    if chart_type == XL_DOUGHNUT:
        try:
            chart.ChartGroups(1).DoughnutHoleSize = 50
        except Exception:
            pass
    if chart_type == XL_BAR_CLUSTERED:
        try:
            chart.Axes(1).ReversePlotOrder = True
        except Exception:
            pass


def create_charts(
    dashboard: object, pts: dict[str, object], cfg: dict[str, Any]
) -> None:
    """PG-01 から PG-04 のピボットグラフを作成する。

    Args:
        dashboard: 売上分析シート。
        pts: ピボットテーブル辞書。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        売上分析シートへ 4 個のピボットグラフを追加する。
    """
    colors = cfg["colors"]

    pg01 = create_chart(
        dashboard, pts["PT-03"], "PG-01", XL_LINE, "A13", "月別推移（万円）"
    )
    style_chart(pg01, XL_LINE, True, colors)
    try:
        pg01.SeriesCollection(1).Format.Line.ForeColor.RGB = rgb(colors["teal"])
        pg01.SeriesCollection(1).Format.Line.Weight = 2.25
        pg01.SeriesCollection(2).Format.Line.ForeColor.RGB = rgb(
            colors["yellow_green"]
        )
        pg01.SeriesCollection(2).Format.Line.Weight = 2.25
    except Exception as exc:
        log(f"PG-01 の系列色設定を省略: {exc}")

    pg02 = create_chart(
        dashboard, pts["PT-08"], "PG-02", XL_COLUMN_CLUSTERED, "H13", "前年比較（万円）"
    )
    style_chart(pg02, XL_COLUMN_CLUSTERED, True, colors)
    try:
        pg02.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgb(
            colors["light_green"]
        )
        pg02.SeriesCollection(2).Format.Fill.ForeColor.RGB = rgb(colors["teal"])
    except Exception as exc:
        log(f"PG-02 の系列色設定を省略: {exc}")

    pg03 = create_chart(
        dashboard, pts["PT-02"], "PG-03", XL_BAR_CLUSTERED, "A29", "商品別ランキング（万円）"
    )
    style_chart(pg03, XL_BAR_CLUSTERED, False, colors)
    try:
        pg03.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgb(colors["teal"])
    except Exception as exc:
        log(f"PG-03 の系列色設定を省略: {exc}")

    pg04 = create_chart(
        dashboard, pts["PT-01"], "PG-04", XL_DOUGHNUT, "H29", "業種構成（万円）"
    )
    style_chart(pg04, XL_DOUGHNUT, True, colors)
    try:
        pg04.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = rgb(
            colors["teal"]
        )
        pg04.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = rgb(
            colors["yellow_green"]
        )
    except Exception as exc:
        log(f"PG-04 の系列色設定を省略: {exc}")


def create_heatmap(dashboard: object, cfg: dict[str, Any]) -> None:
    """商品×月ヒートマップとスパークラインを作成する。

    Args:
        dashboard: 売上分析シート。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        売上分析シートの A44 から Z58 にヒートマップを追加する。
    """
    colors = cfg["colors"]
    dashboard.Range("A44").Value = "商品×月ヒートマップ（万円）"
    dashboard.Range("A44").Font.Size = 12
    dashboard.Range("A44").Font.Bold = True
    dashboard.Range("A44").Font.Color = rgb(colors["dark_green"])

    dashboard.Range("A45").Value = "商品／年月"
    dashboard.Range("Z45").Value = "推移"
    for index, month in enumerate(cfg["months"], start=2):
        dashboard.Cells(45, index).Value = month
    header = dashboard.Range("A45:Z45")
    header.Interior.Color = rgb(colors["dark_green"])
    header.Font.Color = rgb(colors["white"])
    header.Font.Bold = True
    header.Font.Size = 9
    header.HorizontalAlignment = XL_CENTER
    header.VerticalAlignment = XL_CENTER
    header.Borders.LineStyle = 1
    header.Borders.Color = rgb(colors["white"])

    dashboard.Range("A:A").ColumnWidth = 18
    dashboard.Range("B:Y").ColumnWidth = 11
    dashboard.Range("Z:Z").ColumnWidth = 18

    # 商品名は商品マスタの表示順を参照する（現在の 8 商品）
    product_formula = (
        '=IF(ROW()-45>COUNTA(商品マスタ!$B$5:$B$24),"",'
        "INDEX(商品マスタ!$B$5:$B$24,ROW()-45))"
    )
    for row in range(46, 54):
        dashboard.Cells(row, 1).Formula = product_formula
        dashboard.Cells(row, 1).Font.Size = 9
        dashboard.Cells(row, 1).Font.Bold = True
        dashboard.Cells(row, 1).Interior.Color = rgb(colors["light_green"])
        dashboard.Cells(row, 1).VerticalAlignment = XL_CENTER
        dashboard.Cells(row, 1).Borders.LineStyle = 1
        dashboard.Cells(row, 1).Borders.Color = rgb(colors["border"])

    for row in range(46, 54):
        for col in range(2, 26):
            letter = column_letter(col)
            dashboard.Cells(row, col).Formula = (
                f'=IF($A{row}="","",IFERROR(ROUND(GETPIVOTDATA('
                f'"売上（万円）",制御!$H$2,"商品",$A{row},"年月",{letter}$45),0),""))'
            )
            cell = dashboard.Cells(row, col)
            cell.NumberFormat = '0"万円"'
            cell.Font.Size = 9
            cell.HorizontalAlignment = XL_CENTER
            cell.VerticalAlignment = XL_CENTER
            cell.Borders.LineStyle = 1
            cell.Borders.Color = rgb(colors["border"])

    # 3 色スケール（空白は数式ルールで停止し、塗りなし）
    heatmap = dashboard.Range("B46:Y53")
    try:
        scale = heatmap.FormatConditions.AddColorScale(3)
        scale.ColorScaleCriteria(1).Type = 1
        scale.ColorScaleCriteria(1).FormatColor.Color = rgb(colors["light_green"])
        scale.ColorScaleCriteria(2).Type = 5
        scale.ColorScaleCriteria(2).Value = 50
        scale.ColorScaleCriteria(2).FormatColor.Color = rgb(colors["mid_green"])
        scale.ColorScaleCriteria(3).Type = 2
        scale.ColorScaleCriteria(3).FormatColor.Color = rgb(colors["dark_teal"])
    except Exception as exc:
        log(f"ヒートマップの色スケール設定を省略: {exc}")
    try:
        blank_rule = heatmap.FormatConditions.Add(2, pythoncom.Empty, '=B46=""')
        blank_rule.StopIfTrue = True
        blank_rule.Interior.Pattern = XL_NONE
    except Exception as exc:
        log(f"ヒートマップの空白ルール設定を省略: {exc}")

    # スパークライン
    for row in range(46, 54):
        try:
            dashboard.Range(f"Z{row}").SparklineGroups.Add(
                1, f"B{row}:Y{row}"
            )
            group = dashboard.Range(f"Z{row}").SparklineGroups(1)
            group.SeriesColor.Color = rgb(colors["teal"])
            group.LineWeight = 1.5
        except Exception as exc:
            log(f"スパークライン作成を省略（{row} 行）: {exc}")

    dashboard.Range("A55:L55").Merge()
    dashboard.Range("A55").Value = (
        "ヒートマップは万円表示（四捨五入）。空欄はスライサー・タイムラインの"
        "選択範囲外です。"
    )
    dashboard.Range("A55").Font.Size = 9
    dashboard.Range("A55").Font.Color = rgb(colors["gray"])
    dashboard.Range("A55").WrapText = True
    dashboard.Range("A55").VerticalAlignment = XL_CENTER

    dashboard.Range("A57:L58").Merge()
    dashboard.Range("A57").Value = (
        "商品マスタは方式 A のためピボット未接続。"
        "スライサーは売上データの商品列を使用します。"
        "商品を追加する場合は売上データにも登録してください。"
    )
    dashboard.Range("A57").Font.Size = 9
    dashboard.Range("A57").Font.Color = rgb(colors["gray"])
    dashboard.Range("A57").WrapText = True
    dashboard.Range("A57").VerticalAlignment = XL_CENTER
    dashboard.Range("A57").Interior.Color = rgb(colors["light_green"])

def close_to(actual: object, expected: float, tolerance: float = 1e-6) -> bool:
    """数値が許容差以内かを判定する。

    Args:
        actual: 実測値。
        expected: 期待値。
        tolerance: 許容差。

    Returns:
        許容差以内なら True。

    Side Effects:
        なし。
    """
    try:
        return abs(float(actual) - expected) <= tolerance
    except (TypeError, ValueError):
        return False


def finalize_state(
    app: object,
    book: object,
    dashboard: object,
    control: object,
    pts: dict[str, object],
    caches: dict[str, object],
    cfg: dict[str, Any],
) -> dict[str, Any]:
    """制御シートを非表示にし、計算・不変条件・KPI を検証する。

    Args:
        app: Excel Application。
        book: 対象ブック。
        dashboard: 売上分析シート。
        control: 制御シート。
        pts: ピボットテーブル辞書。
        caches: スライサーキャッシュ辞書。
        cfg: 設定値。

    Returns:
        検証結果の要約。

    Raises:
        ValueError: KPI または不変条件が想定外の場合。

    Side Effects:
        制御シートを非表示にし、Excel の再計算を実行する。
    """
    dashboard.Activate()
    control.Visible = XL_SHEET_HIDDEN
    app.Calculate()

    kpi = {
        "sales": dashboard.Range("A9").Value2,
        "diff": dashboard.Range("D9").Value2,
        "rate": dashboard.Range("G9").Value2,
        "yoy": dashboard.Range("J9").Value2,
    }
    if not close_to(kpi["sales"], cfg["expected_sales"]):
        raise ValueError(f"KPI 売上が想定外です: {kpi['sales']!r}")
    if not close_to(kpi["diff"], cfg["expected_diff"]):
        raise ValueError(f"KPI 差額が想定外です: {kpi['diff']!r}")
    if not close_to(kpi["rate"], cfg["expected_rate"], 1e-9):
        raise ValueError(f"KPI 差額率が想定外です: {kpi['rate']!r}")
    if not close_to(kpi["yoy"], cfg["expected_yoy"], 1e-9):
        raise ValueError(f"KPI 前年比が想定外です: {kpi['yoy']!r}")

    pt_totals = {
        "PT-01_万円": pts["PT-06"].GetPivotData("売上合計").Value / 10000,
        "PT-06_売上": pts["PT-06"].GetPivotData("売上合計").Value,
        "PT-06_差額": pts["PT-06"].GetPivotData("差額合計").Value,
        "PT-07_2025": pts["PT-07"].GetPivotData("売上（万円）", "年", 2025).Value,
        "PT-07_2026": pts["PT-07"].GetPivotData("売上（万円）", "年", 2026).Value,
        "PT-08_2026_01": pts["PT-08"]
        .GetPivotData("売上（万円）", "月", 1, "年", 2026)
        .Value,
        "PT-05_テーブル_2026-01": pts["PT-05"]
        .GetPivotData("売上（万円）", "商品", "テーブル", "年月", "2026-01")
        .Value,
    }
    if not close_to(pt_totals["PT-06_売上"], cfg["expected_sales"]):
        raise ValueError(f"PT-06 売上が想定外です: {pt_totals['PT-06_売上']!r}")
    if not close_to(pt_totals["PT-06_差額"], cfg["expected_diff"]):
        raise ValueError(f"PT-06 差額が想定外です: {pt_totals['PT-06_差額']!r}")

    notes = count_notes(book)
    indirect = count_indirect(book)
    if notes != cfg["expected_notes"]:
        raise ValueError(f"メモ件数が変化しました: {notes}")
    if indirect != cfg["expected_indirect"]:
        raise ValueError(f"INDIRECT 件数が変化しました: {indirect}")
    if any(sheet.ProtectContents for sheet in book.Worksheets):
        raise ValueError("シート保護が設定されています")
    if app.ActiveWindow.FreezePanes:
        raise ValueError("ウィンドウ枠が固定されています")
    lock_targets = (
        (dashboard, "A9"),
        (book.Worksheets(cfg["product_master"]), "B5"),
        (book.Worksheets(cfg["data_sheet"]), "J5"),
    )
    for sheet, cell in lock_targets:
        if sheet.Range(cell).Locked:
            raise ValueError(f"セルがロックされています: {sheet.Name}!{cell}")
        if sheet.Range(cell).FormulaHidden:
            raise ValueError(f"数式が非表示です: {sheet.Name}!{cell}")
    if book.HasVBProject:
        raise ValueError("VBA が含まれています")
    if book.LinkSources():
        raise ValueError("外部リンクが含まれています")
    if control.Visible != XL_SHEET_HIDDEN:
        raise ValueError("制御シートが非表示ではありません")

    summary = {
        "kpi": kpi,
        "pt_totals": pt_totals,
        "notes": notes,
        "indirect": indirect,
        "pivot_tables": list(pts.keys()),
        "slicer_items": {
            "industry": caches["industry"].SlicerItems.Count,
            "product": caches["product"].SlicerItems.Count,
        },
        "timeline_pivot_tables": caches["timeline"].PivotTables.Count,
        "control_visible": control.Visible,
    }
    log("final:", json.dumps(summary, ensure_ascii=False, default=str))
    return summary


def execute_logic(
    app: object, source: Path, dest: Path, cfg: dict[str, Any]
) -> dict[str, Any]:
    """v0.4 のビルド本体を実行する。

    Args:
        app: Excel Application。
        source: v0.3 入力。
        dest: v0.4 出力。
        cfg: 設定値。

    Returns:
        ビルド結果の要約。

    Raises:
        RuntimeError: OOXML 後処理で禁止情報が残った場合。

    Side Effects:
        出力 xlsx を生成し、OOXML を後処理する。
    """
    source = source.resolve()
    dest = dest.resolve()
    if source == dest:
        raise ValueError("入力と出力が同じパスです")
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, dest)

    book = app.Workbooks.Open(str(dest), UpdateLinks=0)
    try:
        validate_workbook(book, app, cfg)
        add_helper_columns(book, cfg)
        create_product_master(book, cfg)
        pts, _cache = create_control_sheet(book, cfg)

        dashboard = book.Worksheets(cfg["dashboard"])
        dashboard.Activate()
        while dashboard.Shapes.Count > 0:
            dashboard.Shapes(1).Delete()
        rebuild_dashboard_cells(dashboard, cfg)
        caches = create_slicers(book, dashboard, pts, cfg)
        create_charts(dashboard, pts, cfg)
        create_heatmap(dashboard, cfg)
        summary = finalize_state(
            app, book, dashboard, book.Worksheets(cfg["control"]), pts, caches, cfg
        )
        book.Save()
    finally:
        book.Close(False)

    sanitized = sanitize_package(dest, OOXML_CONFIG)
    summary["sanitize"] = sanitized
    if sanitized["remaining_hits"] != 0:
        raise RuntimeError(f"禁止情報が残っています: {sanitized}")
    return summary


def main() -> int:
    """コマンドラインからビルドを実行する。

    Returns:
        終了コード。

    Side Effects:
        出力 xlsx を生成する。
    """
    if len(sys.argv) != 3:
        log("usage: build_v04_a.py <input.xlsx> <output.xlsx>")
        return 2
    source = Path(sys.argv[1])
    dest = Path(sys.argv[2])
    try:
        validate_input(source)
    except (FileNotFoundError, ValueError) as exc:
        log(str(exc))
        return 2

    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AutomationSecurity = 3
    app.AskToUpdateLinks = False
    try:
        summary = execute_logic(app, source, dest, CONFIG)
    except Exception as exc:
        import traceback

        traceback.print_exc()
        log(f"build failed: {exc}")
        return 1
    finally:
        app.Quit()
    log(json.dumps(summary, ensure_ascii=False, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
