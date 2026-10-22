"""v0.6 デザインバリアント C（売上分析シートの 0 からの再作成）。

目的:
    v0.4 方式 A のブックを複製し、既存の見積・原価・利益シミュレーターを
    維持したまま、`売上分析` シートだけを削除して 0 から再作成する。
    v0.4 とは配色、KPI カード、グラフ構成、スライサー配置を変えた
    Warm Slate Editorial を採用したデザインバリアント C を生成する。

仕様:
    - 入力は v0.4 成果物 `prototype_v04.xlsx` である。
    - `売上分析` を削除し、同名の新規シートを 商品マスタ の直前へ作成する。
    - 既存の `制御` シートの 8 個の PT と共有キャッシュ 1 個を再利用する。
    - v0.6 専用 fixture で売上明細を再生成し、2026 年売上を 2025 年比約 8% 増とする。
    - 旧スライサーキャッシュを削除し、新しい業種・商品スライサーと
      日付タイムラインを作成する。
    - KPI は 2 × 2 のカード配置、グラフは折れ線 2 個、横棒、100% 積み上げ横棒とする。
    - ヒートマップは 0 円、データなし、選択範囲外を区別する。
    - 74 件のメモは矩形・自動サイズとし、4 グラフは Q 列内へ収める。
    - 保存後に OOXML を後処理し、絶対パスと作成者情報を除去する。

前提:
    - Windows と Microsoft Excel デスクトップアプリ、pywin32 が利用できる。
    - 入力のメモ 74 件、INDIRECT 24 件、保護・固定枠なしを維持する。
    - 既存の計算セル（見積・原価・利益）には触れない。

入出力:
    - 入力: 引数 1（v0.4 xlsx）とスクリプト階層の `sales_fixture_v06.json`。
    - 出力: 引数 2（v0.6 xlsx のパス）。

注意点:
    - 旧スライサーキャッシュを削除してから `売上分析` を削除する。
    - PT は `制御` シートに残るため、共有キャッシュと PT を再作成しない。
    - タイムラインは PT-01 から PT-06 へ接続し、PT-07 と PT-08 は非接続とする。
    - ヒートマップの数式は `IFERROR` で一律に空欄化しない。
    - 保存のたびに `finalize_v03_ooxml.sanitize_package` を実行する。

拡張ポイント:
    - 商品マスタを PT へ接続する場合は方式 B へ移行する。
    - ゲージ、目標管理、チェックボックスは P4 の判断後に追加する。

変更履歴:
    2026-10-22: 初版作成（デザインバリアント C）。
"""

from __future__ import annotations

import json
import os
import sys
import zipfile
from datetime import datetime
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
    "fixture_file": "sales_fixture_v06.json",
    # ヒートマップの月（2025-01 から 2026-12）
    "months": tuple(
        f"{year}-{month:02d}" for year in (2025, 2026) for month in range(1, 13)
    ),
    # タイムライン初期範囲と全期間
    "timeline_start": "2026-01-01",
    "timeline_end": "2026-12-31",
    "full_start": "2025-01-01",
    "full_end": "2026-12-31",
    # v0.6 専用 fixture の期待値（全業種・全商品）
    "expected_sales_2025": 48027300,
    "expected_cost_2025": 28999224,
    "expected_sales": 51869100,
    "expected_cost": 31729100,
    "expected_diff": 20140000,
    "expected_rate": 0.38828512544077304,
    "expected_yoy": 0.07999200454741362,
    "expected_heatmap_2025_jan_table": 155,
    "expected_heatmap_2026_jan_table": 85,
    "expected_heatmap_2026_dec_table": 75,
    # 不変条件
    "expected_notes": 74,
    "expected_indirect": 24,
    # 既存 PT と新スライサーキャッシュ
    "pt_names": tuple(f"PT-{index:02d}" for index in range(1, 9)),
    "slicer_cache_names": ("SL_業種_C", "SL_商品_C", "TL_日付_C"),
    "creator": "ExcelBIBasic_DeepSeek_V4.1",
    # 配色（暖色系のスレート + コーラル + アンバー）
    "colors": {
        "navy": (31, 42, 68),
        "navy_soft": (52, 64, 92),
        "coral": (217, 100, 74),
        "amber": (233, 162, 59),
        "sand": (232, 220, 203),
        "paper": (247, 243, 238),
        "white": (255, 255, 255),
        "text": (45, 52, 66),
        "gray": (104, 110, 120),
        "border": (214, 205, 193),
        "light_peach": (246, 227, 216),
        "muted_blue": (91, 124, 153),
        "negative": (184, 74, 58),
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
XL_TOP = -4160
XL_BOTTOM = -4107
XL_LANDSCAPE = 2
XL_PAPER_A4 = 9
XL_SHEET_HIDDEN = 0
XL_SHEET_VISIBLE = -1
XL_LINE = 4
XL_LINE_MARKERS = 65
XL_AREA = 1
XL_PIE = 5
XL_COLUMN_CLUSTERED = 51
XL_BAR_CLUSTERED = 57
XL_BAR_STACKED_100 = 59
MSO_SHAPE_RECTANGLE = 1
XL_NONE = -4142
XL_EDGE_BOTTOM = 9
XL_EDGE_LEFT = 7
XL_EDGE_TOP = 8
XL_EDGE_RIGHT = 10

# OOXML 後処理では v0.6 の正規プロジェクト名を使う。
OOXML_CONFIG = {**OOXML_CONFIG, "creator": CONFIG["creator"]}


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
        >>> rgb((217, 100, 74))
    """
    red, green, blue = color
    return red + (green << 8) + (blue << 16)


def ensure_creator_metadata(path: Path, creator: str) -> None:
    """OOXML core properties に正規の作成者メタデータを設定する。

    Args:
        path: Excel を閉じた後の成果物。
        creator: 成果物へ記録する正規プロジェクト名。

    Returns:
        なし。

    Raises:
        FileNotFoundError: core properties がない場合。
        RuntimeError: creator 要素を作成できない場合。

    Side Effects:
        xlsx パッケージ内の docProps/core.xml を更新する。
    """
    temporary = path.with_name(f"{path.stem}.creator.tmp{path.suffix}")
    creator_element = f"<dc:creator>{creator}</dc:creator>".encode("utf-8")
    found_core = False
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(
        temporary, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for entry in source.infolist():
            data = source.read(entry.filename)
            if entry.filename == "docProps/core.xml":
                found_core = True
                while b"<dc:creator>" in data:
                    start = data.index(b"<dc:creator>")
                    end = data.index(b"</dc:creator>", start) + len(b"</dc:creator>")
                    data = data[:start] + data[end:]
                root_start = data.find(b"<cp:coreProperties")
                root_end = data.find(b">", root_start) + 1
                if root_start < 0 or root_end <= root_start:
                    raise RuntimeError("core properties の XML 形式が不正です")
                data = data[:root_end] + creator_element + data[root_end:]
            target.writestr(entry, data)
    if not found_core:
        temporary.unlink(missing_ok=True)
        raise FileNotFoundError("core properties がありません")
    temporary.replace(path)


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


def normalize_comment_shapes(book: object, expected_count: int) -> int:
    """全メモを四角形へ統一し、文字列に合わせて図形を自動調整する。

    Args:
        book: 対象ブック。
        expected_count: 維持すべきメモ件数。

    Returns:
        整形したメモ件数。

    Raises:
        ValueError: メモ件数、図形種別、表示文字列が想定と異なる場合。

    Side Effects:
        各メモ図形の種類とサイズを更新する。メモ本文は変更しない。
    """
    normalized = 0
    for sheet in book.Worksheets:
        for index in range(1, sheet.Comments.Count + 1):
            comment = sheet.Comments.Item(index)
            shape = comment.Shape
            original_text = str(comment.Text())
            shape.AutoShapeType = MSO_SHAPE_RECTANGLE
            shape.TextFrame.AutoSize = True
            rendered_text = str(shape.TextFrame.Characters().Text)
            if shape.AutoShapeType != MSO_SHAPE_RECTANGLE:
                raise ValueError("メモ図形を四角形に変更できませんでした")
            if rendered_text != original_text:
                raise ValueError("メモ本文の文字数・内容が変化しました")
            normalized += 1
    if normalized != expected_count:
        raise ValueError(f"メモ整形数が想定外です: {normalized}")
    return normalized


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
        raise FileNotFoundError("入力ファイルがありません")
    if path.suffix.lower() != ".xlsx":
        raise ValueError(f"入力は .xlsx です: {path}")


def validate_v04(book: object, app: object, cfg: dict[str, Any]) -> None:
    """v0.4 入力ブックの前提を検証する。

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
    required = (
        cfg["dashboard"],
        cfg["control"],
        cfg["product_master"],
        cfg["data_sheet"],
        "計算",
        "見積入力",
        "設定",
    )
    for name in required:
        if name not in names:
            raise ValueError(f"必須シートがありません: {name}")
    if book.Worksheets.Count != 13:
        raise ValueError(f"シート数が想定外です: {book.Worksheets.Count}")
    if book.PivotCaches().Count != 1:
        raise ValueError(f"共有キャッシュ数が想定外です: {book.PivotCaches().Count}")
    if book.SlicerCaches.Count != 3:
        raise ValueError(f"スライサーキャッシュ数が想定外です: {book.SlicerCaches.Count}")

    control = book.Worksheets(cfg["control"])
    pt_names = tuple(sorted(pt.Name for pt in control.PivotTables()))
    if pt_names != tuple(sorted(cfg["pt_names"])):
        raise ValueError(f"PT 名が想定外です: {pt_names}")
    if control.Visible != XL_SHEET_HIDDEN:
        raise ValueError("制御シートが非表示ではありません")

    data = book.Worksheets(cfg["data_sheet"])
    if cfg["table"] not in [lo.Name for lo in data.ListObjects]:
        raise ValueError(f"テーブルがありません: {cfg['table']}")
    if data.ListObjects(cfg["table"]).Range.Address != "$A$4:$L$580":
        raise ValueError(
            f"SalesHistory の範囲が想定外です: {data.ListObjects(cfg['table']).Range.Address}"
        )

    product = book.Worksheets(cfg["product_master"])
    if cfg["product_table"] not in [lo.Name for lo in product.ListObjects]:
        raise ValueError(f"テーブルがありません: {cfg['product_table']}")
    if product.ListObjects(cfg["product_table"]).Range.Address != "$A$4:$E$12":
        raise ValueError(
            f"ProductMaster の範囲が想定外です: {product.ListObjects(cfg['product_table']).Range.Address}"
        )

    old_cache_names = {book.SlicerCaches(i).Name for i in range(1, book.SlicerCaches.Count + 1)}
    expected_old = {"SL_業種", "SL_商品", "TL_日付"}
    if old_cache_names != expected_old:
        raise ValueError(f"旧スライサーキャッシュ名が想定外です: {old_cache_names}")

    notes = count_notes(book)
    indirect = count_indirect(book)
    if notes != cfg["expected_notes"]:
        raise ValueError(f"メモ件数が想定外です: {notes}")
    if indirect != cfg["expected_indirect"]:
        raise ValueError(f"INDIRECT 件数が想定外です: {indirect}")
    if any(sheet.ProtectContents for sheet in book.Worksheets):
        raise ValueError("入力にシート保護があります")
    if book.HasVBProject:
        raise ValueError("入力に VBA が含まれています")
    if book.LinkSources():
        raise ValueError("入力に外部リンクがあります")
    if book.Queries.Count != 0:
        raise ValueError("入力に Power Query があります")
    if book.FileFormat != 51:
        raise ValueError(f"入力が .xlsx ではありません: {book.FileFormat}")

    calc = book.Worksheets("計算")
    if not close_to(calc.Range("B32").Value2, 72000):
        raise ValueError(f"既存の利益が想定外です: {calc.Range('B32').Value2!r}")
    if not close_to(calc.Range("B39").Value2, 176000):
        raise ValueError(f"既存の税込見積が想定外です: {calc.Range('B39').Value2!r}")


def load_v06_fixture(cfg: dict[str, Any]) -> list[dict[str, Any]]:
    """v0.6 fixture を読み、件数・年度合計・明細の整合性を検証する。

    Args:
        cfg: ビルド設定。

    Returns:
        2025年と2026年の売上明細。

    Raises:
        FileNotFoundError: fixture がない場合。
        ValueError: 行数、売上式、年次合計が期待値と異なる場合。

    Side Effects:
        fixture JSON を読み取る。
    """
    fixture_path = Path(__file__).resolve().parent / cfg["fixture_file"]
    if not fixture_path.is_file():
        raise FileNotFoundError("v0.6 fixture がありません")
    fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
    records = fixture.get("records")
    metadata = fixture.get("meta", {})
    if not isinstance(records, list) or len(records) != 576:
        raise ValueError("v0.6 fixture は576行でなければなりません")

    annual_sales: dict[int, int] = {2025: 0, 2026: 0}
    annual_cost: dict[int, int] = {2025: 0, 2026: 0}
    row_counts: dict[tuple[int, int, str], int] = {}
    for row in records:
        year = int(row["year"])
        month = int(row["month"])
        product = str(row["product"])
        quantity = int(row["qty"])
        price = int(row["price"])
        sales = int(row["sales"])
        cost = int(row["cost"])
        if year not in annual_sales or not 1 <= month <= 12:
            raise ValueError("v0.6 fixture の年月が不正です")
        if quantity <= 0 or price <= 0 or cost < 0 or cost > sales:
            raise ValueError("v0.6 fixture の数量・価格・原価が不正です")
        if sales != quantity * price:
            raise ValueError("v0.6 fixture の売上が数量×単価と一致しません")
        if int(row.get("day", 0)) not in (5, 15, 25):
            raise ValueError("v0.6 fixture の日付が不正です")
        if str(row["industry"]) not in ("木工所", "農家"):
            raise ValueError("v0.6 fixture の業種が不正です")
        annual_sales[year] += sales
        annual_cost[year] += cost
        key = (year, month, product)
        row_counts[key] = row_counts.get(key, 0) + 1

    if any(count != 3 for count in row_counts.values()) or len(row_counts) != 192:
        raise ValueError("v0.6 fixture は商品・年月ごとに3明細必要です")
    if annual_sales[2025] != cfg["expected_sales_2025"]:
        raise ValueError("v0.6 fixture の2025年売上合計が想定外です")
    if annual_cost[2025] != cfg["expected_cost_2025"]:
        raise ValueError("v0.6 fixture の2025年原価合計が想定外です")
    if annual_sales[2026] != cfg["expected_sales"]:
        raise ValueError("v0.6 fixture の2026年売上合計が想定外です")
    if annual_cost[2026] != cfg["expected_cost"]:
        raise ValueError("v0.6 fixture の2026年原価合計が想定外です")
    actual_growth = annual_sales[2026] / annual_sales[2025] - 1
    if not close_to(actual_growth, 0.08, 0.0001):
        raise ValueError("v0.6 fixture の2026年売上成長率が約8%ではありません")
    if int(metadata.get("row_count", 0)) != len(records):
        raise ValueError("v0.6 fixture のメタデータ行数が一致しません")
    return records


def replace_sales_data(
    book: object, app: object, cfg: dict[str, Any]
) -> dict[str, int]:
    """v0.6 コピー内の SalesHistory 明細を fixture で置き換えて集計を更新する。

    Args:
        book: v0.4 から複製した作業ブック。
        app: 専用 Excel Application。
        cfg: ビルド設定。

    Returns:
        年別の売上合計を含む要約。

    Raises:
        ValueError: テーブル件数または更新後の集計が不正な場合。

    Side Effects:
        v0.6 作業ブックの売上データ、共有 PivotCache、PT を更新する。
    """
    records = load_v06_fixture(cfg)
    data_sheet = book.Worksheets(cfg["data_sheet"])
    table = data_sheet.ListObjects(cfg["table"])
    if table.ListRows.Count != len(records):
        raise ValueError("SalesHistory の明細数がfixtureと一致しません")

    input_rows = tuple(
        (
            datetime(int(row["year"]), int(row["month"]), int(row["day"]), 9, 0),
            str(row["industry"]),
            str(row["product"]),
            int(row["qty"]),
            int(row["price"]),
        )
        for row in records
    )
    costs = tuple((int(row["cost"]),) for row in records)
    categories = tuple(("架空サンプル",) for _ in records)
    last_row = 4 + len(records)
    data_sheet.Range(f"A5:E{last_row}").Value = input_rows
    data_sheet.Range(f"G5:G{last_row}").Value = costs
    data_sheet.Range(f"I5:I{last_row}").Value = categories

    # PT-05 と KPI は共有キャッシュを使うため、全 PT を同じキャッシュから更新する。
    control = book.Worksheets(cfg["control"])
    control.Visible = XL_SHEET_VISIBLE
    # 構造化数式の計算値を先に確定してから PivotCache を更新する。
    app.CalculateFullRebuild()
    book.PivotCaches().Item(1).Refresh()
    pivot_names: list[str] = []
    for pivot in control.PivotTables():
        pivot.RefreshTable()
        pivot_names.append(pivot.Name)
    app.CalculateFullRebuild()

    year_sales = {2025: 0, 2026: 0}
    for row in records:
        year_sales[int(row["year"])] += int(row["sales"])
    if tuple(sorted(pivot_names)) != tuple(sorted(cfg["pt_names"])):
        raise ValueError("fixture 更新後の PT 構成が変化しました")
    return {str(year): total for year, total in year_sales.items()}


def delete_old_dashboard(book: object, cfg: dict[str, Any]) -> object:
    """旧スライサーキャッシュと売上分析シートを削除し、新しいシートを作る。

    Args:
        book: 対象ブック。
        cfg: 設定値。

    Returns:
        新しい `売上分析` シート。

    Raises:
        ValueError: 削除後の状態が想定外の場合。

    Side Effects:
        旧スライサーキャッシュ、旧売上分析シート、旧図形を削除する。
    """
    for index in range(book.SlicerCaches.Count, 0, -1):
        cache = book.SlicerCaches.Item(index)
        log(f"旧スライサーキャッシュを削除: {cache.Name}")
        cache.Delete()
    if book.SlicerCaches.Count != 0:
        raise ValueError("旧スライサーキャッシュが残っています")

    old = book.Worksheets(cfg["dashboard"])
    old.Delete()
    if book.Worksheets.Count != 12:
        raise ValueError(f"旧売上分析の削除後のシート数が想定外です: {book.Worksheets.Count}")

    dashboard = book.Worksheets.Add()
    dashboard.Name = cfg["dashboard"]
    dashboard.Move(Before=book.Worksheets(cfg["product_master"]))
    if dashboard.Index != 10:
        raise ValueError(f"売上分析の位置が想定外です: {dashboard.Index}")
    return dashboard


def style_dashboard_shell(dashboard: object, cfg: dict[str, Any]) -> None:
    """売上分析シートの背景、見出し、サイドバー、セクション見出しを設定する。

    Args:
        dashboard: 売上分析シート。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        セルの結合、塗り、罫線、列幅、行高を変更する。
    """
    colors = cfg["colors"]
    dashboard.Cells.UnMerge()
    dashboard.Cells.Clear()
    try:
        dashboard.Cells.Validation.Delete()
    except Exception:
        pass

    dashboard.Cells.Font.Name = "Yu Gothic UI"
    dashboard.Cells.Font.Size = 10
    dashboard.Cells.Font.Color = rgb(colors["text"])
    dashboard.Cells.Interior.Color = rgb(colors["paper"])
    dashboard.Cells.Locked = False
    dashboard.Cells.FormulaHidden = False

    widths = {
        "A": 24,
        "B": 11,
        "C": 11,
        "D": 14,
        "E": 12,
        "F": 12,
        "G": 12,
        "H": 14,
        "I": 12,
        "J": 12,
        "K": 14,
        "L": 12,
        "M": 12,
        "N": 12,
        "O": 12,
        "P": 12,
        "Q": 12,
        "R": 2,
        "S": 2,
        "T": 11,
        "U": 11,
        "V": 11,
        "W": 11,
        "X": 11,
        "Y": 11,
        "Z": 18,
    }
    for column, width in widths.items():
        dashboard.Columns(column).ColumnWidth = width

    heights = {
        2: 36,
        3: 18,
        4: 4,
        5: 24,
        6: 18,
        7: 18,
        8: 18,
        9: 8,
        10: 18,
        11: 18,
        12: 18,
        13: 18,
        14: 22,
        31: 22,
        48: 24,
        49: 22,
        50: 20,
        60: 22,
        61: 20,
        62: 20,
        63: 20,
        64: 20,
    }
    for row, height in heights.items():
        dashboard.Rows(row).RowHeight = height
    for row in range(15, 31):
        dashboard.Rows(row).RowHeight = 18
    for row in range(32, 48):
        dashboard.Rows(row).RowHeight = 18
    for row in range(51, 59):
        dashboard.Rows(row).RowHeight = 20

    # 左サイドバー（フィルタ帯）
    dashboard.Range("A5:C40").Interior.Color = rgb(colors["navy"])
    dashboard.Range("A5:C5").Merge()
    dashboard.Range("A5").Value = "FILTERS"
    dashboard.Range("A5").Font.Size = 11
    dashboard.Range("A5").Font.Bold = True
    dashboard.Range("A5").Font.Color = rgb(colors["white"])
    dashboard.Range("A5").HorizontalAlignment = XL_LEFT
    dashboard.Range("A5").IndentLevel = 1
    dashboard.Range("A5").VerticalAlignment = XL_CENTER

    for cell, label in (("A6", "業種"), ("A13", "商品"), ("A24", "期間")):
        dashboard.Range(f"{cell}:C{cell[1:]}").Merge()
        dashboard.Range(cell).Value = label
        dashboard.Range(cell).Font.Size = 9
        dashboard.Range(cell).Font.Bold = True
        dashboard.Range(cell).Font.Color = rgb(colors["sand"])
        dashboard.Range(cell).HorizontalAlignment = XL_LEFT
        dashboard.Range(cell).IndentLevel = 2
        dashboard.Range(cell).VerticalAlignment = XL_CENTER

    dashboard.Range("A32:C38").Merge()
    dashboard.Range("A32").Value = (
        "前年比・前年比較は\nタイムライン非連動。\n"
        "スライサーは全PTに接続。\n商品マスタはPT未接続。"
    )
    dashboard.Range("A32").Font.Size = 9
    dashboard.Range("A32").Font.Color = rgb(colors["sand"])
    dashboard.Range("A32").WrapText = True
    dashboard.Range("A32").VerticalAlignment = XL_CENTER
    dashboard.Range("A32").HorizontalAlignment = XL_LEFT
    dashboard.Range("A32").IndentLevel = 1

    # 上部アクセントとタイトル
    dashboard.Range("A4:Z4").Interior.Color = rgb(colors["coral"])
    dashboard.Range("A2:F2").Merge()
    dashboard.Range("A2").Value = "売上分析"
    dashboard.Range("A2").Font.Size = 26
    dashboard.Range("A2").Font.Bold = True
    dashboard.Range("A2").Font.Color = rgb(colors["navy"])
    dashboard.Range("A2").VerticalAlignment = XL_CENTER

    dashboard.Range("A3:O3").Merge()
    dashboard.Range("A3").Value = "数字で読む、商品と季節の動き"
    dashboard.Range("A3").Font.Size = 11
    dashboard.Range("A3").Font.Color = rgb(colors["gray"])
    dashboard.Range("A3").VerticalAlignment = XL_CENTER

    dashboard.Range("G2:O2").Merge()
    dashboard.Range("G2").Value = "DESIGN C / WARM SLATE"
    dashboard.Range("G2").Font.Size = 9
    dashboard.Range("G2").Font.Bold = True
    dashboard.Range("G2").Font.Color = rgb(colors["gray"])
    dashboard.Range("G2").HorizontalAlignment = XL_LEFT
    dashboard.Range("G2").VerticalAlignment = XL_CENTER

    dashboard.Range("Y2:Z2").Merge()
    dashboard.Range("Y2").Value = "2025-2026 / 架空データ"
    dashboard.Range("Y2").Font.Size = 9
    dashboard.Range("Y2").Font.Bold = True
    dashboard.Range("Y2").Font.Color = rgb(colors["white"])
    dashboard.Range("Y2").Interior.Color = rgb(colors["negative"])
    dashboard.Range("Y2").HorizontalAlignment = XL_CENTER
    dashboard.Range("Y2").VerticalAlignment = XL_CENTER

    # セクション見出し
    sections = (
        ("D14:J14", "01 月別推移"),
        ("K14:Q14", "02 前年比較"),
        ("D31:J31", "03 商品別"),
        ("K31:Q31", "04 業種構成"),
    )
    for cell_range, title in sections:
        dashboard.Range(cell_range).Merge()
        first = cell_range.split(":")[0]
        dashboard.Range(first).Value = title
        dashboard.Range(cell_range).Interior.Color = rgb(colors["navy"])
        dashboard.Range(first).Font.Size = 13
        dashboard.Range(first).Font.Bold = True
        dashboard.Range(first).Font.Color = rgb(colors["white"])
        dashboard.Range(first).HorizontalAlignment = XL_LEFT
        dashboard.Range(first).IndentLevel = 1
        dashboard.Range(first).VerticalAlignment = XL_CENTER

    dashboard.Tab.Color = rgb(colors["coral"])


def create_kpi_card(
    dashboard: object,
    cfg: dict[str, Any],
    accent_range: str,
    label_range: str,
    value_range: str,
    note_range: str,
    label: str,
    note: str,
    accent_color: tuple[int, int, int],
    value_cell: str,
    formula: str,
    number_format: str,
) -> None:
    """KPI カードを 1 枚作成する。

    Args:
        dashboard: 売上分析シート。
        cfg: 設定値。
        accent_range: 左側のアクセント帯。
        label_range: ラベルの結合範囲。
        value_range: 値の結合範囲。
        note_range: 補足の結合範囲。
        label: ラベル文字列。
        note: 補足文字列。
        accent_color: アクセント色。
        value_cell: 値のセル。
        formula: 値の数式。
        number_format: 数値書式。

    Returns:
        なし。

    Side Effects:
        セルの結合、塗り、数式、書式を変更する。
    """
    colors = cfg["colors"]
    for cell_range in (label_range, value_range, note_range):
        dashboard.Range(cell_range).Merge()
    dashboard.Range(accent_range).Interior.Color = rgb(accent_color)
    for cell_range in (label_range, value_range, note_range):
        dashboard.Range(cell_range).Interior.Color = rgb(colors["white"])

    dashboard.Range(label_range).Value = label
    dashboard.Range(label_range).Font.Size = 10
    dashboard.Range(label_range).Font.Bold = True
    dashboard.Range(label_range).Font.Color = rgb(colors["gray"])
    dashboard.Range(label_range).HorizontalAlignment = XL_LEFT
    dashboard.Range(label_range).IndentLevel = 1
    dashboard.Range(label_range).VerticalAlignment = XL_CENTER

    dashboard.Range(value_cell).Formula = formula
    dashboard.Range(value_cell).Font.Size = 30
    dashboard.Range(value_cell).Font.Bold = True
    dashboard.Range(value_cell).Font.Color = rgb(colors["navy"])
    dashboard.Range(value_cell).NumberFormat = number_format
    dashboard.Range(value_cell).HorizontalAlignment = XL_LEFT
    dashboard.Range(value_cell).IndentLevel = 1
    dashboard.Range(value_cell).VerticalAlignment = XL_CENTER

    dashboard.Range(note_range).Value = note
    dashboard.Range(note_range).Font.Size = 9
    dashboard.Range(note_range).Font.Color = rgb(colors["gray"])
    dashboard.Range(note_range).HorizontalAlignment = XL_LEFT
    dashboard.Range(note_range).IndentLevel = 1
    dashboard.Range(note_range).VerticalAlignment = XL_CENTER

    card_range = f"{accent_range.split(':')[0]}:{note_range.split(':')[1]}"
    for edge in (XL_EDGE_LEFT, XL_EDGE_TOP, XL_EDGE_RIGHT, XL_EDGE_BOTTOM):
        dashboard.Range(card_range).Borders(edge).LineStyle = 1
        dashboard.Range(card_range).Borders(edge).Color = rgb(colors["border"])


def create_kpi_cards(dashboard: object, cfg: dict[str, Any]) -> None:
    """4 枚の KPI カードを 2 × 2 に配置する。

    Args:
        dashboard: 売上分析シート。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        KPI の数式と書式を設定する。
    """
    colors = cfg["colors"]
    create_kpi_card(
        dashboard,
        cfg,
        "D5:D8",
        "E5:G5",
        "E6:G7",
        "E8:G8",
        "税抜売上（円）",
        "スライサー・期間連動",
        colors["coral"],
        "E6",
        '=ROUND(GETPIVOTDATA("売上合計",制御!$H$14),0)',
        "#,##0",
    )
    create_kpi_card(
        dashboard,
        cfg,
        "H5:H8",
        "I5:K5",
        "I6:K7",
        "I8:K8",
        "差額（円）",
        "計上原価控除後",
        colors["amber"],
        "I6",
        '=ROUND(GETPIVOTDATA("差額合計",制御!$H$14),0)',
        "#,##0",
    )
    create_kpi_card(
        dashboard,
        cfg,
        "D10:D13",
        "E10:G10",
        "E11:G12",
        "E13:G13",
        "差額率",
        "差額 ÷ 売上",
        colors["muted_blue"],
        "E11",
        '=IF(E6=0,"算出不可",I6/E6)',
        "0.0%",
    )
    create_kpi_card(
        dashboard,
        cfg,
        "H10:H13",
        "I10:K10",
        "I11:K12",
        "I13:K13",
        "前年比",
        "2026 vs 2025 / 期間非連動",
        colors["navy_soft"],
        "I11",
        '=IFERROR(GETPIVOTDATA("売上（万円）",制御!$H$20,"年",2026)'
        '/GETPIVOTDATA("売上（万円）",制御!$H$20,"年",2025)-1,"算出不可")',
        "0.0%",
    )


def create_insight_panel(dashboard: object, cfg: dict[str, Any]) -> None:
    """右上の読み取りポイントパネルを作成する。

    Args:
        dashboard: 売上分析シート。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        セルの結合、塗り、文字列を変更する。
    """
    colors = cfg["colors"]
    panel = dashboard.Range("L5:Q13")
    panel.Interior.Color = rgb(colors["sand"])
    panel.Borders.LineStyle = 1
    panel.Borders.Color = rgb(colors["border"])

    dashboard.Range("L5:Q5").Merge()
    dashboard.Range("L5").Value = "DATA NOTE"
    dashboard.Range("L5").Interior.Color = rgb(colors["negative"])
    dashboard.Range("L5").Font.Size = 8
    dashboard.Range("L5").Font.Bold = True
    dashboard.Range("L5").Font.Color = rgb(colors["white"])
    dashboard.Range("L5").HorizontalAlignment = XL_LEFT
    dashboard.Range("L5").IndentLevel = 1
    dashboard.Range("L5").VerticalAlignment = XL_CENTER

    dashboard.Range("L6:Q6").Merge()
    dashboard.Range("L6").Value = "読み取りのポイント"
    dashboard.Range("L6").Font.Size = 11
    dashboard.Range("L6").Font.Bold = True
    dashboard.Range("L6").Font.Color = rgb(colors["navy"])
    dashboard.Range("L6").HorizontalAlignment = XL_LEFT
    dashboard.Range("L6").IndentLevel = 1
    dashboard.Range("L6").VerticalAlignment = XL_CENTER

    dashboard.Range("L8:Q12").Merge()
    dashboard.Range("L8").Value = (
        "1. 差額 = 売上 - 計上原価\n"
        "2. 前年比 = 2026年 ÷ 2025年 - 1\n"
        "3. タイムラインは前年比・前年比較に非連動\n"
        "4. 商品マスタは方式AのためPT未接続"
    )
    dashboard.Range("L8").Font.Size = 9
    dashboard.Range("L8").Font.Color = rgb(colors["text"])
    dashboard.Range("L8").WrapText = True
    dashboard.Range("L8").VerticalAlignment = XL_TOP
    dashboard.Range("L8").HorizontalAlignment = XL_LEFT
    dashboard.Range("L8").IndentLevel = 1


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
    """左サイドバーに業種・商品スライサーと日付タイムラインを作成する。

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
    industry_top = dashboard.Range("A7").Top + 4
    industry_left = dashboard.Range("A7").Left + 4
    product_top = dashboard.Range("A14").Top + 4
    product_left = dashboard.Range("A14").Left + 4
    timeline_top = dashboard.Range("A25").Top + 4
    timeline_left = dashboard.Range("A25").Left + 4

    industry = book.SlicerCaches.Add2(pts["PT-01"], "業種", cfg["slicer_cache_names"][0])
    industry_shape = industry.Slicers.Add(
        dashboard,
        pythoncom.Empty,
        "SL_業種_C_Shape",
        "業種",
        industry_top,
        industry_left,
        150,
        90,
    )
    connect_slicer(industry, pts, "PT-01")

    product = book.SlicerCaches.Add2(pts["PT-02"], "商品", cfg["slicer_cache_names"][1])
    product_shape = product.Slicers.Add(
        dashboard,
        pythoncom.Empty,
        "SL_商品_C_Shape",
        "商品",
        product_top,
        product_left,
        180,
        150,
    )
    connect_slicer(product, pts, "PT-02")

    timeline = book.SlicerCaches.Add2(
        pts["PT-03"], "日付", cfg["slicer_cache_names"][2], 2
    )
    timeline_shape = timeline.Slicers.Add(
        dashboard,
        pythoncom.Empty,
        "TL_日付_C_Shape",
        "日付",
        timeline_top,
        timeline_left,
        225,
        100,
    )
    for name in ("PT-01", "PT-02", "PT-04", "PT-05", "PT-06"):
        try:
            timeline.PivotTables.AddPivotTable(pts[name])
        except Exception as exc:
            raise RuntimeError(f"タイムライン接続に失敗: {name}: {exc}") from exc

    for shape in (industry_shape, product_shape, timeline_shape):
        try:
            shape.Placement = 3
        except Exception:
            pass
    try:
        product_shape.NumberOfColumns = 2
    except Exception as exc:
        log(f"商品スライサーの列数設定を省略: {exc}")
    for slicer, style_name in (
        (industry_shape, "SlicerStyleLight2"),
        (product_shape, "SlicerStyleLight2"),
    ):
        try:
            slicer.Style = style_name
        except Exception as exc:
            log(f"スライサースタイル設定を省略: {exc}")
    try:
        timeline_shape.Style = "TimeSlicerStyleLight1"
    except Exception as exc:
        log(f"タイムラインスタイル設定を省略: {exc}")

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
    left: float,
    top: float,
    width: float,
    height: float,
    title: str,
    cfg: dict[str, Any],
    source_range: object | None = None,
) -> object:
    """ピボットテーブルまたは連動補助範囲からデザイン C のグラフを作成する。

    Args:
        dashboard: 売上分析シート。
        pt: 元ピボットテーブル。
        name: グラフ名。
        chart_type: Excel のグラフ種類。
        left: 左位置（ポイント）。
        top: 上位置（ポイント）。
        width: 幅（ポイント）。
        height: 高さ（ポイント）。
        title: グラフタイトル。
        cfg: 設定値。
        source_range: 使用するグラフデータ範囲。省略時はピボット範囲。

    Returns:
        作成した Chart オブジェクト。

    Side Effects:
        売上分析シートへピボットグラフを追加する。
    """
    colors = cfg["colors"]
    shape = dashboard.Shapes.AddChart2(-1, chart_type, left, top, width, height)
    shape.Name = name
    try:
        shape.Placement = 3
    except Exception:
        pass
    chart = shape.Chart
    chart.SetSourceData(source_range if source_range is not None else pt.TableRange1)
    chart.ChartType = chart_type
    try:
        # ピボットグラフのフィールドボタンは UI を乱すため非表示にする。
        chart.ShowAllFieldButtons = False
    except Exception as exc:
        log(f"グラフのフィールドボタン非表示を省略: {exc}")
    chart.HasTitle = True
    chart.ChartTitle.Text = title
    chart.ChartTitle.Font.Size = 11
    chart.ChartTitle.Font.Bold = True
    chart.ChartTitle.Font.Name = "Yu Gothic UI"
    chart.ChartTitle.Font.Color = rgb(colors["navy"])
    chart.ChartArea.Font.Name = "Yu Gothic UI"
    chart.ChartArea.Font.Size = 9
    chart.ChartArea.Font.Color = rgb(colors["text"])
    chart.ChartArea.RoundedCorners = False
    try:
        chart.ChartArea.Format.Fill.ForeColor.RGB = rgb(colors["white"])
        chart.PlotArea.Format.Fill.ForeColor.RGB = rgb(colors["white"])
        chart.ChartArea.Format.Line.Visible = -1
        chart.ChartArea.Format.Line.ForeColor.RGB = rgb(colors["border"])
    except Exception as exc:
        log(f"グラフ背景設定を省略: {exc}")
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

    if chart_type != XL_PIE:
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
        chart.ChartGroups(1).GapWidth = 80
    except Exception:
        pass
    if chart_type == XL_PIE:
        try:
            chart.ChartGroups(1).FirstSliceAngle = 0
        except Exception:
            pass


def create_charts(
    dashboard: object, pts: dict[str, object], cfg: dict[str, Any]
) -> None:
    """デザイン C の 4 グラフを作成する。

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
    margin = 6
    chart1_left = dashboard.Range("D15").Left + margin
    chart1_top = dashboard.Range("D15").Top + margin
    chart1_width = dashboard.Range("K15").Left - chart1_left - margin
    chart1_height = dashboard.Range("D31").Top - chart1_top - margin
    chart2_left = dashboard.Range("K15").Left + margin
    chart2_top = dashboard.Range("K15").Top + margin
    chart2_width = dashboard.Range("R15").Left - chart2_left - margin
    chart2_height = dashboard.Range("K31").Top - chart2_top - margin
    chart3_left = dashboard.Range("D32").Left + margin
    chart3_top = dashboard.Range("D32").Top + margin
    chart3_width = dashboard.Range("K32").Left - chart3_left - margin
    chart3_height = dashboard.Range("D48").Top - chart3_top - margin
    chart4_left = dashboard.Range("K32").Left + margin
    chart4_top = dashboard.Range("K32").Top + margin
    chart4_width = dashboard.Range("R32").Left - chart4_left - margin
    chart4_height = dashboard.Range("K48").Top - chart4_top - margin

    # PT-03 の売上値を参照する通常グラフ用データを作り、差額系列を混ぜない。
    dashboard.Range("AE1").Value = "年月"
    dashboard.Range("AF1").Value = "売上（万円）"
    for row, month in enumerate(cfg["months"][12:], start=2):
        dashboard.Cells(row, 31).Value = month
        dashboard.Cells(row, 32).Formula = (
            f'=IFERROR(GETPIVOTDATA("売上（万円）",制御!$A$26,"年月",AE{row}),0)'
        )

    # 01 月別売上: 売上系列だけを残し、差額との二重表示を避ける。
    pg01 = create_chart(
        dashboard,
        pts["PT-03"],
        "PG-C01",
        XL_LINE,
        chart1_left,
        chart1_top,
        chart1_width,
        chart1_height,
        "月別売上（万円）",
        cfg,
        source_range=dashboard.Range("AE1:AF13"),
    )
    pg01.PlotVisibleOnly = False
    pg01.PlotBy = 2
    style_chart(pg01, XL_LINE, False, colors)
    try:
        series = pg01.SeriesCollection(1)
        series.Format.Line.ForeColor.RGB = rgb(colors["coral"])
        series.Format.Line.Weight = 1.75
        series.MarkerStyle = 8
        series.MarkerSize = 4
    except Exception as exc:
        raise RuntimeError(f"PG-C01 の単一系列設定に失敗: {exc}") from exc

    # 02 前年比較: 2025 / 2026 の系列色を固定し、終点ラベルを付ける。
    pg02 = create_chart(
        dashboard,
        pts["PT-08"],
        "PG-C02",
        XL_LINE_MARKERS,
        chart2_left,
        chart2_top,
        chart2_width,
        chart2_height,
        "2025 年と 2026 年の比較（万円）",
        cfg,
    )
    style_chart(pg02, XL_LINE_MARKERS, True, colors)
    try:
        for index, color in ((1, colors["muted_blue"]), (2, colors["coral"])):
            series = pg02.SeriesCollection(index)
            series.Format.Line.ForeColor.RGB = rgb(color)
            series.Format.Line.Weight = 2.0
            series.MarkerStyle = 8
            series.MarkerSize = 5
            series.MarkerBackgroundColor = rgb(colors["white"])
            series.MarkerForegroundColor = rgb(color)
            end_point = series.Points(series.Points().Count)
            end_point.HasDataLabel = True
            end_point.DataLabel.ShowSeriesName = True
            end_point.DataLabel.Font.Size = 8
            end_point.DataLabel.Font.Name = "Yu Gothic UI"
            end_point.DataLabel.Font.Color = rgb(color)
        pg02.Axes(1).HasTitle = True
        pg02.Axes(1).AxisTitle.Text = "月"
        pg02.Axes(1).AxisTitle.Font.Size = 8
        pg02.Axes(1).AxisTitle.Font.Name = "Yu Gothic UI"
    except Exception as exc:
        log(f"PG-C02 の系列設定を省略: {exc}")

    # 03 商品別: 商品名を縦軸へ置く横棒グラフ。
    pg03 = create_chart(
        dashboard,
        pts["PT-02"],
        "PG-C03",
        XL_BAR_CLUSTERED,
        chart3_left,
        chart3_top,
        chart3_width,
        chart3_height,
        "商品別売上（万円）",
        cfg,
    )
    style_chart(pg03, XL_BAR_CLUSTERED, False, colors)
    try:
        pg03.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgb(colors["amber"])
        pg03.SeriesCollection(1).Format.Line.Visible = 0
        pg03.Axes(1).ReversePlotOrder = True
        pg03.SeriesCollection(1).HasDataLabels = True
        pg03.SeriesCollection(1).DataLabels().ShowValue = True
        pg03.SeriesCollection(1).DataLabels().Position = 2
        pg03.SeriesCollection(1).DataLabels().NumberFormat = '0"万円"'
        pg03.SeriesCollection(1).DataLabels().Font.Size = 8
    except Exception as exc:
        log(f"PG-C03 の系列設定を省略: {exc}")

    # 04 業種構成: PT-01 の値を参照する補助セルで 2 系列の構成比を作る。
    dashboard.Range("AA1").Value = "業種構成"
    dashboard.Range("AB1").Value = "木工所"
    dashboard.Range("AC1").Value = "農家"
    dashboard.Range("AA2").Value = "構成比"
    dashboard.Range("AB2").Formula = "=IFERROR(AB3/SUM($AB$3:$AC$3),0)"
    dashboard.Range("AC2").Formula = "=IFERROR(AC3/SUM($AB$3:$AC$3),0)"
    dashboard.Range("AA3").Value = "売上"
    dashboard.Range("AB3").Formula = (
        '=IFERROR(GETPIVOTDATA("売上（万円）",制御!$A$2,"業種","木工所"),0)'
    )
    dashboard.Range("AC3").Formula = (
        '=IFERROR(GETPIVOTDATA("売上（万円）",制御!$A$2,"業種","農家"),0)'
    )
    dashboard.Columns("AA:AF").Hidden = True
    pg04 = create_chart(
        dashboard,
        pts["PT-01"],
        "PG-C04",
        XL_BAR_STACKED_100,
        chart4_left,
        chart4_top,
        chart4_width,
        chart4_height,
        "業種構成（%）",
        cfg,
        source_range=dashboard.Range("AA1:AC2"),
    )
    pg04.PlotVisibleOnly = False
    pg04.PlotBy = 2
    style_chart(pg04, XL_BAR_STACKED_100, False, colors)
    try:
        for index, color in enumerate((colors["muted_blue"], colors["coral"]), start=1):
            series = pg04.SeriesCollection(index)
            series.Format.Fill.ForeColor.RGB = rgb(color)
            series.Format.Line.Visible = 0
            series.HasDataLabels = True
            labels = series.DataLabels()
            labels.ShowSeriesName = True
            labels.ShowValue = True
            labels.ShowCategoryName = False
            labels.NumberFormat = "0%"
            labels.Position = -4108
            labels.Font.Size = 8
            labels.Font.Name = "Yu Gothic UI"
    except Exception as exc:
        log(f"PG-C04 の系列設定を省略: {exc}")


def fit_charts_to_grid(dashboard: object) -> None:
    """4 グラフを最終列幅に合わせ、各行の右端を Q 列内へ収める。

    Args:
        dashboard: 売上分析シート。

    Returns:
        なし。

    Raises:
        ValueError: グラフが指定グリッドをはみ出す場合。

    Side Effects:
        4 グラフの位置、幅、高さ、セル配置設定を更新する。
    """
    margin = 6
    placements = (
        ("PG-C01", "D15", "K15", "D31"),
        ("PG-C02", "K15", "R15", "K31"),
        ("PG-C03", "D32", "K32", "D48"),
        ("PG-C04", "K32", "R32", "K48"),
    )
    for name, left_cell, right_cell, bottom_cell in placements:
        shape = dashboard.Shapes(name)
        shape.Left = dashboard.Range(left_cell).Left + margin
        shape.Top = dashboard.Range(left_cell).Top + margin
        shape.Width = dashboard.Range(right_cell).Left - shape.Left - margin
        shape.Height = dashboard.Range(bottom_cell).Top - shape.Top - margin
        # Heatmap setup adjusts R/S column widths; floating placement locks final chart size.
        shape.Placement = 3
        right_limit = dashboard.Range(right_cell).Left
        if shape.Left + shape.Width > right_limit:
            raise ValueError(f"{name} が Q 列内へ収まっていません")


def create_heatmap(dashboard: object, cfg: dict[str, Any]) -> None:
    """商品×月ヒートマップと列スパークラインを作成する。

    Args:
        dashboard: 売上分析シート。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        売上分析シートの A48 から Z64 にヒートマップと注記を追加する。
    """
    colors = cfg["colors"]
    dashboard.Range("A48:J48").Merge()
    dashboard.Range("A48").Value = "商品×月ヒートマップ（万円）"
    dashboard.Range("A48").Font.Size = 13
    dashboard.Range("A48").Font.Bold = True
    dashboard.Range("A48").Font.Color = rgb(colors["navy"])
    dashboard.Range("A48").VerticalAlignment = XL_CENTER

    for cell_range, label, fill_color in (
        ("A49:F49", "0万円 = 売上0円", colors["light_peach"]),
        ("G49:L49", "－ = データなし", colors["sand"]),
        ("M49:Z49", "空欄 = 選択範囲外", colors["white"]),
    ):
        dashboard.Range(cell_range).Merge()
        cell = dashboard.Range(cell_range.split(":")[0])
        cell.Value = label
        cell.Interior.Color = rgb(fill_color)
        cell.Font.Size = 9
        cell.Font.Bold = True
        cell.Font.Color = rgb(colors["text"])
        cell.HorizontalAlignment = XL_CENTER
        cell.VerticalAlignment = XL_CENTER
        dashboard.Range(cell_range).Borders.LineStyle = 1
        dashboard.Range(cell_range).Borders.Color = rgb(colors["border"])

    dashboard.Range("B50:Y50").NumberFormat = "@"
    dashboard.Range("A50").Value = "商品 / 年月"
    dashboard.Range("Z50").Value = "推移"
    for index, month in enumerate(cfg["months"], start=2):
        dashboard.Cells(50, index).Value = month
    header = dashboard.Range("A50:Z50")
    header.Interior.Color = rgb(colors["navy"])
    header.Font.Color = rgb(colors["white"])
    header.Font.Bold = True
    header.Font.Size = 9
    header.HorizontalAlignment = XL_CENTER
    header.VerticalAlignment = XL_CENTER
    header.Borders.LineStyle = 1
    header.Borders.Color = rgb(colors["white"])

    dashboard.Range("A:A").ColumnWidth = 24
    dashboard.Range("B:Y").ColumnWidth = 11
    dashboard.Range("Z:Z").ColumnWidth = 18

    product_formula = (
        '=IF(ROW()-50>COUNTA(商品マスタ!$B$5:$B$24),"",'
        "INDEX(商品マスタ!$B$5:$B$24,ROW()-50))"
    )
    for row in range(51, 59):
        dashboard.Cells(row, 1).Formula = product_formula
        dashboard.Cells(row, 1).Font.Size = 9
        dashboard.Cells(row, 1).Font.Bold = True
        dashboard.Cells(row, 1).Font.Color = rgb(colors["navy"])
        dashboard.Cells(row, 1).Interior.Color = rgb(colors["sand"])
        dashboard.Cells(row, 1).VerticalAlignment = XL_CENTER
        dashboard.Cells(row, 1).Borders.LineStyle = 1
        dashboard.Cells(row, 1).Borders.Color = rgb(colors["border"])

    for row in range(51, 59):
        for col in range(2, 26):
            letter = column_letter(col)
            dashboard.Cells(row, col).Formula = (
                f'=IF($A{row}="","",'
                f'IF(ISERROR(GETPIVOTDATA("売上（万円）",制御!$H$2,'
                f'"商品",$A{row},"年月",{letter}$50)),'
                f'IF(COUNTIFS(SalesHistory[商品],$A{row},'
                f'SalesHistory[年月],{letter}$50)=0,"－",""),'
                f'ROUND(GETPIVOTDATA("売上（万円）",制御!$H$2,'
                f'"商品",$A{row},"年月",{letter}$50),0)))'
            )
            cell = dashboard.Cells(row, col)
            cell.NumberFormat = '0"万円"'
            cell.Font.Size = 9
            cell.Font.Color = rgb(colors["text"])
            cell.Interior.Color = rgb(colors["white"])
            cell.HorizontalAlignment = XL_CENTER
            cell.VerticalAlignment = XL_CENTER
            cell.Borders.LineStyle = 1
            cell.Borders.Color = rgb(colors["border"])

    heatmap = dashboard.Range("B51:Y58")
    try:
        blank_rule = heatmap.FormatConditions.Add(2, pythoncom.Empty, "=B51=\"\"")
        blank_rule.StopIfTrue = True
        blank_rule.Interior.Pattern = XL_NONE
        blank_rule.Priority = 1
    except Exception as exc:
        log(f"ヒートマップの空白ルール設定を省略: {exc}")
    try:
        text_rule = heatmap.FormatConditions.Add(2, pythoncom.Empty, "=ISTEXT(B51)")
        text_rule.StopIfTrue = True
        text_rule.Interior.Color = rgb(colors["sand"])
        text_rule.Font.Color = rgb(colors["text"])
        text_rule.Font.Bold = True
        text_rule.Priority = 2
    except Exception as exc:
        log(f"ヒートマップのテキストルール設定を省略: {exc}")
    try:
        scale = heatmap.FormatConditions.AddColorScale(3)
        scale.ColorScaleCriteria(1).Type = 1
        scale.ColorScaleCriteria(1).FormatColor.Color = rgb(colors["light_peach"])
        scale.ColorScaleCriteria(2).Type = 5
        scale.ColorScaleCriteria(2).Value = 50
        scale.ColorScaleCriteria(2).FormatColor.Color = rgb(colors["amber"])
        scale.ColorScaleCriteria(3).Type = 2
        scale.ColorScaleCriteria(3).FormatColor.Color = rgb(colors["coral"])
        scale.Priority = 3
    except Exception as exc:
        log(f"ヒートマップの色スケール設定を省略: {exc}")

    for row in range(51, 59):
        try:
            dashboard.Range(f"Z{row}").SparklineGroups.Add(2, f"B{row}:Y{row}")
            group = dashboard.Range(f"Z{row}").SparklineGroups(1)
            group.SeriesColor.Color = rgb(colors["coral"])
            group.LineWeight = 0.75
        except Exception as exc:
            log(f"列スパークライン作成を省略（{row} 行）: {exc}")

    notes = (
        "差額 = 売上 − 計上原価。売上履歴の計上原価のみを控除し、会計上の粗利・営業利益とは区別する。",
        "売上データは架空データである。商品を追加するときは商品マスタと売上データの両方へ登録する。",
        "業種・商品スライサーは全ピボットへ接続。商品マスタは方式 A のためピボットへ接続していない。",
        "前年比・前年比較はタイムライン非連動である。2025 年のヒートマップは期間を全期間へ広げるまで空欄である。",
        "ヒートマップ: 0万円 = 売上 0 円 / － = データなし / 空欄 = 選択範囲外。状態は色と文字で区別する。",
    )
    for row, text in zip(range(60, 65), notes):
        cell_range = f"A{row}:Z{row}"
        dashboard.Range(cell_range).Merge()
        cell = dashboard.Range(f"A{row}")
        cell.Value = text
        cell.Font.Size = 9
        cell.Font.Color = rgb(colors["gray"])
        cell.WrapText = True
        cell.VerticalAlignment = XL_CENTER
        cell.IndentLevel = 1
        if row % 2 == 0:
            dashboard.Range(cell_range).Interior.Color = rgb(colors["white"])

    try:
        dashboard.PageSetup.PrintArea = "$A$1:$Z$64"
        dashboard.PageSetup.Orientation = XL_LANDSCAPE
        dashboard.PageSetup.PaperSize = XL_PAPER_A4
        dashboard.PageSetup.LeftMargin = dashboard.Application.InchesToPoints(0.4)
        dashboard.PageSetup.RightMargin = dashboard.Application.InchesToPoints(0.4)
        dashboard.PageSetup.TopMargin = dashboard.Application.InchesToPoints(0.4)
        dashboard.PageSetup.BottomMargin = dashboard.Application.InchesToPoints(0.4)
        dashboard.PageSetup.LeftFooter = "ExcelBIBasic_DeepSeek_V4.1"
        dashboard.PageSetup.CenterFooter = "&A"
        dashboard.PageSetup.RightFooter = "Page &P / &N"
        dashboard.PageSetup.Zoom = False
        dashboard.PageSetup.FitToPagesWide = 1
        dashboard.PageSetup.FitToPagesTall = 1
    except Exception as exc:
        log(f"印刷設定を省略: {exc}")


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
        "sales": dashboard.Range("E6").Value2,
        "diff": dashboard.Range("I6").Value2,
        "rate": dashboard.Range("E11").Value2,
        "yoy": dashboard.Range("I11").Value2,
    }
    if not close_to(kpi["sales"], cfg["expected_sales"]):
        raise ValueError(f"KPI 売上が想定外です: {kpi['sales']!r}")
    if not close_to(kpi["diff"], cfg["expected_diff"]):
        raise ValueError(f"KPI 差額が想定外です: {kpi['diff']!r}")
    if not close_to(kpi["rate"], cfg["expected_rate"], 1e-9):
        raise ValueError(f"KPI 差額率が想定外です: {kpi['rate']!r}")
    if not close_to(kpi["yoy"], cfg["expected_yoy"], 1e-9):
        raise ValueError(f"KPI 前年比が想定外です: {kpi['yoy']!r}")

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
    if book.HasVBProject:
        raise ValueError("VBA が含まれています")
    if book.LinkSources():
        raise ValueError("外部リンクが含まれています")
    if book.Queries.Count != 0:
        raise ValueError("Power Query が含まれています")
    if control.Visible != XL_SHEET_HIDDEN:
        raise ValueError("制御シートが非表示ではありません")
    if caches["industry"].PivotTables.Count != 8:
        raise ValueError("業種スライサーの接続数が想定外です")
    if caches["product"].PivotTables.Count != 8:
        raise ValueError("商品スライサーの接続数が想定外です")
    if caches["timeline"].PivotTables.Count != 6:
        raise ValueError("タイムラインの接続数が想定外です")
    if dashboard.Shapes.Count != 7:
        raise ValueError(f"売上分析の図形数が想定外です: {dashboard.Shapes.Count}")
    if dashboard.Range("B51").Value not in (None, ""):
        raise ValueError(f"2025-01 のヒートマップが空欄ではありません: {dashboard.Range('B51').Value!r}")
    if not dashboard.Range("N51").Value:
        raise ValueError("2026-01 のヒートマップが空です")
    if "COUNTIFS" not in dashboard.Range("B51").Formula:
        raise ValueError("ヒートマップの 3 状態判定がありません")
    for cell in dashboard.Range("B51:Y58"):
        if cell.HasFormula and "IFERROR" in cell.Formula.upper():
            raise ValueError(f"ヒートマップに IFERROR が残っています: {cell.Address}")

    summary = {
        "kpi": kpi,
        "notes": notes,
        "indirect": indirect,
        "pivot_tables": list(pts.keys()),
        "slicer_items": {
            "industry": caches["industry"].SlicerItems.Count,
            "product": caches["product"].SlicerItems.Count,
        },
        "timeline_pivot_tables": caches["timeline"].PivotTables.Count,
        "control_visible": control.Visible,
        "dashboard_shapes": dashboard.Shapes.Count,
    }
    log("final:", json.dumps(summary, ensure_ascii=False, default=str))
    return summary


def execute_logic(
    app: object, source: Path, dest: Path, cfg: dict[str, Any]
) -> dict[str, Any]:
    """v0.6 デザインバリアント C のビルド本体を実行する。

    Args:
        app: Excel Application。
        source: v0.4 入力。
        dest: v0.6 出力。
        cfg: 設定値。

    Returns:
        ビルド結果の要約。

    Raises:
        RuntimeError: OOXML 後処理で禁止情報が残った場合。

    Side Effects:
        出力 xlsx を生成し、OOXML を後処理する。
    """
    source = source.resolve()
    # pathlib.Path.resolve() が未作成の相対出力先を絶対化しない Python 版にも対応する。
    dest = Path(os.path.abspath(dest))
    if source == dest:
        raise ValueError("入力と出力が同じパスです")
    if source.name.lower() in {"prototype_v05_design_b.xlsx", "prototype_v04.xlsx"} and dest.name.lower() in {
        "prototype_v05_design_b.xlsx",
        "prototype_v04.xlsx",
    }:
        raise ValueError("v0.4 / v0.5 の参照ブックは上書きできません")
    if dest.exists():
        raise FileExistsError("出力ファイルが既にあります。既存成果物を確認してください")
    dest.parent.mkdir(parents=True, exist_ok=True)
    book = app.Workbooks.Open(Filename=str(source), UpdateLinks=0, ReadOnly=True)
    try:
        validate_v04(book, app, cfg)
        # Excel の SaveAs で新しい作業ブックを作り、v0.4 入力には書き込まない。
        book.SaveAs(Filename=str(dest), FileFormat=51, AddToMru=False)
        dashboard = delete_old_dashboard(book, cfg)
        fixture_summary = replace_sales_data(book, app, cfg)
        comments_normalized = normalize_comment_shapes(book, cfg["expected_notes"])
        control = book.Worksheets(cfg["control"])
        control.Visible = XL_SHEET_VISIBLE
        pts = {pt.Name: pt for pt in control.PivotTables()}
        style_dashboard_shell(dashboard, cfg)
        create_kpi_cards(dashboard, cfg)
        create_insight_panel(dashboard, cfg)
        caches = create_slicers(book, dashboard, pts, cfg)
        create_charts(dashboard, pts, cfg)
        create_heatmap(dashboard, cfg)
        fit_charts_to_grid(dashboard)
        summary = finalize_state(app, book, dashboard, control, pts, caches, cfg)
        summary["fixture"] = fixture_summary
        summary["comments_normalized"] = comments_normalized
        book.Save()
    finally:
        book.Close(False)

    sanitized = sanitize_package(dest, OOXML_CONFIG)
    ensure_creator_metadata(dest, cfg["creator"])
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
        log("usage: build_v06_design_c.py <input.xlsx> <output.xlsx>")
        return 2
    source = Path(sys.argv[1])
    dest = Path(sys.argv[2])
    try:
        validate_input(source)
    except (FileNotFoundError, ValueError) as exc:
        log(str(exc))
        return 2

    if dest.exists():
        log("出力ファイルが既にあります。既存成果物を確認してください")
        return 2

    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.AutomationSecurity = 3
        app.AskToUpdateLinks = False
        summary = execute_logic(app, source, dest, CONFIG)
    except Exception as exc:
        import traceback

        traceback.print_exc()
        log(f"build failed: {exc}")
        return 1
    finally:
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()
    log(json.dumps(summary, ensure_ascii=False, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
