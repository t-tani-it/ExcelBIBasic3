"""v0.6 デザインバリアント C の Windows 検証。

目的:
    `prototype_v06_design_c.xlsx` が v0.4 の機能・不変条件を維持し、
    売上分析シートがデザイン C の要件を満たすことを検証する。

仕様:
    - シート、テーブル、ピボット、スライサー、タイムライン、グラフ、
      ヒートマップ、スパークライン、KPI、既存計算、禁止事項を確認する。
    - ヒートマップの 0 円・データなし・選択範囲外を一時コピーで確認する。
    - 結果を JSON へ保存し、失敗時は終了コード 1 を返す。

前提:
    - Windows と Microsoft Excel デスクトップアプリ、pywin32 が利用できる。
    - 検証対象はビルド済みの `.xlsx` である。

入出力:
    - 入力: 引数 1（検証対象 xlsx）。
    - 出力: 引数 2（結果 JSON、省略時は windows_v06_results.json）。

注意点:
    - 検証対象は読み取り専用で開く。
    - 0 円・データなしの検証は一時コピーでのみ行い、保存しない。
    - 絶対パスや個人情報を結果 JSON に含めない。

拡張ポイント:
    - 方式 B へ移行した場合は、データモデルとリレーションの検証を追加する。

変更履歴:
    2026-10-22: 初版作成（デザインバリアント C）。
"""

from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import zipfile
from datetime import date
from pathlib import Path
from typing import Any, Callable

import win32com.client

# ---------------------------------------------------------------------------
# 設定値
# ---------------------------------------------------------------------------
CONFIG: dict[str, Any] = {
    "dashboard": "売上分析",
    "control": "制御",
    "product_master": "商品マスタ",
    "data_sheet": "売上データ",
    "calc_sheet": "計算",
    "table": "SalesHistory",
    "product_table": "ProductMaster",
    "fixture_file": "sales_fixture_v06.json",
    "dashboard_index": 10,
    "sheet_count": 13,
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
    "expected_profit": 72000,
    "expected_quote": 176000,
    "expected_notes": 74,
    "expected_indirect": 24,
    "expected_creator": "ExcelBIBasic_DeepSeek_V4.1",
    "months": tuple(
        f"{year}-{month:02d}" for year in (2025, 2026) for month in range(1, 13)
    ),
    "heatmap_start_row": 51,
    "heatmap_end_row": 58,
    "heatmap_start_col": 2,
    "heatmap_end_col": 25,
    "slicer_cache_names": ("SL_業種_C", "SL_商品_C", "TL_日付_C"),
    "chart_names": ("PG-C01", "PG-C02", "PG-C03", "PG-C04"),
    "chart_types": (65, 65, 57, 59),
    "pt_ranges": {
        "PT-01": "$A$2:$B$4",
        "PT-02": "$A$12:$B$20",
        "PT-03": "$A$26:$C$38",
        "PT-04": "$A$56:$C$66",
        "PT-05": "$H$2:$T$11",
        "PT-06": "$H$14:$I$15",
        "PT-07": "$H$20:$I$22",
        "PT-08": "$H$28:$J$41",
    },
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
    },
}

CHECKS: list[dict[str, Any]] = []


def rgb(color: tuple[int, int, int]) -> int:
    """RGB タプルを Excel COM の色値へ変換する。

    Args:
        color: (赤, 緑, 青) の 0 から 255。

    Returns:
        Excel COM の色値。

    Side Effects:
        なし。
    """
    red, green, blue = color
    return red + (green << 8) + (blue << 16)


def check(
    name: str,
    expected: Any,
    actual: Any,
    predicate: Callable[[Any], bool] | None = None,
) -> bool:
    """1 件の検証結果を記録する。

    Args:
        name: 検証名。
        expected: 期待値。
        actual: 実測値。
        predicate: 実測値を評価する関数。省略時は等価比較。

    Returns:
        合格なら True。

    Side Effects:
        CHECKS へ結果を追加する。
    """
    try:
        ok = predicate(actual) if predicate is not None else actual == expected
    except Exception as exc:
        ok = False
        actual = f"ERROR: {exc}"
    CHECKS.append(
        {
            "name": name,
            "expected": expected,
            "actual": actual,
            "ok": bool(ok),
        }
    )
    return bool(ok)


def check_close(name: str, expected: float, actual: Any, tolerance: float = 1e-9) -> bool:
    """数値の近傍一致を検証する。

    Args:
        name: 検証名。
        expected: 期待値。
        actual: 実測値。
        tolerance: 許容差。

    Returns:
        合格なら True。

    Side Effects:
        CHECKS へ結果を追加する。
    """
    try:
        ok = abs(float(actual) - expected) <= tolerance
        value = float(actual)
    except (TypeError, ValueError):
        ok = False
        value = actual
    CHECKS.append(
        {
            "name": name,
            "expected": expected,
            "actual": value,
            "ok": bool(ok),
        }
    )
    return bool(ok)


def check_true(name: str, condition: bool, detail: Any = None) -> bool:
    """真偽条件を検証する。

    Args:
        name: 検証名。
        condition: 条件。
        detail: 実測値の表示用。

    Returns:
        合格なら True。

    Side Effects:
        CHECKS へ結果を追加する。
    """
    return check(name, True, detail if detail is not None else condition, lambda _: bool(condition))


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


def count_formula_token(book: object, token: str) -> int:
    """数式に指定文字列を含むセル数を数える。

    Args:
        book: 対象ブック。
        token: 検索文字列（大文字小文字を区別しない）。

    Returns:
        該当セル数。

    Side Effects:
        なし。
    """
    total = 0
    needle = token.upper()
    for sheet in book.Worksheets:
        for cell in sheet.UsedRange:
            try:
                if cell.HasFormula and needle in cell.Formula.upper():
                    total += 1
            except Exception:
                continue
    return total


def find_slicer_cache(book: object, name: str) -> object:
    """名前でスライサーキャッシュを取得する。

    Args:
        book: 対象ブック。
        name: キャッシュ名。

    Returns:
        スライサーキャッシュ。

    Raises:
        LookupError: 見つからない場合。

    Side Effects:
        なし。
    """
    for index in range(1, book.SlicerCaches.Count + 1):
        cache = book.SlicerCaches.Item(index)
        if cache.Name == name:
            return cache
    raise LookupError(f"スライサーキャッシュがありません: {name}")


def find_chart(dashboard: object, name: str) -> object:
    """名前でグラフを取得する。

    Args:
        dashboard: 売上分析シート。
        name: 図形名。

    Returns:
        Chart オブジェクト。

    Raises:
        LookupError: 見つからない場合。

    Side Effects:
        なし。
    """
    shape = dashboard.Shapes(name)
    if not shape.HasChart:
        raise LookupError(f"グラフではありません: {name}")
    return shape.Chart


def find_heatmap_row(dashboard: object, product: str) -> int:
    """ヒートマップの商品行を取得する。

    Args:
        dashboard: 売上分析シート。
        product: 商品名。

    Returns:
        行番号。

    Raises:
        LookupError: 見つからない場合。

    Side Effects:
        なし。
    """
    for row in range(CONFIG["heatmap_start_row"], CONFIG["heatmap_end_row"] + 1):
        if dashboard.Cells(row, 1).Value == product:
            return row
    raise LookupError(f"商品行がありません: {product}")


def find_heatmap_column(dashboard: object, month: str) -> int:
    """ヒートマップの月列を取得する。

    Args:
        dashboard: 売上分析シート。
        month: 年月文字列。

    Returns:
        列番号。

    Raises:
        LookupError: 見つからない場合。

    Side Effects:
        なし。
    """
    for col in range(CONFIG["heatmap_start_col"], CONFIG["heatmap_end_col"] + 1):
        if dashboard.Cells(50, col).Value == month:
            return col
    raise LookupError(f"月列がありません: {month}")


def heatmap_logic_formula(product_cell: str, month_cell: str) -> str:
    """ヒートマップと同じ 3 状態判定の検証用数式を返す。

    Args:
        product_cell: 商品名のセル。
        month_cell: 年月のセル。

    Returns:
        検証用数式。

    Side Effects:
        なし。
    """
    base = (
        'GETPIVOTDATA("売上（万円）",制御!$H$2,'
        f'"商品",{product_cell},"年月",{month_cell})'
    )
    return (
        f'=IF(ISERROR({base}),'
        f'IF(COUNTIFS(SalesHistory[商品],{product_cell},'
        f'SalesHistory[年月],{month_cell})=0,"－",""),'
        f"ROUND({base},0))"
    )


def scan_package(path: Path) -> dict[str, Any]:
    """xlsx パッケージの禁止情報とデータモデルを確認する。

    Args:
        path: 対象 xlsx。

    Returns:
        検査結果。

    Side Effects:
        なし。
    """
    # 検出語は分割して組み立てる。ソース自体にローカルパスや個人名を残さないためである。
    needles = (
        b":" + b"\\" + b"Users",
        b":" + b"/" + b"Users",
        b"abs" + b"Path",
        b"file:" + b"//",
    )
    hits: list[str] = []
    model_parts: list[str] = []
    vba_parts: list[str] = []
    creator = ""
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.lower().startswith("xl/model/"):
                model_parts.append(name)
            if name.lower().endswith("vbaproject.bin"):
                vba_parts.append(name)
            data = archive.read(name)
            if name == "docProps/core.xml":
                core_xml = data.decode("utf-8", "ignore")
                marker = "<dc:creator>"
                if marker in core_xml:
                    creator = core_xml.split(marker, 1)[1].split("</dc:creator>", 1)[0]
            for needle in needles:
                if needle in data:
                    hits.append(f"{name}:{needle.decode('utf-8', errors='replace')}")
    return {
        "hits": hits,
        "model_parts": model_parts,
        "vba_parts": vba_parts,
        "creator": creator,
    }


def load_v06_fixture(cfg: dict[str, Any]) -> dict[str, Any]:
    """v0.6 専用の架空データ fixture を読み込む。

    Args:
        cfg: 検証設定。

    Returns:
        metadata と売上明細を含む fixture 辞書。

    Raises:
        FileNotFoundError: fixture がない場合。
        ValueError: fixture の行数が想定外の場合。

    Side Effects:
        JSON ファイルを読み取る。
    """
    fixture_path = Path(__file__).resolve().parent / cfg["fixture_file"]
    if not fixture_path.is_file():
        raise FileNotFoundError("v0.6 fixture がありません")
    fixture = json.loads(fixture_path.read_text(encoding="utf-8"))
    records = fixture.get("records")
    if not isinstance(records, list) or len(records) != 576:
        raise ValueError("v0.6 fixture の行数が576件ではありません")
    return fixture


def test_zero_state(source: Path, cfg: dict[str, Any]) -> None:
    """一時コピーで 0 円セルが 0万円 と表示されることを検証する。

    Args:
        source: 検証対象 xlsx。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        一時コピーの売上データを 0 に変更し、保存せずに閉じる。
    """
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    try:
        with tempfile.TemporaryDirectory(prefix="v06_zero_") as temp_dir:
            target = Path(temp_dir) / "zero.xlsx"
            shutil.copy2(source, target)
            book = app.Workbooks.Open(str(target), UpdateLinks=0)
            try:
                data = book.Worksheets(cfg["data_sheet"])
                table = data.ListObjects(cfg["table"])
                changed = 0
                for row in range(2, table.ListRows.Count + 2):
                    product = table.Range.Cells(row, 3).Value
                    month = table.Range.Cells(row, 10).Value
                    if product == "テーブル" and month == "2026-01":
                        table.Range.Cells(row, 6).Value = 0
                        changed += 1
                check_true("zero test source rows", changed > 0, changed)
                for sheet in book.Worksheets:
                    try:
                        for pt in sheet.PivotTables():
                            pt.RefreshTable()
                    except Exception:
                        pass
                app.Calculate()
                dashboard = book.Worksheets(cfg["dashboard"])
                row = find_heatmap_row(dashboard, "テーブル")
                col = find_heatmap_column(dashboard, "2026-01")
                cell = dashboard.Cells(row, col)
                check_close("zero state value", 0, cell.Value2, 1e-9)
                check("zero state display", "0万円", cell.Text)
                check_true(
                    "zero state is numeric",
                    not cell.HasFormula or cell.Value2 == 0,
                    cell.Value2,
                )
            finally:
                book.Close(False)
    except Exception as exc:
        check_true("zero state test", False, str(exc))
    finally:
        app.Quit()


def test_no_data_state(source: Path, cfg: dict[str, Any]) -> None:
    """一時コピーでデータなしと選択範囲外の表示を検証する。

    Args:
        source: 検証対象 xlsx。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        一時コピーの空きセルへ検証用数式を書き込み、保存せずに閉じる。
    """
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    try:
        with tempfile.TemporaryDirectory(prefix="v06_state_") as temp_dir:
            target = Path(temp_dir) / "states.xlsx"
            shutil.copy2(source, target)
            book = app.Workbooks.Open(str(target), UpdateLinks=0)
            try:
                dashboard = book.Worksheets(cfg["dashboard"])
                dashboard.Range("AA1").Value = "未登録商品"
                dashboard.Range("AA2").Value = "2026-01"
                dashboard.Range("AA3").Formula = heatmap_logic_formula("AA1", "AA2")
                dashboard.Range("AA4").Value = "テーブル"
                dashboard.Range("AA5").Value = "2026-01"
                dashboard.Range("AA6").Formula = heatmap_logic_formula("AA4", "AA5")
                dashboard.Range("AA7").Value = "2025-01"
                dashboard.Range("AA8").Formula = heatmap_logic_formula("AA4", "AA7")
                dashboard.Range("AA3:AA8").NumberFormat = '0"万円"'
                app.Calculate()
                check("no data state", "－", dashboard.Range("AA3").Text)
                check_close(
                    "existing data state",
                    cfg["expected_heatmap_2026_jan_table"],
                    dashboard.Range("AA6").Value2,
                    1e-9,
                )
                check_true(
                    "out of range state",
                    dashboard.Range("AA8").Value2 in (None, ""),
                    dashboard.Range("AA8").Value2,
                )
            finally:
                book.Close(False)
    except Exception as exc:
        check_true("no data state test", False, str(exc))
    finally:
        app.Quit()


def test_timeline_heatmap_year_selection(source: Path, cfg: dict[str, Any]) -> None:
    """一時コピーで年別タイムライン選択がヒートマップへ反映されることを検証する。

    Args:
        source: 検証対象 xlsx。
        cfg: 検証設定。

    Returns:
        なし。

    Side Effects:
        一時コピーのタイムラインを2025年、2026年、両年へ切り替える。
    """
    temp_root = Path(os.environ["TEMP"]) / "opencode"
    target = temp_root / "v06_timeline_test.xlsx"
    if target.exists():
        check_true("timeline temporary path available", False, "temporary file already exists")
        return

    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    book = None
    try:
        shutil.copy2(source, target)
        book = app.Workbooks.Open(str(target), UpdateLinks=0)
        dashboard = book.Worksheets(cfg["dashboard"])
        timeline = find_slicer_cache(book, "TL_日付_C")
        control = book.Worksheets(cfg["control"])

        def set_period(start: str, end: str) -> tuple[int, int]:
            """期間を設定して、両年の数値セル数を返す。"""
            timeline.TimelineState.SetFilterDateRange(start, end)
            for pivot in control.PivotTables():
                pivot.RefreshTable()
            app.CalculateFullRebuild()
            count_2025 = sum(
                1
                for row in range(51, 59)
                for col in range(2, 14)
                if isinstance(dashboard.Cells(row, col).Value2, (int, float))
            )
            count_2026 = sum(
                1
                for row in range(51, 59)
                for col in range(14, 26)
                if isinstance(dashboard.Cells(row, col).Value2, (int, float))
            )
            return count_2025, count_2026

        counts_2025 = set_period("2025-01-01", "2025-12-31")
        check_true(
            "timeline 2025 selection drives heatmap",
            counts_2025 == (96, 0)
            and dashboard.Range("B51").Value2 == cfg["expected_heatmap_2025_jan_table"],
            {"numeric_cells": counts_2025, "jan_value": dashboard.Range("B51").Value2},
        )

        counts_2026 = set_period("2026-01-01", "2026-12-31")
        check_true(
            "timeline 2026 selection drives heatmap",
            counts_2026 == (0, 96)
            and dashboard.Range("N51").Value2 == cfg["expected_heatmap_2026_jan_table"],
            {"numeric_cells": counts_2026, "jan_value": dashboard.Range("N51").Value2},
        )

        counts_both = set_period("2025-01-01", "2026-12-31")
        check("timeline both years drive heatmap", (96, 96), counts_both)
    except Exception as exc:
        check_true("timeline heatmap year-selection test", False, type(exc).__name__)
    finally:
        if book is not None:
            book.Close(False)
        app.Quit()
        if target.exists():
            target.unlink()


def run_workbook_checks(source: Path, cfg: dict[str, Any]) -> None:
    """検証対象ブックの主要チェックを実行する。

    Args:
        source: 検証対象 xlsx。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        Excel を起動し、読み取り専用でブックを開く。
    """
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    book = app.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True)
    try:
        sheet_names = [book.Worksheets(i).Name for i in range(1, book.Worksheets.Count + 1)]
        check("sheet count", cfg["sheet_count"], book.Worksheets.Count)
        check("dashboard index", cfg["dashboard_index"], book.Worksheets(cfg["dashboard"]).Index)
        check("dashboard visible", -1, book.Worksheets(cfg["dashboard"]).Visible)
        check("control hidden", 0, book.Worksheets(cfg["control"]).Visible)
        check("active sheet", cfg["dashboard"], book.ActiveSheet.Name)
        check("dashboard before product master", True, sheet_names.index(cfg["dashboard"]) < sheet_names.index(cfg["product_master"]))
        check("required sheets", True, all(name in sheet_names for name in ("計算", "見積入力", "設定", "採算シミュレーション")))

        # テーブル
        data = book.Worksheets(cfg["data_sheet"])
        product = book.Worksheets(cfg["product_master"])
        data_table = data.ListObjects(cfg["table"])
        product_table = product.ListObjects(cfg["product_table"])
        check("SalesHistory range", "$A$4:$L$580", data_table.Range.Address)
        check("SalesHistory rows", 576, data_table.ListRows.Count)
        check("SalesHistory columns", 12, data_table.ListColumns.Count)
        check("SalesHistory totals", False, bool(data_table.ShowTotals))
        headers = [data_table.Range.Cells(1, i).Value for i in range(1, 13)]
        check(
            "SalesHistory headers",
            ["日付", "業種", "商品", "数量", "税抜単価", "税抜売上", "計上原価", "計上原価控除後差額", "データ区分", "年月", "年", "月"],
            headers,
        )
        check("SalesHistory helper 年月", "2025-01", data_table.Range.Cells(2, 10).Value)
        check("SalesHistory helper 年", 2025, data_table.Range.Cells(2, 11).Value)
        check("SalesHistory helper 月", 1, data_table.Range.Cells(2, 12).Value)
        fixture_records = load_v06_fixture(cfg)["records"]
        actual_rows = data_table.DataBodyRange.Value
        fixture_mismatches: list[int] = []
        fixture_sales = {2025: 0, 2026: 0}
        fixture_costs = {2025: 0, 2026: 0}
        monthly_totals = {(year, month): 0 for year in (2025, 2026) for month in range(1, 13)}
        product_totals = {(year, product): 0 for year in (2025, 2026) for product in {str(row["product"]) for row in fixture_records}}
        for row_index, (expected, actual) in enumerate(zip(fixture_records, actual_rows), start=1):
            year = int(expected["year"])
            month = int(expected["month"])
            sales = int(expected["sales"])
            cost = int(expected["cost"])
            actual_date = actual[0]
            date_matches = (
                getattr(actual_date, "year", None) == year
                and getattr(actual_date, "month", None) == month
                and getattr(actual_date, "day", None) == int(expected["day"])
            )
            expected_values = (
                str(expected["industry"]),
                str(expected["product"]),
                int(expected["qty"]),
                int(expected["price"]),
                sales,
                cost,
                sales - cost,
                "架空サンプル",
                f"{year}-{month:02d}",
                year,
                month,
            )
            if not date_matches or tuple(actual[1:]) != expected_values:
                fixture_mismatches.append(row_index)
            fixture_sales[year] += sales
            fixture_costs[year] += cost
            monthly_totals[(year, month)] += sales
            product_totals[(year, str(expected["product"]))] += sales
        check_true(
            "SalesHistory matches v0.6 fixture",
            not fixture_mismatches,
            {"mismatch_count": len(fixture_mismatches), "sample_rows": fixture_mismatches[:5]},
        )
        check("fixture 2025 sales", cfg["expected_sales_2025"], fixture_sales[2025])
        check("fixture 2025 cost", cfg["expected_cost_2025"], fixture_costs[2025])
        check("fixture 2026 sales", cfg["expected_sales"], fixture_sales[2026])
        check("fixture 2026 cost", cfg["expected_cost"], fixture_costs[2026])
        annual_growth = fixture_sales[2026] / fixture_sales[2025] - 1
        check_close("fixture annual sales growth", 0.08, annual_growth, 0.0001)
        check_true(
            "fixture monthly profiles differ",
            [monthly_totals[(2025, month)] for month in range(1, 13)]
            != [monthly_totals[(2026, month)] for month in range(1, 13)],
        )
        check_true(
            "fixture product mix differs",
            any(
                product_totals[(2025, product)] != product_totals[(2026, product)]
                for product in {str(row["product"]) for row in fixture_records}
            ),
        )
        check("ProductMaster range", "$A$4:$E$12", product_table.Range.Address)
        check("ProductMaster rows", 8, product_table.ListRows.Count)
        check("ProductMaster columns", 5, product_table.ListColumns.Count)
        product_headers = [product_table.Range.Cells(1, i).Value for i in range(1, 6)]
        check("ProductMaster headers", ["商品 ID", "商品名", "業種", "表示順", "有効"], product_headers)

        # ピボット
        control = book.Worksheets(cfg["control"])
        check("pivot cache count", 1, book.PivotCaches().Count)
        check("pivot table count", 8, control.PivotTables().Count)
        pts = {pt.Name: pt for pt in control.PivotTables()}
        check("pivot names", set(cfg["pt_ranges"]), set(pts))
        for name, expected_range in cfg["pt_ranges"].items():
            pt = pts[name]
            check(f"{name} range", expected_range, pt.TableRange1.Address)
            check(f"{name} cache index", 1, pt.CacheIndex)
        source_data = str(book.PivotCaches().Item(1).SourceData)
        check_true(
            "pivot source is SalesHistory",
            ("SalesHistory" in source_data or "売上データ" in source_data)
            and "ProductMaster" not in source_data,
            source_data,
        )

        # スライサー
        check("slicer cache count", 3, book.SlicerCaches.Count)
        cache_names = {book.SlicerCaches(i).Name for i in range(1, book.SlicerCaches.Count + 1)}
        check("slicer cache names", set(cfg["slicer_cache_names"]), cache_names)
        check("old slicer caches absent", True, {"SL_業種", "SL_商品", "TL_日付"}.isdisjoint(cache_names))
        industry = find_slicer_cache(book, "SL_業種_C")
        product_cache = find_slicer_cache(book, "SL_商品_C")
        timeline = find_slicer_cache(book, "TL_日付_C")
        check("industry source field", "業種", industry.SourceName)
        check("product source field", "商品", product_cache.SourceName)
        check("timeline source field", "日付", timeline.SourceName)
        check("industry PT connections", 8, industry.PivotTables.Count)
        check("product PT connections", 8, product_cache.PivotTables.Count)
        check("timeline PT connections", 6, timeline.PivotTables.Count)
        check("industry items", 2, industry.SlicerItems.Count)
        check("product items", 8, product_cache.SlicerItems.Count)
        timeline_pt_names = {timeline.PivotTables(i).Name for i in range(1, timeline.PivotTables.Count + 1)}
        check("timeline current-period PTs", {"PT-01", "PT-02", "PT-03", "PT-04", "PT-05", "PT-06"}, timeline_pt_names)
        check_true("timeline excludes PT-07", "PT-07" not in timeline_pt_names, timeline_pt_names)
        check_true("timeline excludes PT-08", "PT-08" not in timeline_pt_names, timeline_pt_names)
        check("timeline start", date(2026, 1, 1), timeline.TimelineState.StartDate.date())
        check("timeline end", date(2026, 12, 31), timeline.TimelineState.EndDate.date())

        # スライサー図形
        dashboard = book.Worksheets(cfg["dashboard"])
        slicer_shapes = []
        for cache in (industry, product_cache, timeline):
            check(f"{cache.Name} slicer shape count", 1, cache.Slicers.Count)
            slicer = cache.Slicers(1)
            slicer_shapes.append(slicer)
            check(f"{cache.Name} parent sheet", cfg["dashboard"], slicer.Parent.Name)
        check_true("slicer rail left", all(s.Left < 20 for s in slicer_shapes), [s.Left for s in slicer_shapes])
        check_true("slicer rail stacked", slicer_shapes[0].Top < slicer_shapes[1].Top < slicer_shapes[2].Top, [s.Top for s in slicer_shapes])
        check_true("slicer rail fits", slicer_shapes[2].Left + slicer_shapes[2].Width < dashboard.Range("D15").Left, (slicer_shapes[2].Left, slicer_shapes[2].Width, dashboard.Range("D15").Left))
        check("industry slicer style", "SlicerStyleLight2", str(slicer_shapes[0].Style))
        check("product slicer style", "SlicerStyleLight2", str(slicer_shapes[1].Style))
        check("timeline slicer style", "TimeSlicerStyleLight1", str(slicer_shapes[2].Style))

        # KPI
        check_close("KPI sales", cfg["expected_sales"], dashboard.Range("E6").Value2)
        check_close("KPI diff", cfg["expected_diff"], dashboard.Range("I6").Value2)
        check_close("KPI rate", cfg["expected_rate"], dashboard.Range("E11").Value2)
        check_close("KPI yoy", cfg["expected_yoy"], dashboard.Range("I11").Value2)
        check("KPI sales formula", True, "GETPIVOTDATA" in dashboard.Range("E6").Formula)
        check("KPI diff formula", True, "GETPIVOTDATA" in dashboard.Range("I6").Formula)
        check("KPI rate formula", True, "I6/E6" in dashboard.Range("E11").Formula)
        check("KPI yoy formula", True, "2026" in dashboard.Range("I11").Formula and "2025" in dashboard.Range("I11").Formula)
        check("KPI sales label", "税抜売上（円）", dashboard.Range("E5").Value)
        check("KPI diff label", "差額（円）", dashboard.Range("I5").Value)
        check("KPI rate label", "差額率", dashboard.Range("E10").Value)
        check("KPI yoy label", "前年比", dashboard.Range("I10").Value)
        check("KPI sales format", "#,##0", dashboard.Range("E6").NumberFormat)
        check("KPI rate format", "0.0%", dashboard.Range("E11").NumberFormat)
        check("KPI sales merge", True, bool(dashboard.Range("E6").MergeCells))
        check("KPI diff merge", True, bool(dashboard.Range("I6").MergeCells))
        check("KPI cards 2x2", True, bool(dashboard.Range("E5").MergeCells and dashboard.Range("I5").MergeCells and dashboard.Range("E10").MergeCells and dashboard.Range("I10").MergeCells))
        check("KPI accent 1", rgb(cfg["colors"]["coral"]), dashboard.Range("D5").Interior.Color)
        check("KPI accent 2", rgb(cfg["colors"]["amber"]), dashboard.Range("H5").Interior.Color)
        check("KPI accent 3", rgb(cfg["colors"]["muted_blue"]), dashboard.Range("D10").Interior.Color)
        check("KPI accent 4", rgb(cfg["colors"]["navy_soft"]), dashboard.Range("H10").Interior.Color)
        check("KPI sales font size", 30, dashboard.Range("E6").Font.Size)
        check("KPI diff font size", 30, dashboard.Range("I6").Font.Size)
        check("KPI rate font size", 30, dashboard.Range("E11").Font.Size)
        check("KPI yoy font size", 30, dashboard.Range("I11").Font.Size)

        # シェル
        check("title", "売上分析", dashboard.Range("A2").Value)
        check("title size", 26, dashboard.Range("A2").Font.Size)
        check("subtitle", "数字で読む、商品と季節の動き", dashboard.Range("A3").Value)
        check("design label", "DESIGN C / WARM SLATE", dashboard.Range("G2").Value)
        check("badge", "2025-2026 / 架空データ", dashboard.Range("Y2").Value)
        check("accent line", rgb(cfg["colors"]["coral"]), dashboard.Range("A4").Interior.Color)
        check("sidebar", rgb(cfg["colors"]["navy"]), dashboard.Range("A5").Interior.Color)
        check("sidebar title", "FILTERS", dashboard.Range("A5").Value)
        check("sidebar industry label", "業種", dashboard.Range("A6").Value)
        check("sidebar product label", "商品", dashboard.Range("A13").Value)
        check("sidebar period label", "期間", dashboard.Range("A24").Value)
        check_true("sidebar note", "タイムライン非連動" in str(dashboard.Range("A32").Value), dashboard.Range("A32").Value)
        check("insight strip", "DATA NOTE", dashboard.Range("L5").Value)
        check("insight title", "読み取りのポイント", dashboard.Range("L6").Value)
        check("insight fill", rgb(cfg["colors"]["sand"]), dashboard.Range("L6").Interior.Color)
        check_true("insight body", "商品マスタ" in str(dashboard.Range("L8").Value), dashboard.Range("L8").Value)
        check("tab color", rgb(cfg["colors"]["coral"]), dashboard.Tab.Color)

        # グラフ
        chart_shapes = [shape for shape in dashboard.Shapes if shape.HasChart]
        check("dashboard shape count", 7, dashboard.Shapes.Count)
        check("dashboard chart count", 4, len(chart_shapes))
        for name, expected_type in zip(cfg["chart_names"], cfg["chart_types"]):
            chart = find_chart(dashboard, name)
            check(f"{name} chart type", expected_type, chart.ChartType)
            check(f"{name} title", True, bool(chart.HasTitle))
            check(f"{name} field buttons hidden", False, bool(chart.ShowAllFieldButtons))
        check("PG-C01 series", 1, find_chart(dashboard, "PG-C01").SeriesCollection().Count)
        check("PG-C02 series", 2, find_chart(dashboard, "PG-C02").SeriesCollection().Count)
        check("PG-C03 series", 1, find_chart(dashboard, "PG-C03").SeriesCollection().Count)
        check("PG-C04 series", 2, find_chart(dashboard, "PG-C04").SeriesCollection().Count)
        check("PG-C01 title text", "月別売上（万円）", find_chart(dashboard, "PG-C01").ChartTitle.Text)
        check("PG-C02 title text", "2025 年と 2026 年の比較（万円）", find_chart(dashboard, "PG-C02").ChartTitle.Text)
        check("PG-C03 title text", "商品別売上（万円）", find_chart(dashboard, "PG-C03").ChartTitle.Text)
        check("PG-C04 title text", "業種構成（%）", find_chart(dashboard, "PG-C04").ChartTitle.Text)
        check("PG-C02 axis title", "月", find_chart(dashboard, "PG-C02").Axes(1).AxisTitle.Text)
        check_true(
            "PG-C01 left of charts",
            all(dashboard.Shapes(name).Left > dashboard.Range("D15").Left for name in cfg["chart_names"]),
            [dashboard.Shapes(name).Left for name in cfg["chart_names"]],
        )
        check_true(
            "PG-C01 above PG-C03",
            dashboard.Shapes("PG-C01").Top < dashboard.Shapes("PG-C03").Top,
            (dashboard.Shapes("PG-C01").Top, dashboard.Shapes("PG-C03").Top),
        )
        for name, right_boundary in (
            ("PG-C01", "K15"),
            ("PG-C02", "R15"),
            ("PG-C03", "K32"),
            ("PG-C04", "R32"),
        ):
            shape = dashboard.Shapes(name)
            limit = dashboard.Range(right_boundary).Left
            check_true(
                f"{name} fits through Q column",
                shape.Left + shape.Width <= limit,
                {"right": shape.Left + shape.Width, "limit": limit},
            )
        check("existing 採算 charts", 2, sum(1 for shape in book.Worksheets("採算シミュレーション").Shapes if shape.HasChart))

        # ヒートマップ
        check("heatmap title", "商品×月ヒートマップ（万円）", dashboard.Range("A48").Value)
        check("heatmap header 1", "2025-01", dashboard.Range("B50").Value)
        check("heatmap header 13", "2026-01", dashboard.Range("N50").Value)
        check("heatmap header 24", "2026-12", dashboard.Range("Y50").Value)
        check("heatmap sparkline header", "推移", dashboard.Range("Z50").Value)
        check("heatmap product 1", "テーブル", dashboard.Range("A51").Value)
        check("heatmap product 8", "季節の果物", dashboard.Range("A58").Value)
        formula_cells = 0
        iferror_cells = 0
        countifs_cells = 0
        iserror_cells = 0
        number_format_errors = 0
        for row in range(cfg["heatmap_start_row"], cfg["heatmap_end_row"] + 1):
            for col in range(cfg["heatmap_start_col"], cfg["heatmap_end_col"] + 1):
                cell = dashboard.Cells(row, col)
                formula = str(cell.Formula or "")
                if cell.HasFormula:
                    formula_cells += 1
                if "IFERROR" in formula.upper():
                    iferror_cells += 1
                if "COUNTIFS" in formula.upper():
                    countifs_cells += 1
                if "ISERROR" in formula.upper():
                    iserror_cells += 1
                if cell.NumberFormat != '0"万円"':
                    number_format_errors += 1
        check("heatmap formula cells", 192, formula_cells)
        check("heatmap IFERROR absent", 0, iferror_cells)
        check("heatmap COUNTIFS cells", 192, countifs_cells)
        check("heatmap ISERROR cells", 192, iserror_cells)
        check("heatmap number format", 0, number_format_errors)
        check_true(
            "heatmap 2025-01 blank",
            dashboard.Range("B51").Value2 in (None, ""),
            dashboard.Range("B51").Value2,
        )
        first_product = str(fixture_records[0]["product"])
        fixture_2025_jan = sum(
            int(row["sales"])
            for row in fixture_records
            if int(row["year"]) == 2025
            and int(row["month"]) == 1
            and str(row["product"]) == first_product
        )
        check(
            "fixture heatmap 2025-01 expected",
            cfg["expected_heatmap_2025_jan_table"],
            round(fixture_2025_jan / 10000),
        )
        check("heatmap 2026-01 value", cfg["expected_heatmap_2026_jan_table"], dashboard.Range("N51").Value2)
        check("heatmap 2026-12 value", cfg["expected_heatmap_2026_dec_table"], dashboard.Range("Y51").Value2)
        check("heatmap 2025 block blank", True, all(dashboard.Cells(r, c).Value2 in (None, "") for r in range(51, 59) for c in range(2, 14)))
        check("heatmap 2026 block numeric", 96, sum(1 for r in range(51, 59) for c in range(14, 26) if isinstance(dashboard.Cells(r, c).Value2, (int, float))))
        check("heatmap legend 0", True, "0万円" in str(dashboard.Range("A49").Value))
        check("heatmap legend no data", True, "－" in str(dashboard.Range("G49").Value))
        check("heatmap legend blank", True, "空欄" in str(dashboard.Range("M49").Value))
        check("heatmap note product master", True, "商品マスタ" in str(dashboard.Range("A62").Value))

        # 条件付き書式
        conditions = dashboard.Range("B51:Y58").FormatConditions
        check("heatmap conditional format count", 3, conditions.Count)
        check("blank rule type", 2, conditions(1).Type)
        check("blank rule formula", "=B51=\"\"", conditions(1).Formula1)
        check("blank rule stop", True, bool(conditions(1).StopIfTrue))
        check("blank rule priority", 1, conditions(1).Priority)
        check("blank rule no fill", -4142, conditions(1).Interior.Pattern)
        check("text rule type", 2, conditions(2).Type)
        check("text rule formula", "=ISTEXT(B51)", conditions(2).Formula1)
        check("text rule stop", True, bool(conditions(2).StopIfTrue))
        check("text rule priority", 2, conditions(2).Priority)
        check("text rule fill", rgb(cfg["colors"]["sand"]), conditions(2).Interior.Color)
        check("text rule font contrast", rgb(cfg["colors"]["text"]), conditions(2).Font.Color)
        check("scale rule type", 3, conditions(3).Type)
        check("scale rule priority", 3, conditions(3).Priority)
        check("scale criteria count", 3, conditions(3).ColorScaleCriteria.Count)
        check("scale low color", rgb(cfg["colors"]["light_peach"]), conditions(3).ColorScaleCriteria(1).FormatColor.Color)
        check("scale mid type", 5, conditions(3).ColorScaleCriteria(2).Type)
        check("scale mid value", 50, conditions(3).ColorScaleCriteria(2).Value)
        check("scale mid color", rgb(cfg["colors"]["amber"]), conditions(3).ColorScaleCriteria(2).FormatColor.Color)
        check("scale high color", rgb(cfg["colors"]["coral"]), conditions(3).ColorScaleCriteria(3).FormatColor.Color)

        # スパークライン
        for row in range(51, 59):
            groups = dashboard.Range(f"Z{row}").SparklineGroups
            check(f"sparkline {row} count", 1, groups.Count)
            check(f"sparkline {row} type", 2, groups(1).Type)
            check(f"sparkline {row} source", f"B{row}:Y{row}", groups(1).SourceData)
            check(f"sparkline {row} color", rgb(cfg["colors"]["coral"]), groups(1).SeriesColor.Color)

        # 印刷
        check("print area", "$A$1:$Z$64", dashboard.PageSetup.PrintArea)
        check("landscape", 2, dashboard.PageSetup.Orientation)
        check("A4 paper", 9, dashboard.PageSetup.PaperSize)
        check("fit wide", 1, dashboard.PageSetup.FitToPagesWide)
        check("fit tall", 1, dashboard.PageSetup.FitToPagesTall)
        check("zoom off", False, bool(dashboard.PageSetup.Zoom))

        # 既存計算・不変条件
        calc = book.Worksheets(cfg["calc_sheet"])
        check("profit", cfg["expected_profit"], calc.Range("B32").Value2)
        check("quote total", cfg["expected_quote"], calc.Range("B39").Value2)
        check("notes", cfg["expected_notes"], count_notes(book))
        comment_shapes: list[int] = []
        comment_autosize: list[bool] = []
        comment_text_matches: list[bool] = []
        for sheet in book.Worksheets:
            for index in range(1, sheet.Comments.Count + 1):
                comment = sheet.Comments.Item(index)
                shape = comment.Shape
                comment_shapes.append(int(shape.AutoShapeType))
                comment_autosize.append(bool(shape.TextFrame.AutoSize))
                comment_text_matches.append(
                    str(comment.Text()) == str(shape.TextFrame.Characters().Text)
                )
        check("comment shapes are rectangles", cfg["expected_notes"], sum(1 for kind in comment_shapes if kind == 1))
        check("comment text autofit enabled", cfg["expected_notes"], sum(1 for enabled in comment_autosize if enabled))
        check_true("comment text is not cut", all(comment_text_matches), comment_text_matches.count(False))
        check("INDIRECT", cfg["expected_indirect"], count_formula_token(book, "INDIRECT"))
        check("CUBE absent", 0, count_formula_token(book, "CUBE"))
        check("sheet protection absent", True, all(not sheet.ProtectContents for sheet in book.Worksheets))
        check("freeze panes absent", False, bool(app.ActiveWindow.FreezePanes))
        check("VBA absent", False, bool(book.HasVBProject))
        check("external links absent", None, book.LinkSources())
        check("queries absent", 0, book.Queries.Count)
        check("file format", 51, book.FileFormat)
        check("cell protection absent", True, all(not dashboard.Range(address).Locked for address in ("A1", "D5", "E6", "A51", "Z51")))

        # パッケージ
        package = scan_package(source)
        check("prohibited info absent", [], package["hits"])
        check("creator metadata", cfg["expected_creator"], package["creator"])
        check("data model parts absent", [], package["model_parts"])
        check("VBA parts absent", [], package["vba_parts"])
        check("macro extension absent", True, source.suffix.lower() == ".xlsx")

        # 旧デザインの残存なし
        check("old chart names absent", True, all(not any(shape.Name == name for shape in dashboard.Shapes) for name in ("PG-B01", "PG-B02", "PG-B03", "PG-B04", "PG-01", "PG-02", "PG-03", "PG-04")))
        check("new chart names present", True, all(any(shape.Name == name for shape in dashboard.Shapes) for name in cfg["chart_names"]))
    finally:
        book.Close(False)
        app.Quit()


def json_default(value: Any) -> Any:
    """JSON へ変換できない値を読みやすい形へ変換する。

    Args:
        value: 変換対象。

    Returns:
        JSON 互換値。

    Side Effects:
        なし。
    """
    if isinstance(value, set):
        return sorted(value)
    if isinstance(value, (date,)):
        return value.isoformat()
    try:
        return str(value)
    except Exception:
        return f"<{type(value).__name__}>"


def main() -> int:
    """検証を実行し、結果 JSON を保存する。

    Returns:
        終了コード（0 は全件合格）。

    Side Effects:
        Excel を起動し、結果 JSON を書き込む。
    """
    if len(sys.argv) < 2:
        print("usage: verify_v06_windows.py <workbook.xlsx> [results.json]")
        return 2
    source = Path(sys.argv[1]).resolve()
    results_path = Path(sys.argv[2]) if len(sys.argv) >= 3 else Path("windows_v06_results.json")
    if not source.is_file():
        print(f"file not found: {source.name}")
        return 2

    try:
        run_workbook_checks(source, CONFIG)
        test_zero_state(source, CONFIG)
        test_no_data_state(source, CONFIG)
        test_timeline_heatmap_year_selection(source, CONFIG)
    except Exception as exc:
        check_true("verification execution", False, str(exc))

    total = len(CHECKS)
    passed = sum(1 for item in CHECKS if item["ok"])
    failed = total - passed
    summary = {
        "workbook": source.name,
        "total": total,
        "passed": passed,
        "failed": failed,
        "checks": CHECKS,
    }
    results_path.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2, default=json_default),
        encoding="utf-8",
    )
    print(f"total={total} passed={passed} failed={failed}")
    for item in CHECKS:
        if not item["ok"]:
            print(f"FAIL: {item['name']} expected={item['expected']!r} actual={item['actual']!r}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
