"""v0.4 方式 A の実機検証。

目的:
    `prototype_v04.xlsx` が方式 A（通常ピボット・共有キャッシュ・スライサー・
    タイムライン・ピボットグラフ）の受入条件を満たすか検証する。

仕様:
    - 公開禁止情報（絶対パス・ローカルパス・作成者）の不在を確認する。
    - シート、テーブル、共有キャッシュ、PT-01 から PT-08 の構成を確認する。
    - 業種・商品スライサー、日付タイムラインの接続と初期状態を確認する。
    - KPI、ピボットグラフ 4 個、商品×月ヒートマップ、スパークラインを確認する。
    - スライサー・タイムライン操作後の KPI と PT 値を確認する。
    - v0.3 の不変条件（メモ 74、INDIRECT 24、保護・固定枠なし、VBA なし）を確認する。

前提:
    - Windows と Microsoft Excel デスクトップアプリ、pywin32 が利用できる。
    - 引数 1 に検証対象 xlsx を指定する。

入出力:
    - 入力: 引数 1（v0.4 xlsx）。
    - 出力: 標準出力の PASS/FAIL と `windows_v04_results.json`。

注意点:
    - 検証は一時コピーに対して行い、元ファイルを変更しない。
    - Excel の Calculation プロパティは環境により設定できないため使用しない。
    - タイムラインの期間設定は文字列日付で行う。

拡張ポイント:
    - ゲージ・目標管理を追加した場合は、対応する KPI と非連動範囲を追加する。

変更履歴:
    2026-10-22: 初版作成（方式 A の受入検証）。
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import win32com.client

# ---------------------------------------------------------------------------
# 設定値
# ---------------------------------------------------------------------------
CONFIG: dict[str, object] = {
    "expected_notes": 74,
    "expected_indirect": 24,
    "expected_sales": 47492700,
    "expected_diff": 18852426,
    "expected_rate": 0.3969541845378342,
    "expected_yoy": -0.011131169147547193,
    "expected_sales_2025": 48027300,
    "expected_sales_full": 95520000,
    "expected_wood_full": 73300000,
    "expected_wood_2026": 36327500,
    "expected_table_2026": 17450000,
    "expected_pt01_wood": 3632.75,
    "expected_pt01_farm": 1116.52,
    "expected_pt07_wood_2025": 3697.25,
    "expected_pt07_wood_2026": 3632.75,
    "expected_pt08_2026_01": 433.33,
    "expected_pt05_table_2026_01": 170.0,
    "expected_heatmap_table_2026_01": 170.0,
    "expected_heatmap_chair_2026_01": 39.0,
    "expected_profit": 72000,
    "expected_quote": 176000,
    "expected_sheet_count": 13,
    "expected_pt_count": 8,
    "expected_slicer_count": 3,
    "expected_timeline_pt_count": 6,
    "expected_heatmap_cf": 2,
    "expected_sparklines": 8,
}

RESULTS: list[dict[str, object]] = []


def json_safe(value: object) -> object:
    """JSON へ保存できる形へ値を変換する。

    Args:
        value: 任意の判定値。

    Returns:
        set はソート済みリスト、tuple はリスト、その他はそのまま。

    Side Effects:
        なし。

    Examples:
        >>> json_safe({"b", "a"})
        ['a', 'b']
    """
    if isinstance(value, set):
        return sorted(value, key=str)
    if isinstance(value, tuple):
        return list(value)
    return value


def check(name: str, expected: object, actual: object) -> None:
    """判定を 1 件記録する。

    Args:
        name: 判定名。
        expected: 期待値。
        actual: 実測値。

    Returns:
        なし。

    Side Effects:
        RESULTS へ追記する。

    Examples:
        >>> check("notes", 74, 74)
    """
    passed = expected == actual
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        passed = abs(float(actual) - float(expected)) <= 1e-6
    RESULTS.append(
        {
            "name": name,
            "expected": json_safe(expected),
            "actual": json_safe(actual),
            "passed": passed,
        }
    )


def calculate(app: object) -> None:
    """同期再計算を実行する。

    Args:
        app: Excel Application。

    Returns:
        なし。

    Side Effects:
        Excel の再計算を実行する。

    Examples:
        >>> calculate(app)
    """
    app.Calculate()


def scan_package(path: Path) -> dict[str, object]:
    """xlsx パッケージの公開禁止情報を検査する。

    Args:
        path: 検査対象 xlsx。

    Returns:
        絶対パス、ローカルパス、作成者情報の実測値。

    Side Effects:
        なし。

    Examples:
        >>> scan_package(Path("prototype_v04.xlsx"))
    """
    with zipfile.ZipFile(path, "r") as z:
        names = z.namelist()
        workbook = z.read("xl/workbook.xml").decode("utf-8", "ignore")
        core = (
            z.read("docProps/core.xml").decode("utf-8", "ignore")
            if "docProps/core.xml" in names
            else ""
        )
    local_paths = 0
    with zipfile.ZipFile(path, "r") as z:
        for name in z.namelist():
            data = z.read(name).decode("utf-8", "ignore")
            local_paths += data.count("C:\\Users")
    creators = re.findall(r"<(?:dc:creator|cp:lastModifiedBy)>([^<]*)<", core)
    return {
        "abs_path": workbook.count("x15ac:absPath"),
        "local_paths": local_paths,
        "creators": creators,
    }


def scan_timeline(path: Path) -> dict[str, object]:
    """タイムラインキャッシュの初期期間と接続 PT を検査する。

    Args:
        path: 検査対象 xlsx。

    Returns:
        タイムライン部品数、期間設定、接続 PT 名。

    Side Effects:
        なし。

    Examples:
        >>> scan_timeline(Path("prototype_v04.xlsx"))
    """
    with zipfile.ZipFile(path, "r") as z:
        parts = [
            name
            for name in z.namelist()
            if "timelineCache" in name and name.endswith(".xml")
        ]
        text = "\n".join(
            z.read(name).decode("utf-8", "ignore") for name in parts
        )
    state = re.search(
        r'<state[^>]*filterType="dateBetween"[^>]*>(.*?)</state>', text, re.S
    )
    state_text = state.group(1) if state else ""
    pivot_names = re.findall(r'<pivotTable[^>]*name="([^"]+)"', text)
    return {
        "parts": len(parts),
        "date_between": 'filterType="dateBetween"' in text,
        "start_2026": 'startDate="2026-01-01T00:00:00"' in state_text,
        "end_2026": 'endDate="2026-12-31T00:00:00"' in state_text,
        "pivot_names": pivot_names,
    }


def count_indirect(book: object) -> int:
    """INDIRECT を含む数式の件数を数える。

    Args:
        book: 対象ブック。

    Returns:
        INDIRECT を含む数式のセル数。

    Side Effects:
        なし。

    Examples:
        >>> count_indirect(book)
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


def find_slicer(book: object, name: str) -> object:
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

    Examples:
        >>> find_slicer(book, "SL_業種")
    """
    caches = book.SlicerCaches
    for index in range(1, caches.Count + 1):
        cache = caches.Item(index)
        if cache.Name == name:
            return cache
    raise LookupError(f"スライサーキャッシュがありません: {name}")


def set_slicer_items(cache: object, names: set[str], app: object) -> None:
    """スライサーの選択項目を指定集合へ変更する。

    Args:
        cache: 対象スライサーキャッシュ。
        names: 選択する項目名。
        app: Excel Application。

    Returns:
        なし。

    Side Effects:
        スライサーの選択状態とピボットテーブルの絞り込みを変更する。

    Examples:
        >>> set_slicer_items(industry, {"木工所"}, app)
    """
    for index in range(1, cache.SlicerItems.Count + 1):
        item = cache.SlicerItems(index)
        item.Selected = item.Name in names
    calculate(app)


def set_timeline(cache: object, start: str, end: str, app: object) -> None:
    """タイムラインの期間を文字列日付で設定する。

    Args:
        cache: タイムラインキャッシュ。
        start: 開始日（yyyy-mm-dd）。
        end: 終了日（yyyy-mm-dd）。
        app: Excel Application。

    Returns:
        なし。

    Side Effects:
        タイムラインの期間と接続 PT の絞り込みを変更する。

    Examples:
        >>> set_timeline(timeline, "2026-01-01", "2026-12-31", app)
    """
    cache.TimelineState.SetFilterDateRange(start, end)
    calculate(app)


def is_blank(value: object) -> bool:
    """セル値が空欄相当かを判定する。

    Args:
        value: セル値。

    Returns:
        空欄相当なら True。

    Side Effects:
        なし。

    Examples:
        >>> is_blank("")
        True
    """
    return value in (None, "")


def pt_map(book: object) -> dict[str, object]:
    """制御シートのピボットテーブルを名前で取得する。

    Args:
        book: 対象ブック。

    Returns:
        PT 名から COM オブジェクトへの辞書。

    Side Effects:
        なし。

    Examples:
        >>> pts = pt_map(book)
    """
    control = book.Worksheets("制御")
    return {
        control.PivotTables(index).Name: control.PivotTables(index)
        for index in range(1, control.PivotTables().Count + 1)
    }


def verify_package(source: Path, cfg: dict[str, object]) -> None:
    """パッケージとタイムライン XML を検証する。

    Args:
        source: 対象 xlsx。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        RESULTS を更新する。

    Examples:
        >>> verify_package(Path("prototype_v04.xlsx"), CONFIG)
    """
    package = scan_package(source)
    check("absPath absent", 0, package["abs_path"])
    check("local path absent", 0, package["local_paths"])
    check(
        "creator fixed",
        True,
        all(c == "ExcelBIBasic3" for c in package["creators"]),
    )

    timeline = scan_timeline(source)
    check("timeline part", 1, timeline["parts"])
    check("timeline dateBetween", True, timeline["date_between"])
    check("timeline start 2026", True, timeline["start_2026"])
    check("timeline end 2026", True, timeline["end_2026"])
    expected_names = {f"PT-{index:02d}" for index in range(1, 7)}
    check(
        "timeline PT names",
        expected_names,
        set(timeline["pivot_names"]),
    )
    check("PT-07 timeline off", False, "PT-07" in timeline["pivot_names"])
    check("PT-08 timeline off", False, "PT-08" in timeline["pivot_names"])


def verify_tables(book: object, cfg: dict[str, object]) -> None:
    """シート・テーブル・共有キャッシュ・PT の構成を検証する。

    Args:
        book: 対象ブック。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        RESULTS を更新する。

    Examples:
        >>> verify_tables(book, CONFIG)
    """
    check("sheet count", cfg["expected_sheet_count"], book.Worksheets.Count)
    names = {book.Worksheets(i).Name for i in range(1, book.Worksheets.Count + 1)}
    for required in ("売上データ", "売上分析", "商品マスタ", "制御", "計算"):
        check(f"sheet {required}", True, required in names)
    check("control hidden", 0, book.Worksheets("制御").Visible)
    check("product master visible", -1, book.Worksheets("商品マスタ").Visible)
    check("dashboard visible", -1, book.Worksheets("売上分析").Visible)
    check("file format xlsx", 51, book.FileFormat)

    data = book.Worksheets("売上データ")
    table = data.ListObjects("SalesHistory")
    check("SalesHistory range", "$A$4:$L$580", table.Range.Address)
    check("SalesHistory columns", 12, table.ListColumns.Count)
    check("SalesHistory rows", 576, table.ListRows.Count)
    check("SalesHistory J4", "年月", data.Range("J4").Value)
    check("SalesHistory K4", "年", data.Range("K4").Value)
    check("SalesHistory L4", "月", data.Range("L4").Value)
    check("SalesHistory J5", "2025-01", data.Range("J5").Value)
    check("SalesHistory K5", 2025, data.Range("K5").Value)
    check("SalesHistory L5", 1, data.Range("L5").Value)
    check("SalesHistory J5 unlocked", False, bool(data.Range("J5").Locked))

    master = book.Worksheets("商品マスタ")
    product_table = master.ListObjects("ProductMaster")
    check("ProductMaster range", "$A$4:$E$12", product_table.Range.Address)
    check("ProductMaster rows", 8, product_table.ListRows.Count)
    check("ProductMaster A5", "P-01", master.Range("A5").Value)
    check("ProductMaster B5", "テーブル", master.Range("B5").Value)
    check("ProductMaster E5", True, bool(master.Range("E5").Value))
    check("ProductMaster B5 unlocked", False, bool(master.Range("B5").Locked))

    caches = book.PivotCaches()
    check("pivot cache count", 1, caches.Count)
    check("pivot cache source", "SalesHistory", caches.Item(1).SourceData)

    control = book.Worksheets("制御")
    check("control PT count", cfg["expected_pt_count"], control.PivotTables().Count)
    pts = pt_map(book)
    check(
        "PT names",
        {f"PT-{index:02d}" for index in range(1, 9)},
        set(pts),
    )
    check("shared cache indexes", {1}, {pt.CacheIndex for pt in pts.values()})
    check("PT-01 table", "$A$2:$B$4", pts["PT-01"].TableRange1.Address)
    check("PT-05 table", "$H$2:$T$11", pts["PT-05"].TableRange1.Address)
    check("PT-06 table", "$H$14:$I$15", pts["PT-06"].TableRange1.Address)
    check("PT-01 data field", "売上（万円）", pts["PT-01"].DataFields(1).Name)
    check("PT-03 data fields", 2, pts["PT-03"].DataFields.Count)
    check("PT-06 data fields", 2, pts["PT-06"].DataFields.Count)
    check(
        "PT-01 万円 format",
        True,
        "万円" in pts["PT-01"].DataFields(1).NumberFormat,
    )
    check(
        "PT-01 木工所",
        cfg["expected_pt01_wood"],
        pts["PT-01"].GetPivotData("売上（万円）", "業種", "木工所").Value,
    )
    check(
        "PT-01 農家",
        cfg["expected_pt01_farm"],
        pts["PT-01"].GetPivotData("売上（万円）", "業種", "農家").Value,
    )
    check(
        "PT-06 sales",
        cfg["expected_sales"],
        pts["PT-06"].GetPivotData("売上合計").Value,
    )
    check(
        "PT-06 diff",
        cfg["expected_diff"],
        pts["PT-06"].GetPivotData("差額合計").Value,
    )
    check(
        "PT-07 2025",
        4802.73,
        pts["PT-07"].GetPivotData("売上（万円）", "年", 2025).Value,
    )
    check(
        "PT-07 2026",
        4749.27,
        pts["PT-07"].GetPivotData("売上（万円）", "年", 2026).Value,
    )
    check(
        "PT-08 2026-01",
        cfg["expected_pt08_2026_01"],
        pts["PT-08"].GetPivotData("売上（万円）", "月", 1, "年", 2026).Value,
    )
    check(
        "PT-05 table 2026-01",
        cfg["expected_pt05_table_2026_01"],
        pts["PT-05"]
        .GetPivotData("売上（万円）", "商品", "テーブル", "年月", "2026-01")
        .Value,
    )
    check("control A2 unlocked", False, bool(control.Range("A2").Locked))


def verify_slicers_and_charts(book: object, cfg: dict[str, object]) -> None:
    """スライサー・タイムライン・ピボットグラフを検証する。

    Args:
        book: 対象ブック。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        RESULTS を更新する。

    Examples:
        >>> verify_slicers_and_charts(book, CONFIG)
    """
    check("slicer cache count", cfg["expected_slicer_count"], book.SlicerCaches.Count)
    industry = find_slicer(book, "SL_業種")
    product = find_slicer(book, "SL_商品")
    timeline = find_slicer(book, "TL_日付")
    check("industry source", "業種", industry.SourceName)
    check("product source", "商品", product.SourceName)
    check("timeline source", "日付", timeline.SourceName)
    check("industry type", 1, industry.SlicerCacheType)
    check("product type", 1, product.SlicerCacheType)
    check("timeline type", 2, timeline.SlicerCacheType)
    check("industry items", 2, industry.SlicerItems.Count)
    check("product items", 8, product.SlicerItems.Count)

    expected_all = {f"PT-{index:02d}" for index in range(1, 9)}
    industry_pts = {
        industry.PivotTables(i).Name for i in range(1, industry.PivotTables.Count + 1)
    }
    product_pts = {
        product.PivotTables(i).Name for i in range(1, product.PivotTables.Count + 1)
    }
    timeline_pts = {
        timeline.PivotTables(i).Name
        for i in range(1, timeline.PivotTables.Count + 1)
    }
    check("industry PTs", expected_all, industry_pts)
    check("product PTs", expected_all, product_pts)
    check("timeline PT count", cfg["expected_timeline_pt_count"], len(timeline_pts))
    check(
        "timeline PTs",
        {f"PT-{index:02d}" for index in range(1, 7)},
        timeline_pts,
    )
    check("PT-07 slicer on", True, "PT-07" in industry_pts)
    check("PT-08 slicer on", True, "PT-08" in product_pts)
    for index in range(1, industry.SlicerItems.Count + 1):
        check(
            f"industry selected {industry.SlicerItems(index).Name}",
            True,
            bool(industry.SlicerItems(index).Selected),
        )
    for index in range(1, product.SlicerItems.Count + 1):
        check(
            f"product selected {product.SlicerItems(index).Name}",
            True,
            bool(product.SlicerItems(index).Selected),
        )

    analysis = book.Worksheets("売上分析")
    check("chart object count", 4, analysis.ChartObjects().Count)
    expected = {
        "PG-01": (4, "月別推移（万円）"),
        "PG-02": (51, "前年比較（万円）"),
        "PG-03": (57, "商品別ランキング（万円）"),
        "PG-04": (-4120, "業種構成（万円）"),
    }
    actual_names = set()
    for index in range(1, analysis.ChartObjects().Count + 1):
        chart_object = analysis.ChartObjects(index)
        name = chart_object.Name
        actual_names.add(name)
        chart = chart_object.Chart
        check(f"{name} type", expected[name][0], chart.ChartType)
        check(f"{name} title", expected[name][1], chart.ChartTitle.Text)
        check(f"{name} pivot layout", True, chart.PivotLayout is not None)
    check("chart names", set(expected), actual_names)
    check("dashboard shape count", 7, analysis.Shapes.Count)


def verify_dashboard(book: object, app: object, cfg: dict[str, object]) -> None:
    """KPI・ヒートマップ・注記・セル状態を検証する。

    Args:
        book: 対象ブック。
        app: Excel Application。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        RESULTS を更新する。

    Examples:
        >>> verify_dashboard(book, app, CONFIG)
    """
    calculate(app)
    analysis = book.Worksheets("売上分析")
    check("KPI sales", cfg["expected_sales"], analysis.Range("A9").Value2)
    check("KPI diff", cfg["expected_diff"], analysis.Range("D9").Value2)
    check("KPI rate", cfg["expected_rate"], analysis.Range("G9").Value2)
    check("KPI yoy", cfg["expected_yoy"], analysis.Range("J9").Value2)
    check("KPI sales label", "税抜売上（円）", analysis.Range("A8").Value)
    check("KPI diff label", "計上原価控除後差額（円）", analysis.Range("D8").Value)
    check("KPI rate label", "差額率", analysis.Range("G8").Value)
    check("KPI yoy label", "前年比", analysis.Range("J8").Value)
    check("KPI sales formula", True, "GETPIVOTDATA" in analysis.Range("A9").Formula)
    check("KPI diff formula", True, "GETPIVOTDATA" in analysis.Range("D9").Formula)
    check(
        "KPI yoy formula years",
        True,
        "2025" in analysis.Range("J9").Formula
        and "2026" in analysis.Range("J9").Formula,
    )
    check("old KPI B5 blank", True, is_blank(analysis.Range("B5").Value))
    check("old KPI E5 blank", True, is_blank(analysis.Range("E5").Value))
    check("old KPI H5 blank", True, is_blank(analysis.Range("H5").Value))
    check("old KPI K5 blank", True, is_blank(analysis.Range("K5").Value))
    check("old matrix A28 blank", True, is_blank(analysis.Range("A28").Value))
    sumifs = 0
    for cell in analysis.UsedRange:
        try:
            if cell.HasFormula and "SUMIFS" in cell.Formula.upper():
                sumifs += 1
        except Exception:
            continue
    check("SUMIFS absent", 0, sumifs)

    check("heatmap title", "商品×月ヒートマップ（万円）", analysis.Range("A44").Value)
    check("heatmap corner", "商品／年月", analysis.Range("A45").Value)
    check("heatmap spark header", "推移", analysis.Range("Z45").Value)
    check("heatmap first month", "2025-01", analysis.Range("B45").Value)
    check("heatmap last month", "2026-12", analysis.Range("Y45").Value)
    check("heatmap first product", "テーブル", analysis.Range("A46").Value)
    check("heatmap last product", "季節の果物", analysis.Range("A53").Value)
    check(
        "heatmap table 2026-01",
        cfg["expected_heatmap_table_2026_01"],
        analysis.Range("N46").Value2,
    )
    check(
        "heatmap chair 2026-01",
        cfg["expected_heatmap_chair_2026_01"],
        analysis.Range("N47").Value2,
    )
    check("heatmap 2025 blank", True, is_blank(analysis.Range("B46").Value2))
    check("heatmap 2025 blank 2", True, is_blank(analysis.Range("C46").Value2))
    check(
        "heatmap CF count",
        cfg["expected_heatmap_cf"],
        analysis.Range("B46:Y53").FormatConditions.Count,
    )
    check(
        "heatmap sparklines",
        cfg["expected_sparklines"],
        analysis.Range("Z46:Z53").SparklineGroups.Count,
    )
    check(
        "heatmap 万円 format",
        True,
        "万円" in analysis.Range("N46").NumberFormat,
    )
    check("note A3", True, "架空" in str(analysis.Range("A3").Value))
    check("note A6", True, "前年" in str(analysis.Range("A6").Value))
    check("note A11", True, "計上原価" in str(analysis.Range("A11").Value))
    check("note A55", True, "万円" in str(analysis.Range("A55").Value))
    check("note A57", True, "商品マスタ" in str(analysis.Range("A57").Value))
    check("KPI A9 unlocked", False, bool(analysis.Range("A9").Locked))
    check("KPI A9 formula visible", False, bool(analysis.Range("A9").FormulaHidden))
    check("heatmap N46 unlocked", False, bool(analysis.Range("N46").Locked))


def verify_invariants(book: object, app: object, cfg: dict[str, object]) -> None:
    """v0.3 の不変条件と保護・固定枠の不在を検証する。

    Args:
        book: 対象ブック。
        app: Excel Application。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        RESULTS を更新する。可視シートを順にアクティブ化する。

    Examples:
        >>> verify_invariants(book, app, CONFIG)
    """
    quote = book.Worksheets("見積入力")
    sim = book.Worksheets("採算シミュレーション")
    calc = book.Worksheets("計算")
    check("B12 merge", "$B$12:$F$12", quote.Range("B12").MergeArea.Address)
    check("B12 note", "試作品の架空データです", quote.Range("B12").Value)
    check("D12 empty", None, quote.Range("D12").Value)
    check("D12:F12 CF count", 0, quote.Range("D12:F12").FormatConditions.Count)
    check("採算 D12 status present", True, sim.Range("D12").Value not in (None, ""))
    check("profit", cfg["expected_profit"], calc.Range("B32").Value2)
    check("quote total", cfg["expected_quote"], calc.Range("B39").Value2)
    check(
        "notes",
        cfg["expected_notes"],
        sum(sheet.Comments.Count for sheet in book.Worksheets),
    )
    check("INDIRECT", cfg["expected_indirect"], count_indirect(book))
    check(
        "protection",
        False,
        any(bool(sheet.ProtectContents) for sheet in book.Worksheets),
    )
    check("VBA absent", False, bool(book.HasVBProject))
    check("external links", None, book.LinkSources())
    check("queries absent", 0, book.Queries.Count)
    for sheet in book.Worksheets:
        if sheet.Visible != -1:
            continue
        sheet.Activate()
        check(f"{sheet.Name} freeze", False, bool(app.ActiveWindow.FreezePanes))
    book.Worksheets("売上分析").Activate()


def verify_interactions(
    book: object, app: object, cfg: dict[str, object]
) -> None:
    """スライサー・タイムライン操作後の KPI と PT 値を検証する。

    Args:
        book: 対象ブック。
        app: Excel Application。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        RESULTS を更新する。検証後に初期状態（全選択・2026 年）へ戻す。

    Examples:
        >>> verify_interactions(book, app, CONFIG)
    """
    analysis = book.Worksheets("売上分析")
    industry = find_slicer(book, "SL_業種")
    product = find_slicer(book, "SL_商品")
    timeline = find_slicer(book, "TL_日付")
    pts = pt_map(book)
    industry_all = {
        industry.SlicerItems(i).Name for i in range(1, industry.SlicerItems.Count + 1)
    }
    product_all = {
        product.SlicerItems(i).Name for i in range(1, product.SlicerItems.Count + 1)
    }

    set_timeline(timeline, "2025-01-01", "2026-12-31", app)
    check("KPI full period", cfg["expected_sales_full"], analysis.Range("A9").Value2)
    set_slicer_items(industry, {"木工所"}, app)
    check("KPI wood full", cfg["expected_wood_full"], analysis.Range("A9").Value2)
    check(
        "PT-07 wood 2025",
        cfg["expected_pt07_wood_2025"],
        pts["PT-07"].GetPivotData("売上（万円）", "年", 2025).Value,
    )
    check(
        "PT-07 wood 2026",
        cfg["expected_pt07_wood_2026"],
        pts["PT-07"].GetPivotData("売上（万円）", "年", 2026).Value,
    )
    set_timeline(timeline, "2026-01-01", "2026-12-31", app)
    check("KPI wood 2026", cfg["expected_wood_2026"], analysis.Range("A9").Value2)
    check(
        "PT-07 wood 2025 after timeline",
        cfg["expected_pt07_wood_2025"],
        pts["PT-07"].GetPivotData("売上（万円）", "年", 2025).Value,
    )
    set_slicer_items(industry, industry_all, app)
    check("KPI default after industry reset", cfg["expected_sales"], analysis.Range("A9").Value2)

    set_timeline(timeline, "2025-01-01", "2025-12-31", app)
    check("KPI 2025", cfg["expected_sales_2025"], analysis.Range("A9").Value2)
    check(
        "PT-08 timeline off",
        cfg["expected_pt08_2026_01"],
        pts["PT-08"].GetPivotData("売上（万円）", "月", 1, "年", 2026).Value,
    )
    set_timeline(timeline, "2026-01-01", "2026-12-31", app)
    set_slicer_items(product, {"テーブル"}, app)
    check(
        "KPI table 2026",
        cfg["expected_table_2026"],
        analysis.Range("A9").Value2,
    )
    check(
        "PT-05 table after product",
        cfg["expected_pt05_table_2026_01"],
        pts["PT-05"]
        .GetPivotData("売上（万円）", "商品", "テーブル", "年月", "2026-01")
        .Value,
    )
    set_slicer_items(product, product_all, app)
    check("KPI default after product reset", cfg["expected_sales"], analysis.Range("A9").Value2)


def execute_logic(app: object, source: Path, cfg: dict[str, object]) -> None:
    """検証本体を実行する。

    Args:
        app: Excel Application。
        source: 検証対象 xlsx。
        cfg: 設定値。

    Returns:
        なし。

    Side Effects:
        RESULTS を更新する。一時コピーを開いて閉じる。

    Examples:
        >>> execute_logic(app, Path("prototype_v04.xlsx"), CONFIG)
    """
    verify_package(source, cfg)
    with tempfile.TemporaryDirectory(prefix="v04_check_") as tmp:
        dest = Path(tmp) / "check.xlsx"
        shutil.copy2(source, dest)
        book = app.Workbooks.Open(str(dest), UpdateLinks=0, ReadOnly=True)
        try:
            verify_tables(book, cfg)
            verify_slicers_and_charts(book, cfg)
            verify_dashboard(book, app, cfg)
            verify_invariants(book, app, cfg)
            verify_interactions(book, app, cfg)
        finally:
            book.Close(False)


def format_output() -> str:
    """検証結果を整形する。

    Returns:
        整形済み文字列。

    Side Effects:
        なし。

    Examples:
        >>> print(format_output())
    """
    lines = []
    for item in RESULTS:
        mark = "PASS" if item["passed"] else "FAIL"
        lines.append(
            f"{mark} {item['name']}: expected={item['expected']!r} actual={item['actual']!r}"
        )
    passed = sum(1 for item in RESULTS if item["passed"])
    lines.append(f"PASS {passed} / {len(RESULTS)}")
    return "\n".join(lines)


def main() -> int:
    """エントリポイント。

    Returns:
        終了コード（0 は全件合格）。

    Side Effects:
        Excel を起動し、結果 JSON を書き出す。

    Examples:
        >>> raise SystemExit(main())
    """
    if len(sys.argv) != 2:
        print("usage: verify_v04_windows.py <xlsx>")
        return 2
    source = Path(sys.argv[1]).resolve()
    if not source.is_file():
        print("入力ファイルが見つかりません")
        return 2
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AutomationSecurity = 3
    try:
        execute_logic(app, source, CONFIG)
    finally:
        app.Quit()
    print(format_output())
    report = {"results": RESULTS}
    Path("windows_v04_results.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return 0 if all(item["passed"] for item in RESULTS) else 1


if __name__ == "__main__":
    raise SystemExit(main())
