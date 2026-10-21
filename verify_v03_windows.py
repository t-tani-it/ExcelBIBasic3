"""v0.3 差分の実機検証（Windows + Excel + pywin32）。

目的:
    P5 で変更した見積入力の備考欄復帰と状態表示撤去を検証し、
    不変条件（メモ 74 件、INDIRECT 24 件、保護・固定枠なし）と
    主要計算値の維持を確認する。

仕様:
    - 入力ブックのコピーを一時フォルダで開き、読み取り専用の検証を行う。
    - 計算は Calculate() を使用する（CalculateFullRebuild は環境により完了しないため）。
    - 結果は windows_v03_results.json へ保存する。

前提:
    - Windows と Microsoft Excel デスクトップアプリ、pywin32。
    - 入力は v0.3 成果物 prototype_v03.xlsx。

入出力:
    - 入力: 引数 1（xlsx のパス）。
    - 出力: windows_v03_results.json（判定結果）。

注意点:
    - 元ファイルは変更しない。
    - 画面ロック中は Excel の計算が完了しない場合がある。

拡張ポイント:
    - P1〜P4 実装後に、PQ・モデル・ピボット・スライサーの検証を追加する。

変更履歴:
    2026-10-21: 初版作成（P5 差分）。
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import win32com.client

# ---------------------------------------------------------------------------
# 設定値
# ---------------------------------------------------------------------------
CONFIG: dict[str, object] = {
    "expected_notes": 74,
    "expected_indirect": 24,
    "expected_sales_total": 47492700,
    "expected_cost_total": 18852426,
    "expected_diff_total": 18852426,
    "expected_diff_rate": 0.3969541845378342,
    "expected_profit": 72000,
    "expected_quote": 176000,
    "calc_timeout": 60,
}

RESULTS: list[dict[str, object]] = []


def check(name: str, expected: object, actual: object) -> None:
    """判定を 1 件記録する。

    Args:
        name: 判定名。
        expected: 期待値。
        actual: 実測値。

    Side Effects:
        RESULTS へ追記する。

    Examples:
        >>> check("notes", 74, 74)
    """
    passed = expected == actual
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        passed = abs(float(actual) - float(expected)) < 1e-6
    RESULTS.append({"name": name, "expected": expected, "actual": actual, "passed": passed})


def calculate(app: object, cfg: dict[str, object]) -> None:
    """計算を実行する。

    Args:
        app: Excel Application。
        cfg: 設定値。

    Side Effects:
        Excel の再計算。

    Notes:
        CalculationState は非表示セッションで 0 に戻らない場合があるため、
        同期実行の Calculate() のみを使用する。

    Examples:
        >>> calculate(app, CONFIG)
    """
    app.Calculate()


def scan_package(path: Path) -> dict[str, object]:
    """xlsx パッケージの公開禁止情報を検査する。

    Args:
        path: 検査対象の xlsx。

    Returns:
        絶対パス、ローカルパス、作成者情報の実測値。

    Side Effects:
        なし。

    Examples:
        >>> scan_package(Path("prototype_v03.xlsx"))
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


def execute_logic(app: object, source: Path, cfg: dict[str, object]) -> None:
    """検証本体を実行する。

    Args:
        app: Excel Application。
        source: 検証対象 xlsx。
        cfg: 設定値。

    Side Effects:
        RESULTS を更新する。

    Examples:
        >>> execute_logic(app, Path("prototype_v03.xlsx"), CONFIG)
    """
    with tempfile.TemporaryDirectory(prefix="v03_check_") as tmp:
        dest = Path(tmp) / "check.xlsx"
        shutil.copy2(source, dest)
        book = app.Workbooks.Open(str(dest), UpdateLinks=0, ReadOnly=True)
        try:
            package = scan_package(source)
            check("absPath absent", 0, package["abs_path"])
            check("local path absent", 0, package["local_paths"])
            check(
                "creator fixed",
                True,
                all(c == "ExcelBIBasic3" for c in package["creators"]),
            )
            calculate(app, cfg)
            quote = book.Worksheets("見積入力")
            sim = book.Worksheets("採算シミュレーション")
            calc = book.Worksheets("計算")

            check("B12 merge", "$B$12:$F$12", quote.Range("B12").MergeArea.Address)
            check("B12 note", "試作品の架空データです", quote.Range("B12").Value)
            check("D12 empty", None, quote.Range("D12").Value)
            check("D12:F12 CF count", 0, quote.Range("D12:F12").FormatConditions.Count)
            check("採算 D12 status present", True, sim.Range("D12").Value not in (None, ""))
            check("notes", cfg["expected_notes"], sum(s.Comments.Count for s in book.Worksheets))
            check("INDIRECT", cfg["expected_indirect"], count_indirect(book))
            analysis = book.Worksheets("売上分析")
            check("KPI 売上", cfg["expected_sales_total"], analysis.Range("A9").Value2)
            check("KPI 差額", cfg["expected_diff_total"], analysis.Range("F9").Value2)
            check("KPI 差額率", cfg["expected_diff_rate"], analysis.Range("K9").Value2)
            check("profit", cfg["expected_profit"], calc.Range("B32").Value2)
            check("quote total", cfg["expected_quote"], calc.Range("B39").Value2)
            check("VBA absent", False, bool(book.HasVBProject))
            check("external links", None, book.LinkSources())
            check("protection", False, any(bool(s.ProtectContents) for s in book.Worksheets))
            for sheet in book.Worksheets:
                sheet.Activate()
                check(f"{sheet.Name} freeze", False, bool(app.ActiveWindow.FreezePanes))
            check("queries (P1 pending)", 0, book.Queries.Count)
        finally:
            book.Close(False)


def format_output(cfg: dict[str, object]) -> str:
    """結果を整形する。

    Args:
        cfg: 設定値。

    Returns:
        整形済み文字列。

    Side Effects:
        なし。

    Examples:
        >>> print(format_output(CONFIG))
    """
    lines = []
    for item in RESULTS:
        mark = "PASS" if item["passed"] else "FAIL"
        lines.append(f"{mark} {item['name']}: expected={item['expected']!r} actual={item['actual']!r}")
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
        print("usage: verify_v03_windows.py <xlsx>")
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
    print(format_output(CONFIG))
    report = {"results": RESULTS}
    Path("windows_v03_results.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return 0 if all(item["passed"] for item in RESULTS) else 1


if __name__ == "__main__":
    raise SystemExit(main())
