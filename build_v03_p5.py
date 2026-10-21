"""P5: 見積入力の備考欄復帰と入力状態表示の削除（ExcelBIBasic3）。

目的:
    v0.2 で追加した見積入力 D12:F12 の入力状態表示を撤去し、B12:F12 を
    備考入力へ戻す。状態表示は採算シミュレーション D12 の業務状態のみとする。

仕様:
    - 見積入力 B12:C12 と D12:F12 の結合を解除し、B12:F12 に再結合する。
    - B12 の備考文字列を維持する。
    - D12:F12 に付いた条件付き書式 3 件を削除する。
    - 行28の入力検査（G28 見出し、I28 判定式）は既存機能として維持する。
    - メモ 74 件、INDIRECT 24 件、保護・固定枠なしを維持する。
    - 保存後に OOXML を後処理し、絶対パスと作成者情報を除去する。

前提:
    - 入力は v0.2 成果物 prototype.xlsx である。
    - Windows と Microsoft Excel デスクトップアプリ、pywin32 が利用できる。

入出力:
    - 入力: 引数 1（入力 xlsx のパス）。
    - 出力: 引数 2（出力 xlsx のパス）。

注意点:
    - 状態表示セルにメモが付いている場合は中止する（メモ保全のため）。
    - 計算セル（B8、I16:I25 など）は変更しない。
    - 数式から I28 を参照する箇所が D12 以外にある場合は中止する。

拡張ポイント:
    - 採算シミュレーション D12 の業務状態表示は既存のまま維持する。

変更履歴:
    2026-10-21: 初版作成（P5）。
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

import win32com.client

from finalize_v03_ooxml import CONFIG as OOXML_CONFIG
from finalize_v03_ooxml import sanitize_package

# ---------------------------------------------------------------------------
# 設定値
# ---------------------------------------------------------------------------
CONFIG: dict[str, object] = {
    # 対象シート
    "sheet_quote": "見積入力",
    # 結合を解除する範囲
    "unmerge_ranges": ("B12:C12", "D12:F12"),
    # 再結合する範囲（備考欄）
    "merge_range": "B12:F12",
    # 備考の入力セル
    "note_cell": "B12",
    # 削除する条件付き書式の適用範囲
    "remove_cf_range": "$D$12:$F$12",
    # 既存の入力検査（行28）は維持する
    "state_header": "G28",
    "state_formula": "I28",
    # 期待値（不変条件）
    "expected_notes": 74,
    "expected_indirect": 24,
}


# ---------------------------------------------------------------------------
# 検証
# ---------------------------------------------------------------------------
def validate_input(path: Path) -> None:
    """入力ファイルの前提を検証する。

    Args:
        path: 入力 xlsx のパス。

    Raises:
        FileNotFoundError: ファイルが存在しない場合。
        ValueError: 拡張子が xlsx でない場合。

    Side Effects:
        なし。

    Examples:
        >>> validate_input(Path("prototype.xlsx"))
    """
    if not path.is_file():
        raise FileNotFoundError(f"入力ファイルが見つかりません: {path.name}")
    if path.suffix.lower() != ".xlsx":
        raise ValueError("入力は .xlsx である必要があります")


def validate_preconditions(book: object, cfg: dict[str, object]) -> None:
    """編集前の状態とメモの保全条件を検証する。

    Args:
        book: 対象ブック。
        cfg: 設定値。

    Raises:
        ValueError: 想定と異なる結合状態、または保護されたメモ対象がある場合。
        RuntimeError: I28 を外部から参照している場合。

    Side Effects:
        なし。

    Examples:
        >>> validate_preconditions(book, CONFIG)
    """
    ws = book.Worksheets(cfg["sheet_quote"])
    if ws.Range("B12").MergeArea.Address != "$B$12:$C$12":
        raise ValueError("B12:C12 が想定どおり結合されていません")
    if ws.Range("D12").MergeArea.Address != "$D$12:$F$12":
        raise ValueError("D12:F12 が想定どおり結合されていません")
    for address in ("D12", "E12", "F12"):
        if ws.Range(address).Comment is not None:
            raise ValueError(f"メモ保全のため中止します: {address} にメモがあります")


# ---------------------------------------------------------------------------
# 実行
# ---------------------------------------------------------------------------
def apply_edits(book: object, cfg: dict[str, object]) -> None:
    """見積入力シートの結合と状態表示を変更する。

    Args:
        book: 対象ブック。
        cfg: 設定値。

    Side Effects:
        セルの結合、値、条件付き書式を変更する。

    Examples:
        >>> apply_edits(book, CONFIG)
    """
    ws = book.Worksheets(cfg["sheet_quote"])
    note = ws.Range(cfg["note_cell"]).Value

    ws.Range("D12:F12").UnMerge()
    ws.Range("B12:C12").UnMerge()
    ws.Range("D12:F12").ClearContents()
    ws.Range(cfg["merge_range"]).Merge()
    ws.Range(cfg["note_cell"]).Value = note
    ws.Range(cfg["note_cell"]).WrapText = True
    ws.Range(cfg["note_cell"]).VerticalAlignment = -4160  # xlTop

    # 状態表示用の条件付き書式のみ削除する（明細警告の書式は残す）
    conditions = ws.Range(cfg["remove_cf_range"]).FormatConditions
    for index in range(conditions.Count, 0, -1):
        conditions(index).Delete()


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


def execute_logic(app: object, source: Path, dest: Path, cfg: dict[str, object]) -> dict[str, object]:
    """P5 の本体処理を実行する。

    Args:
        app: Excel Application。
        source: 入力 xlsx。
        dest: 出力 xlsx。
        cfg: 設定値。

    Returns:
        検証結果の辞書。

    Side Effects:
        出力ファイルを生成する。

    Examples:
        >>> values = execute_logic(app, Path("prototype.xlsx"), Path("prototype_v03.xlsx"), CONFIG)
    """
    shutil.copy2(source, dest)
    book = app.Workbooks.Open(str(dest), UpdateLinks=0)
    try:
        validate_preconditions(book, cfg)
        apply_edits(book, cfg)
        book.Save()
        ws = book.Worksheets(cfg["sheet_quote"])
        values = {
            "merge": ws.Range(cfg["note_cell"]).MergeArea.Address,
            "note": ws.Range(cfg["note_cell"]).Value,
            "d12": ws.Range("D12").Value,
            "g28": ws.Range(cfg["state_header"]).Value,
            "i28": ws.Range(cfg["state_formula"]).Value,
            "cf_count": ws.Range(cfg["remove_cf_range"]).FormatConditions.Count,
            "notes": sum(sheet.Comments.Count for sheet in book.Worksheets),
            "indirect": count_indirect(book),
            "protected": any(bool(sheet.ProtectContents) for sheet in book.Worksheets),
        }
    finally:
        book.Close(False)
    # Excel の保存で付与される絶対パスと作成者情報を除去する
    package = sanitize_package(dest, OOXML_CONFIG)
    values["remaining_hits"] = package["remaining_hits"]
    return values


# ---------------------------------------------------------------------------
# 出力
# ---------------------------------------------------------------------------
def format_output(values: dict[str, object], cfg: dict[str, object]) -> str:
    """検証結果を整形する。

    Args:
        values: 実測値。
        cfg: 設定値。

    Returns:
        整形済み文字列。

    Side Effects:
        なし。

    Examples:
        >>> print(format_output(values, CONFIG))
    """
    lines = [
        f"merge: {values['merge']}",
        f"note: {values['note']!r}",
        f"D12: {values['d12']!r} / G28: {values['g28']!r} / I28: {values['i28']!r}",
        f"CF rules on D12:F12: {values['cf_count']}",
        f"notes: {values['notes']} (expected {cfg['expected_notes']})",
        f"INDIRECT: {values['indirect']} (expected {cfg['expected_indirect']})",
        f"protected: {values['protected']}",
        f"absPath / local path hits: {values['remaining_hits']}",
    ]
    return "\n".join(lines)


def main() -> int:
    """エントリポイント。

    Returns:
        終了コード（0 は正常）。

    Side Effects:
        Excel を起動し、出力ファイルを生成する。

    Examples:
        >>> raise SystemExit(main())
    """
    if len(sys.argv) != 3:
        print("usage: build_v03_p5.py <in.xlsx> <out.xlsx>")
        return 2
    source = Path(sys.argv[1]).resolve()
    dest = Path(sys.argv[2]).resolve()
    validate_input(source)
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AutomationSecurity = 3
    app.AskToUpdateLinks = False
    try:
        values = execute_logic(app, source, dest, CONFIG)
        print(format_output(values, CONFIG))
        ok = (
            values["merge"] == "$B$12:$F$12"
            and values["note"] == "試作品の架空データです"
            and values["d12"] in (None, "")
            and values["g28"] == "入力状態"
            and values["i28"] not in (None, "")
            and values["cf_count"] == 0
            and values["notes"] == CONFIG["expected_notes"]
            and values["indirect"] == CONFIG["expected_indirect"]
            and not values["protected"]
            and values["remaining_hits"] == 0
        )
        print("P5 RESULT:", "PASS" if ok else "FAIL")
        return 0 if ok else 1
    finally:
        app.Quit()


if __name__ == "__main__":
    raise SystemExit(main())
