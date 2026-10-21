"""v0.3 成果物の OOXML 後処理（ローカルパスと作成者情報の除去）。

目的:
    Excel の COM 保存で混入する `x15ac:absPath`（絶対パス）と、
    docProps の作成者情報を成果物から除去する。

仕様:
    - `xl/workbook.xml` から `x15ac:absPath` を含む `mc:AlternateContent` ブロックを削除する。
    - `docProps/core.xml` の作成者と最終更新者を固定値へ置換する。
    - `docProps/app.xml` の Company と Manager を削除する。
    - ZIP のエントリ順と内容を維持して再書き出しする。

前提:
    - 入力は Excel で保存済みの `.xlsx` であり、Excel が閉じていること。
    - Python 標準ライブラリのみを使用する。

入出力:
    - 入力: 引数 1（xlsx のパス）。
    - 出力: なし（入力ファイルを後処理する）。

注意点:
    - Excel で再保存すると absPath が再び付与されるため、保存のたびに実行する。
    - 数式・書式・グラフには触れない。

拡張ポイント:
    - カスタムプロパティが増えた場合は、対象パートを追加する。

変更履歴:
    2026-10-21: 初版作成（P5 の後処理）。
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

# ---------------------------------------------------------------------------
# 設定値
# ---------------------------------------------------------------------------
CONFIG: dict[str, str] = {
    # ブック本体のパート
    "workbook_part": "xl/workbook.xml",
    # 文書プロパティのパート
    "core_part": "docProps/core.xml",
    "app_part": "docProps/app.xml",
    # 作成者情報の固定値
    "creator": "ExcelBIBasic3",
    # 除去対象の目印
    "abs_path_marker": "x15ac:absPath",
}

ABS_PATH_BLOCK = re.compile(
    r"<mc:AlternateContent\b[^>]*>.*?</mc:AlternateContent>", re.DOTALL
)


def strip_abs_path(xml: str, cfg: dict[str, str]) -> tuple[str, int]:
    """`x15ac:absPath` を含む代替コンテンツブロックを削除する。

    Args:
        xml: workbook.xml の文字列。
        cfg: 設定値。

    Returns:
        (除去後の文字列, 除去したブロック数)。

    Side Effects:
        なし。

    Examples:
        >>> strip_abs_path(xml, CONFIG)
    """
    removed = 0

    def replace(match: re.Match[str]) -> str:
        nonlocal removed
        if cfg["abs_path_marker"] in match.group(0):
            removed += 1
            return ""
        return match.group(0)

    return ABS_PATH_BLOCK.sub(replace, xml), removed


def sanitize_core(xml: str, cfg: dict[str, str]) -> str:
    """作成者と最終更新者を固定値へ置換する。

    Args:
        xml: core.xml の文字列。
        cfg: 設定値。

    Returns:
        置換後の文字列。

    Side Effects:
        なし。

    Examples:
        >>> sanitize_core(xml, CONFIG)
    """
    creator = cfg["creator"]
    xml = re.sub(
        r"(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)",
        rf"\g<1>{creator}\g<2>",
        xml,
    )
    xml = re.sub(
        r"(<dc:creator>)[^<]*(</dc:creator>)",
        rf"\g<1>{creator}\g<2>",
        xml,
    )
    return xml


def sanitize_app(xml: str) -> str:
    """Company と Manager の要素を削除する。

    Args:
        xml: app.xml の文字列。

    Returns:
        削除後の文字列。

    Side Effects:
        なし。

    Examples:
        >>> sanitize_app(xml)
    """
    xml = re.sub(r"<Company>[^<]*</Company>", "", xml)
    xml = re.sub(r"<Manager>[^<]*</Manager>", "", xml)
    return xml


def rewrite_package(path: Path, replacements: dict[str, bytes]) -> None:
    """ZIP を再書き出しして対象パートを置換する。

    Args:
        path: 対象 xlsx。
        replacements: パート名と置換後のバイト列。

    Side Effects:
        入力ファイルを上書きする。

    Examples:
        >>> rewrite_package(Path("prototype_v03.xlsx"), {"xl/workbook.xml": data})
    """
    tmp = path.with_name(path.stem + ".tmp" + path.suffix)
    with zipfile.ZipFile(path, "r") as src, zipfile.ZipFile(
        tmp, "w", zipfile.ZIP_DEFLATED
    ) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename in replacements:
                data = replacements[item.filename]
            dst.writestr(item, data)
    tmp.replace(path)


def sanitize_package(path: Path, cfg: dict[str, str]) -> dict[str, object]:
    """パッケージ全体を後処理する。

    Args:
        path: 対象 xlsx。
        cfg: 設定値。

    Returns:
        処理結果（absPath 除去数、作成者置換の有無、残存パス件数）。

    Side Effects:
        入力ファイルを上書きする。

    Examples:
        >>> sanitize_package(Path("prototype_v03.xlsx"), CONFIG)
    """
    replacements: dict[str, bytes] = {}
    removed_blocks = 0
    with zipfile.ZipFile(path, "r") as z:
        names = set(z.namelist())
        workbook = z.read(cfg["workbook_part"]).decode("utf-8")
        workbook, removed_blocks = strip_abs_path(workbook, cfg)
        replacements[cfg["workbook_part"]] = workbook.encode("utf-8")

        if cfg["core_part"] in names:
            core = z.read(cfg["core_part"]).decode("utf-8")
            replacements[cfg["core_part"]] = sanitize_core(core, cfg).encode("utf-8")

        if cfg["app_part"] in names:
            app = z.read(cfg["app_part"]).decode("utf-8")
            replacements[cfg["app_part"]] = sanitize_app(app).encode("utf-8")

    rewrite_package(path, replacements)

    remaining = 0
    with zipfile.ZipFile(path, "r") as z:
        for name in z.namelist():
            data = z.read(name).decode("utf-8", "ignore")
            if cfg["abs_path_marker"] in data or "C:\\Users" in data:
                remaining += 1
    return {
        "removed_blocks": removed_blocks,
        "creator": cfg["creator"],
        "remaining_hits": remaining,
    }


def main() -> int:
    """エントリポイント。

    Returns:
        終了コード（0 は正常）。

    Side Effects:
        入力ファイルを後処理する。

    Examples:
        >>> raise SystemExit(main())
    """
    if len(sys.argv) != 2:
        print("usage: finalize_v03_ooxml.py <xlsx>")
        return 2
    path = Path(sys.argv[1]).resolve()
    if not path.is_file():
        print("入力ファイルが見つかりません")
        return 2
    result = sanitize_package(path, CONFIG)
    print(result)
    return 0 if result["remaining_hits"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
