"""v0.6 用の決定的な架空売上データを生成する。

Purpose:
    共通 fixture を変更せず、2025年の基準と差が見える v0.6 専用 fixture を作る。
Specification:
    2025年の既存値を維持し、2026年を年次売上約8%増、商品・月別に異なる構成で生成する。
Assumptions / Dependencies:
    入力はプロジェクト共通の sales_fixture.json であり、Python 標準ライブラリのみを使う。
I/O:
    sales_fixture.json を読み、sales_fixture_v06.json を出力する。
Caution:
    乱数を使わない。v0.4 / v0.5 のブックと共有 fixture は変更しない。
Future Work:
    商品マスタが変更された場合は、商品別係数と入力検査を更新する。
Change Log:
    2026-10-22: v0.6 用に年次成長、季節性、商品構成を決定的に生成する。
"""

from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# 設定値
# ---------------------------------------------------------------------------
PROJECT_DIR = Path(__file__).resolve().parent
SOURCE_PATH = PROJECT_DIR / "sales_fixture.json"
OUTPUT_PATH = PROJECT_DIR / "sales_fixture_v06.json"
TARGET_2026_GROWTH = 0.08
EXPECTED_ROWS = 576
ROWS_PER_PRODUCT_MONTH = 3

# 月ごとの係数で、2026年は春・初夏・秋に別の売上ピークを持たせる。
MONTH_FACTORS: dict[int, float] = {
    1: 0.78,
    2: 0.84,
    3: 1.08,
    4: 1.18,
    5: 1.32,
    6: 1.20,
    7: 0.83,
    8: 0.74,
    9: 0.92,
    10: 1.24,
    11: 1.12,
    12: 0.75,
}

# 商品構成を変え、年比較グラフとヒートマップで商品の違いを読み取れるようにする。
PRODUCT_FACTORS: dict[str, float] = {
    "テーブル": 0.78,
    "椅子": 1.58,
    "収納棚": 1.15,
    "木製小物": 1.65,
    "野菜5kg箱": 1.34,
    "トマト箱": 1.72,
    "米5kg": 0.82,
    "季節の果物": 1.95,
}

# 2026年における商品ごとの計上原価率の変化。値は比率の差である。
COST_RATE_SHIFTS: dict[str, float] = {
    "テーブル": 0.010,
    "椅子": -0.030,
    "収納棚": 0.025,
    "木製小物": -0.020,
    "野菜5kg箱": 0.040,
    "トマト箱": -0.025,
    "米5kg": 0.015,
    "季節の果物": -0.040,
}

MONTH_COST_SHIFTS: tuple[float, ...] = (
    -0.010, 0.000, 0.010, 0.015, 0.020, 0.010,
    -0.005, -0.015, 0.000, 0.020, 0.010, -0.010,
)
TRANSACTION_DAYS: tuple[int, ...] = (5, 15, 25)


def validate_input(records: list[dict[str, Any]]) -> None:
    """共通 fixture の行数、年度、商品構成を検証する。

    Args:
        records: 共通売上 fixture の行。

    Returns:
        なし。

    Raises:
        ValueError: 件数または商品構成が仕様と一致しない場合。

    Side Effects:
        なし。
    """
    years = {int(row["year"]) for row in records}
    products = {str(row["product"]) for row in records}
    if len(records) != EXPECTED_ROWS:
        raise ValueError("sales_fixture.json の行数が576件ではない")
    if years != {2025, 2026}:
        raise ValueError("sales_fixture.json は2025年・2026年の2年分が必要")
    if products != set(PRODUCT_FACTORS):
        raise ValueError("fixture の商品構成と v0.6 商品設定が一致しない")
    group_records(records)


def group_records(records: list[dict[str, Any]]) -> dict[tuple[int, str], list[dict[str, Any]]]:
    """同一年月・商品ごとに fixture の明細をまとめる。

    Args:
        records: 売上 fixture の行。

    Returns:
        (月, 商品) をキーとする2025年の明細辞書。

    Side Effects:
        なし。
    """
    grouped: dict[tuple[int, str], list[dict[str, Any]]] = defaultdict(list)
    for row in records:
        if int(row["year"]) == 2025:
            grouped[(int(row["month"]), str(row["product"]))].append(dict(row))
    for key, rows in grouped.items():
        rows.sort(key=lambda row: int(row["sales"]))
        if len(rows) != ROWS_PER_PRODUCT_MONTH:
            raise ValueError(f"2025年 {key} の明細数が3件ではない")
    return grouped


def summarize_sales(records: list[dict[str, Any]]) -> dict[int, int]:
    """年度ごとの税抜売上合計を計算する。

    Args:
        records: 売上 fixture の行。

    Returns:
        年をキーとした税抜売上合計。

    Side Effects:
        なし。
    """
    totals: dict[int, int] = defaultdict(int)
    for row in records:
        totals[int(row["year"])] += int(row["sales"])
    return dict(totals)


def balance_annual_sales(
    records: list[dict[str, Any]], target_sales: int
) -> None:
    """数量を1単位ずつ調整し、2026年の税抜売上を目標に近づける。

    Args:
        records: 2026年の生成行。数量と売上を書き換える。
        target_sales: 2026年の目標税抜売上。

    Returns:
        なし。

    Side Effects:
        対象行の qty と sales を更新する。
    """
    current = sum(int(row["sales"]) for row in records)
    for _ in range(10000):
        difference = target_sales - current
        if difference == 0:
            return
        direction = 1 if difference > 0 else -1
        eligible = [
            row
            for row in records
            if int(row["price"]) <= abs(difference)
            and (direction > 0 or int(row["qty"]) > 1)
        ]
        if not eligible:
            return
        # 一度の調整で目標へ最も近づく価格を優先し、順序は固定する。
        row = max(eligible, key=lambda item: int(item["price"]))
        row["qty"] = int(row["qty"]) + direction
        row["sales"] = int(row["qty"]) * int(row["price"])
        current += direction * int(row["price"])


def create_2026_records(
    base_2025: list[dict[str, Any]], annual_target: int
) -> list[dict[str, Any]]:
    """2025年を基準に、変動のある2026年架空データを作成する。

    Args:
        base_2025: 元の2025年 fixture。
        annual_target: 2026年の税抜売上目標。

    Returns:
        2026年の売上明細288件。

    Raises:
        ValueError: 商品・月ごとの明細が不足している場合。

    Side Effects:
        なし。
    """
    product_order = {name: index for index, name in enumerate(PRODUCT_FACTORS)}
    base_cost_rates: dict[str, float] = defaultdict(float)
    base_sales: dict[str, int] = defaultdict(int)
    for row in base_2025:
        product = str(row["product"])
        base_sales[product] += int(row["sales"])
        base_cost_rates[product] += int(row["cost"])
    for product in base_sales:
        base_cost_rates[product] /= base_sales[product]

    generated: list[dict[str, Any]] = []
    for row_index, base in enumerate(base_2025):
        month = int(base["month"])
        product = str(base["product"])
        product_index = product_order[product]
        transaction_index = row_index % ROWS_PER_PRODUCT_MONTH
        interaction = 1.0 + (((month * 7 + product_index * 3 + transaction_index * 5) % 9) - 4) * 0.04
        raw_quantity = max(
            1,
            round(
                int(base["qty"])
                * PRODUCT_FACTORS[product]
                * MONTH_FACTORS[month]
                * interaction
            ),
        )
        item = dict(base)
        item["year"] = 2026
        item["day"] = TRANSACTION_DAYS[transaction_index]
        item["qty"] = raw_quantity
        item["sales"] = raw_quantity * int(item["price"])
        generated.append(item)

    raw_total = sum(int(row["sales"]) for row in generated)
    scale = annual_target / raw_total
    for row_index, row in enumerate(generated):
        product = str(row["product"])
        month = int(row["month"])
        transaction_index = row_index % ROWS_PER_PRODUCT_MONTH
        row["qty"] = max(1, round(int(row["qty"]) * scale))
        row["sales"] = int(row["qty"]) * int(row["price"])

        cost_ratio = (
            base_cost_rates[product]
            + COST_RATE_SHIFTS[product]
            + MONTH_COST_SHIFTS[month - 1]
            + (transaction_index - 1) * 0.008
        )
        cost_ratio = min(0.78, max(0.35, cost_ratio))
        row["cost"] = round(int(row["sales"]) * cost_ratio / 100) * 100
    balance_annual_sales(generated, annual_target)

    # 数量調整後もコストは同じ決定的な商品・月・明細原価率で再計算する。
    for row_index, row in enumerate(generated):
        product = str(row["product"])
        month = int(row["month"])
        transaction_index = row_index % ROWS_PER_PRODUCT_MONTH
        cost_ratio = min(
            0.78,
            max(
                0.35,
                base_cost_rates[product]
                + COST_RATE_SHIFTS[product]
                + MONTH_COST_SHIFTS[month - 1]
                + (transaction_index - 1) * 0.008,
            ),
        )
        row["cost"] = round(int(row["sales"]) * cost_ratio / 100) * 100
    return generated


def build_fixture(source_records: list[dict[str, Any]]) -> dict[str, Any]:
    """年度成長、件数、変動のメタデータを含む v0.6 fixture を作る。

    Args:
        source_records: 共通 fixture の全行。

    Returns:
        metadata と v0.6 records を含む辞書。

    Side Effects:
        なし。
    """
    validate_input(source_records)
    baseline_2025 = [dict(row) for row in source_records if int(row["year"]) == 2025]
    baseline_totals = summarize_sales(baseline_2025)
    target_2026 = round(baseline_totals[2025] * (1 + TARGET_2026_GROWTH) / 1000) * 1000
    generated_2026 = create_2026_records(baseline_2025, target_2026)

    records: list[dict[str, Any]] = []
    for year_records in (baseline_2025, generated_2026):
        for row_index, row in enumerate(year_records):
            record = dict(row)
            record["day"] = TRANSACTION_DAYS[row_index % ROWS_PER_PRODUCT_MONTH]
            records.append(record)

    if len(records) != EXPECTED_ROWS:
        raise ValueError("v0.6 fixture の行数が576件ではない")
    totals = summarize_sales(records)
    actual_growth = totals[2026] / totals[2025] - 1
    return {
        "meta": {
            "name": "ExcelBIBasic_DeepSeek_V4.1 v0.6 synthetic sales fixture",
            "version": "1.0",
            "source": SOURCE_PATH.name,
            "row_count": EXPECTED_ROWS,
            "rows_per_product_month": ROWS_PER_PRODUCT_MONTH,
            "target_2026_growth": TARGET_2026_GROWTH,
            "actual_2026_growth": actual_growth,
            "annual_sales": totals,
            "generation": "deterministic_month_product_transaction_factors_no_randomness",
        },
        "records": records,
    }


def main() -> int:
    """共通 fixture から v0.6 専用 JSON を生成する。

    Returns:
        正常終了時は 0。

    Raises:
        FileNotFoundError: 入力 fixture がない場合。
        ValueError: 入力構造が要件と一致しない場合。

    Side Effects:
        `sales_fixture_v06.json` を作成または更新する。
    """
    source_records = json.loads(SOURCE_PATH.read_text(encoding="utf-8"))
    fixture = build_fixture(source_records)
    OUTPUT_PATH.write_text(
        json.dumps(fixture, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    meta = fixture["meta"]
    print(
        "fixture rows=%d sales_2025=%d sales_2026=%d growth=%.4f%%"
        % (
            meta["row_count"],
            meta["annual_sales"][2025],
            meta["annual_sales"][2026],
            meta["actual_2026_growth"] * 100,
        ),
        flush=True,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
