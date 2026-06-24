"""Policy-cost backtest harness.

For a grid of conservatism knobs ``(service_level_offset, alpha)``, replays each
SKU's recent history and reports the three numbers the operator trades between:
units shipped, simulated stockout (lost-sale) units, and simulated leftover
(overstock) units. The operator can then pick a setting on the efficient
frontier instead of guessing.

This complements ``eval_forecast.py`` (which measures forecast accuracy). Here the
shipping target is the service-level QUANTILE of demand over the holdout window
(``forecast_stocking_period_sales``), scaled by ``alpha`` — the actual order-up-to
target — NOT the predictive mean.

Deliberate simplification: each holdout window is assumed to start from ZERO
available stock and be served by a single shipment of ``target`` units. The repo
has no historical stock-position snapshots, so an exact multi-day inventory replay
is impossible. This makes the ABSOLUTE cost numbers optimistic, but the RELATIVE
ordering across knob settings — which is what the operator is choosing between — is
sound.

Outputs:
  - per-(sku, knob) CSV with target/shipped/actual/lost/leftover
  - grid JSON with per-cell totals and fill rate
"""

from __future__ import annotations

import argparse
import csv
import json
import math
from collections.abc import Sequence
from dataclasses import asdict, dataclass
from pathlib import Path

from .engine import compute_forecast_metrics
from .models import SalesRecord
from .parsers import (
    SALES_REQUIRED_COLUMNS,
    assert_required_columns,
    assert_temu_daily_sales_columns,
    assert_xlsx,
    parse_sales,
    parse_temu_daily_sales,
)
from .xlsx_reader import read_xlsx_table

DEFAULT_HOLDOUT_DAYS = 7
DEFAULT_MIN_TRAIN_DAYS = 13
DEFAULT_OFFSETS = "0.0,-0.05,-0.10,-0.15,-0.20"
DEFAULT_ALPHAS = "1.0,0.9,0.8,0.7,0.6"


@dataclass(frozen=True)
class PolicyEvalRow:
    skc: str
    skuid: str
    offset: float
    alpha: float
    train_days: int
    holdout_days: int
    target: float
    actual: int
    shipped: int
    lost: int
    leftover: int


def simulate_sku(
    *,
    train: tuple[int, ...],
    test: tuple[int, ...],
    stock_in_warehouse: float,
    offset: float,
    alpha: float,
    apply_stockout_mask: bool = True,
) -> tuple[float, int, int, int, int]:
    """Return (target, actual, shipped, lost, leftover) for one knob setting.

    ``target`` is the service-level quantile of demand over the holdout window
    (``forecast_stocking_period_sales``) scaled by ``alpha``. The window is assumed
    to start from zero stock and be served by a single shipment of ``shipped`` units.

    ponytail: start-at-zero / single-cycle model — no inventory state to replay, so
    we deliberately do not reconstruct true stock positions.
    """
    metrics = compute_forecast_metrics(
        daily_sales=train,
        stocking_days=len(test),
        stock_in_warehouse=stock_in_warehouse,
        apply_stockout_mask=apply_stockout_mask,
        service_level_offset=offset,
    )
    target = alpha * metrics.forecast_stocking_period_sales
    shipped = round(target)
    actual = int(sum(test))
    lost = max(0, actual - shipped)
    leftover = max(0, shipped - actual)
    return target, actual, shipped, lost, leftover


def run_policy_grid(
    *,
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]],
    sales_by_key: dict[tuple[str, str], SalesRecord],
    offsets: Sequence[float],
    alphas: Sequence[float],
    holdout_days: int = DEFAULT_HOLDOUT_DAYS,
    min_train_days: int = DEFAULT_MIN_TRAIN_DAYS,
    apply_stockout_mask: bool = True,
) -> tuple[list[PolicyEvalRow], int, list[dict[str, object]]]:
    """Returns (per-(sku,knob) rows, skipped_count, grid_summary).

    grid_summary has one dict per (offset, alpha) cell with totals:
    offset, alpha, evaluated_skus, ship_units, lost_units, leftover_units,
    fill_rate ( = served / actual = 1 - lost/actual ).
    """
    rows: list[PolicyEvalRow] = []
    skipped = 0
    cells: dict[tuple[float, float], dict[str, float]] = {
        (offset, alpha): {
            "evaluated_skus": 0,
            "ship_units": 0,
            "lost_units": 0,
            "leftover_units": 0,
            "actual_units": 0,
        }
        for offset in offsets
        for alpha in alphas
    }

    for key, daily in daily_sales_by_key.items():
        if len(daily) < holdout_days + min_train_days:
            skipped += 1
            continue

        train = daily[:-holdout_days]
        test = daily[-holdout_days:]
        sales = sales_by_key.get(key)
        stock_in_warehouse = (
            getattr(sales, "stock_in_warehouse", 0.0) if sales is not None else 0.0
        )

        for offset in offsets:
            for alpha in alphas:
                target, actual, shipped, lost, leftover = simulate_sku(
                    train=train,
                    test=test,
                    stock_in_warehouse=stock_in_warehouse,
                    offset=offset,
                    alpha=alpha,
                    apply_stockout_mask=apply_stockout_mask,
                )
                rows.append(
                    PolicyEvalRow(
                        skc=key[0],
                        skuid=key[1],
                        offset=offset,
                        alpha=alpha,
                        train_days=len(train),
                        holdout_days=holdout_days,
                        target=_round(target),
                        actual=actual,
                        shipped=shipped,
                        lost=lost,
                        leftover=leftover,
                    )
                )
                cell = cells[(offset, alpha)]
                cell["evaluated_skus"] += 1
                cell["ship_units"] += shipped
                cell["lost_units"] += lost
                cell["leftover_units"] += leftover
                cell["actual_units"] += actual

    grid_summary: list[dict[str, object]] = []
    for offset in offsets:
        for alpha in alphas:
            cell = cells[(offset, alpha)]
            actual_units = cell["actual_units"]
            fill_rate = (
                _round(1 - cell["lost_units"] / actual_units)
                if actual_units > 0
                else None
            )
            grid_summary.append(
                {
                    "offset": offset,
                    "alpha": alpha,
                    "evaluated_skus": int(cell["evaluated_skus"]),
                    "ship_units": int(cell["ship_units"]),
                    "lost_units": int(cell["lost_units"]),
                    "leftover_units": int(cell["leftover_units"]),
                    "fill_rate": fill_rate,
                }
            )
    return rows, skipped, grid_summary


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Backtest the policy cost (ship/lost/leftover) over a knob grid."
    )
    parser.add_argument("--sales", required=True, help="Sales xlsx (for stock)")
    parser.add_argument("--temu-sales", required=True, help="Temu daily sales xlsx")
    parser.add_argument("--out-dir", default="data/output", help="Output directory (default: data/output)")
    parser.add_argument(
        "--holdout-days",
        type=int,
        default=DEFAULT_HOLDOUT_DAYS,
        help="Trailing days to hold out (default: 7)",
    )
    parser.add_argument(
        "--min-train-days",
        type=int,
        default=DEFAULT_MIN_TRAIN_DAYS,
        help="Skip SKUs with fewer training days than this (default: 13)",
    )
    parser.add_argument(
        "--offsets",
        default=DEFAULT_OFFSETS,
        help=f"Comma-separated service-level offsets (default: {DEFAULT_OFFSETS})",
    )
    parser.add_argument(
        "--alphas",
        default=DEFAULT_ALPHAS,
        help=f"Comma-separated protection-interval factors (default: {DEFAULT_ALPHAS})",
    )
    parser.add_argument(
        "--no-stockout-mask",
        action="store_true",
        help="Disable the OOS trailing-zero mask (evaluate raw history)",
    )
    return parser


def _parse_float_list(raw: str) -> list[float]:
    return [float(part.strip()) for part in raw.split(",") if part.strip()]


def main(argv: list[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    sales_path = Path(args.sales)
    temu_path = Path(args.temu_sales)
    assert_xlsx(sales_path)
    assert_xlsx(temu_path)

    if args.holdout_days <= 0:
        raise ValueError("--holdout-days must be > 0")
    if args.min_train_days <= 0:
        raise ValueError("--min-train-days must be > 0")

    offsets = _parse_float_list(args.offsets)
    alphas = _parse_float_list(args.alphas)
    if not offsets:
        raise ValueError("--offsets must contain at least one value")
    if not alphas:
        raise ValueError("--alphas must contain at least one value")

    sales_header, sales_rows = read_xlsx_table(sales_path)
    assert_required_columns(sales_header, SALES_REQUIRED_COLUMNS, "sales file")
    sales_records = parse_sales(sales_rows)

    temu_header, temu_rows = read_xlsx_table(temu_path)
    assert_temu_daily_sales_columns(temu_header, "Temu daily sales file")
    daily_sales_by_key = parse_temu_daily_sales(temu_rows)

    sales_by_key: dict[tuple[str, str], SalesRecord] = {
        (record.skc, record.skuid): record for record in sales_records
    }

    rows, skipped, grid_summary = run_policy_grid(
        daily_sales_by_key=daily_sales_by_key,
        sales_by_key=sales_by_key,
        offsets=offsets,
        alphas=alphas,
        holdout_days=args.holdout_days,
        min_train_days=args.min_train_days,
        apply_stockout_mask=not args.no_stockout_mask,
    )

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    detail_path = out_dir / "policy_eval_detail.csv"
    grid_path = out_dir / "policy_eval_grid.json"

    _write_detail_csv(detail_path, rows)
    grid_path.write_text(
        json.dumps(grid_summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    evaluated_skus = grid_summary[0]["evaluated_skus"] if grid_summary else 0
    print(f"Evaluated SKUs: {evaluated_skus}  Skipped: {skipped}")
    print(
        f"{'offset':>8}  {'alpha':>6}  {'ship_units':>11}  "
        f"{'lost_units':>11}  {'leftover_units':>15}  {'fill_rate':>9}"
    )
    for cell in grid_summary:
        fill_rate = cell["fill_rate"]
        fill_rate_str = "None" if fill_rate is None else f"{fill_rate:.4f}"
        print(
            f"{cell['offset']:>8}  {cell['alpha']:>6}  {cell['ship_units']:>11}  "
            f"{cell['lost_units']:>11}  {cell['leftover_units']:>15}  {fill_rate_str:>9}"
        )
    print(f"Outputs: {detail_path}")
    print(f"         {grid_path}")
    return 0


def _write_detail_csv(path: Path, rows: Sequence[PolicyEvalRow]) -> None:
    fieldnames = [
        "skc",
        "skuid",
        "offset",
        "alpha",
        "train_days",
        "holdout_days",
        "target",
        "actual",
        "shipped",
        "lost",
        "leftover",
    ]
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(asdict(row) for row in rows)


def _round(value: float) -> float:
    if math.isnan(value) or math.isinf(value):
        return value
    return round(value, 4)


if __name__ == "__main__":
    raise SystemExit(main())
