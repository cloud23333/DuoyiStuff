"""Forecast eval harness.

Holds out the last N days of each SKU's daily sales, runs the forecast
pipeline on the remaining history, and compares the predicted period
total against the actual holdout sum. Because the operator reruns daily
with fresh data, the useful signal is "how well does my model predict
the next ``holdout_days``-window of demand given the data I had ``holdout_days``
ago?"

Outputs:
  - per-SKU CSV with predicted/actual totals and signed error
  - summary JSON with WAPE, MAE, bias, per-strategy breakdown
"""

from __future__ import annotations

import argparse
import csv
import json
import math
from collections import defaultdict
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

from .engine import compute_forecast_metrics
from .forecast_benchmark import BenchmarkRow, run_forecast_benchmark
from .models import SalesRecord
from .parsers import (
    SALES_REQUIRED_COLUMNS,
    TEMU_DAILY_REQUIRED_COLUMNS,
    assert_required_columns,
    assert_temu_daily_sales_columns,
    assert_xlsx,
    parse_sales,
    parse_temu_daily_sales,
)
from .xlsx_reader import read_xlsx_table

DEFAULT_HOLDOUT_DAYS = 7
DEFAULT_MIN_TRAIN_DAYS = 14


@dataclass(frozen=True)
class EvalRow:
    skc: str
    skuid: str
    strategy: str
    train_days: int
    holdout_days: int
    forecast_daily_sales: float
    forecast_holdout_total: float
    actual_holdout_total: int
    abs_error: float
    signed_error: float


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Evaluate forecast accuracy over a holdout window.")
    parser.add_argument("--sales", required=True, help="Sales xlsx (for stocking_days, is_hot_style, stock)")
    parser.add_argument("--temu-sales", required=True, help="Temu daily sales xlsx")
    parser.add_argument("--holdout-days", type=int, default=7, help="Trailing days to hold out (default: 7)")
    parser.add_argument("--out-dir", default="data/output", help="Output directory (default: data/output)")
    parser.add_argument(
        "--min-train-days",
        type=int,
        default=14,
        help="Skip SKUs with fewer training days than this (default: 14)",
    )
    parser.add_argument(
        "--no-stockout-mask",
        action="store_true",
        help="Disable the OOS trailing-zero mask (evaluate raw history)",
    )
    parser.add_argument(
        "--benchmark",
        action="store_true",
        help="Also compare current forecast against intermittent-demand benchmark models",
    )
    return parser


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

    sales_header, sales_rows = read_xlsx_table(sales_path)
    assert_required_columns(sales_header, SALES_REQUIRED_COLUMNS, "sales file")
    sales_records = parse_sales(sales_rows)

    temu_header, temu_rows = read_xlsx_table(temu_path)
    assert_temu_daily_sales_columns(temu_header, "Temu daily sales file")
    assert_required_columns(temu_header, TEMU_DAILY_REQUIRED_COLUMNS, "Temu daily sales file")
    daily_sales_by_key = parse_temu_daily_sales(temu_rows)

    rows, skipped, summary = run_holdout_eval(
        sales_records=sales_records,
        daily_sales_by_key=daily_sales_by_key,
        holdout_days=args.holdout_days,
        min_train_days=args.min_train_days,
        apply_stockout_mask=not args.no_stockout_mask,
    )

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    detail_path = out_dir / "forecast_eval_detail.csv"
    summary_path = out_dir / "forecast_eval_summary.json"

    _write_detail_csv(detail_path, rows)
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Evaluated SKUs: {summary['evaluated_skus']}")
    print(f"Skipped (insufficient history): {summary['skipped_insufficient_history']}")
    print(f"WAPE: {summary['wape']}  MAE: {summary['mae']}  Bias: {summary['bias']}")
    print("By strategy:")
    for strategy, stats in summary["by_strategy"].items():
        print(
            f"  {strategy}: n={stats['count']} "
            f"WAPE={stats['wape']} MAE={stats['mae']} bias={stats['bias']}"
        )
    if args.benchmark:
        benchmark_rows, _, benchmark_summary = run_forecast_benchmark(
            sales_records=sales_records,
            daily_sales_by_key=daily_sales_by_key,
            holdout_days=args.holdout_days,
            min_train_days=args.min_train_days,
            apply_stockout_mask=not args.no_stockout_mask,
        )
        benchmark_detail_path = out_dir / "forecast_benchmark_detail.csv"
        benchmark_summary_path = out_dir / "forecast_benchmark_summary.json"
        _write_benchmark_detail_csv(benchmark_detail_path, benchmark_rows)
        benchmark_summary_path.write_text(
            json.dumps(benchmark_summary, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print("Benchmark by model:")
        for model, stats in benchmark_summary["by_model"].items():
            print(
                f"  {model}: n={stats['count']} "
                f"WAPE={stats['wape']} MAE={stats['mae']} bias={stats['bias']}"
            )
        print(f"         {benchmark_detail_path}")
        print(f"         {benchmark_summary_path}")
    print(f"Outputs: {detail_path}")
    print(f"         {summary_path}")
    return 0


def run_holdout_eval(
    *,
    sales_records: Sequence[SalesRecord],
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]],
    holdout_days: int = DEFAULT_HOLDOUT_DAYS,
    min_train_days: int = DEFAULT_MIN_TRAIN_DAYS,
    apply_stockout_mask: bool = True,
) -> tuple[list[EvalRow], int, dict[str, object]]:
    """Evaluate forecast accuracy on a trailing holdout window.

    Returns (per-SKU rows, skipped-due-to-insufficient-history, summary dict).
    Reusable from the main CLI so recommendations can cite per-SKU error.
    """
    sales_by_key: dict[tuple[str, str], SalesRecord] = {
        (record.skc, record.skuid): record for record in sales_records
    }
    rows, skipped = _evaluate(
        daily_sales_by_key=daily_sales_by_key,
        sales_by_key=sales_by_key,
        holdout_days=holdout_days,
        min_train_days=min_train_days,
        apply_stockout_mask=apply_stockout_mask,
    )
    summary = _build_summary(
        rows,
        holdout_days=holdout_days,
        skipped_insufficient_history=skipped,
    )
    return rows, skipped, summary


def _evaluate(
    *,
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]],
    sales_by_key: dict[tuple[str, str], SalesRecord],
    holdout_days: int,
    min_train_days: int,
    apply_stockout_mask: bool,
) -> tuple[list[EvalRow], int]:
    rows: list[EvalRow] = []
    skipped = 0
    for key, daily in daily_sales_by_key.items():
        if len(daily) < holdout_days + min_train_days:
            skipped += 1
            continue

        train = daily[:-holdout_days]
        test = daily[-holdout_days:]

        sales = sales_by_key.get(key)
        is_hot_style = getattr(sales, "is_hot_style", False) if sales is not None else False
        stock_in_warehouse = (
            getattr(sales, "stock_in_warehouse", 0.0) if sales is not None else 0.0
        )

        metrics = compute_forecast_metrics(
            daily_sales=train,
            stocking_days=holdout_days,
            is_hot_style=is_hot_style,
            stock_in_warehouse=stock_in_warehouse,
            apply_stockout_mask=apply_stockout_mask,
        )

        forecast_holdout_total = metrics.forecast_daily_sales * holdout_days
        actual_total = int(sum(test))
        signed_error = forecast_holdout_total - actual_total

        rows.append(
            EvalRow(
                skc=key[0],
                skuid=key[1],
                strategy=metrics.strategy,
                train_days=len(train),
                holdout_days=holdout_days,
                forecast_daily_sales=_round(metrics.forecast_daily_sales),
                forecast_holdout_total=_round(forecast_holdout_total),
                actual_holdout_total=actual_total,
                abs_error=_round(abs(signed_error)),
                signed_error=_round(signed_error),
            )
        )
    return rows, skipped


def _write_detail_csv(path: Path, rows: Sequence[EvalRow]) -> None:
    fieldnames = [
        "skc",
        "skuid",
        "strategy",
        "train_days",
        "holdout_days",
        "forecast_daily_sales",
        "forecast_holdout_total",
        "actual_holdout_total",
        "abs_error",
        "signed_error",
    ]
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    "skc": row.skc,
                    "skuid": row.skuid,
                    "strategy": row.strategy,
                    "train_days": row.train_days,
                    "holdout_days": row.holdout_days,
                    "forecast_daily_sales": row.forecast_daily_sales,
                    "forecast_holdout_total": row.forecast_holdout_total,
                    "actual_holdout_total": row.actual_holdout_total,
                    "abs_error": row.abs_error,
                    "signed_error": row.signed_error,
                }
            )


def _write_benchmark_detail_csv(path: Path, rows: Sequence[BenchmarkRow]) -> None:
    fieldnames = [
        "skc",
        "skuid",
        "model",
        "train_days",
        "holdout_days",
        "forecast_daily_sales",
        "forecast_holdout_total",
        "actual_holdout_total",
        "abs_error",
        "signed_error",
    ]
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    "skc": row.skc,
                    "skuid": row.skuid,
                    "model": row.model,
                    "train_days": row.train_days,
                    "holdout_days": row.holdout_days,
                    "forecast_daily_sales": row.forecast_daily_sales,
                    "forecast_holdout_total": row.forecast_holdout_total,
                    "actual_holdout_total": row.actual_holdout_total,
                    "abs_error": row.abs_error,
                    "signed_error": row.signed_error,
                }
            )


def _build_summary(
    rows: Sequence[EvalRow],
    *,
    holdout_days: int,
    skipped_insufficient_history: int,
) -> dict[str, object]:
    if not rows:
        return {
            "evaluated_skus": 0,
            "skipped_insufficient_history": skipped_insufficient_history,
            "holdout_days": holdout_days,
            "wape": None,
            "mae": None,
            "bias": None,
            "by_strategy": {},
        }

    abs_total = sum(row.abs_error for row in rows)
    actual_total = sum(row.actual_holdout_total for row in rows)
    signed_total = sum(row.signed_error for row in rows)
    count = len(rows)

    by_strategy_rows: dict[str, list[EvalRow]] = defaultdict(list)
    for row in rows:
        by_strategy_rows[row.strategy].append(row)

    by_strategy: dict[str, dict[str, object]] = {}
    for strategy, strategy_rows in by_strategy_rows.items():
        strategy_abs = sum(r.abs_error for r in strategy_rows)
        strategy_actual = sum(r.actual_holdout_total for r in strategy_rows)
        strategy_signed = sum(r.signed_error for r in strategy_rows)
        strategy_count = len(strategy_rows)
        by_strategy[strategy] = {
            "count": strategy_count,
            "mae": _round(strategy_abs / strategy_count),
            "wape": _round(strategy_abs / strategy_actual) if strategy_actual > 0 else None,
            "bias": _round(strategy_signed / strategy_count),
        }

    return {
        "evaluated_skus": count,
        "skipped_insufficient_history": skipped_insufficient_history,
        "holdout_days": holdout_days,
        "mae": _round(abs_total / count),
        "wape": _round(abs_total / actual_total) if actual_total > 0 else None,
        "bias": _round(signed_total / count),
        "by_strategy": by_strategy,
    }


def _round(value: float) -> float:
    if math.isnan(value) or math.isinf(value):
        return value
    return round(value, 4)


if __name__ == "__main__":
    raise SystemExit(main())
