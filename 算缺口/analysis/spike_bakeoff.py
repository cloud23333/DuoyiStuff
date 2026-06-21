"""Walk-forward bake-off for spike-robust daily-sales forecasting.

Reproduces the evidence behind the 2026-06-21 spike-robust-level refactor:
profiles intermittency, measures spike reversion, and scores candidate level
estimators + the full stocking decision against a trailing holdout window.

Run:
    ./.venv/bin/python analysis/spike_bakeoff.py --data /path/to/data-dir

Reads the Temu daily-sales detail xlsx (平台SKC_ID/平台SKU_ID + 日销量 columns)
directly so it works regardless of the production parser's header expectations.
"""

from __future__ import annotations

import argparse
import statistics
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from shipment_planner.engine import compute_forecast_metrics  # noqa: E402
from shipment_planner.forecast_distribution import horizon_quantile  # noqa: E402
from shipment_planner.forecast_level import (  # noqa: E402
    robust_level,
    weighted_mean,
    weighted_variance,
)
from shipment_planner.parsers import parse_int, temu_daily_sales_columns  # noqa: E402
from shipment_planner.xlsx_reader import read_xlsx_table  # noqa: E402

HOLDOUT = 7
MIN_TRAIN = 7
BASE_FLOOR = 2


def _col(row: dict[str, str], *names: str) -> str:
    for name in names:
        if name in row:
            return row[name]
    return ""


def load_series(data_dir: Path) -> dict[tuple[str, str], list[float]]:
    detail = next(data_dir.glob("Temu销售明细*xlsx"))
    header, rows = read_xlsx_table(detail)
    date_cols = temu_daily_sales_columns(header)
    series: dict[tuple[str, str], list[float]] = {}
    for row in rows:
        skc = _col(row, "平台SKC_ID", "平台SKCID").strip()
        sku = _col(row, "平台SKU_ID", "平台SKUID").strip()
        if not skc or not sku:
            continue
        daily = [float(parse_int(row.get(c))) for c in date_cols]
        key = (skc, sku)
        if key in series:
            series[key] = [a + b for a, b in zip(series[key], daily)]
        else:
            series[key] = daily
    return {k: v for k, v in series.items() if sum(v) > 0}


def recommended_decision(train: list[float], horizon: int, service_level: float) -> float:
    level = robust_level(train)
    raw_mean = weighted_mean(train, 5.0)
    variance = max(weighted_variance(train, 5.0, mean=raw_mean), level)
    quantile = horizon_quantile(
        mean=level * horizon, variance=variance * horizon, probability=service_level
    )
    return max(quantile, BASE_FLOOR)


def current_decision(train: list[float], horizon: int) -> float:
    metrics = compute_forecast_metrics(
        daily_sales=tuple(int(x) for x in train),
        stocking_days=horizon,
        is_hot_style=False,
        stock_in_warehouse=0.0,
        apply_stockout_mask=False,
    )
    return max(metrics.forecast_stocking_period_sales, BASE_FLOOR)


def is_spike_origin(train: list[float]) -> bool:
    base_pos = [x for x in train[:-1] if x > 0]
    base = statistics.median(base_pos) if base_pos else 0.0
    threshold = max(3 * max(1.0, base), 10)
    last_two_avg = statistics.fmean(train[-2:]) if len(train) >= 2 else train[-1]
    return train[-1] >= threshold or last_two_avg >= max(3 * max(1.0, base), 8)


def build_samples(series: dict) -> list[tuple[bool, list[float], int]]:
    samples = []
    for daily in series.values():
        for t in range(MIN_TRAIN, len(daily) - HOLDOUT + 1):
            train = daily[:t]
            actual = int(sum(daily[t : t + HOLDOUT]))
            if sum(train) <= 0:
                continue
            samples.append((is_spike_origin(train), train, actual))
    return samples


def report(series: dict) -> None:
    samples = build_samples(series)
    print(f"active SKUs={len(series)}  origins={len(samples)}  holdout={HOLDOUT}d\n")

    def line(label, decider):
        over = under = served = demand = stockout = 0.0
        spike_over = spike_under = 0.0
        for spike, train, actual in samples:
            d = decider(train)
            over += max(0, d - actual)
            under += max(0, actual - d)
            served += min(d, actual)
            demand += actual
            stockout += 1 if d < actual else 0
            if spike:
                spike_over += max(0, d - actual)
                spike_under += max(0, actual - d)
        n = len(samples)
        print(
            f"{label:30} overstock={over:8.0f}  fill={served/max(demand,1)*100:4.1f}%  "
            f"stockout={stockout/n*100:4.1f}%  | spike over={spike_over:5.0f} under={spike_under:5.0f}"
        )

    line("CURRENT (prod)", lambda t: current_decision(t, HOLDOUT))
    for sl in (0.40, 0.45, 0.50, 0.60):
        line(f"RECOMMENDED @ SL={sl}", lambda t, sl=sl: recommended_decision(t, HOLDOUT, sl))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data", required=True, help="Directory holding the Temu xlsx exports")
    args = parser.parse_args()
    report(load_series(Path(args.data)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
