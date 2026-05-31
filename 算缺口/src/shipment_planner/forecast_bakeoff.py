# src/shipment_planner/forecast_bakeoff.py
from __future__ import annotations

from collections import defaultdict
from collections.abc import Sequence

from .eval_forecast import (
    classify_holdout_segment,
    fill_rate,
    overstock_units,
    pinball_loss,
)
from .forecast_distribution import (
    hurdle_distribution,
    negbin_ewma_distribution,
    poisson_ewma_distribution,
)
from .forecast_level import has_recent_drop

CANDIDATE_ESTIMATORS = ("negbin_ewma", "poisson_ewma", "hurdle")

_ESTIMATORS = {
    "negbin_ewma": negbin_ewma_distribution,
    "poisson_ewma": poisson_ewma_distribution,
    "hurdle": hurdle_distribution,
}


def run_bakeoff(
    *,
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]],
    stock_by_key: dict[tuple[str, str], float],
    hot_by_key: dict[tuple[str, str], bool],
    holdout_days: int,
    min_train_days: int,
    service_level: float,
) -> dict[str, object]:
    rows: dict[str, list[dict[str, float]]] = defaultdict(list)
    for key, daily in daily_sales_by_key.items():
        if len(daily) < holdout_days + min_train_days:
            continue
        train = list(daily[:-holdout_days])
        test = list(daily[-holdout_days:])
        actual = float(sum(test))
        train_mean = sum(train) / len(train) if train else 0.0
        segment = classify_holdout_segment(
            train_mean=train_mean, actual_total=actual, holdout_days=holdout_days
        )
        stock = stock_by_key.get(key, 0.0)
        recent_drop = has_recent_drop(train, stock_in_warehouse=stock)
        for name, estimator in _ESTIMATORS.items():
            dist = estimator(train, horizon=holdout_days, recent_drop=recent_drop)
            quantile = dist.quantile(service_level)
            rows[name].append(
                {
                    "mean": dist.mean,
                    "quantile": quantile,
                    "actual": actual,
                    "abs_error": abs(dist.mean - actual),
                    "pinball": pinball_loss(
                        forecast=quantile, actual=actual, quantile=service_level
                    ),
                    "fill": fill_rate(forecast_quantile=quantile, actual=actual),
                    "overstock": overstock_units(
                        forecast_quantile=quantile, actual=actual
                    ),
                    "segment": segment,
                }
            )
    return {"by_estimator": {name: _aggregate(group) for name, group in rows.items()}}


def _aggregate(group: Sequence[dict[str, float]]) -> dict[str, float]:
    if not group:
        return {}
    actual_total = sum(r["actual"] for r in group) or 1.0
    death = [r for r in group if r["segment"] == "death"]
    spike = [r for r in group if r["segment"] == "spike"]
    return {
        "count": len(group),
        "wape": round(sum(r["abs_error"] for r in group) / actual_total, 4),
        "mae": round(sum(r["abs_error"] for r in group) / len(group), 4),
        "pinball": round(sum(r["pinball"] for r in group) / len(group), 4),
        "death_fill_rate": round(
            sum(r["fill"] for r in death) / len(death), 4
        ) if death else None,
        "death_overstock": round(
            sum(r["overstock"] for r in death) / len(death), 4
        ) if death else None,
        "spike_fill_rate": round(
            sum(r["fill"] for r in spike) / len(spike), 4
        ) if spike else None,
    }
