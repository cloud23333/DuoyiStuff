"""Forecast benchmark helpers for candidate model comparison.

This module is intentionally separate from the production recommendation path.
It lets us compare alternate demand models on holdout windows before deciding
whether any model should replace the current forecast.
"""

from __future__ import annotations

import math
from collections import defaultdict
from collections.abc import Sequence
from dataclasses import dataclass

from .engine import compute_forecast_metrics
from .forecast_models import (
    MODEL_ADIDA,
    MODEL_CROSTON_SBA,
    MODEL_CURRENT,
    MODEL_IMAPA,
    MODEL_MEAN,
    MODEL_TSB,
    forecast_candidate_model,
)
from .models import SalesRecord

DEFAULT_BENCHMARK_MODELS = (
    MODEL_CURRENT,
    MODEL_MEAN,
    MODEL_CROSTON_SBA,
    MODEL_TSB,
    MODEL_ADIDA,
    MODEL_IMAPA,
)


@dataclass(frozen=True)
class BenchmarkRow:
    skc: str
    skuid: str
    model: str
    train_days: int
    holdout_days: int
    forecast_daily_sales: float
    forecast_holdout_total: float
    actual_holdout_total: int
    abs_error: float
    signed_error: float


def run_forecast_benchmark(
    *,
    sales_records: Sequence[SalesRecord],
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]],
    holdout_days: int,
    min_train_days: int,
    models: Sequence[str] = DEFAULT_BENCHMARK_MODELS,
    apply_stockout_mask: bool = True,
) -> tuple[list[BenchmarkRow], int, dict[str, object]]:
    sales_by_key = {(record.skc, record.skuid): record for record in sales_records}
    rows: list[BenchmarkRow] = []
    skipped = 0

    for key, daily in daily_sales_by_key.items():
        if len(daily) < holdout_days + min_train_days:
            skipped += 1
            continue

        train = daily[:-holdout_days]
        test = daily[-holdout_days:]
        actual_total = int(sum(test))
        sales = sales_by_key.get(key)
        is_hot_style = getattr(sales, "is_hot_style", False) if sales is not None else False
        stock_in_warehouse = (
            getattr(sales, "stock_in_warehouse", 0.0) if sales is not None else 0.0
        )

        for model in models:
            forecast_daily_sales = forecast_with_model(
                model,
                train,
                holdout_days=holdout_days,
                is_hot_style=is_hot_style,
                stock_in_warehouse=stock_in_warehouse,
                apply_stockout_mask=apply_stockout_mask,
            )
            forecast_total = forecast_daily_sales * holdout_days
            signed_error = forecast_total - actual_total
            rows.append(
                BenchmarkRow(
                    skc=key[0],
                    skuid=key[1],
                    model=model,
                    train_days=len(train),
                    holdout_days=holdout_days,
                    forecast_daily_sales=_round(forecast_daily_sales),
                    forecast_holdout_total=_round(forecast_total),
                    actual_holdout_total=actual_total,
                    abs_error=_round(abs(signed_error)),
                    signed_error=_round(signed_error),
                )
            )

    summary = _build_benchmark_summary(
        rows,
        holdout_days=holdout_days,
        skipped_insufficient_history=skipped,
    )
    return rows, skipped, summary


def forecast_with_model(
    model: str,
    daily_sales: Sequence[int],
    *,
    holdout_days: int,
    is_hot_style: bool = False,
    stock_in_warehouse: float = 0.0,
    apply_stockout_mask: bool = True,
) -> float:
    values = tuple(max(0, int(value)) for value in daily_sales)
    if model == MODEL_CURRENT:
        metrics = compute_forecast_metrics(
            daily_sales=values,
            stocking_days=holdout_days,
            is_hot_style=is_hot_style,
            stock_in_warehouse=stock_in_warehouse,
            apply_stockout_mask=apply_stockout_mask,
        )
        return metrics.forecast_daily_sales
    return forecast_candidate_model(model, values, horizon_days=holdout_days)


def _build_benchmark_summary(
    rows: Sequence[BenchmarkRow],
    *,
    holdout_days: int,
    skipped_insufficient_history: int,
) -> dict[str, object]:
    if not rows:
        return {
            "evaluated_skus": 0,
            "model_rows": 0,
            "skipped_insufficient_history": skipped_insufficient_history,
            "holdout_days": holdout_days,
            "by_model": {},
        }

    by_model_rows: dict[str, list[BenchmarkRow]] = defaultdict(list)
    sku_keys: set[tuple[str, str]] = set()
    for row in rows:
        by_model_rows[row.model].append(row)
        sku_keys.add((row.skc, row.skuid))

    by_model = {
        model: _model_summary(model_rows)
        for model, model_rows in sorted(by_model_rows.items())
    }
    return {
        "evaluated_skus": len(sku_keys),
        "model_rows": len(rows),
        "skipped_insufficient_history": skipped_insufficient_history,
        "holdout_days": holdout_days,
        "best_model_by_wape": _best_model_by_metric(by_model, "wape"),
        "best_model_by_mae": _best_model_by_metric(by_model, "mae"),
        "by_model": by_model,
    }


def _model_summary(rows: Sequence[BenchmarkRow]) -> dict[str, object]:
    abs_total = sum(row.abs_error for row in rows)
    actual_total = sum(row.actual_holdout_total for row in rows)
    signed_total = sum(row.signed_error for row in rows)
    count = len(rows)
    return {
        "count": count,
        "mae": _round(abs_total / count),
        "wape": _round(abs_total / actual_total) if actual_total > 0 else None,
        "bias": _round(signed_total / count),
    }


def _best_model_by_metric(
    by_model: dict[str, dict[str, object]],
    metric: str,
) -> str | None:
    ranked = [
        (str(model), stats.get(metric))
        for model, stats in by_model.items()
        if isinstance(stats.get(metric), int | float)
    ]
    if not ranked:
        return None
    return min(ranked, key=lambda item: float(item[1]))[0]


def _round(value: float) -> float:
    if math.isnan(value) or math.isinf(value):
        return value
    return round(value, 4)
