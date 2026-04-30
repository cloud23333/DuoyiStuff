from __future__ import annotations

import statistics
from collections.abc import Sequence

MODEL_CURRENT = "current"
MODEL_MEAN = "mean"
MODEL_CROSTON_SBA = "croston_sba"
MODEL_TSB = "tsb"
MODEL_ADIDA = "adida"
MODEL_IMAPA = "imapa"

_DEFAULT_ALPHA = 0.2


def forecast_candidate_model(
    model: str,
    daily_sales: Sequence[float],
    *,
    horizon_days: int,
) -> float:
    values = tuple(max(0, int(value)) for value in daily_sales)
    if model == MODEL_MEAN:
        return _mean_forecast(values)
    if model == MODEL_CROSTON_SBA:
        return _croston_sba_forecast(values)
    if model == MODEL_TSB:
        return _tsb_forecast(values)
    if model == MODEL_ADIDA:
        return _adida_forecast(values, aggregation_period=horizon_days)
    if model == MODEL_IMAPA:
        return _imapa_forecast(values, max_aggregation_period=horizon_days)
    raise ValueError(f"Unknown candidate forecast model: {model}")


def _mean_forecast(values: Sequence[int]) -> float:
    return statistics.fmean(values) if values else 0.0


def _croston_sba_forecast(values: Sequence[int], alpha: float = _DEFAULT_ALPHA) -> float:
    demand_estimate: float | None = None
    interval_estimate: float | None = None
    interval = 0

    for value in values:
        interval += 1
        if value <= 0:
            continue
        if demand_estimate is None or interval_estimate is None:
            demand_estimate = float(value)
            interval_estimate = float(interval)
        else:
            demand_estimate += alpha * (value - demand_estimate)
            interval_estimate += alpha * (interval - interval_estimate)
        interval = 0

    if demand_estimate is None or interval_estimate is None or interval_estimate <= 0:
        return 0.0
    return max(0.0, (1.0 - (alpha / 2.0)) * demand_estimate / interval_estimate)


def _tsb_forecast(values: Sequence[int], alpha: float = _DEFAULT_ALPHA) -> float:
    if not values:
        return 0.0

    first_nonzero = next((float(value) for value in values if value > 0), 0.0)
    if first_nonzero <= 0:
        return 0.0

    probability = 1.0 if values[0] > 0 else 0.0
    demand_size = first_nonzero
    for value in values:
        occurrence = 1.0 if value > 0 else 0.0
        probability += alpha * (occurrence - probability)
        if value > 0:
            demand_size += alpha * (value - demand_size)
    return max(0.0, probability * demand_size)


def _adida_forecast(values: Sequence[int], aggregation_period: int) -> float:
    if not values:
        return 0.0

    period = min(max(1, aggregation_period), len(values))
    aggregated = [
        sum(values[index : index + period])
        for index in range(0, len(values), period)
        if values[index : index + period]
    ]
    if not aggregated:
        return 0.0
    return max(0.0, statistics.fmean(aggregated) / period)


def _imapa_forecast(values: Sequence[int], max_aggregation_period: int) -> float:
    if not values:
        return 0.0

    max_period = min(max(1, max_aggregation_period), len(values))
    forecasts = [
        _adida_forecast(values, aggregation_period=period)
        for period in range(1, max_period + 1)
    ]
    return statistics.fmean(forecasts) if forecasts else 0.0
