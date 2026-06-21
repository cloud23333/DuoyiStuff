from __future__ import annotations

import math

from .engine import (
    FORECAST_STRATEGY_AGGRESSIVE,
    FORECAST_STRATEGY_CONSERVATIVE,
)


def build_forecast_curve(
    strategy: str,
    forecast_daily_sales: float,
    stocking_days: float,
) -> tuple[float, ...]:
    """Per-day forecast curve over the stocking period.

    The curve sums to ``forecast_daily_sales * stocking_days`` so it stays
    consistent with the tabular "预测备货期销量" figure (display only).
    Shape depends on strategy per 2C:
      - 正常: flat
      - 保守: monotonically decaying (1.3x → 0.7x of daily avg)
      - 激进: gentle ramp up to plateau (0.7x → 1.3x over first half)
    """
    day_count = max(0, math.ceil(stocking_days))
    if day_count == 0:
        return ()

    total = forecast_daily_sales * stocking_days
    if total <= 0:
        return tuple(0.0 for _ in range(day_count))

    weights = _strategy_weights(strategy, day_count)
    weight_sum = sum(weights)
    if weight_sum <= 0:
        return tuple(0.0 for _ in range(day_count))

    scale = total / weight_sum
    return tuple(w * scale for w in weights)


def _strategy_weights(strategy: str, day_count: int) -> list[float]:
    if day_count <= 1:
        return [1.0] * day_count

    if strategy == FORECAST_STRATEGY_CONSERVATIVE:
        step = 0.6 / (day_count - 1)
        return [1.3 - step * i for i in range(day_count)]

    if strategy == FORECAST_STRATEGY_AGGRESSIVE:
        half = (day_count - 1) / 2.0
        if half <= 0:
            return [1.0] * day_count
        return [0.7 + 0.6 * min(1.0, i / half) for i in range(day_count)]

    return [1.0] * day_count
