# src/shipment_planner/forecast_decision.py
from __future__ import annotations

import math

from .forecast_distribution import DemandDistribution

SERVICE_LEVEL_FLOOR = 0.50
SERVICE_LEVEL_CEIL = 0.95
SL_DEATH = 0.50
SL_BASE = 0.55
SL_HOT = 0.70


def resolve_service_level(
    *,
    is_hot_style: bool,
    recent_drop: bool,
    sustained_rise: bool,
    offset: float = 0.0,
) -> float:
    """Conservative-by-default service level (near median).

    Overstock is the more expensive error, so the base sits at the median and
    only genuine hot/rising SKUs get a buffer; a confirmed collapse pins to the
    floor so dying SKUs stop being restocked.
    """
    if recent_drop:
        level = SL_DEATH
    elif is_hot_style or sustained_rise:
        level = SL_HOT
    else:
        level = SL_BASE
    return min(SERVICE_LEVEL_CEIL, max(SERVICE_LEVEL_FLOOR, level + offset))


def horizon_quantile_units(
    distribution: DemandDistribution, service_level: float
) -> float:
    """Chosen-quantile demand over the whole horizon (drives the gap)."""
    return distribution.quantile(service_level)


def gap_from_distribution(
    distribution: DemandDistribution,
    available_stock: float,
    service_level: float,
) -> int:
    quantile = distribution.quantile(service_level)
    return math.ceil(max(0.0, quantile - available_stock))
