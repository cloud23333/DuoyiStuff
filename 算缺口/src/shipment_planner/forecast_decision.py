# src/shipment_planner/forecast_decision.py
from __future__ import annotations

SERVICE_LEVEL_FLOOR = 0.40
SERVICE_LEVEL_CEIL = 0.95
SL_DEATH = 0.50
SL_BASE = 0.45
SL_HOT = 0.50


def resolve_service_level(
    *,
    recent_drop: bool,
    sustained_rise: bool,
    offset: float = 0.0,
) -> float:
    """Conservative-by-default service level (near median).

    Overstock is the more expensive error, so the base sits at the median and
    only a genuine data-driven sustained rise gets a buffer; a confirmed collapse
    pins to the floor so dying SKUs stop being restocked.
    """
    if recent_drop:
        level = SL_DEATH
    elif sustained_rise:
        level = SL_HOT
    else:
        level = SL_BASE
    return min(SERVICE_LEVEL_CEIL, max(SERVICE_LEVEL_FLOOR, level + offset))
