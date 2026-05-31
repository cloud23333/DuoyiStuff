from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(slots=True)
class OrderLine:
    row_number: int
    internal_order_id: str
    skc: str
    skuid: str
    product_code: str
    order_sku: str
    status: str
    order_time: datetime
    quantity: int


@dataclass(slots=True)
class SalesRecord:
    row_number: int
    skc: str
    skuid: str
    system_sku: str
    is_hot_style: bool
    stocking_days: float
    stock_in_warehouse: float
    pending_receive: float
    pending_ship: float


@dataclass(slots=True)
class ForecastMetrics:
    strategy: str
    forecast_daily_sales: float
    forecast_stocking_period_sales: float
    demand_profile: str = ""
    anomaly_flags: str = "无"
    service_level: float = 0.0
    forecast_model: str = "current"
    effective_daily_sales: float = 0.0
    predictive_mean: float = 0.0
    dispersion: float = 0.0
    chosen_quantile: float = 0.0


@dataclass(slots=True)
class KeyState:
    skc: str
    skuid: str
    system_sku: str
    order_qty_total: int
    gap: int
    recommended_qty_total: int
    forecast_metrics: ForecastMetrics
    min_order_ship_qty_exempt_eligible: bool
    base_stock_qty: int = 0
    base_stock_gap: int = 0
    base_stock_triggered: bool = False
