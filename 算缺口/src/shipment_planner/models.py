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
    sold30: int
    sold7: int
    stocking_days: float
    stock_in_warehouse: float
    pending_receive: float
    pending_ship: float


@dataclass(slots=True)
class KeyState:
    skc: str
    skuid: str
    system_sku: str
    order_qty_total: int
    gap: int
    recommended_qty_total: int
    min_order_ship_qty_exempt_eligible: bool
