from __future__ import annotations

from datetime import datetime

from shipment_planner.engine import build_recommendations
from shipment_planner.models import OrderLine, SalesRecord


def _order(qty=20):
    return OrderLine(
        row_number=1, internal_order_id="o1", skc="skc-1", skuid="sku-1",
        product_code="ZX1", order_sku="ZX1", status="待发货",
        order_time=datetime(2026, 4, 20, 9, 0, 0), quantity=qty,
    )


def _sales():
    return SalesRecord(
        row_number=2, skc="skc-1", skuid="sku-1", system_sku="ZX1",
        is_hot_style=False, stocking_days=7, stock_in_warehouse=0,
        pending_receive=0, pending_ship=0,
    )


def test_build_recommendations_accepts_service_level_offset_not_multiplier():
    recs, _, summary = build_recommendations(
        [_order()], [_sales()],
        daily_sales_by_key={("skc-1", "sku-1"): (5, 5, 5, 5, 5, 5, 5, 5, 5, 5)},
        service_level_offset=0.0,
    )
    assert "service_level_offset" in summary
    assert "global_gap_multiplier" not in summary
