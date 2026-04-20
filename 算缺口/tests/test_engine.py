from __future__ import annotations

from datetime import datetime, timedelta

from shipment_planner.engine import build_recommendations
from shipment_planner.models import OrderLine, SalesRecord


def _order_line(
    *,
    row_number: int = 2,
    internal_order_id: str = "order-1",
    quantity: int = 20,
    order_sku: str = "SKU-1",
) -> OrderLine:
    return OrderLine(
        row_number=row_number,
        internal_order_id=internal_order_id,
        skc="skc-1",
        skuid="sku-id-1",
        product_code=order_sku,
        order_sku=order_sku,
        status="缺货",
        order_time=datetime(2026, 1, 1) + timedelta(minutes=row_number),
        quantity=quantity,
    )


def _sales_record(
    *,
    sold30: int = 0,
    sold7: int = 70,
    stock_in_warehouse: float = 0.0,
    pending_receive: float = 0.0,
    pending_ship: float = 0.0,
    system_sku: str = "SKU-1",
) -> SalesRecord:
    return SalesRecord(
        row_number=2,
        skc="skc-1",
        skuid="sku-id-1",
        system_sku=system_sku,
        is_hot_style=False,
        sold30=sold30,
        sold7=sold7,
        stocking_days=10.0,
        stock_in_warehouse=stock_in_warehouse,
        pending_receive=pending_receive,
        pending_ship=pending_ship,
    )


def test_zero_sold7_stockout_cap_is_exempt_from_min_order_threshold() -> None:
    recommendations, quality_rows, summary = build_recommendations(
        [_order_line(quantity=20)],
        [_sales_record(sold30=30, sold7=0)],
        min_order_ship_qty=10,
        zero_sold7_with_sold30_stockout_max_qty=5,
    )

    assert quality_rows == []
    assert recommendations[0]["recommended_ship"] == 5
    assert recommendations[0]["min_order_ship_qty_exempt_applied"] is True
    assert summary["low_qty_orders_exempted"] == 1


def test_sku_order_limit_caps_total_for_same_order_and_sku() -> None:
    recommendations, _, summary = build_recommendations(
        [
            _order_line(row_number=2, quantity=5, internal_order_id="order-1"),
            _order_line(row_number=3, quantity=5, internal_order_id="order-1"),
        ],
        [_sales_record()],
        min_order_ship_qty=0,
        sku_order_max_qty={"sku-1": 6},
    )

    assert sum(int(row["recommended_ship"]) for row in recommendations) == 6
    assert summary["sku_order_limit_capped_lines"] == 1
