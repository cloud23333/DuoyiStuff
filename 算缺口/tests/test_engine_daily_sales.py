from __future__ import annotations

from datetime import datetime

from shipment_planner.engine import build_recommendations
from shipment_planner.models import OrderLine, SalesRecord


def test_missing_daily_sales_row_warns_and_holds_recommendation():
    order_lines = [
        OrderLine(
            row_number=2,
            internal_order_id="order-1",
            skc="1064826604",
            skuid="8482832088",
            product_code="SKU-1",
            order_sku="SKU-1",
            status="待发货",
            order_time=datetime(2026, 4, 21, 10, 0),
            quantity=12,
        )
    ]
    sales_records = [
        SalesRecord(
            row_number=3,
            skc="1064826604",
            skuid="8482832088",
            system_sku="SKU-1",
            is_hot_style=False,
            stocking_days=7,
            stock_in_warehouse=0,
            pending_receive=0,
            pending_ship=0,
        )
    ]

    recommendations, quality_rows, summary = build_recommendations(
        order_lines=order_lines,
        sales_records=sales_records,
        daily_sales_by_key={},
    )

    assert recommendations[0]["forecast_daily_sales"] == 0
    assert recommendations[0]["forecast_stocking_period_sales"] == 0
    assert recommendations[0]["gap"] == 0
    assert recommendations[0]["recommended_ship"] == 0
    assert quality_rows == [
        {
            "type": "missing_daily_sales",
            "row_number": 2,
            "internal_order_id": "order-1",
            "skc": "1064826604",
            "skuid": "8482832088",
            "order_sku": "SKU-1",
            "system_sku": "SKU-1",
            "message": "Missing daily sales data for (SKC, SKUID)",
        }
    ]
    assert summary["quality_issue_rows"] == 1
