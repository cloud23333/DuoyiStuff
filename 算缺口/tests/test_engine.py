from __future__ import annotations

from datetime import datetime, timedelta

import pytest

from shipment_planner.engine import build_recommendations
from shipment_planner.models import OrderLine, SalesRecord
from shipment_planner.reports import RECOMMENDATION_FIELDS


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
    is_hot_style: bool = False,
) -> SalesRecord:
    return SalesRecord(
        row_number=2,
        skc="skc-1",
        skuid="sku-id-1",
        system_sku=system_sku,
        is_hot_style=is_hot_style,
        sold30=sold30,
        sold7=sold7,
        stocking_days=10.0,
        stock_in_warehouse=stock_in_warehouse,
        pending_receive=pending_receive,
        pending_ship=pending_ship,
    )


def _daily_sales(values: tuple[int, ...] = (8, 8, 8, 8, 8, 8, 8)) -> dict[tuple[str, str], tuple[int, ...]]:
    return {("skc-1", "sku-id-1"): values}


def test_build_recommendations_requires_daily_sales_data() -> None:
    with pytest.raises(ValueError, match="daily sales data is required"):
        build_recommendations(
            [_order_line(quantity=20)],
            [_sales_record()],
            min_order_ship_qty=0,
            daily_sales_by_key=None,
        )


def test_missing_daily_sales_sku_errors_without_fallback() -> None:
    with pytest.raises(ValueError, match="Missing daily sales data"):
        build_recommendations(
            [_order_line(quantity=20)],
            [_sales_record()],
            min_order_ship_qty=0,
            daily_sales_by_key={},
        )


def test_empty_daily_sales_sequence_errors_without_fallback() -> None:
    with pytest.raises(ValueError, match="Empty daily sales sequence"):
        build_recommendations(
            [_order_line(quantity=20)],
            [_sales_record()],
            min_order_ship_qty=0,
            daily_sales_by_key=_daily_sales(()),
        )


def test_stable_daily_sales_outputs_normal_strategy() -> None:
    recommendations, quality_rows, summary = build_recommendations(
        [_order_line(quantity=20)],
        [_sales_record(sold30=30, sold7=0)],
        min_order_ship_qty=0,
        daily_sales_by_key=_daily_sales((5, 5, 5, 5, 5, 5, 5)),
    )

    assert quality_rows == []
    assert recommendations[0]["forecast_strategy"] == "正常"
    assert recommendations[0]["forecast_daily_sales"] == 5
    assert recommendations[0]["forecast_stocking_period_sales"] == 50
    assert recommendations[0]["recommended_ship"] == 20
    assert summary["total_recommended_qty"] == 20


def test_recently_rising_daily_sales_outputs_aggressive_strategy() -> None:
    recommendations, _, _ = build_recommendations(
        [_order_line(quantity=100)],
        [_sales_record()],
        min_order_ship_qty=0,
        daily_sales_by_key=_daily_sales((2, 2, 2, 3, 3, 4, 5, 6, 7, 8)),
    )

    assert recommendations[0]["forecast_strategy"] == "激进"
    assert recommendations[0]["forecast_daily_sales"] > 5


def test_recently_falling_daily_sales_outputs_conservative_strategy() -> None:
    recommendations, _, _ = build_recommendations(
        [_order_line(quantity=100)],
        [_sales_record()],
        min_order_ship_qty=0,
        daily_sales_by_key=_daily_sales((10, 10, 10, 9, 8, 4, 3, 2)),
    )

    assert recommendations[0]["forecast_strategy"] == "保守"


def test_daily_sales_forecast_drives_gap_and_recommended_ship() -> None:
    recommendations, _, _ = build_recommendations(
        [_order_line(quantity=100)],
        [_sales_record(stock_in_warehouse=5)],
        min_order_ship_qty=0,
        daily_sales_by_key=_daily_sales((2, 2, 2, 2, 2, 2, 2)),
    )

    row = recommendations[0]
    assert row["forecast_strategy"] == "正常"
    assert row["forecast_daily_sales"] == 2
    assert row["forecast_stocking_period_sales"] == 20
    assert row["gap"] == 15
    assert row["recommended_ship"] == 15


def test_recommendation_report_includes_forecast_columns_near_sales_fields() -> None:
    field_names = [target for _, target in RECOMMENDATION_FIELDS]

    sold7_index = field_names.index("近7日销量")
    assert field_names[sold7_index + 1 : sold7_index + 4] == [
        "预测策略",
        "预测日均销量",
        "预测备货期销量",
    ]


def test_sku_order_limit_caps_total_for_same_order_and_sku() -> None:
    recommendations, _, summary = build_recommendations(
        [
            _order_line(row_number=2, quantity=5, internal_order_id="order-1"),
            _order_line(row_number=3, quantity=5, internal_order_id="order-1"),
        ],
        [_sales_record()],
        min_order_ship_qty=0,
        sku_order_max_qty={"sku-1": 6},
        daily_sales_by_key=_daily_sales(),
    )

    assert sum(int(row["recommended_ship"]) for row in recommendations) == 6
    assert summary["sku_order_limit_capped_lines"] == 1


def test_small_change_keep_can_exceed_gap_but_not_sku_order_limit() -> None:
    recommendations, _, summary = build_recommendations(
        [_order_line(quantity=100)],
        [_sales_record()],
        min_order_ship_qty=0,
        sku_order_max_qty={"sku-1": 90},
        daily_sales_by_key=_daily_sales(),
    )

    row = recommendations[0]
    assert row["gap"] == 80
    assert row["recommended_ship_before_small_change_rule"] == 80
    assert row["recommended_ship"] == 90
    assert row["key_recommended_total"] == 90
    assert summary["sku_order_limit_capped_lines"] == 1
