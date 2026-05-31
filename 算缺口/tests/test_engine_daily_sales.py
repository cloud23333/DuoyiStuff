from __future__ import annotations

from datetime import datetime

from shipment_planner.engine import build_recommendations
from shipment_planner.models import OrderLine, SalesRecord
from shipment_planner.summary import build_summary


def _order_line(
    *,
    quantity: int = 12,
    skc: str = "1064826604",
    skuid: str = "8482832088",
    order_sku: str = "SKU-1",
) -> OrderLine:
    return OrderLine(
        row_number=2,
        internal_order_id="order-1",
        skc=skc,
        skuid=skuid,
        product_code=order_sku,
        order_sku=order_sku,
        status="待发货",
        order_time=datetime(2026, 4, 21, 10, 0),
        quantity=quantity,
    )


def _sales_record(
    *,
    stock_in_warehouse: float = 0,
    pending_receive: float = 0,
    pending_ship: float = 0,
    is_hot_style: bool = False,
    skc: str = "1064826604",
    skuid: str = "8482832088",
    system_sku: str = "SKU-1",
    stocking_days: float = 7,
) -> SalesRecord:
    return SalesRecord(
        row_number=3,
        skc=skc,
        skuid=skuid,
        system_sku=system_sku,
        is_hot_style=is_hot_style,
        stocking_days=stocking_days,
        stock_in_warehouse=stock_in_warehouse,
        pending_receive=pending_receive,
        pending_ship=pending_ship,
    )


def test_missing_daily_sales_row_warns_and_uses_base_stock():
    order_lines = [
        _order_line()
    ]
    sales_records = [
        _sales_record()
    ]

    recommendations, quality_rows, summary = build_recommendations(
        order_lines=order_lines,
        sales_records=sales_records,
        daily_sales_by_key={},
    )

    assert recommendations[0]["forecast_daily_sales"] == 0
    assert recommendations[0]["forecast_stocking_period_sales"] == 0
    assert recommendations[0]["gap"] == 2
    assert recommendations[0]["base_stock_qty"] == 2
    assert recommendations[0]["base_stock_gap"] == 2
    assert recommendations[0]["base_stock_triggered_warning"] == "yes"
    assert recommendations[0]["recommended_ship"] == 2
    assert recommendations[0]["min_order_ship_qty_exempt_applied_warning"] == "yes"
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
    assert summary["base_stock_triggered_skus"] == 1
    assert summary["base_stock_triggered_lines"] == 1


def test_all_zero_sales_gets_default_base_stock_recommendation():
    recommendations, _quality_rows, summary = build_recommendations(
        order_lines=[_order_line(quantity=8)],
        sales_records=[_sales_record()],
        daily_sales_by_key={("1064826604", "8482832088"): (0, 0, 0, 0, 0)},
    )

    row = recommendations[0]
    assert row["demand_profile"] == "无销量款"
    assert row["forecast_daily_sales"] == 0
    assert row["gap"] == 2
    assert row["recommended_ship"] == 2
    assert row["base_stock_triggered_warning"] == "yes"
    assert row["min_order_ship_qty_exempt_applied_warning"] == "yes"
    assert summary["base_stock_triggered_skus"] == 1
    assert summary["base_stock_triggered_lines"] == 1


def test_single_day_big_order_without_small_sales_uses_base_stock_only():
    recommendations, _quality_rows, summary = build_recommendations(
        order_lines=[_order_line(quantity=36)],
        sales_records=[_sales_record(is_hot_style=True, stocking_days=17)],
        daily_sales_by_key={
            ("1064826604", "8482832088"): (
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                47,
                0,
                0,
                0,
                0,
                0,
                0,
            )
        },
    )

    row = recommendations[0]
    assert row["anomaly_flags"] == "单日大单、疑似缺货尾部"
    assert row["forecast_daily_sales"] == 0
    assert row["forecast_stocking_period_sales"] == 0
    assert row["gap"] == 2
    assert row["recommended_ship"] == 2
    assert row["base_stock_triggered_warning"] == "yes"
    assert row["min_order_ship_qty_exempt_applied_warning"] == "yes"
    assert summary["base_stock_triggered_skus"] == 1
    assert summary["base_stock_triggered_lines"] == 1


def test_base_stock_does_not_trigger_when_available_stock_reaches_target():
    recommendations, _quality_rows, summary = build_recommendations(
        order_lines=[_order_line(quantity=8)],
        sales_records=[_sales_record(stock_in_warehouse=1, pending_receive=1)],
        daily_sales_by_key={("1064826604", "8482832088"): (0, 0, 0, 0, 0)},
    )

    row = recommendations[0]
    assert row["gap"] == 0
    assert row["recommended_ship"] == 0
    assert row["base_stock_gap"] == 0
    assert row["base_stock_triggered_warning"] == "no"
    assert summary["base_stock_triggered_skus"] == 0
    assert summary["base_stock_triggered_lines"] == 0


def test_forecast_gap_takes_precedence_when_larger_than_base_stock_gap():
    recommendations, _quality_rows, _summary = build_recommendations(
        order_lines=[_order_line(quantity=50)],
        sales_records=[_sales_record()],
        daily_sales_by_key={("1064826604", "8482832088"): (5, 5, 5, 5, 5, 5, 5)},
    )

    row = recommendations[0]
    assert row["base_stock_gap"] == 2
    assert row["base_stock_triggered_warning"] == "no"
    assert row["gap"] > row["base_stock_gap"]
    assert row["recommended_ship"] > row["base_stock_gap"]


def test_base_stock_can_be_disabled():
    recommendations, _quality_rows, summary = build_recommendations(
        order_lines=[_order_line(quantity=8)],
        sales_records=[_sales_record()],
        daily_sales_by_key={("1064826604", "8482832088"): (0, 0, 0, 0, 0)},
        base_stock_qty=0,
    )

    row = recommendations[0]
    assert row["gap"] == 0
    assert row["recommended_ship"] == 0
    assert row["base_stock_qty"] == 0
    assert row["base_stock_triggered_warning"] == "no"
    assert summary["base_stock_triggered_skus"] == 0


def test_excluded_skc_still_blocks_base_stock_recommendation():
    recommendations, _quality_rows, summary = build_recommendations(
        order_lines=[_order_line(quantity=8)],
        sales_records=[_sales_record()],
        daily_sales_by_key={("1064826604", "8482832088"): (0, 0, 0, 0, 0)},
        exclude_skc={"1064826604"},
    )

    row = recommendations[0]
    assert row["base_stock_triggered_warning"] == "yes"
    assert row["intercept_reason"] == "skc"
    assert row["recommended_ship"] == 0
    assert summary["base_stock_triggered_skus"] == 0
    assert summary["base_stock_triggered_lines"] == 0


def test_summary_includes_forecast_explanation_distributions():
    summary = build_summary(
        order_lines=[],
        sales_records=[],
        recommendations=[
            {
                "decision_reason": "ship_partial",
                "sku_code_check": "exact_match",
                "recommended_ship": 3,
                "forecast_strategy": "regular",
                "demand_profile": "稳定款",
                "anomaly_flags": "无",
                "service_level": 0.75,
                "forecast_model": "tsb",
            },
            {
                "decision_reason": "ship_partial",
                "sku_code_check": "exact_match",
                "recommended_ship": 2,
                "forecast_strategy": "slow_mover",
                "demand_profile": "慢销/间歇款",
                "anomaly_flags": "无",
                "service_level": 0.65,
                "forecast_model": "imapa",
            },
            {
                "decision_reason": "hold",
                "sku_code_check": "exact_match",
                "recommended_ship": 0,
                "forecast_strategy": "conservative",
                "demand_profile": "波动款",
                "anomaly_flags": "孤立爆单、连续暴跌",
                "service_level": 0.70,
                "forecast_model": "current",
            },
            {
                "decision_reason": "hold",
                "sku_code_check": "missing_key",
                "recommended_ship": 0,
                "forecast_strategy": "",
                "demand_profile": "",
                "anomaly_flags": "无",
                "service_level": 0.0,
                "forecast_model": "current",
            },
            {
                "decision_reason": "ship_partial",
                "sku_code_check": "exact_match",
                "recommended_ship": 1,
                "forecast_strategy": "conservative",
                "demand_profile": "稳定款",
                "anomaly_flags": "疑似缺货尾部",
                "service_level": 0.60,
                "forecast_model": "current",
            },
        ],
        quality_rows=[],
        duplicate_keys=set(),
        min_order_ship_qty=10,
        threshold_stats={},
        sku_order_limit_rule_count=0,
        sku_order_limit_capped_lines=0,
        excluded_skc_rule_count=0,
        excluded_skuid_rule_count=0,
        intercepted_order_lines=0,
        intercepted_orders=0,
        small_change_kept_lines=0,
        service_level_offset=0.0,
        base_stock_qty=2,
    )

    assert summary["demand_profile_summary"] == "稳定款 2，波动款 1，慢销/间歇款 1"
    assert summary["anomaly_flag_summary"] == "孤立爆单 1，疑似缺货尾部 1，连续暴跌 1"
    assert summary["service_level_summary"] == "P75 1，P70 1，P65 1，P60 1"
    assert summary["forecast_model_summary"] == "tsb 1，imapa 1，current 2"
    assert summary["base_stock_qty"] == 2
