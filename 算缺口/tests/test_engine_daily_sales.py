from __future__ import annotations

from datetime import datetime

from shipment_planner.engine import build_recommendations
from shipment_planner.models import OrderLine, SalesRecord
from shipment_planner.summary import build_summary


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
        global_gap_multiplier=1.0,
    )

    assert summary["demand_profile_summary"] == "稳定款 1，波动款 1，慢销/间歇款 1"
    assert summary["anomaly_flag_summary"] == "孤立爆单 1，连续暴跌 1"
    assert summary["service_level_summary"] == "P75 1，P70 1，P65 1"
    assert summary["forecast_model_summary"] == "tsb 1，imapa 1，current 1"
