from __future__ import annotations

import pytest

from shipment_planner.engine import (
    FORECAST_STRATEGY_CONSERVATIVE,
    FORECAST_STRATEGY_SLOW_MOVER,
    compute_forecast_metrics,
)
from shipment_planner.reports import RECOMMENDATION_FIELDS


def test_isolated_spike_is_marked_and_downweighted():
    metrics = compute_forecast_metrics(
        daily_sales=(4, 5, 5, 42, 4, 5, 4),
        stocking_days=7,
        is_hot_style=False,
        stock_in_warehouse=10,
    )

    assert metrics.demand_profile == "波动款"
    assert "孤立爆单" in metrics.anomaly_flags
    assert metrics.service_level == pytest.approx(0.70)
    assert metrics.effective_daily_sales < 10
    assert metrics.forecast_daily_sales < 10


def test_small_intermittent_sales_are_not_marked_as_isolated_spike():
    metrics = compute_forecast_metrics(
        daily_sales=(0, 2, 0, 8, 0, 2, 0),
        stocking_days=7,
        is_hot_style=False,
        stock_in_warehouse=10,
    )

    assert "孤立爆单" not in metrics.anomaly_flags


def test_recent_sales_drop_with_stock_is_conservative_but_not_zero():
    metrics = compute_forecast_metrics(
        daily_sales=(10, 11, 10, 12, 10, 2, 1, 2, 1, 2),
        stocking_days=7,
        is_hot_style=False,
        stock_in_warehouse=20,
    )

    assert metrics.strategy == FORECAST_STRATEGY_CONSERVATIVE
    assert "连续暴跌" in metrics.anomaly_flags
    assert metrics.service_level == pytest.approx(0.60)
    assert 0 < metrics.forecast_daily_sales < 10


def test_three_day_gap_does_not_trigger_recent_drop_when_five_day_window_is_healthy():
    metrics = compute_forecast_metrics(
        daily_sales=(10, 11, 10, 9, 10, 12, 12, 0, 0, 0),
        stocking_days=7,
        is_hot_style=False,
        stock_in_warehouse=20,
    )

    assert "连续暴跌" not in metrics.anomaly_flags


def test_recent_sales_collapse_to_zero_with_stock_forecasts_zero():
    metrics = compute_forecast_metrics(
        daily_sales=(10, 11, 10, 12, 10, 9, 0, 0, 0, 0, 0),
        stocking_days=7,
        is_hot_style=False,
        stock_in_warehouse=20,
    )

    assert metrics.strategy == FORECAST_STRATEGY_CONSERVATIVE
    assert "连续暴跌" in metrics.anomaly_flags
    assert metrics.forecast_daily_sales == 0
    assert metrics.forecast_stocking_period_sales == 0


def test_intermittent_sales_use_slow_mover_profile_without_collapsing():
    metrics = compute_forecast_metrics(
        daily_sales=(0, 0, 3, 0, 0, 0, 4, 0, 0, 2, 0, 0),
        stocking_days=7,
        is_hot_style=False,
        stock_in_warehouse=5,
    )

    assert metrics.strategy == FORECAST_STRATEGY_SLOW_MOVER
    assert metrics.demand_profile == "慢销/间歇款"
    assert metrics.forecast_model == "imapa"
    assert metrics.service_level == pytest.approx(0.65)
    assert metrics.effective_daily_sales > 0
    assert metrics.forecast_daily_sales > 0


def test_single_big_order_with_small_sales_uses_small_sale_baseline():
    metrics = compute_forecast_metrics(
        daily_sales=(0, 1, 0, 2, 0, 50, 0, 1, 0, 2, 0),
        stocking_days=7,
        is_hot_style=False,
        stock_in_warehouse=10,
    )

    assert "单日大单" in metrics.anomaly_flags
    assert metrics.demand_profile == "慢销/间歇款"
    assert metrics.forecast_daily_sales > 0
    assert metrics.forecast_daily_sales < 1


def test_all_zero_sales_are_classified_as_no_sales():
    metrics = compute_forecast_metrics(
        daily_sales=(0, 0, 0, 0, 0),
        stocking_days=7,
        is_hot_style=False,
        stock_in_warehouse=5,
    )

    assert metrics.demand_profile == "无销量款"
    assert metrics.anomaly_flags == "无"
    assert metrics.service_level == pytest.approx(0.50)
    assert metrics.effective_daily_sales == 0
    assert metrics.forecast_daily_sales == 0


def test_stockout_tail_zeros_are_not_treated_as_sales_drop():
    metrics = compute_forecast_metrics(
        daily_sales=(8, 9, 8, 7, 0, 0, 0),
        stocking_days=7,
        is_hot_style=False,
        stock_in_warehouse=0,
    )

    assert "连续暴跌" not in metrics.anomaly_flags
    assert metrics.demand_profile == "稳定款"
    assert metrics.forecast_daily_sales > 0


def test_stable_hot_style_uses_aggressive_service_level():
    metrics = compute_forecast_metrics(
        daily_sales=(8, 9, 8, 9, 8, 9, 8),
        stocking_days=7,
        is_hot_style=True,
        stock_in_warehouse=20,
    )

    assert metrics.demand_profile == "稳定款"
    assert metrics.forecast_model == "tsb"
    assert metrics.service_level == pytest.approx(0.85)
    assert metrics.forecast_daily_sales >= metrics.effective_daily_sales


def test_stable_non_hot_sales_use_tsb_model():
    metrics = compute_forecast_metrics(
        daily_sales=(8, 9, 8, 9, 8, 9, 8),
        stocking_days=7,
        is_hot_style=False,
        stock_in_warehouse=20,
    )

    assert metrics.demand_profile == "稳定款"
    assert metrics.forecast_model == "tsb"
    assert metrics.forecast_daily_sales > 0


def test_recommendation_report_includes_forecast_explanation_columns():
    field_labels = dict(RECOMMENDATION_FIELDS)

    assert field_labels["demand_profile"] == "需求类型"
    assert field_labels["anomaly_flags"] == "异常标记"
    assert field_labels["service_level"] == "服务水平"
    assert field_labels["forecast_model"] == "预测模型"
    assert field_labels["effective_daily_sales"] == "异常调整后日均销量"
