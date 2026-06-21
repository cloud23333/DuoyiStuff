from __future__ import annotations

import pytest

from shipment_planner.engine import (
    FORECAST_STRATEGY_AGGRESSIVE,
    FORECAST_STRATEGY_CONSERVATIVE,
    FORECAST_STRATEGY_SLOW_MOVER,
    compute_forecast_metrics,
)
from shipment_planner.forecast_decision import SL_BASE, SL_DEATH, SL_HOT
from shipment_planner.reports import RECOMMENDATION_FIELDS


def test_one_off_spike_does_not_inflate_daily_mean():
    flat = compute_forecast_metrics(
        daily_sales=(5, 5, 5, 5, 5, 5, 5), stocking_days=7,
        stock_in_warehouse=10,
    )
    spiked = compute_forecast_metrics(
        daily_sales=(5, 5, 5, 40, 5, 5, 5), stocking_days=7,
        stock_in_warehouse=10,
    )
    assert "孤立爆单" in spiked.anomaly_flags
    # spike feeds dispersion, not the level → effective daily mean stays near 5
    assert spiked.effective_daily_sales == pytest.approx(flat.effective_daily_sales, rel=0.3)
    assert spiked.dispersion > flat.dispersion


def test_recent_collapse_drives_forecast_down_fast():
    metrics = compute_forecast_metrics(
        daily_sales=(10, 11, 10, 12, 10, 9, 0, 0, 0, 0, 0), stocking_days=7,
        stock_in_warehouse=20,
    )
    assert metrics.strategy == FORECAST_STRATEGY_CONSERVATIVE
    assert "连续暴跌" in metrics.anomaly_flags
    assert metrics.service_level == pytest.approx(SL_DEATH)
    assert metrics.forecast_daily_sales < 3


def test_recent_drop_partial_stays_low_but_positive():
    metrics = compute_forecast_metrics(
        daily_sales=(10, 11, 10, 12, 10, 2, 1, 2, 1, 2), stocking_days=7,
        stock_in_warehouse=20,
    )
    assert "连续暴跌" in metrics.anomaly_flags
    assert metrics.service_level == pytest.approx(SL_DEATH)
    assert 0 < metrics.forecast_daily_sales < 10


def test_intermittent_uses_slow_mover_label():
    metrics = compute_forecast_metrics(
        daily_sales=(0, 0, 3, 0, 0, 0, 4, 0, 0, 2, 0, 0), stocking_days=7,
        stock_in_warehouse=5,
    )
    assert metrics.strategy == FORECAST_STRATEGY_SLOW_MOVER
    assert metrics.demand_profile == "慢销/间歇款"
    assert metrics.forecast_daily_sales > 0


def test_all_zero_is_no_sales():
    metrics = compute_forecast_metrics(
        daily_sales=(0, 0, 0, 0, 0), stocking_days=7,
        stock_in_warehouse=5,
    )
    assert metrics.demand_profile == "无销量款"
    assert metrics.forecast_daily_sales == 0
    assert metrics.forecast_stocking_period_sales == 0


def test_sustained_rise_uses_higher_service_level():
    metrics = compute_forecast_metrics(
        daily_sales=(1, 1, 2, 2, 3, 4, 6, 8, 10), stocking_days=7,
        stock_in_warehouse=20,
    )
    assert metrics.strategy == FORECAST_STRATEGY_AGGRESSIVE
    assert metrics.service_level == pytest.approx(SL_HOT)


def test_stable_uses_base_service_level():
    metrics = compute_forecast_metrics(
        daily_sales=(8, 9, 8, 9, 8, 9, 8), stocking_days=7,
        stock_in_warehouse=20,
    )
    assert metrics.demand_profile == "稳定款"
    assert metrics.service_level == pytest.approx(SL_BASE)
    assert metrics.forecast_daily_sales > 0


def test_stockout_tail_not_treated_as_drop():
    metrics = compute_forecast_metrics(
        daily_sales=(8, 9, 8, 7, 0, 0, 0), stocking_days=7,
        stock_in_warehouse=0,
    )
    assert "连续暴跌" not in metrics.anomaly_flags
    assert "疑似缺货尾部" in metrics.anomaly_flags
    assert metrics.forecast_daily_sales > 0


def test_single_big_order_flagged_and_small_forecast():
    metrics = compute_forecast_metrics(
        daily_sales=(0, 1, 0, 2, 0, 50, 0, 1, 0, 2, 0), stocking_days=7,
        stock_in_warehouse=10,
    )
    assert "单日大单" in metrics.anomaly_flags
    assert metrics.forecast_daily_sales < 1


def test_recommendation_report_includes_forecast_explanation_columns():
    field_labels = dict(RECOMMENDATION_FIELDS)
    assert field_labels["demand_profile"] == "需求类型"
    assert field_labels["anomaly_flags"] == "异常标记"
    assert field_labels["service_level"] == "服务水平"
    assert field_labels["forecast_model"] == "预测模型"
    assert field_labels["effective_daily_sales"] == "异常调整后日均销量"
