# tests/test_eval_metrics.py
from __future__ import annotations

import pytest

from shipment_planner.eval_forecast import (
    classify_holdout_segment,
    fill_rate,
    overstock_units,
    pinball_loss,
)


def test_pinball_penalizes_underforecast_more_at_high_quantile():
    # actual above forecast at q=0.9 → heavy underage penalty
    under = pinball_loss(forecast=5.0, actual=10.0, quantile=0.9)
    over = pinball_loss(forecast=10.0, actual=5.0, quantile=0.9)
    assert under > over


def test_fill_rate_is_capped_at_one():
    assert fill_rate(forecast_quantile=20.0, actual=10.0) == pytest.approx(1.0)
    assert fill_rate(forecast_quantile=5.0, actual=10.0) == pytest.approx(0.5)


def test_overstock_units_counts_excess_only():
    assert overstock_units(forecast_quantile=15.0, actual=10.0) == pytest.approx(5.0)
    assert overstock_units(forecast_quantile=8.0, actual=10.0) == pytest.approx(0.0)


def test_segment_classification():
    assert classify_holdout_segment(train_mean=1.0, actual_total=70.0, holdout_days=7) == "spike"
    assert classify_holdout_segment(train_mean=10.0, actual_total=7.0, holdout_days=7) == "death"
    assert classify_holdout_segment(train_mean=5.0, actual_total=35.0, holdout_days=7) == "normal"
