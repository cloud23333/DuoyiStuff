# tests/test_forecast_level.py
from __future__ import annotations

import pytest

from shipment_planner.forecast_level import (
    clean_isolated_spikes,
    has_recent_drop,
    has_sustained_rise,
    recent_mean,
    weighted_mean,
    weighted_variance,
)


def test_weighted_mean_favors_recent_values():
    rising = weighted_mean([1, 1, 1, 9], half_life=2.0)
    flat = sum([1, 1, 1, 9]) / 4
    assert rising > flat  # recent 9 pulls the recency-weighted mean above the flat mean


def test_weighted_variance_nonnegative_and_zero_for_constant():
    assert weighted_variance([5, 5, 5, 5], half_life=5.0) == pytest.approx(0.0)
    assert weighted_variance([0, 0, 9, 0], half_life=5.0) > 0.0


def test_recent_mean_uses_last_n():
    assert recent_mean([10, 10, 0, 0, 0], days=3) == pytest.approx(0.0)


def test_clean_isolated_spikes_flattens_one_off():
    cleaned, changed = clean_isolated_spikes([4, 5, 5, 42, 4, 5, 4])
    assert changed is True
    assert cleaned[3] < 42


def test_clean_isolated_spikes_cleans_last_day():
    cleaned, changed = clean_isolated_spikes([4, 5, 5, 4, 5, 4, 42])
    assert changed is True
    assert cleaned[-1] < 42


def test_clean_isolated_spikes_cleans_first_day():
    cleaned, changed = clean_isolated_spikes([42, 5, 5, 4, 5, 4, 5])
    assert changed is True
    assert cleaned[0] < 42


def test_has_sustained_rise_false_for_trailing_one_off_spike():
    assert has_sustained_rise([4, 5, 5, 4, 5, 4, 5, 4, 5, 42]) is False


def test_has_recent_drop_detects_collapse():
    assert has_recent_drop([10, 11, 10, 12, 10, 2, 1, 2, 1, 2], stock_in_warehouse=20) is True


def test_has_sustained_rise_detects_trend():
    assert has_sustained_rise([2, 2, 2, 2, 8, 9, 10]) is True
    assert has_sustained_rise([4, 5, 5, 42, 4, 5, 4]) is False  # one-off spike, not sustained
