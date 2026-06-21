# tests/test_forecast_level.py
from __future__ import annotations

import pytest

from shipment_planner.forecast_level import (
    clean_isolated_spikes,
    has_recent_drop,
    has_sustained_rise,
    recent_mean,
    robust_level,
    trimmed_mean,
    weighted_mean,
    weighted_variance,
    winsorize,
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


def test_trimmed_mean_ignores_extremes():
    assert trimmed_mean([0, 5, 5, 5, 100], trim_fraction=0.2) == pytest.approx(5.0)


def test_winsorize_caps_top_values():
    capped = winsorize([1, 2, 3, 4, 100], upper_quantile=0.8)
    assert max(capped) <= 4
    assert capped[:4] == [1, 2, 3, 4]


def test_robust_level_ignores_one_off_spike():
    baseline = [5, 5, 6, 5, 4, 5, 6, 5, 4, 5]
    spiked = baseline + [40, 38]
    base = robust_level(baseline)
    after_spike = robust_level(spiked)
    assert after_spike < base * 1.25


def test_robust_level_follows_genuine_rise_but_capped():
    rising = [10, 12, 15, 18, 22, 26, 30, 34, 40, 46]
    level = robust_level(rising)
    assert level > 15
    assert level < max(rising)


def test_robust_estimators_handle_empty_and_degenerate_input():
    assert winsorize([]) == []
    assert trimmed_mean([]) == pytest.approx(0.0)
    assert robust_level([]) == pytest.approx(0.0)
    assert robust_level([7]) == pytest.approx(7.0)
    assert robust_level([0, 0, 0, 0]) == pytest.approx(0.0)
    assert robust_level([-3, -2, -5]) == pytest.approx(0.0)
