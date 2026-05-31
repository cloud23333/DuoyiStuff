# tests/test_forecast_estimators.py
from __future__ import annotations

import pytest

from shipment_planner.forecast_distribution import (
    hurdle_distribution,
    negbin_ewma_distribution,
    poisson_ewma_distribution,
)


def test_negbin_one_off_spike_raises_variance_not_mean():
    flat = negbin_ewma_distribution([5, 5, 5, 5, 5, 5, 5], horizon=7)
    spiked = negbin_ewma_distribution([5, 5, 5, 40, 5, 5, 5], horizon=7)
    # one-off spike is cleaned out of the level, so the mean barely moves...
    assert spiked.mean == pytest.approx(flat.mean, rel=0.25)
    # ...but it inflates dispersion, widening the upper quantile
    assert spiked.variance > flat.variance
    assert spiked.quantile(0.95) > flat.quantile(0.95)


def test_negbin_recent_drop_collapses_mean_fast():
    dist = negbin_ewma_distribution(
        [10, 10, 10, 10, 10, 1, 1, 1], horizon=7, recent_drop=True
    )
    assert dist.mean < 5 * 7  # level dragged down to the recent ~1/day, not the old 10/day


def test_sustained_rise_lifts_mean():
    flat = negbin_ewma_distribution([3, 3, 3, 3, 3, 3, 3], horizon=7)
    rising = negbin_ewma_distribution([3, 3, 3, 5, 7, 9, 11], horizon=7)
    assert rising.mean > flat.mean


def test_poisson_distribution_has_variance_equal_mean():
    dist = poisson_ewma_distribution([4, 5, 4, 5, 4, 5], horizon=7)
    assert dist.variance == pytest.approx(dist.mean)


def test_hurdle_handles_intermittent_without_zero():
    dist = hurdle_distribution([0, 0, 3, 0, 0, 4, 0, 0, 2, 0], horizon=7)
    assert dist.mean > 0
    assert dist.quantile(0.55) >= 0


def test_all_zero_series_returns_zero_distribution():
    assert negbin_ewma_distribution([0, 0, 0, 0], horizon=7).mean == 0.0
    assert hurdle_distribution([0, 0, 0, 0], horizon=7).mean == 0.0
