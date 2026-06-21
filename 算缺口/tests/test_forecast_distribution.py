from __future__ import annotations

import pytest

from shipment_planner.forecast_distribution import (
    DemandDistribution,
    horizon_quantile,
    negbin_ewma_distribution,
)


def test_poisson_quantile_median_near_mean():
    # var == mean → Poisson path; median of Poisson(10) is 10
    assert horizon_quantile(mean=10.0, variance=10.0, probability=0.5) == pytest.approx(10.0)


def test_quantile_is_monotonic_in_probability():
    lows = horizon_quantile(mean=10.0, variance=10.0, probability=0.5)
    highs = horizon_quantile(mean=10.0, variance=10.0, probability=0.95)
    assert highs >= lows


def test_overdispersed_widens_upper_quantile():
    poisson = horizon_quantile(mean=10.0, variance=10.0, probability=0.95)
    negbin = horizon_quantile(mean=10.0, variance=30.0, probability=0.95)
    assert negbin > poisson


def test_zero_mean_returns_zero():
    assert horizon_quantile(mean=0.0, variance=0.0, probability=0.95) == 0.0


def test_distribution_quantile_delegates():
    dist = DemandDistribution(horizon=7, mean=10.0, variance=10.0)
    assert dist.quantile(0.5) == pytest.approx(10.0)


def test_negbin_level_ignores_recent_spike():
    baseline = [5, 5, 6, 5, 4, 5, 6, 5, 4, 5]
    spiked = baseline + [40, 38]
    base_daily = negbin_ewma_distribution(baseline, horizon=7).mean / 7
    spiked_daily = negbin_ewma_distribution(spiked, horizon=7).mean / 7
    assert spiked_daily < base_daily * 1.8
