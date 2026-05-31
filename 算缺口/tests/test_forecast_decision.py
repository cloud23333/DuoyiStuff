# tests/test_forecast_decision.py
from __future__ import annotations

import pytest

from shipment_planner.forecast_distribution import DemandDistribution
from shipment_planner.forecast_decision import (
    SL_BASE,
    SL_DEATH,
    SL_HOT,
    gap_from_distribution,
    resolve_service_level,
)


def test_death_uses_lowest_service_level():
    assert resolve_service_level(
        is_hot_style=False, recent_drop=True, sustained_rise=False
    ) == pytest.approx(SL_DEATH)


def test_hot_style_uses_higher_tier():
    assert resolve_service_level(
        is_hot_style=True, recent_drop=False, sustained_rise=False
    ) == pytest.approx(SL_HOT)


def test_default_is_conservative_base():
    assert resolve_service_level(
        is_hot_style=False, recent_drop=False, sustained_rise=False
    ) == pytest.approx(SL_BASE)


def test_offset_is_clamped():
    assert resolve_service_level(
        is_hot_style=False, recent_drop=False, sustained_rise=False, offset=0.9
    ) == pytest.approx(0.95)
    assert resolve_service_level(
        is_hot_style=False, recent_drop=True, sustained_rise=False, offset=-0.9
    ) == pytest.approx(0.50)


def test_gap_subtracts_stock_and_ceils():
    dist = DemandDistribution(horizon=7, mean=10.0, variance=10.0)
    # quantile(0.5)=10 ; minus 4 stock = 6
    assert gap_from_distribution(dist, available_stock=4.0, service_level=0.5) == 6


def test_gap_never_negative():
    dist = DemandDistribution(horizon=7, mean=10.0, variance=10.0)
    assert gap_from_distribution(dist, available_stock=999.0, service_level=0.5) == 0
