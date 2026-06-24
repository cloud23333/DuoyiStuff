# tests/test_forecast_decision.py
from __future__ import annotations

import pytest

from shipment_planner.forecast_decision import (
    SERVICE_LEVEL_FLOOR,
    SL_BASE,
    SL_DEATH,
    SL_HOT,
    critical_ratio_from_aversion,
    resolve_service_level,
)


def test_death_uses_lowest_service_level():
    assert resolve_service_level(
        recent_drop=True, sustained_rise=False
    ) == pytest.approx(SL_DEATH)


def test_sustained_rise_uses_higher_tier():
    assert resolve_service_level(
        recent_drop=False, sustained_rise=True
    ) == pytest.approx(SL_HOT)


def test_default_is_conservative_base():
    assert resolve_service_level(
        recent_drop=False, sustained_rise=False
    ) == pytest.approx(SL_BASE)


def test_base_service_level_is_conservative_default():
    assert SL_BASE == pytest.approx(0.45)


def test_offset_is_clamped():
    assert resolve_service_level(
        recent_drop=False, sustained_rise=False, offset=0.9
    ) == pytest.approx(0.95)
    assert resolve_service_level(
        recent_drop=True, sustained_rise=False, offset=-0.9
    ) == pytest.approx(SERVICE_LEVEL_FLOOR)


def test_critical_ratio_from_aversion():
    assert critical_ratio_from_aversion(11 / 9) == pytest.approx(0.45)
    assert critical_ratio_from_aversion(3.0) == pytest.approx(0.25)
