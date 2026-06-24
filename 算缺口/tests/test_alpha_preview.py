from __future__ import annotations

from datetime import datetime

from shipment_planner.alpha_preview import (
    FALLBACK_SUGGESTED_ALPHA,
    AlphaPoint,
    _curve_from_parsed,
    interpolate_ship_lost,
    suggest_alpha,
)
from shipment_planner.models import OrderLine, SalesRecord


def _point(alpha, ship_units=0, lost_units=0, leftover_units=0):
    return AlphaPoint(
        alpha=alpha,
        ship_units=ship_units,
        lost_units=lost_units,
        leftover_units=leftover_units,
        fill_rate=None,
    )


def test_suggest_alpha_picks_the_knee_where_marginal_crosses_threshold():
    points = (
        _point(1.0, lost_units=0, leftover_units=100),
        _point(0.9, lost_units=10, leftover_units=80),
        _point(0.8, lost_units=20, leftover_units=60),
        _point(0.7, lost_units=80, leftover_units=40),
    )
    suggested = suggest_alpha(points)
    assert suggested == 0.8
    assert 0.7 <= suggested <= 0.95


def test_suggest_alpha_falls_back_when_no_leftover_saved():
    points = (
        _point(1.0, lost_units=0, leftover_units=50),
        _point(0.9, lost_units=5, leftover_units=50),
        _point(0.8, lost_units=12, leftover_units=50),
    )
    assert suggest_alpha(points) == FALLBACK_SUGGESTED_ALPHA


def _order(qty=20):
    return OrderLine(
        row_number=1, internal_order_id="o1", skc="skc-1", skuid="sku-1",
        product_code="ZX1", order_sku="ZX1", status="待发货",
        order_time=datetime(2026, 4, 20, 9, 0, 0), quantity=qty,
    )


def _sales():
    return SalesRecord(
        row_number=2, skc="skc-1", skuid="sku-1", system_sku="ZX1",
        stocking_days=7, stock_in_warehouse=0,
        pending_receive=0, pending_ship=0,
    )


def test_curve_from_parsed_is_non_increasing_and_default_matches_alpha_one():
    daily = (5,) * 20
    curve = _curve_from_parsed(
        order_lines=[_order()],
        sales_records=[_sales()],
        shipping_in_progress={},
        daily_sales_by_key={("skc-1", "sku-1"): daily},
        alphas=(1.0, 0.9, 0.8, 0.7, 0.6),
    )

    ships = [point.ship_units for point in curve.points]
    assert ships == sorted(ships, reverse=True)

    alpha_one = next(point for point in curve.points if point.alpha == 1.0)
    assert alpha_one.ship_units == curve.default_ship_units


def test_interpolate_ship_lost_returns_exact_grid_point_as_floats():
    points = (
        _point(1.0, ship_units=100, lost_units=10),
        _point(0.8, ship_units=60, lost_units=40),
    )
    assert interpolate_ship_lost(points, 0.8) == (60.0, 40.0)


def test_interpolate_ship_lost_averages_at_midpoint():
    points = (
        _point(1.0, ship_units=100, lost_units=10),
        _point(0.8, ship_units=60, lost_units=40),
    )
    assert interpolate_ship_lost(points, 0.9) == (80.0, 25.0)


def test_interpolate_ship_lost_clamps_below_smallest_alpha():
    points = (
        _point(1.0, ship_units=100, lost_units=10),
        _point(0.8, ship_units=60, lost_units=40),
    )
    assert interpolate_ship_lost(points, 0.5) == (60.0, 40.0)


def test_interpolate_ship_lost_clamps_above_largest_alpha():
    points = (
        _point(1.0, ship_units=100, lost_units=10),
        _point(0.8, ship_units=60, lost_units=40),
    )
    assert interpolate_ship_lost(points, 1.2) == (100.0, 10.0)
