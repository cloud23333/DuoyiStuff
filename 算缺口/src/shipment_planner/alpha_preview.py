"""In-process preview of how ship volume and stockout risk vary with alpha.

For a grid of conservatism factors ``alpha`` (``protection_interval_factor``),
this recomputes today's recommended ship total in-process (no file writes) and
joins the stockout-risk backtest from :mod:`eval_policy`. The result drives an
operator-facing slider + plot: pick an ``alpha`` and see the trade-off between
units shipped today and simulated stockout risk.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

from . import eval_policy
from .engine import build_recommendations
from .models import OrderLine, SalesRecord
from .parsers import (
    ORDER_REQUIRED_COLUMNS,
    SALES_REQUIRED_COLUMNS,
    assert_required_columns,
    assert_temu_daily_sales_columns,
    parse_orders,
    parse_sales,
    parse_temu_daily_sales,
)
from .xlsx_reader import read_xlsx_table

DEFAULT_ALPHAS = (1.0, 0.95, 0.9, 0.85, 0.8, 0.75, 0.7, 0.65, 0.6, 0.55, 0.5)
SUGGESTED_MARGINAL_MAX = 2.0
FALLBACK_SUGGESTED_ALPHA = 0.9
_SUGGESTED_ALPHA_FLOOR = 0.7
_SUGGESTED_ALPHA_CEILING = 0.95


@dataclass(frozen=True, slots=True)
class AlphaPoint:
    alpha: float
    ship_units: int
    lost_units: int
    leftover_units: int
    fill_rate: float | None


@dataclass(frozen=True, slots=True)
class AlphaCurve:
    points: tuple[AlphaPoint, ...]
    suggested_alpha: float
    default_ship_units: int


def compute_alpha_curve(
    *,
    orders_path: str | Path,
    sales_path: str | Path,
    temu_path: str | Path,
    alphas: Sequence[float] = DEFAULT_ALPHAS,
) -> AlphaCurve:
    order_header, order_rows = read_xlsx_table(orders_path)
    assert_required_columns(order_header, ORDER_REQUIRED_COLUMNS, "orders file")
    order_lines, shipping_in_progress = parse_orders(order_rows)

    sales_header, sales_rows = read_xlsx_table(sales_path)
    assert_required_columns(sales_header, SALES_REQUIRED_COLUMNS, "sales file")
    sales_records = parse_sales(sales_rows)

    temu_header, temu_rows = read_xlsx_table(temu_path)
    assert_temu_daily_sales_columns(temu_header, "Temu daily sales file")
    daily_sales_by_key = parse_temu_daily_sales(temu_rows)

    return _curve_from_parsed(
        order_lines=order_lines,
        sales_records=sales_records,
        shipping_in_progress=shipping_in_progress,
        daily_sales_by_key=daily_sales_by_key,
        alphas=alphas,
    )


def _curve_from_parsed(
    *,
    order_lines: list[OrderLine],
    sales_records: list[SalesRecord],
    shipping_in_progress: dict[tuple[str, str], int],
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]],
    alphas: Sequence[float],
) -> AlphaCurve:
    ordered_alphas = _ordered_alphas(alphas)
    ship_by_alpha = {
        alpha: _ship_units_for_alpha(
            order_lines=order_lines,
            sales_records=sales_records,
            shipping_in_progress=shipping_in_progress,
            daily_sales_by_key=daily_sales_by_key,
            alpha=alpha,
        )
        for alpha in ordered_alphas
    }

    risk_by_alpha = _risk_by_alpha(
        daily_sales_by_key=daily_sales_by_key,
        sales_records=sales_records,
        alphas=ordered_alphas,
    )

    points = tuple(
        AlphaPoint(
            alpha=alpha,
            ship_units=ship_by_alpha[alpha],
            lost_units=risk_by_alpha[alpha]["lost_units"],
            leftover_units=risk_by_alpha[alpha]["leftover_units"],
            fill_rate=risk_by_alpha[alpha]["fill_rate"],
        )
        for alpha in ordered_alphas
    )

    default_ship_units = ship_by_alpha.get(1.0, points[0].ship_units if points else 0)

    return AlphaCurve(
        points=points,
        suggested_alpha=suggest_alpha(points),
        default_ship_units=default_ship_units,
    )


def _ordered_alphas(alphas: Sequence[float]) -> tuple[float, ...]:
    return tuple(sorted(dict.fromkeys(float(alpha) for alpha in alphas), reverse=True))


def _ship_units_for_alpha(
    *,
    order_lines: list[OrderLine],
    sales_records: list[SalesRecord],
    shipping_in_progress: dict[tuple[str, str], int],
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]],
    alpha: float,
) -> int:
    recommendations, _, _ = build_recommendations(
        order_lines,
        sales_records,
        shipping_in_progress_by_key=shipping_in_progress,
        daily_sales_by_key=daily_sales_by_key,
        protection_interval_factor=alpha,
    )
    return int(sum(int(rec["recommended_ship"]) for rec in recommendations))


def _risk_by_alpha(
    *,
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]],
    sales_records: list[SalesRecord],
    alphas: Sequence[float],
) -> dict[float, dict[str, object]]:
    sales_by_key: dict[tuple[str, str], SalesRecord] = {
        (record.skc, record.skuid): record for record in sales_records
    }
    _, _, grid_summary = eval_policy.run_policy_grid(
        daily_sales_by_key=daily_sales_by_key,
        sales_by_key=sales_by_key,
        offsets=[0.0],
        alphas=list(alphas),
    )
    return {float(cell["alpha"]): cell for cell in grid_summary}


def suggest_alpha(points: Sequence[AlphaPoint]) -> float:
    ordered = sorted(points, key=lambda point: point.alpha, reverse=True)
    if len(ordered) < 2:
        return FALLBACK_SUGGESTED_ALPHA

    chosen = ordered[0].alpha
    any_leftover_saved = False
    for previous, current in zip(ordered, ordered[1:]):
        leftover_saved = previous.leftover_units - current.leftover_units
        extra_lost = current.lost_units - previous.lost_units
        if leftover_saved <= 0:
            break
        any_leftover_saved = True
        if extra_lost / leftover_saved > SUGGESTED_MARGINAL_MAX:
            break
        chosen = current.alpha

    if not any_leftover_saved:
        return FALLBACK_SUGGESTED_ALPHA
    return _clamp(chosen, _SUGGESTED_ALPHA_FLOOR, _SUGGESTED_ALPHA_CEILING)


def _clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))
