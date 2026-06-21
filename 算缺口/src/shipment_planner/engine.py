from __future__ import annotations

import math
import statistics
from collections import defaultdict
from collections.abc import Sequence
from dataclasses import dataclass

from .allocation import allocate_recommendation_quantities
from .forecast_decision import resolve_service_level
from .forecast_distribution import (
    hurdle_distribution,
    negbin_ewma_distribution,
    poisson_ewma_distribution,
)
from .forecast_level import (
    clean_isolated_spikes,
    has_recent_drop,
    has_sustained_rise,
)
from .models import ForecastMetrics, KeyState, OrderLine, SalesRecord
from .parsers import normalize_sku_code
from .post_processing import (
    apply_small_change_keep_rule,
    assign_order_decision_reasons,
    assign_order_intercept_warnings,
    decision_reason,
    flag_min_order_ship_qty,
    line_change_ratio,
    refresh_key_recommended_totals,
    refresh_line_decision_reasons,
    round_qty,
)
from .summary import build_summary

DEFAULT_BASE_STOCK_QTY = 2
# Demand estimator in use. One of: negbin_ewma / poisson_ewma / hurdle.
SELECTED_ESTIMATOR = "hurdle"
MODEL_CURRENT = "current"
SMALL_CHANGE_KEEP_RATIO = 0.3
FORECAST_STRATEGY_CONSERVATIVE = "保守"
FORECAST_STRATEGY_NORMAL = "正常"
FORECAST_STRATEGY_AGGRESSIVE = "激进"
FORECAST_STRATEGY_SLOW_MOVER = "慢销"
FORECAST_STRATEGY_NO_SALES = "无销量"

DEMAND_PROFILE_STABLE = "稳定款"
DEMAND_PROFILE_VOLATILE = "波动款"
DEMAND_PROFILE_INTERMITTENT = "慢销/间歇款"
DEMAND_PROFILE_NO_SALES = "无销量款"

ANOMALY_NONE = "无"
ANOMALY_ISOLATED_SPIKE = "孤立爆单"
ANOMALY_RECENT_DROP = "连续暴跌"
ANOMALY_SINGLE_DAY_BIG_ORDER = "单日大单"
ANOMALY_STOCKOUT_TAIL = "疑似缺货尾部"

# Single-day big-order detection thresholds (used by _adjust_single_day_big_order).
SINGLE_DAY_BIG_ORDER_MIN_QTY = 10.0
SINGLE_DAY_BIG_ORDER_SHARE = 0.70
SINGLE_DAY_BIG_ORDER_PEER_MULTIPLIER = 5.0
SINGLE_DAY_BIG_ORDER_MIN_SMALL_SALE_DAYS = 2
# Demand-profile classification thresholds (used by _demand_profile).
INTERMITTENT_ZERO_RATIO = 0.5
STABLE_VOLATILITY_THRESHOLD = 0.8
# Service level used by the single-day big-order base-stock fallback.
SERVICE_LEVEL_CONSERVATIVE = 0.60

_ESTIMATORS = {
    "negbin_ewma": negbin_ewma_distribution,
    "poisson_ewma": poisson_ewma_distribution,
    "hurdle": hurdle_distribution,
}


@dataclass(frozen=True, slots=True)
class _SalesExplanation:
    daily_sales: tuple[float, ...]
    single_day_big_order: bool
    base_stock_only: bool
    stockout_tail: bool


def build_recommendations(
    order_lines: list[OrderLine],
    sales_records: list[SalesRecord],
    min_order_ship_qty: int = 10,
    sku_order_max_qty: dict[str, int] | None = None,
    exclude_skc: set[str] | None = None,
    exclude_skuid: set[str] | None = None,
    shipping_in_progress_by_key: dict[tuple[str, str], int] | None = None,
    service_level_offset: float = 0.0,
    base_stock_qty: int = DEFAULT_BASE_STOCK_QTY,
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]] | None = None,
) -> tuple[list[dict[str, object]], list[dict[str, object]], dict[str, object]]:
    if base_stock_qty < 0:
        raise ValueError("base_stock_qty must be greater than or equal to 0.")

    ordered_lines = sorted(order_lines, key=lambda line: line.row_number)
    normalized_sku_limits = _normalize_sku_limits(sku_order_max_qty)
    normalized_exclude_skc = _normalize_excluded_codes(exclude_skc)
    normalized_exclude_skuid = _normalize_excluded_codes(exclude_skuid)
    sales_by_key, duplicate_keys = _build_sales_lookup(sales_records)
    key_demand = _build_key_demand(ordered_lines)
    missing_daily_sales_keys = _missing_daily_sales_keys(key_demand, daily_sales_by_key)
    shipping_in_progress_lookup = shipping_in_progress_by_key or {}
    key_states = _build_key_states(
        key_demand=key_demand,
        sales_by_key=sales_by_key,
        daily_sales_by_key=daily_sales_by_key or {},
        missing_daily_sales_keys=missing_daily_sales_keys,
        shipping_in_progress_by_key=shipping_in_progress_lookup,
        service_level_offset=service_level_offset,
        base_stock_qty=base_stock_qty,
    )
    suggested_by_row, sku_order_limit_capped_rows = (
        allocate_recommendation_quantities(
            order_lines=ordered_lines,
            key_states=key_states,
            sales_by_key=sales_by_key,
            sku_order_max_qty=normalized_sku_limits,
        )
    )
    suggested_by_row_before_intercept = dict(suggested_by_row)
    intercept_reason_by_row = _build_intercept_reason_by_row(
        ordered_lines,
        exclude_skc=normalized_exclude_skc,
        exclude_skuid=normalized_exclude_skuid,
    )
    _apply_intercepts_to_suggestions(suggested_by_row, intercept_reason_by_row)
    key_recommended_totals = _build_key_recommended_totals(
        ordered_lines, suggested_by_row
    )
    intercepted_order_lines = len(intercept_reason_by_row)

    recommendations: list[dict[str, object]] = []
    quality_rows: list[dict[str, object]] = []

    for line in ordered_lines:
        key = (line.skc, line.skuid)
        state = key_states.get(key)
        sales = sales_by_key.get(key)

        system_sku = sales.system_sku if sales is not None else ""
        display_order_sku = _display_sku_with_source_order(
            line.order_sku, line.product_code
        )
        display_system_sku = _display_sku_with_source_order(
            system_sku, line.product_code
        )
        key_order_qty = state.order_qty_total if state is not None else line.quantity
        key_recommended_total = key_recommended_totals.get(key, 0)
        gap = state.gap if state is not None else 0
        forecast_metrics = (
            state.forecast_metrics
            if state is not None
            else ForecastMetrics("", 0.0, 0.0)
        )
        base_stock_target = state.base_stock_qty if state is not None else 0
        base_stock_gap = state.base_stock_gap if state is not None else 0
        base_stock_triggered = (
            state.base_stock_triggered if state is not None else False
        )
        is_min_order_ship_qty_exempt_eligible = (
            state.min_order_ship_qty_exempt_eligible if state is not None else False
        )
        suggested_qty = suggested_by_row.get(line.row_number, 0)
        intercept_reason = intercept_reason_by_row.get(line.row_number, "")

        sku_code_check, quality_issue_row = _evaluate_sku_code(line, sales)
        if quality_issue_row is not None:
            quality_rows.append(quality_issue_row)
        if sales is not None and key in missing_daily_sales_keys:
            quality_rows.append(
                _quality_issue_row(
                    line,
                    issue_type="missing_daily_sales",
                    system_sku=system_sku,
                    message="Missing daily sales data for (SKC, SKUID)",
                )
            )

        (
            stocking_days,
            stock_in_warehouse,
            pending_receive,
            pending_ship,
        ) = _sales_metrics(sales)
        shipping_in_progress = shipping_in_progress_lookup.get(key, 0)

        recommendations.append(
            {
                "row_number": line.row_number,
                "internal_order_id": line.internal_order_id,
                "店铺款式编码": line.skc,
                "店铺商品编码": line.skuid,
                "原始商品编码": display_order_sku,
                "系统商品编码": display_system_sku,
                "line_order_qty": line.quantity,
                "key_order_qty": key_order_qty,
                "forecast_strategy": forecast_metrics.strategy,
                "demand_profile": forecast_metrics.demand_profile,
                "anomaly_flags": forecast_metrics.anomaly_flags,
                "service_level": round_qty(forecast_metrics.service_level),
                "forecast_model": forecast_metrics.forecast_model,
                "effective_daily_sales": round_qty(
                    forecast_metrics.effective_daily_sales
                ),
                "forecast_daily_sales": round_qty(
                    forecast_metrics.forecast_daily_sales
                ),
                "forecast_stocking_period_sales": round_qty(
                    forecast_metrics.forecast_stocking_period_sales
                ),
                "stocking_days": round_qty(stocking_days),
                "wh": round_qty(stock_in_warehouse),
                "pending_recv": round_qty(pending_receive),
                "pending_ship": round_qty(pending_ship),
                "shipping_in_progress": shipping_in_progress,
                "gap": gap,
                "base_stock_qty": base_stock_target,
                "base_stock_gap": base_stock_gap,
                "base_stock_triggered_warning": (
                    "yes" if base_stock_triggered else "no"
                ),
                "recommended_ship": suggested_qty,
                "recommended_ship_before_small_change_rule": suggested_qty,
                "small_change_ratio_before_rule": round_qty(
                    line_change_ratio(line.quantity, suggested_qty)
                ),
                "small_change_keep_warning": "no",
                "key_recommended_total": key_recommended_total,
                "decision_reason": decision_reason(line.quantity, suggested_qty),
                "order_decision_reason": "",
                "sku_code_check": sku_code_check,
                "intercept_reason": intercept_reason,
                "order_recommended_ship_total_before_threshold": 0,
                "min_order_ship_qty_threshold": min_order_ship_qty,
                "order_low_qty_warning": "no",
                "min_order_ship_qty_exempt_eligible": (
                    is_min_order_ship_qty_exempt_eligible
                ),
                "min_order_ship_qty_exempt_applied": False,
                "min_order_ship_qty_exempt_warning": (
                    "yes" if is_min_order_ship_qty_exempt_eligible else "no"
                ),
                "min_order_ship_qty_exempt_applied_warning": "no",
                "order_intercept_warning": "no",
            }
        )

    intercept_stats = assign_order_intercept_warnings(
        recommendations,
        suggested_by_row_before_intercept=suggested_by_row_before_intercept,
    )
    small_change_stats = apply_small_change_keep_rule(
        recommendations,
        order_lines=ordered_lines,
        keep_change_ratio=SMALL_CHANGE_KEEP_RATIO,
        sales_by_key=sales_by_key,
        sku_order_max_qty=normalized_sku_limits,
    )
    sku_order_limit_capped_rows.update(
        small_change_stats.get("sku_order_limit_capped_rows", set())
    )
    threshold_stats = flag_min_order_ship_qty(recommendations, min_order_ship_qty)
    refresh_key_recommended_totals(recommendations)
    refresh_line_decision_reasons(recommendations)
    assign_order_decision_reasons(recommendations, ordered_lines)
    summary = build_summary(
        ordered_lines,
        sales_records,
        recommendations,
        quality_rows,
        duplicate_keys,
        min_order_ship_qty=min_order_ship_qty,
        threshold_stats=threshold_stats,
        sku_order_limit_rule_count=len(normalized_sku_limits),
        sku_order_limit_capped_lines=len(sku_order_limit_capped_rows),
        excluded_skc_rule_count=len(normalized_exclude_skc),
        excluded_skuid_rule_count=len(normalized_exclude_skuid),
        intercepted_order_lines=intercepted_order_lines,
        intercepted_orders=intercept_stats.get("intercepted_orders", 0),
        small_change_kept_lines=small_change_stats.get("small_change_kept_lines", 0),
        service_level_offset=service_level_offset,
        base_stock_qty=base_stock_qty,
    )
    return recommendations, quality_rows, summary


def _normalize_sku_limits(
    sku_order_max_qty: dict[str, int] | None,
) -> dict[str, int]:
    normalized_sku_limits: dict[str, int] = {}
    for sku, limit in (sku_order_max_qty or {}).items():
        normalized_sku = normalize_sku_code(sku)
        if not normalized_sku:
            continue
        normalized_sku_limits[normalized_sku] = limit
    return normalized_sku_limits


def _normalize_excluded_codes(codes: set[str] | None) -> set[str]:
    return {code.strip() for code in (codes or set()) if code.strip()}


def _missing_daily_sales_keys(
    key_demand: dict[tuple[str, str], int],
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]] | None,
) -> set[tuple[str, str]]:
    if daily_sales_by_key is None:
        raise ValueError("daily sales data is required.")

    missing_keys: set[tuple[str, str]] = set()
    for skc, skuid in key_demand:
        key = (skc, skuid)
        if not daily_sales_by_key.get(key):
            missing_keys.add(key)
    return missing_keys


def _build_intercept_reason_by_row(
    order_lines: list[OrderLine],
    *,
    exclude_skc: set[str],
    exclude_skuid: set[str],
) -> dict[int, str]:
    reasons: dict[int, str] = {}
    for line in order_lines:
        skc_hit = line.skc.strip() in exclude_skc
        skuid_hit = line.skuid.strip() in exclude_skuid
        if skc_hit and skuid_hit:
            reasons[line.row_number] = "skc_and_skuid"
        elif skc_hit:
            reasons[line.row_number] = "skc"
        elif skuid_hit:
            reasons[line.row_number] = "skuid"
    return reasons


def _apply_intercepts_to_suggestions(
    suggested_by_row: dict[int, int],
    intercept_reason_by_row: dict[int, str],
) -> None:
    for row_number in intercept_reason_by_row:
        suggested_by_row[row_number] = 0


def _build_key_recommended_totals(
    order_lines: list[OrderLine],
    suggested_by_row: dict[int, int],
) -> dict[tuple[str, str], int]:
    totals: dict[tuple[str, str], int] = defaultdict(int)
    for line in order_lines:
        key = (line.skc, line.skuid)
        totals[key] += suggested_by_row.get(line.row_number, 0)
    return dict(totals)


def _display_sku_with_source_order(sku_value: str, source_product_code: str) -> str:
    if not sku_value:
        return ""
    if normalize_sku_code(sku_value) == normalize_sku_code(source_product_code):
        return source_product_code
    return sku_value


def _evaluate_sku_code(
    line: OrderLine,
    sales: SalesRecord | None,
) -> tuple[str, dict[str, object] | None]:
    if sales is None:
        return (
            "missing_key",
            _quality_issue_row(
                line,
                issue_type="missing_sales_key",
                system_sku="",
                message="No sales row found for (SKC, SKUID)",
            ),
        )

    system_sku = sales.system_sku
    if line.order_sku == system_sku:
        return "exact_match", None

    if normalize_sku_code(line.order_sku) == normalize_sku_code(system_sku):
        return "normalized_match", None

    return (
        "diff",
        _quality_issue_row(
            line,
            issue_type="sku_code_diff",
            system_sku=system_sku,
            message="Order SKU and system SKU differ after normalization",
        ),
    )


def _sales_metrics(
    sales: SalesRecord | None,
) -> tuple[float, float, float, float]:
    if sales is None:
        return 0.0, 0.0, 0.0, 0.0

    return (
        sales.stocking_days,
        sales.stock_in_warehouse,
        sales.pending_receive,
        sales.pending_ship,
    )


def _quality_issue_row(
    line: OrderLine,
    issue_type: str,
    system_sku: str,
    message: str,
) -> dict[str, object]:
    return {
        "type": issue_type,
        "row_number": line.row_number,
        "internal_order_id": line.internal_order_id,
        "skc": line.skc,
        "skuid": line.skuid,
        "order_sku": line.order_sku,
        "system_sku": system_sku,
        "message": message,
    }


def _build_sales_lookup(
    sales_records: list[SalesRecord],
) -> tuple[dict[tuple[str, str], SalesRecord], set[tuple[str, str]]]:
    # Accumulate totals into plain intermediate dicts — never mutate SalesRecord instances
    first_seen: dict[tuple[str, str], SalesRecord] = {}
    accum: dict[tuple[str, str], dict] = {}
    duplicate_keys: set[tuple[str, str]] = set()

    for record in sales_records:
        key = (record.skc, record.skuid)
        if key not in accum:
            first_seen[key] = record
            accum[key] = {
                "stocking_days": record.stocking_days,
                "stock_in_warehouse": record.stock_in_warehouse,
                "pending_receive": record.pending_receive,
                "pending_ship": record.pending_ship,
                "system_sku": record.system_sku,
            }
        else:
            duplicate_keys.add(key)
            acc = accum[key]
            accum[key] = {
                "stocking_days": max(acc["stocking_days"], record.stocking_days),
                "stock_in_warehouse": acc["stock_in_warehouse"] + record.stock_in_warehouse,
                "pending_receive": acc["pending_receive"] + record.pending_receive,
                "pending_ship": acc["pending_ship"] + record.pending_ship,
                "system_sku": acc["system_sku"] if acc["system_sku"] else record.system_sku,
            }

    # Construct new SalesRecord instances from accumulated data
    lookup: dict[tuple[str, str], SalesRecord] = {
        key: SalesRecord(
            row_number=first_seen[key].row_number,
            skc=first_seen[key].skc,
            skuid=first_seen[key].skuid,
            stocking_days=acc["stocking_days"],
            stock_in_warehouse=acc["stock_in_warehouse"],
            pending_receive=acc["pending_receive"],
            pending_ship=acc["pending_ship"],
            system_sku=acc["system_sku"],
        )
        for key, acc in accum.items()
    }

    return lookup, duplicate_keys


def _build_key_demand(order_lines: list[OrderLine]) -> dict[tuple[str, str], int]:
    demand: dict[tuple[str, str], int] = defaultdict(int)
    for line in order_lines:
        demand[(line.skc, line.skuid)] += line.quantity
    return dict(demand)


def _build_key_states(
    key_demand: dict[tuple[str, str], int],
    sales_by_key: dict[tuple[str, str], SalesRecord],
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]],
    missing_daily_sales_keys: set[tuple[str, str]],
    shipping_in_progress_by_key: dict[tuple[str, str], int],
    service_level_offset: float,
    base_stock_qty: int,
) -> dict[tuple[str, str], KeyState]:
    states: dict[tuple[str, str], KeyState] = {}
    for key, order_qty_total in key_demand.items():
        skc, skuid = key
        sales = sales_by_key.get(key)
        if sales is None:
            states[key] = KeyState(
                skc=skc,
                skuid=skuid,
                system_sku="",
                order_qty_total=order_qty_total,
                gap=0,
                recommended_qty_total=0,
                forecast_metrics=ForecastMetrics("", 0.0, 0.0),
                min_order_ship_qty_exempt_eligible=False,
            )
            continue

        shipping_in_progress = shipping_in_progress_by_key.get(key, 0)
        available_stock = (
            sales.stock_in_warehouse + sales.pending_receive + shipping_in_progress
        )
        if key in missing_daily_sales_keys:
            base_stock_gap_raw = _base_stock_gap(
                base_stock_qty=base_stock_qty,
                available_stock=available_stock,
            )
            gap = math.ceil(base_stock_gap_raw)
            base_stock_triggered = base_stock_gap_raw > 0
            states[key] = KeyState(
                skc=skc,
                skuid=skuid,
                system_sku=sales.system_sku,
                order_qty_total=order_qty_total,
                gap=gap,
                recommended_qty_total=min(order_qty_total, gap),
                forecast_metrics=ForecastMetrics("", 0.0, 0.0),
                min_order_ship_qty_exempt_eligible=base_stock_triggered,
                base_stock_qty=base_stock_qty,
                base_stock_gap=gap,
                base_stock_triggered=base_stock_triggered,
            )
            continue

        forecast_metrics = compute_forecast_metrics(
            daily_sales=daily_sales_by_key[key],
            stocking_days=sales.stocking_days,
            stock_in_warehouse=sales.stock_in_warehouse,
            service_level_offset=service_level_offset,
        )
        forecast_gap = max(
            0.0,
            forecast_metrics.forecast_stocking_period_sales - available_stock,
        )
        base_stock_gap_raw = _base_stock_gap(
            base_stock_qty=base_stock_qty,
            available_stock=available_stock,
        )
        base_stock_below_floor = base_stock_gap_raw > 0
        base_stock_triggered = (
            base_stock_below_floor and base_stock_gap_raw >= forecast_gap
        )
        raw_gap = max(forecast_gap, base_stock_gap_raw)
        gap = math.ceil(raw_gap)
        recommended_qty_total = min(order_qty_total, gap)
        states[key] = KeyState(
            skc=skc,
            skuid=skuid,
            system_sku=sales.system_sku,
            order_qty_total=order_qty_total,
            gap=gap,
            recommended_qty_total=recommended_qty_total,
            forecast_metrics=forecast_metrics,
            min_order_ship_qty_exempt_eligible=base_stock_below_floor,
            base_stock_qty=base_stock_qty,
            base_stock_gap=math.ceil(base_stock_gap_raw),
            base_stock_triggered=base_stock_triggered,
        )
    return states


def _base_stock_gap(*, base_stock_qty: int, available_stock: float) -> float:
    if base_stock_qty <= 0:
        return 0.0
    return max(0.0, float(base_stock_qty) - available_stock)


def _adjust_single_day_big_order(
    daily_sales: Sequence[float],
) -> tuple[tuple[float, ...], bool, bool]:
    values = tuple(max(0.0, float(value)) for value in daily_sales)
    if not values:
        return values, False, False

    max_qty = max(values)
    if max_qty < SINGLE_DAY_BIG_ORDER_MIN_QTY:
        return values, False, False
    if values.count(max_qty) != 1:
        return values, False, False

    max_index = values.index(max_qty)
    peer_positive = [
        value for index, value in enumerate(values) if index != max_index and value > 0
    ]
    positive_day_count = len(peer_positive) + 1
    if positive_day_count == 1:
        is_single_day_big_order = True
    else:
        total_sales = sum(values)
        peer_median = float(statistics.median(peer_positive))
        is_single_day_big_order = (
            total_sales > 0
            and max_qty / total_sales >= SINGLE_DAY_BIG_ORDER_SHARE
            and max_qty >= peer_median * SINGLE_DAY_BIG_ORDER_PEER_MULTIPLIER
        )

    if not is_single_day_big_order:
        return values, False, False

    has_small_sale_baseline = (
        len(peer_positive) >= SINGLE_DAY_BIG_ORDER_MIN_SMALL_SALE_DAYS
    )
    if not has_small_sale_baseline:
        return values, True, True

    adjusted = list(values)
    adjusted[max_index] = float(statistics.median(peer_positive))
    return tuple(adjusted), True, False


def _explain_daily_sales(
    *,
    daily_sales: Sequence[float],
    stock_in_warehouse: float,
    apply_stockout_mask: bool,
) -> _SalesExplanation:
    normalized = tuple(max(0.0, float(value)) for value in daily_sales)
    (
        adjusted_daily_sales,
        single_day_big_order,
        base_stock_only,
    ) = _adjust_single_day_big_order(normalized)
    stockout_cut = (
        stockout_mask_cut(
            daily_sales=adjusted_daily_sales,
            stock_in_warehouse=stock_in_warehouse,
        )
        if apply_stockout_mask
        else None
    )
    effective_daily_sales = (
        adjusted_daily_sales[:stockout_cut]
        if stockout_cut is not None
        else adjusted_daily_sales
    )
    return _SalesExplanation(
        daily_sales=effective_daily_sales,
        single_day_big_order=single_day_big_order,
        base_stock_only=base_stock_only,
        stockout_tail=stockout_cut is not None,
    )


def _single_day_big_order_base_stock_metrics(
    *,
    stockout_tail: bool,
) -> ForecastMetrics:
    return ForecastMetrics(
        strategy=FORECAST_STRATEGY_CONSERVATIVE,
        forecast_daily_sales=0.0,
        forecast_stocking_period_sales=0.0,
        demand_profile=DEMAND_PROFILE_INTERMITTENT,
        anomaly_flags=_anomaly_flags(
            isolated_spike=False,
            recent_drop=False,
            single_day_big_order=True,
            stockout_tail=stockout_tail,
        ),
        service_level=SERVICE_LEVEL_CONSERVATIVE,
        forecast_model=MODEL_CURRENT,
        effective_daily_sales=0.0,
    )


def compute_forecast_metrics(
    *,
    daily_sales: tuple[int, ...],
    stocking_days: float,
    stock_in_warehouse: float = 0.0,
    apply_stockout_mask: bool = True,
    service_level_offset: float = 0.0,
) -> ForecastMetrics:
    """Public wrapper for the forecast pipeline (used by eval harness + engine)."""
    sales_explanation = _explain_daily_sales(
        daily_sales=daily_sales,
        stock_in_warehouse=stock_in_warehouse,
        apply_stockout_mask=apply_stockout_mask,
    )
    if sales_explanation.base_stock_only:
        return _single_day_big_order_base_stock_metrics(
            stockout_tail=sales_explanation.stockout_tail,
        )

    return _forecast_metrics(
        daily_sales=sales_explanation.daily_sales,
        stocking_days=stocking_days,
        stock_in_warehouse=stock_in_warehouse,
        single_day_big_order=sales_explanation.single_day_big_order,
        stockout_tail=sales_explanation.stockout_tail,
        service_level_offset=service_level_offset,
    )


def stockout_mask_cut(
    *,
    daily_sales: tuple[float, ...],
    stock_in_warehouse: float,
) -> int | None:
    """Return the index where the stockout mask starts, or None if no mask applies.

    When a cut ``c`` is returned, ``daily_sales[:c]`` is the kept history and
    ``daily_sales[c:]`` are the trailing zeros attributed to being out of stock.
    """
    if stock_in_warehouse > 0 or not daily_sales:
        return None

    last_nonzero = -1
    for index in range(len(daily_sales) - 1, -1, -1):
        if daily_sales[index] > 0:
            last_nonzero = index
            break

    if last_nonzero < 0 or last_nonzero == len(daily_sales) - 1:
        return None

    return last_nonzero + 1


def _forecast_metrics(
    *,
    daily_sales: Sequence[float],
    stocking_days: float,
    stock_in_warehouse: float = 0.0,
    single_day_big_order: bool = False,
    stockout_tail: bool = False,
    service_level_offset: float = 0.0,
) -> ForecastMetrics:
    values = [max(0.0, float(value)) for value in daily_sales]
    cleaned_values, isolated_spike = clean_isolated_spikes(values)
    recent_drop = has_recent_drop(values, stock_in_warehouse=stock_in_warehouse)
    sustained_rise = has_sustained_rise(values)
    demand_profile = _demand_profile(
        values=cleaned_values,
        volatility=_simple_volatility(cleaned_values),
        isolated_spike=isolated_spike,
    )
    anomaly_flags = _anomaly_flags(
        isolated_spike=isolated_spike,
        recent_drop=recent_drop,
        single_day_big_order=single_day_big_order,
        stockout_tail=stockout_tail,
    )

    horizon = max(1, math.ceil(stocking_days)) if stocking_days > 0 else 1
    estimator = _ESTIMATORS[SELECTED_ESTIMATOR]
    distribution = estimator(values, horizon=horizon, recent_drop=recent_drop)

    service_level = resolve_service_level(
        recent_drop=recent_drop,
        sustained_rise=sustained_rise,
        offset=service_level_offset,
    )
    quantile_units = distribution.quantile(service_level)  # over `horizon` days
    period_days = max(0.0, stocking_days)
    # Rescale the horizon quantile/mean to the actual stocking period.
    scale = (period_days / horizon) if horizon > 0 else 0.0
    forecast_stocking_period_sales = quantile_units * scale
    predictive_mean = distribution.mean * scale
    forecast_daily_sales = (
        forecast_stocking_period_sales / period_days if period_days > 0 else 0.0
    )
    effective_daily_sales = distribution.mean / horizon if horizon > 0 else 0.0

    strategy = _derive_strategy_label(
        demand_profile=demand_profile,
        recent_drop=recent_drop,
        sustained_rise=sustained_rise,
    )
    return ForecastMetrics(
        strategy=strategy,
        forecast_daily_sales=forecast_daily_sales,
        forecast_stocking_period_sales=forecast_stocking_period_sales,
        demand_profile=demand_profile,
        anomaly_flags=anomaly_flags,
        service_level=service_level,
        forecast_model=SELECTED_ESTIMATOR,
        effective_daily_sales=effective_daily_sales,
        predictive_mean=predictive_mean,
        dispersion=distribution.variance,
    )


def _simple_volatility(values: Sequence[float]) -> float:
    if len(values) < 2:
        return 0.0
    mean = statistics.fmean(values)
    if mean <= 0:
        return 0.0
    return statistics.pstdev(values) / mean


def _derive_strategy_label(
    *,
    demand_profile: str,
    recent_drop: bool,
    sustained_rise: bool,
) -> str:
    if demand_profile == DEMAND_PROFILE_NO_SALES:
        return FORECAST_STRATEGY_NO_SALES
    if demand_profile == DEMAND_PROFILE_INTERMITTENT:
        return FORECAST_STRATEGY_SLOW_MOVER
    if recent_drop:
        return FORECAST_STRATEGY_CONSERVATIVE
    if sustained_rise:
        return FORECAST_STRATEGY_AGGRESSIVE
    return FORECAST_STRATEGY_NORMAL


def _demand_profile(
    *,
    values: Sequence[float],
    volatility: float,
    isolated_spike: bool,
) -> str:
    if not values or sum(values) <= 0:
        return DEMAND_PROFILE_NO_SALES

    zero_ratio = sum(1 for value in values if value <= 0) / len(values)
    median_daily = float(statistics.median(values))
    if zero_ratio >= INTERMITTENT_ZERO_RATIO or median_daily == 0.0:
        return DEMAND_PROFILE_INTERMITTENT

    if isolated_spike or volatility > STABLE_VOLATILITY_THRESHOLD:
        return DEMAND_PROFILE_VOLATILE

    return DEMAND_PROFILE_STABLE


def _anomaly_flags(
    *,
    isolated_spike: bool,
    recent_drop: bool,
    single_day_big_order: bool,
    stockout_tail: bool,
) -> str:
    flags: list[str] = []
    if single_day_big_order:
        flags.append(ANOMALY_SINGLE_DAY_BIG_ORDER)
    if isolated_spike:
        flags.append(ANOMALY_ISOLATED_SPIKE)
    if recent_drop:
        flags.append(ANOMALY_RECENT_DROP)
    if stockout_tail:
        flags.append(ANOMALY_STOCKOUT_TAIL)
    return "、".join(flags) if flags else ANOMALY_NONE


