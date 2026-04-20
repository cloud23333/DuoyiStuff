from __future__ import annotations

from collections import defaultdict
import math
import statistics

from .allocation import allocate_recommendation_quantities
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

HOT_STYLE_GAP_MULTIPLIER = 1.2
DEFAULT_GLOBAL_GAP_MULTIPLIER = 1.0
SMALL_CHANGE_KEEP_RATIO = 0.3
FORECAST_STRATEGY_CONSERVATIVE = "保守"
FORECAST_STRATEGY_NORMAL = "正常"
FORECAST_STRATEGY_AGGRESSIVE = "激进"


def build_recommendations(
    order_lines: list[OrderLine],
    sales_records: list[SalesRecord],
    min_order_ship_qty: int = 10,
    sku_order_max_qty: dict[str, int] | None = None,
    exclude_skc: set[str] | None = None,
    exclude_skuid: set[str] | None = None,
    shipping_in_progress_by_key: dict[tuple[str, str], int] | None = None,
    global_gap_multiplier: float = DEFAULT_GLOBAL_GAP_MULTIPLIER,
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]] | None = None,
) -> tuple[list[dict[str, object]], list[dict[str, object]], dict[str, object]]:
    if global_gap_multiplier <= 0:
        raise ValueError("global_gap_multiplier must be greater than 0.")

    ordered_lines = sorted(order_lines, key=lambda line: line.row_number)
    normalized_sku_limits = _normalize_sku_limits(sku_order_max_qty)
    normalized_exclude_skc = _normalize_excluded_codes(exclude_skc)
    normalized_exclude_skuid = _normalize_excluded_codes(exclude_skuid)
    sales_by_key, duplicate_keys = _build_sales_lookup(sales_records)
    key_demand = _build_key_demand(ordered_lines)
    _validate_daily_sales_for_demand(key_demand, daily_sales_by_key)
    shipping_in_progress_lookup = shipping_in_progress_by_key or {}
    key_states = _build_key_states(
        key_demand=key_demand,
        sales_by_key=sales_by_key,
        daily_sales_by_key=daily_sales_by_key or {},
        shipping_in_progress_by_key=shipping_in_progress_lookup,
        global_gap_multiplier=global_gap_multiplier,
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
        is_min_order_ship_qty_exempt_eligible = (
            state.min_order_ship_qty_exempt_eligible if state is not None else False
        )
        suggested_qty = suggested_by_row.get(line.row_number, 0)
        intercept_reason = intercept_reason_by_row.get(line.row_number, "")

        sku_code_check, quality_issue_row = _evaluate_sku_code(line, sales)
        if quality_issue_row is not None:
            quality_rows.append(quality_issue_row)

        (
            sold30,
            sold7,
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
                "sold30": sold30,
                "sold7": sold7,
                "forecast_strategy": forecast_metrics.strategy,
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
        global_gap_multiplier=global_gap_multiplier,
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


def _validate_daily_sales_for_demand(
    key_demand: dict[tuple[str, str], int],
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]] | None,
) -> None:
    if daily_sales_by_key is None:
        raise ValueError("daily sales data is required.")

    for skc, skuid in key_demand:
        if (skc, skuid) not in daily_sales_by_key:
            raise ValueError(f"Missing daily sales data for (SKC, SKUID): {skc}, {skuid}")
        if not daily_sales_by_key[(skc, skuid)]:
            raise ValueError(f"Empty daily sales sequence for (SKC, SKUID): {skc}, {skuid}")


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
) -> tuple[int, int, float, float, float, float]:
    if sales is None:
        return 0, 0, 0.0, 0.0, 0.0, 0.0

    return (
        sales.sold30,
        sales.sold7,
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
                "sold30": record.sold30,
                "sold7": record.sold7,
                "stocking_days": record.stocking_days,
                "stock_in_warehouse": record.stock_in_warehouse,
                "pending_receive": record.pending_receive,
                "pending_ship": record.pending_ship,
                "is_hot_style": record.is_hot_style,
                "system_sku": record.system_sku,
            }
        else:
            duplicate_keys.add(key)
            acc = accum[key]
            accum[key] = {
                "sold30": acc["sold30"] + record.sold30,
                "sold7": acc["sold7"] + record.sold7,
                "stocking_days": max(acc["stocking_days"], record.stocking_days),
                "stock_in_warehouse": acc["stock_in_warehouse"] + record.stock_in_warehouse,
                "pending_receive": acc["pending_receive"] + record.pending_receive,
                "pending_ship": acc["pending_ship"] + record.pending_ship,
                "is_hot_style": acc["is_hot_style"] or record.is_hot_style,
                "system_sku": acc["system_sku"] if acc["system_sku"] else record.system_sku,
            }

    # Construct new SalesRecord instances from accumulated data
    lookup: dict[tuple[str, str], SalesRecord] = {
        key: SalesRecord(
            row_number=first_seen[key].row_number,
            skc=first_seen[key].skc,
            skuid=first_seen[key].skuid,
            sold30=acc["sold30"],
            sold7=acc["sold7"],
            stocking_days=acc["stocking_days"],
            stock_in_warehouse=acc["stock_in_warehouse"],
            pending_receive=acc["pending_receive"],
            pending_ship=acc["pending_ship"],
            is_hot_style=acc["is_hot_style"],
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
    shipping_in_progress_by_key: dict[tuple[str, str], int],
    global_gap_multiplier: float,
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
        forecast_metrics = _forecast_metrics(
            daily_sales=daily_sales_by_key[key],
            stocking_days=sales.stocking_days,
            is_hot_style=sales.is_hot_style,
        )
        available_stock = (
            sales.stock_in_warehouse + sales.pending_receive + shipping_in_progress
        )
        raw_gap = max(
            0.0,
            forecast_metrics.forecast_stocking_period_sales - available_stock,
        )
        if sales.is_hot_style:
            raw_gap *= HOT_STYLE_GAP_MULTIPLIER
        raw_gap *= global_gap_multiplier
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
            min_order_ship_qty_exempt_eligible=False,
        )
    return states


def _forecast_metrics(
    *,
    daily_sales: tuple[int, ...],
    stocking_days: float,
    is_hot_style: bool,
) -> ForecastMetrics:
    values = [max(0, int(value)) for value in daily_sales]
    mean_daily = statistics.fmean(values)
    median_daily = float(statistics.median(values))
    recent_7_avg = _average_recent(values, 7)
    recent_3_avg = _average_recent(values, 3)
    trimmed_mean = _trimmed_mean(values)
    volatility = _volatility(values, mean_daily)
    isolated_spike = _has_isolated_recent_spike(values, median_daily, recent_7_avg)

    strategy = _forecast_strategy(
        median_daily=median_daily,
        recent_7_avg=recent_7_avg,
        recent_3_avg=recent_3_avg,
        volatility=volatility,
        isolated_spike=isolated_spike,
        is_hot_style=is_hot_style,
    )
    forecast_daily_sales = _forecast_daily_sales(
        strategy=strategy,
        median_daily=median_daily,
        recent_7_avg=recent_7_avg,
        recent_3_avg=recent_3_avg,
        trimmed_mean=trimmed_mean,
    )
    forecast_stocking_period_sales = forecast_daily_sales * max(0.0, stocking_days)
    return ForecastMetrics(
        strategy=strategy,
        forecast_daily_sales=forecast_daily_sales,
        forecast_stocking_period_sales=forecast_stocking_period_sales,
    )


def _average_recent(values: list[int], days: int) -> float:
    window = values[-days:]
    return statistics.fmean(window) if window else 0.0


def _trimmed_mean(values: list[int]) -> float:
    if len(values) < 5:
        return statistics.fmean(values)
    ordered = sorted(values)
    return statistics.fmean(ordered[1:-1])


def _volatility(values: list[int], mean_daily: float) -> float:
    if len(values) < 2 or mean_daily <= 0:
        return 0.0
    return statistics.pstdev(values) / mean_daily


def _has_isolated_recent_spike(
    values: list[int],
    median_daily: float,
    recent_7_avg: float,
) -> bool:
    if len(values) < 3:
        return False
    recent = values[-3:]
    recent_max = max(recent)
    baseline = max(1.0, median_daily, recent_7_avg)
    if recent_max < baseline * 3:
        return False
    remaining = list(recent)
    remaining.remove(recent_max)
    return statistics.fmean(remaining) <= max(1.0, median_daily * 1.2)


def _forecast_strategy(
    *,
    median_daily: float,
    recent_7_avg: float,
    recent_3_avg: float,
    volatility: float,
    isolated_spike: bool,
    is_hot_style: bool,
) -> str:
    baseline = max(1.0, median_daily)
    if recent_3_avg < baseline * 0.6:
        return FORECAST_STRATEGY_CONSERVATIVE
    if volatility > 1.2 or isolated_spike:
        return FORECAST_STRATEGY_CONSERVATIVE
    if recent_3_avg > baseline * 1.5:
        return FORECAST_STRATEGY_AGGRESSIVE
    if is_hot_style and recent_7_avg >= baseline:
        return FORECAST_STRATEGY_AGGRESSIVE
    return FORECAST_STRATEGY_NORMAL


def _forecast_daily_sales(
    *,
    strategy: str,
    median_daily: float,
    recent_7_avg: float,
    recent_3_avg: float,
    trimmed_mean: float,
) -> float:
    if strategy == FORECAST_STRATEGY_CONSERVATIVE:
        return min(median_daily, recent_7_avg, trimmed_mean)
    if strategy == FORECAST_STRATEGY_AGGRESSIVE:
        forecast = (recent_7_avg * 0.7) + (recent_3_avg * 0.3)
        cap = max(median_daily * 2.0, recent_7_avg * 1.2)
        return min(forecast, cap)
    return (median_daily * 0.5) + (recent_7_avg * 0.5)
