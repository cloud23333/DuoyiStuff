from __future__ import annotations

from collections import defaultdict
from datetime import datetime
import math

from .allocation import _allocate_recommendation_quantities
from .models import KeyState, OrderLine, SalesRecord
from .parsers import normalize_sku_code
from .post_processing import (
    _apply_small_change_keep_rule,
    _assign_order_decision_reasons,
    _assign_order_intercept_warnings,
    _decision_reason,
    _flag_min_order_ship_qty,
    _line_change_ratio,
    _refresh_key_recommended_totals,
    _refresh_line_decision_reasons,
    _round_qty,
    SALES_SPIKE_WARNING_DECISION,
)
from .summary import _build_summary

DEFAULT_SOLD30_WEIGHT = 0.2
DEFAULT_SOLD7_WEIGHT = 0.8
DEFAULT_ZERO_SOLD7_WITH_SOLD30_STOCKOUT_MAX_QTY = 5
SOLD30_WINDOW_DAYS = 30.0
SOLD7_WINDOW_DAYS = 7.0
HOT_STYLE_GAP_MULTIPLIER = 1.2
DEFAULT_GLOBAL_GAP_MULTIPLIER = 1.0
SMALL_CHANGE_KEEP_RATIO = 0.3


def build_recommendations(
    order_lines: list[OrderLine],
    sales_records: list[SalesRecord],
    min_order_ship_qty: int = 10,
    zero_sold7_with_sold30_stockout_max_qty: int = (
        DEFAULT_ZERO_SOLD7_WITH_SOLD30_STOCKOUT_MAX_QTY
    ),
    sku_order_max_qty: dict[str, int] | None = None,
    exclude_skc: set[str] | None = None,
    exclude_skuid: set[str] | None = None,
    shipping_in_progress_by_key: dict[tuple[str, str], int] | None = None,
    global_gap_multiplier: float = DEFAULT_GLOBAL_GAP_MULTIPLIER,
    sold30_weight: float = DEFAULT_SOLD30_WEIGHT,
    sold7_weight: float = DEFAULT_SOLD7_WEIGHT,
) -> tuple[list[dict[str, object]], list[dict[str, object]], dict[str, object]]:
    if global_gap_multiplier <= 0:
        raise ValueError("global_gap_multiplier must be greater than 0.")
    if zero_sold7_with_sold30_stockout_max_qty < 0:
        raise ValueError(
            "zero_sold7_with_sold30_stockout_max_qty must be non-negative."
        )
    normalized_sold30_weight, normalized_sold7_weight = _normalize_sales_weights(
        sold30_weight,
        sold7_weight,
    )

    ordered_lines = sorted(order_lines, key=lambda line: line.row_number)
    normalized_sku_limits = _normalize_sku_limits(sku_order_max_qty)
    normalized_exclude_skc = _normalize_excluded_codes(exclude_skc)
    normalized_exclude_skuid = _normalize_excluded_codes(exclude_skuid)
    sales_by_key, duplicate_keys = _build_sales_lookup(sales_records)
    key_demand = _build_key_demand(ordered_lines)
    shipping_in_progress_lookup = shipping_in_progress_by_key or {}
    key_states = _build_key_states(
        key_demand=key_demand,
        sales_by_key=sales_by_key,
        shipping_in_progress_by_key=shipping_in_progress_lookup,
        global_gap_multiplier=global_gap_multiplier,
        sold30_weight=normalized_sold30_weight,
        sold7_weight=normalized_sold7_weight,
        zero_sold7_with_sold30_stockout_max_qty=(
            zero_sold7_with_sold30_stockout_max_qty
        ),
    )
    suggested_by_row, sku_order_limit_capped_lines = (
        _allocate_recommendation_quantities(
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
                "stocking_days": _round_qty(stocking_days),
                "wh": _round_qty(stock_in_warehouse),
                "pending_recv": _round_qty(pending_receive),
                "pending_ship": _round_qty(pending_ship),
                "shipping_in_progress": shipping_in_progress,
                "gap": gap,
                "recommended_ship": suggested_qty,
                "recommended_ship_before_small_change_rule": suggested_qty,
                "small_change_ratio_before_rule": _round_qty(
                    _line_change_ratio(line.quantity, suggested_qty)
                ),
                "small_change_keep_warning": "no",
                "key_recommended_total": key_recommended_total,
                "decision_reason": _decision_reason(line.quantity, suggested_qty),
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

    intercept_stats = _assign_order_intercept_warnings(
        recommendations,
        suggested_by_row_before_intercept=suggested_by_row_before_intercept,
    )
    small_change_stats = _apply_small_change_keep_rule(
        recommendations,
        order_lines=ordered_lines,
        keep_change_ratio=SMALL_CHANGE_KEEP_RATIO,
    )
    threshold_stats = _flag_min_order_ship_qty(recommendations, min_order_ship_qty)
    _refresh_key_recommended_totals(recommendations)
    _refresh_line_decision_reasons(recommendations)
    _assign_order_decision_reasons(recommendations, ordered_lines)
    summary = _build_summary(
        ordered_lines,
        sales_records,
        recommendations,
        quality_rows,
        duplicate_keys,
        min_order_ship_qty=min_order_ship_qty,
        threshold_stats=threshold_stats,
        sku_order_limit_rule_count=len(normalized_sku_limits),
        sku_order_limit_capped_lines=sku_order_limit_capped_lines,
        excluded_skc_rule_count=len(normalized_exclude_skc),
        excluded_skuid_rule_count=len(normalized_exclude_skuid),
        intercepted_order_lines=intercepted_order_lines,
        intercepted_orders=intercept_stats.get("intercepted_orders", 0),
        small_change_kept_lines=small_change_stats.get("small_change_kept_lines", 0),
        global_gap_multiplier=global_gap_multiplier,
        sold30_weight=normalized_sold30_weight,
        sold7_weight=normalized_sold7_weight,
        zero_sold7_with_sold30_stockout_max_qty=(
            zero_sold7_with_sold30_stockout_max_qty
        ),
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


def _normalize_sales_weights(
    sold30_weight: float,
    sold7_weight: float,
) -> tuple[float, float]:
    if sold30_weight < 0 or sold7_weight < 0:
        raise ValueError("sold30_weight and sold7_weight must be non-negative.")
    weight_total = sold30_weight + sold7_weight
    if weight_total <= 0:
        raise ValueError("sold30_weight and sold7_weight cannot both be 0.")
    return sold30_weight / weight_total, sold7_weight / weight_total


def _build_intercept_reason_by_row(
    order_lines: list[OrderLine],
    *,
    exclude_skc: set[str],
    exclude_skuid: set[str],
) -> dict[int, str]:
    reasons: dict[int, str] = {}
    for line in order_lines:
        skc_hit = line.skc in exclude_skc
        skuid_hit = line.skuid in exclude_skuid
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
    shipping_in_progress_by_key: dict[tuple[str, str], int],
    global_gap_multiplier: float,
    sold30_weight: float,
    sold7_weight: float,
    zero_sold7_with_sold30_stockout_max_qty: int,
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
                min_order_ship_qty_exempt_eligible=False,
            )
            continue

        shipping_in_progress = shipping_in_progress_by_key.get(key, 0)
        target_ship_qty = _target_ship_qty(
            sold30=sales.sold30,
            sold7=sales.sold7,
            stocking_days=sales.stocking_days,
            sold30_weight=sold30_weight,
            sold7_weight=sold7_weight,
        )
        available_stock = (
            sales.stock_in_warehouse + sales.pending_receive + shipping_in_progress
        )
        raw_gap = max(
            0.0,
            target_ship_qty - available_stock,
        )
        if sales.is_hot_style:
            raw_gap *= HOT_STYLE_GAP_MULTIPLIER
        raw_gap *= global_gap_multiplier
        gap = math.ceil(raw_gap)
        recommended_qty_total = min(order_qty_total, gap)
        is_min_order_ship_qty_exempt_eligible = (
            sales.sold7 == 0 and sales.sold30 > 0 and available_stock == 0
        )
        if is_min_order_ship_qty_exempt_eligible:
            recommended_qty_total = min(
                order_qty_total,
                zero_sold7_with_sold30_stockout_max_qty,
            )
        states[key] = KeyState(
            skc=skc,
            skuid=skuid,
            system_sku=sales.system_sku,
            order_qty_total=order_qty_total,
            gap=gap,
            recommended_qty_total=recommended_qty_total,
            min_order_ship_qty_exempt_eligible=is_min_order_ship_qty_exempt_eligible,
        )
    return states


def _target_ship_qty(
    sold30: int,
    sold7: int,
    stocking_days: float,
    sold30_weight: float,
    sold7_weight: float,
) -> float:
    sold30_daily = (sold30_weight * sold30) / SOLD30_WINDOW_DAYS
    sold7_daily = (sold7_weight * sold7) / SOLD7_WINDOW_DAYS
    return (sold30_daily + sold7_daily) * stocking_days



