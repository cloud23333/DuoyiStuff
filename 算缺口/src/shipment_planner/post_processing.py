from __future__ import annotations

from collections import defaultdict

from .allocation import _allocation_sort_key
from .models import OrderLine

SALES_SPIKE_MIN_SOLD30 = 30
SALES_SPIKE_RATIO_THRESHOLD = 0.9
SALES_SPIKE_WARNING_DECISION = "sales_spike_warning"


def _round_qty(value: float) -> float:
    return round(value, 4)


def _line_change_ratio(line_qty: float, suggested_qty: float) -> float:
    if line_qty <= 0:
        return 0.0
    return abs(suggested_qty - line_qty) / line_qty


def _decision_reason(line_qty: int, suggested_qty: int) -> str:
    if suggested_qty <= 0:
        return "hold"
    if suggested_qty >= line_qty:
        return "ship_all"
    return "ship_partial"


def _is_sales_spike_warning(sold30: int, sold7: int) -> bool:
    if sold30 < SALES_SPIKE_MIN_SOLD30:
        return False
    if sold30 <= 0:
        return False
    return (sold7 / sold30) >= SALES_SPIKE_RATIO_THRESHOLD


def _recommendation_key(row: dict[str, object]) -> tuple[str, str]:
    return (str(row["店铺款式编码"]), str(row["店铺商品编码"]))


def _sum_order_qty_by_order_id(order_lines: list[OrderLine]) -> dict[str, int]:
    totals: dict[str, int] = defaultdict(int)
    for line in order_lines:
        totals[line.internal_order_id] += line.quantity
    return dict(totals)


def _sum_recommended_by_order(
    recommendations: list[dict[str, object]],
) -> dict[str, int]:
    totals: dict[str, int] = defaultdict(int)
    for row in recommendations:
        totals[str(row["internal_order_id"])] += int(row["recommended_ship"])
    return dict(totals)


def _order_all_sales_spike_warning(
    recommendations: list[dict[str, object]],
) -> dict[str, bool]:
    order_all_warning: dict[str, bool] = {}
    for row in recommendations:
        order_id = str(row["internal_order_id"])
        is_warning = (
            str(row.get("decision_reason", "")) == SALES_SPIKE_WARNING_DECISION
        )
        if order_id not in order_all_warning:
            order_all_warning[order_id] = is_warning
            continue
        order_all_warning[order_id] = order_all_warning[order_id] and is_warning
    return order_all_warning


def _initialize_small_change_fields(row: dict[str, object]) -> None:
    line_qty = int(row["line_order_qty"])
    suggested_qty = int(row["recommended_ship"])
    row["recommended_ship_before_small_change_rule"] = suggested_qty
    row["small_change_ratio_before_rule"] = _round_qty(
        _line_change_ratio(line_qty, suggested_qty)
    )
    row["small_change_keep_warning"] = "no"


def _mark_small_change_kept(
    row: dict[str, object],
    *,
    line_qty: int,
    row_number: int,
    locked_rows: set[int],
) -> None:
    row["recommended_ship"] = line_qty
    row["small_change_keep_warning"] = "yes"
    locked_rows.add(row_number)


def _apply_small_change_keep_by_key(
    *,
    lines: list[OrderLine],
    rows_by_number: dict[int, dict[str, object]],
    keep_change_ratio: float,
    order_totals_before_small_change: dict[str, int],
) -> int:
    if not lines:
        return 0

    prioritized_lines = sorted(lines, key=_allocation_sort_key)
    allocation_rank = {
        line.row_number: idx for idx, line in enumerate(prioritized_lines)
    }
    candidate_rows: list[int] = []
    for line in prioritized_lines:
        row = rows_by_number[line.row_number]
        suggested_qty = int(row["recommended_ship"])
        if line.quantity <= 0:
            continue
        if suggested_qty >= line.quantity:
            continue

        change_ratio = _line_change_ratio(line.quantity, suggested_qty)
        if change_ratio <= keep_change_ratio:
            candidate_rows.append(line.row_number)

    if not candidate_rows:
        return 0

    # Keep deterministic trigger order for reporting consistency.
    candidate_rows.sort(
        key=lambda row_number: (
            -order_totals_before_small_change.get(
                str(rows_by_number[row_number]["internal_order_id"]),
                0,
            ),
            allocation_rank.get(row_number, 0),
            row_number,
        )
    )
    locked_rows: set[int] = set()
    kept_rows = 0

    for candidate_row_number in candidate_rows:
        candidate_row = rows_by_number[candidate_row_number]
        line_qty = int(candidate_row["line_order_qty"])
        _mark_small_change_kept(
            candidate_row,
            line_qty=line_qty,
            row_number=candidate_row_number,
            locked_rows=locked_rows,
        )
        kept_rows += 1

    return kept_rows


def _apply_small_change_keep_rule(
    recommendations: list[dict[str, object]],
    *,
    order_lines: list[OrderLine],
    keep_change_ratio: float,
) -> dict[str, int]:
    rows_by_number: dict[int, dict[str, object]] = {
        int(row["row_number"]): row for row in recommendations
    }

    for row in recommendations:
        _initialize_small_change_fields(row)

    order_totals_before_small_change = _sum_recommended_by_order(recommendations)

    grouped_lines: dict[tuple[str, str], list[OrderLine]] = defaultdict(list)
    for line in order_lines:
        if line.row_number in rows_by_number:
            grouped_lines[(line.skc, line.skuid)].append(line)

    kept_rows = 0
    for lines in grouped_lines.values():
        kept_rows += _apply_small_change_keep_by_key(
            lines=lines,
            rows_by_number=rows_by_number,
            keep_change_ratio=keep_change_ratio,
            order_totals_before_small_change=order_totals_before_small_change,
        )

    return {"small_change_kept_lines": kept_rows}


def _flag_min_order_ship_qty(
    recommendations: list[dict[str, object]],
    min_order_ship_qty: int,
) -> dict[str, int]:
    order_totals = _sum_recommended_by_order(recommendations)
    low_qty_orders: set[str] = set()
    if min_order_ship_qty > 0:
        low_qty_orders = {
            order_id
            for order_id, total in order_totals.items()
            if 0 < total < min_order_ship_qty
        }

    flagged_lines = 0
    affected_orders: set[str] = set()
    exempted_lines = 0
    exempted_orders: set[str] = set()
    low_qty_lines_before_exempt = 0
    for row in recommendations:
        order_id = str(row["internal_order_id"])
        before_total = order_totals.get(order_id, 0)
        row["order_recommended_ship_total_before_threshold"] = before_total
        row["min_order_ship_qty_threshold"] = min_order_ship_qty
        is_low_qty_order = order_id in low_qty_orders
        if is_low_qty_order:
            low_qty_lines_before_exempt += 1

        is_min_order_ship_qty_exempt_eligible = bool(
            row.get("min_order_ship_qty_exempt_eligible", False)
        )
        should_block_by_threshold = (
            is_low_qty_order and not is_min_order_ship_qty_exempt_eligible
        )
        should_apply_exemption = (
            is_low_qty_order and is_min_order_ship_qty_exempt_eligible
        )
        row["order_low_qty_warning"] = "yes" if should_block_by_threshold else "no"
        row["min_order_ship_qty_exempt_eligible"] = (
            is_min_order_ship_qty_exempt_eligible
        )
        row["min_order_ship_qty_exempt_applied"] = should_apply_exemption
        row["min_order_ship_qty_exempt_warning"] = (
            "yes" if is_min_order_ship_qty_exempt_eligible else "no"
        )
        row["min_order_ship_qty_exempt_applied_warning"] = (
            "yes" if should_apply_exemption else "no"
        )

        if should_apply_exemption:
            exempted_lines += 1
            exempted_orders.add(order_id)
            continue
        if not should_block_by_threshold:
            continue
        row["recommended_ship"] = 0
        flagged_lines += 1
        affected_orders.add(order_id)

    return {
        "low_qty_orders_before_exempt": len(low_qty_orders),
        "low_qty_order_lines_before_exempt": low_qty_lines_before_exempt,
        "low_qty_orders": len(affected_orders),
        "low_qty_order_lines": flagged_lines,
        "low_qty_orders_exempted": len(exempted_orders),
        "low_qty_order_lines_exempted": exempted_lines,
    }


def _assign_order_intercept_warnings(
    recommendations: list[dict[str, object]],
    *,
    suggested_by_row_before_intercept: dict[int, int],
) -> dict[str, int]:
    order_totals_after: dict[str, int] = defaultdict(int)
    order_totals_before: dict[str, int] = defaultdict(int)
    order_has_intercept: dict[str, bool] = defaultdict(bool)
    for row in recommendations:
        order_id = str(row["internal_order_id"])
        row_number = int(row["row_number"])
        order_totals_after[order_id] += int(row["recommended_ship"])
        order_totals_before[order_id] += suggested_by_row_before_intercept.get(
            row_number, 0
        )
        if str(row.get("intercept_reason", "")):
            order_has_intercept[order_id] = True

    intercepted_orders = {
        order_id
        for order_id, total_after in order_totals_after.items()
        if total_after <= 0
        and order_totals_before.get(order_id, 0) > 0
        and order_has_intercept.get(order_id, False)
    }

    flagged_lines = 0
    for row in recommendations:
        order_id = str(row["internal_order_id"])
        is_intercepted_order = order_id in intercepted_orders
        row["order_intercept_warning"] = "yes" if is_intercepted_order else "no"
        if is_intercepted_order:
            flagged_lines += 1

    return {
        "intercepted_orders": len(intercepted_orders),
        "intercepted_order_lines": flagged_lines,
    }


def _refresh_key_recommended_totals(recommendations: list[dict[str, object]]) -> None:
    key_totals: dict[tuple[str, str], int] = defaultdict(int)
    for row in recommendations:
        key = _recommendation_key(row)
        key_totals[key] += int(row["recommended_ship"])

    for row in recommendations:
        key = _recommendation_key(row)
        row["key_recommended_total"] = key_totals.get(key, 0)


def _refresh_line_decision_reasons(recommendations: list[dict[str, object]]) -> None:
    for row in recommendations:
        base_decision = _decision_reason(
            int(row["line_order_qty"]),
            int(row["recommended_ship"]),
        )
        if _is_sales_spike_warning(int(row["sold30"]), int(row["sold7"])):
            row["decision_reason"] = SALES_SPIKE_WARNING_DECISION
            continue
        row["decision_reason"] = base_decision


def _assign_order_decision_reasons(
    recommendations: list[dict[str, object]],
    order_lines: list[OrderLine],
) -> None:
    order_qty_totals = _sum_order_qty_by_order_id(order_lines)
    order_recommended_totals = _sum_recommended_by_order(recommendations)
    order_all_sales_spike_warning = _order_all_sales_spike_warning(recommendations)

    for row in recommendations:
        order_id = str(row["internal_order_id"])
        base_order_decision = _decision_reason(
            order_qty_totals.get(order_id, 0),
            order_recommended_totals.get(order_id, 0),
        )
        if order_all_sales_spike_warning.get(order_id, False):
            row["order_decision_reason"] = SALES_SPIKE_WARNING_DECISION
            continue
        row["order_decision_reason"] = base_order_decision
