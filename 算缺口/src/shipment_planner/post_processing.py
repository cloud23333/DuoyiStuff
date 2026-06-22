from __future__ import annotations

from collections import defaultdict

from .allocation import allocation_sort_key, pick_matching_constraint_sku
from .models import OrderLine, SalesRecord


def round_qty(value: float) -> float:
    return round(value, 4)


def line_change_ratio(line_qty: float, suggested_qty: float) -> float:
    if line_qty <= 0:
        return 0.0
    return abs(suggested_qty - line_qty) / line_qty


def decision_reason(line_qty: int, suggested_qty: int) -> str:
    if suggested_qty <= 0:
        return "hold"
    if suggested_qty >= line_qty:
        return "ship_all"
    return "ship_partial"


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


def _initialize_small_change_fields(row: dict[str, object]) -> None:
    line_qty = int(row["line_order_qty"])
    suggested_qty = int(row["recommended_ship"])
    row["recommended_ship_before_small_change_rule"] = suggested_qty
    row["small_change_ratio_before_rule"] = round_qty(
        line_change_ratio(line_qty, suggested_qty)
    )
    row["small_change_keep_warning"] = "no"


def _mark_small_change_kept(
    row: dict[str, object],
    *,
    suggested_qty: int,
    adjusted_qty: int,
    line_qty: int,
    row_number: int,
    locked_rows: set[int],
) -> None:
    if adjusted_qty <= suggested_qty:
        return

    row["recommended_ship"] = adjusted_qty
    if adjusted_qty >= line_qty:
        row["small_change_keep_warning"] = "yes"
        locked_rows.add(row_number)


def _apply_small_change_keep_by_key(
    *,
    lines: list[OrderLine],
    rows_by_number: dict[int, dict[str, object]],
    keep_change_ratio: float,
    order_totals_before_small_change: dict[str, int],
    sales_by_key: dict[tuple[str, str], SalesRecord],
    sku_order_max_qty: dict[str, int],
    order_sku_recommended_totals: dict[tuple[str, str], int],
    sku_order_limit_capped_rows: set[int],
) -> int:
    if not lines:
        return 0

    prioritized_lines = sorted(lines, key=allocation_sort_key)
    allocation_rank = {
        line.row_number: idx for idx, line in enumerate(prioritized_lines)
    }
    lines_by_number = {line.row_number: line for line in prioritized_lines}
    candidate_rows: list[int] = []
    for line in prioritized_lines:
        row = rows_by_number[line.row_number]
        suggested_qty = int(row["recommended_ship"])
        if line.quantity <= 0:
            continue
        if suggested_qty >= line.quantity:
            continue

        change_ratio = line_change_ratio(line.quantity, suggested_qty)
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
        line = lines_by_number[candidate_row_number]
        line_qty = int(candidate_row["line_order_qty"])
        suggested_qty = int(candidate_row["recommended_ship"])
        adjusted_qty = _small_change_adjusted_qty_with_sku_limit(
            line=line,
            line_qty=line_qty,
            suggested_qty=suggested_qty,
            sales_by_key=sales_by_key,
            sku_order_max_qty=sku_order_max_qty,
            order_sku_recommended_totals=order_sku_recommended_totals,
        )
        _mark_small_change_kept(
            candidate_row,
            suggested_qty=suggested_qty,
            adjusted_qty=adjusted_qty,
            line_qty=line_qty,
            row_number=candidate_row_number,
            locked_rows=locked_rows,
        )
        _track_small_change_sku_limit_result(
            line=line,
            line_qty=line_qty,
            suggested_qty=suggested_qty,
            adjusted_qty=adjusted_qty,
            sales_by_key=sales_by_key,
            sku_order_max_qty=sku_order_max_qty,
            order_sku_recommended_totals=order_sku_recommended_totals,
            sku_order_limit_capped_rows=sku_order_limit_capped_rows,
        )
        if adjusted_qty >= line_qty:
            kept_rows += 1

    return kept_rows


def apply_small_change_keep_rule(
    recommendations: list[dict[str, object]],
    *,
    order_lines: list[OrderLine],
    keep_change_ratio: float,
    sales_by_key: dict[tuple[str, str], SalesRecord] | None = None,
    sku_order_max_qty: dict[str, int] | None = None,
) -> dict[str, object]:
    rows_by_number: dict[int, dict[str, object]] = {
        int(row["row_number"]): row for row in recommendations
    }
    sales_lookup = sales_by_key or {}
    sku_limits = sku_order_max_qty or {}

    for row in recommendations:
        _initialize_small_change_fields(row)

    order_totals_before_small_change = _sum_recommended_by_order(recommendations)
    order_sku_recommended_totals = _sum_recommended_by_order_sku(
        order_lines=order_lines,
        rows_by_number=rows_by_number,
        sales_by_key=sales_lookup,
        sku_order_max_qty=sku_limits,
    )
    sku_order_limit_capped_rows: set[int] = set()

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
            sales_by_key=sales_lookup,
            sku_order_max_qty=sku_limits,
            order_sku_recommended_totals=order_sku_recommended_totals,
            sku_order_limit_capped_rows=sku_order_limit_capped_rows,
        )

    return {
        "small_change_kept_lines": kept_rows,
        "sku_order_limit_capped_rows": sku_order_limit_capped_rows,
    }


def _constraint_key_for_line(
    *,
    line: OrderLine,
    sales_by_key: dict[tuple[str, str], SalesRecord],
    sku_order_max_qty: dict[str, int],
) -> tuple[str, str] | None:
    if not sku_order_max_qty:
        return None

    sales = sales_by_key.get((line.skc, line.skuid))
    system_sku = sales.system_sku if sales is not None else ""
    constraint_sku = pick_matching_constraint_sku(
        line.order_sku,
        system_sku,
        sku_order_max_qty,
    )
    if not constraint_sku:
        return None
    return (line.internal_order_id, constraint_sku)


def _sum_recommended_by_order_sku(
    *,
    order_lines: list[OrderLine],
    rows_by_number: dict[int, dict[str, object]],
    sales_by_key: dict[tuple[str, str], SalesRecord],
    sku_order_max_qty: dict[str, int],
) -> dict[tuple[str, str], int]:
    totals: dict[tuple[str, str], int] = defaultdict(int)
    if not sku_order_max_qty:
        return totals

    for line in order_lines:
        row = rows_by_number.get(line.row_number)
        if row is None:
            continue
        constraint_key = _constraint_key_for_line(
            line=line,
            sales_by_key=sales_by_key,
            sku_order_max_qty=sku_order_max_qty,
        )
        if constraint_key is None:
            continue
        totals[constraint_key] += int(row["recommended_ship"])
    return totals


def _small_change_adjusted_qty_with_sku_limit(
    *,
    line: OrderLine,
    line_qty: int,
    suggested_qty: int,
    sales_by_key: dict[tuple[str, str], SalesRecord],
    sku_order_max_qty: dict[str, int],
    order_sku_recommended_totals: dict[tuple[str, str], int],
) -> int:
    constraint_key = _constraint_key_for_line(
        line=line,
        sales_by_key=sales_by_key,
        sku_order_max_qty=sku_order_max_qty,
    )
    if constraint_key is None:
        return line_qty

    constraint_sku = constraint_key[1]
    limit = sku_order_max_qty[constraint_sku]
    current_total = order_sku_recommended_totals.get(constraint_key, 0)
    extra_allowed = max(0, limit - current_total)
    extra_needed = max(0, line_qty - suggested_qty)
    return suggested_qty + min(extra_needed, extra_allowed)


def _track_small_change_sku_limit_result(
    *,
    line: OrderLine,
    line_qty: int,
    suggested_qty: int,
    adjusted_qty: int,
    sales_by_key: dict[tuple[str, str], SalesRecord],
    sku_order_max_qty: dict[str, int],
    order_sku_recommended_totals: dict[tuple[str, str], int],
    sku_order_limit_capped_rows: set[int],
) -> None:
    constraint_key = _constraint_key_for_line(
        line=line,
        sales_by_key=sales_by_key,
        sku_order_max_qty=sku_order_max_qty,
    )
    if constraint_key is None:
        return

    if adjusted_qty < line_qty:
        sku_order_limit_capped_rows.add(line.row_number)
    if adjusted_qty > suggested_qty:
        order_sku_recommended_totals[constraint_key] += adjusted_qty - suggested_qty


def flag_min_order_ship_qty(
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


def assign_order_intercept_warnings(
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

    for row in recommendations:
        order_id = str(row["internal_order_id"])
        is_intercepted_order = order_id in intercepted_orders
        row["order_intercept_warning"] = "yes" if is_intercepted_order else "no"

    return {
        "intercepted_orders": len(intercepted_orders),
    }


def refresh_key_recommended_totals(recommendations: list[dict[str, object]]) -> None:
    key_totals: dict[tuple[str, str], int] = defaultdict(int)
    for row in recommendations:
        key = _recommendation_key(row)
        key_totals[key] += int(row["recommended_ship"])

    for row in recommendations:
        key = _recommendation_key(row)
        row["key_recommended_total"] = key_totals.get(key, 0)


def refresh_line_decision_reasons(recommendations: list[dict[str, object]]) -> None:
    for row in recommendations:
        row["decision_reason"] = decision_reason(
            int(row["line_order_qty"]),
            int(row["recommended_ship"]),
        )


def assign_order_decision_reasons(
    recommendations: list[dict[str, object]],
    order_lines: list[OrderLine],
) -> None:
    order_qty_totals = _sum_order_qty_by_order_id(order_lines)
    order_recommended_totals = _sum_recommended_by_order(recommendations)

    for row in recommendations:
        order_id = str(row["internal_order_id"])
        row["order_decision_reason"] = decision_reason(
            order_qty_totals.get(order_id, 0),
            order_recommended_totals.get(order_id, 0),
        )
