from __future__ import annotations

from collections import defaultdict
from datetime import datetime

from .models import KeyState, OrderLine, SalesRecord
from .parsers import SHORTAGE_STATUS, normalize_sku_code


def _allocate_recommendation_quantities(
    order_lines: list[OrderLine],
    key_states: dict[tuple[str, str], KeyState],
    sales_by_key: dict[tuple[str, str], SalesRecord],
    sku_order_max_qty: dict[str, int],
) -> tuple[dict[int, int], int]:
    suggested_by_row: dict[int, int] = {}
    order_sku_shipped_totals: dict[tuple[str, str], int] = defaultdict(int)
    capped_lines = 0

    grouped_lines: dict[tuple[str, str], list[OrderLine]] = defaultdict(list)
    for line in order_lines:
        grouped_lines[(line.skc, line.skuid)].append(line)

    for key, lines in grouped_lines.items():
        state = key_states.get(key)
        if state is None:
            for line in lines:
                suggested_by_row[line.row_number] = 0
            continue

        remaining = state.recommended_qty_total
        sales = sales_by_key.get(key)
        system_sku = sales.system_sku if sales is not None else ""

        for line in sorted(lines, key=_allocation_sort_key):
            base_suggested_qty = min(line.quantity, remaining)
            suggested_qty, was_capped = _apply_order_sku_limit(
                suggested_qty=base_suggested_qty,
                line=line,
                system_sku=system_sku,
                sku_order_max_qty=sku_order_max_qty,
                order_sku_shipped_totals=order_sku_shipped_totals,
            )
            suggested_by_row[line.row_number] = suggested_qty
            if was_capped:
                capped_lines += 1
            remaining -= suggested_qty

    return suggested_by_row, capped_lines


def _allocation_sort_key(line: OrderLine) -> tuple[int, datetime, int]:
    # Reduce "缺货" first when supply is insufficient.
    status_priority = 1 if line.status == SHORTAGE_STATUS else 0
    return (status_priority, line.order_time, line.row_number)


def _apply_order_sku_limit(
    suggested_qty: int,
    line: OrderLine,
    system_sku: str,
    sku_order_max_qty: dict[str, int],
    order_sku_shipped_totals: dict[tuple[str, str], int],
) -> tuple[int, bool]:
    constraint_sku = _pick_matching_constraint_sku(
        line.order_sku, system_sku, sku_order_max_qty
    )
    if not constraint_sku:
        return suggested_qty, False

    limit = sku_order_max_qty[constraint_sku]
    order_key = (line.internal_order_id, constraint_sku)
    already_suggested = order_sku_shipped_totals.get(order_key, 0)
    remaining_allowed = max(0, limit - already_suggested)
    capped_qty = min(suggested_qty, remaining_allowed)
    order_sku_shipped_totals[order_key] = already_suggested + capped_qty
    was_capped = capped_qty < suggested_qty
    return capped_qty, was_capped


def _pick_matching_constraint_sku(
    order_sku: str,
    system_sku: str,
    sku_order_max_qty: dict[str, int],
) -> str:
    for raw_sku in (order_sku, system_sku):
        normalized_sku = normalize_sku_code(raw_sku)
        if normalized_sku in sku_order_max_qty:
            return normalized_sku
    return ""
