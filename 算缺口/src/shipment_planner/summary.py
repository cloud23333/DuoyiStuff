from __future__ import annotations

from collections import Counter

from .models import OrderLine, SalesRecord
from .post_processing import _round_qty


def _build_summary(
    order_lines: list[OrderLine],
    sales_records: list[SalesRecord],
    recommendations: list[dict[str, object]],
    quality_rows: list[dict[str, object]],
    duplicate_keys: set[tuple[str, str]],
    min_order_ship_qty: int,
    threshold_stats: dict[str, int],
    sku_order_limit_rule_count: int,
    sku_order_limit_capped_lines: int,
    excluded_skc_rule_count: int,
    excluded_skuid_rule_count: int,
    intercepted_order_lines: int,
    intercepted_orders: int,
    small_change_kept_lines: int,
    global_gap_multiplier: float,
    sold30_weight: float,
    sold7_weight: float,
    zero_sold7_with_sold30_stockout_max_qty: int,
) -> dict[str, object]:
    order_line_count = len(order_lines)
    matched_count = sum(
        1 for row in recommendations if row["sku_code_check"] != "missing_key"
    )
    decision_counter = Counter(str(row["decision_reason"]) for row in recommendations)
    sku_check_counter = Counter(str(row["sku_code_check"]) for row in recommendations)
    join_coverage_pct = (
        (matched_count / order_line_count * 100) if order_line_count else 0.0
    )
    total_order_qty = sum(line.quantity for line in order_lines)
    total_recommended_qty = sum(int(row["recommended_ship"]) for row in recommendations)

    return {
        "order_lines": order_line_count,
        "sales_rows": len(sales_records),
        "matched_order_lines": matched_count,
        "join_coverage_pct": _round_qty(join_coverage_pct),
        "total_order_qty": total_order_qty,
        "total_recommended_qty": total_recommended_qty,
        "decision_ship_all": decision_counter.get("ship_all", 0),
        "decision_ship_partial": decision_counter.get("ship_partial", 0),
        "decision_hold": decision_counter.get("hold", 0),
        "sku_check_exact_match": sku_check_counter.get("exact_match", 0),
        "sku_check_normalized_match": sku_check_counter.get("normalized_match", 0),
        "sku_check_diff": sku_check_counter.get("diff", 0),
        "sku_check_missing_key": sku_check_counter.get("missing_key", 0),
        "quality_issue_rows": len(quality_rows),
        "duplicate_sales_keys": len(duplicate_keys),
        "min_order_ship_qty_threshold": min_order_ship_qty,
        "low_qty_orders": threshold_stats.get("low_qty_orders", 0),
        "low_qty_order_lines": threshold_stats.get("low_qty_order_lines", 0),
        "sku_order_limit_rule_count": sku_order_limit_rule_count,
        "sku_order_limit_capped_lines": sku_order_limit_capped_lines,
        "excluded_skc_rule_count": excluded_skc_rule_count,
        "excluded_skuid_rule_count": excluded_skuid_rule_count,
        "intercepted_order_lines": intercepted_order_lines,
        "intercepted_orders": intercepted_orders,
        "small_change_kept_lines": small_change_kept_lines,
        "global_gap_multiplier": _round_qty(global_gap_multiplier),
        "sold30_weight": _round_qty(sold30_weight),
        "sold7_weight": _round_qty(sold7_weight),
        "zero_sold7_with_sold30_stockout_max_qty": (
            zero_sold7_with_sold30_stockout_max_qty
        ),
        "low_qty_orders_before_exempt": threshold_stats.get(
            "low_qty_orders_before_exempt",
            0,
        ),
        "low_qty_order_lines_before_exempt": threshold_stats.get(
            "low_qty_order_lines_before_exempt",
            0,
        ),
        "low_qty_orders_exempted": threshold_stats.get("low_qty_orders_exempted", 0),
        "low_qty_order_lines_exempted": threshold_stats.get(
            "low_qty_order_lines_exempted",
            0,
        ),
    }
