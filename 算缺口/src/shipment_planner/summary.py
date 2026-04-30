from __future__ import annotations

from collections import Counter

from .models import OrderLine, SalesRecord
from .post_processing import round_qty

DEMAND_PROFILE_ORDER = ("稳定款", "波动款", "慢销/间歇款", "无销量款")
FORECAST_MODEL_ORDER = ("tsb", "imapa", "current", "mean", "croston_sba", "adida")


def build_summary(
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
) -> dict[str, object]:
    order_line_count = len(order_lines)
    matched_count = sum(
        1 for row in recommendations if row["sku_code_check"] != "missing_key"
    )
    decision_counter = Counter(str(row["decision_reason"]) for row in recommendations)
    sku_check_counter = Counter(str(row["sku_code_check"]) for row in recommendations)
    forecast_rows = [
        row for row in recommendations if str(row.get("forecast_strategy", "")).strip()
    ]
    join_coverage_pct = (
        (matched_count / order_line_count * 100) if order_line_count else 0.0
    )
    total_order_qty = sum(line.quantity for line in order_lines)
    total_recommended_qty = sum(int(row["recommended_ship"]) for row in recommendations)

    return {
        "order_lines": order_line_count,
        "sales_rows": len(sales_records),
        "matched_order_lines": matched_count,
        "join_coverage_pct": round_qty(join_coverage_pct),
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
        "global_gap_multiplier": round_qty(global_gap_multiplier),
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
        "demand_profile_summary": _format_counter(
            _count_values(forecast_rows, "demand_profile"),
            ordered=DEMAND_PROFILE_ORDER,
        ),
        "anomaly_flag_summary": _format_counter(
            _count_anomaly_flags(forecast_rows),
        ),
        "service_level_summary": _format_counter(
            _count_service_levels(forecast_rows),
            ordered=("P85", "P75", "P70", "P65", "P60", "P50"),
        ),
        "forecast_model_summary": _format_counter(
            _count_values(forecast_rows, "forecast_model"),
            ordered=FORECAST_MODEL_ORDER,
        ),
    }


def _count_values(rows: list[dict[str, object]], key: str) -> Counter[str]:
    counter: Counter[str] = Counter()
    for row in rows:
        value = str(row.get(key, "")).strip()
        if value:
            counter[value] += 1
    return counter


def _count_anomaly_flags(rows: list[dict[str, object]]) -> Counter[str]:
    counter: Counter[str] = Counter()
    for row in rows:
        raw_flags = str(row.get("anomaly_flags", "")).strip()
        if not raw_flags or raw_flags == "无":
            continue
        for flag in raw_flags.split("、"):
            clean_flag = flag.strip()
            if clean_flag and clean_flag != "无":
                counter[clean_flag] += 1
    return counter


def _count_service_levels(rows: list[dict[str, object]]) -> Counter[str]:
    counter: Counter[str] = Counter()
    for row in rows:
        label = _service_level_label(row.get("service_level"))
        if label:
            counter[label] += 1
    return counter


def _service_level_label(value: object) -> str:
    if isinstance(value, bool):
        return ""
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return ""
    if numeric <= 0:
        return ""
    return f"P{int(round(numeric * 100))}"


def _format_counter(
    counter: Counter[str],
    *,
    ordered: tuple[str, ...] = (),
) -> str:
    if not counter:
        return "无"

    parts: list[str] = []
    for key in ordered:
        count = counter.pop(key, 0)
        if count:
            parts.append(f"{key} {count}")
    for key, count in sorted(counter.items()):
        parts.append(f"{key} {count}")
    return "，".join(parts)
