from __future__ import annotations

from collections.abc import Callable
import csv
import json
from pathlib import Path

RECOMMENDATION_FIELDS = [
    ("internal_order_id", "内部订单号"),
    ("店铺款式编码", "店铺款式编码"),
    ("店铺商品编码", "店铺商品编码"),
    ("原始商品编码", "原始商品编码"),
    ("系统商品编码", "系统商品编码"),
    ("sku_code_check", "SKU编码校验"),
    ("line_order_qty", "订单行数量"),
    ("key_order_qty", "同SKC_SKUID总下单量"),
    ("sold30", "近30日销量"),
    ("sold7", "近7日销量"),
    ("stocking_days", "备货逻辑天数"),
    ("wh", "平台仓内库存"),
    ("pending_ship", "平台待发货库存"),
    ("shipping_in_progress", "发货中数量"),
    ("pending_recv", "平台待收货库存"),
    ("gap", "缺口"),
    ("key_recommended_total", "同SKC_SKUID建议总量"),
    ("recommended_ship_before_small_change_rule", "30%规则前建议发货量"),
    ("small_change_ratio_before_rule", "30%规则前变动比例"),
    ("small_change_keep_warning", "30%内免改数量提示"),
    ("recommended_ship", "建议发货量"),
    ("decision_reason", "SKU建议类型"),
    ("order_decision_reason", "订单建议类型"),
    ("intercept_reason", "拦截原因"),
    ("order_intercept_warning", "订单拦截导致不发提示"),
    ("order_recommended_ship_total_before_threshold", "订单阈值前建议总量"),
    ("min_order_ship_qty_threshold", "最小发货阈值"),
    ("order_low_qty_warning", "订单低于起发量提示"),
    ("min_order_ship_qty_exempt_warning", "小于10不发豁免资格提示"),
    ("min_order_ship_qty_exempt_applied_warning", "小于10不发豁免生效提示"),
]

QUALITY_FIELDS = [
    ("type", "问题类型"),
    ("internal_order_id", "内部订单号"),
    ("skc", "店铺款式编码"),
    ("skuid", "店铺商品编码"),
    ("order_sku", "原始商品编码"),
    ("system_sku", "系统商品编码"),
    ("message", "问题说明"),
]

DECISION_REASON_MAP = {
    "ship_all": "全发",
    "ship_partial": "部分发",
    "hold": "暂不发",
    "sales_spike_warning": "销量突增预警",
}

INTERCEPT_REASON_MAP = {
    "skc": "命中SKC拦截",
    "skuid": "命中SKUID拦截",
    "skc_and_skuid": "命中SKC和SKUID拦截",
}

SKU_CHECK_MAP = {
    "exact_match": "完全一致",
    "normalized_match": "标准化一致",
    "diff": "不一致",
    "missing_key": "缺少销售匹配",
}

QUALITY_TYPE_MAP = {
    "sku_code_diff": "SKU编码不一致",
    "missing_sales_key": "缺少销售匹配",
}

QUALITY_MESSAGE_MAP = {
    "Order SKU and system SKU differ after normalization": "原始商品编码与系统商品编码在标准化后仍不一致",
    "No sales row found for (SKC, SKUID)": "未找到对应销售键 (SKC, SKUID)",
}

INT_FORMAT_FIELDS = {
    "line_order_qty",
    "key_order_qty",
    "sold30",
    "sold7",
    "gap",
    "key_recommended_total",
    "recommended_ship_before_small_change_rule",
    "recommended_ship",
    "order_recommended_ship_total_before_threshold",
    "min_order_ship_qty_threshold",
}

SUMMARY_INT_FORMAT_FIELDS = {
    "order_lines",
    "sales_rows",
    "matched_order_lines",
    "total_order_qty",
    "total_recommended_qty",
    "small_change_kept_lines",
    "quality_issue_rows",
    "duplicate_sales_keys",
    "min_order_ship_qty_threshold",
    "zero_sold7_with_sold30_stockout_max_qty",
    "low_qty_orders_before_exempt",
    "low_qty_order_lines_before_exempt",
    "low_qty_orders",
    "low_qty_order_lines",
    "low_qty_orders_exempted",
    "low_qty_order_lines_exempted",
    "sku_order_limit_rule_count",
    "sku_order_limit_capped_lines",
    "excluded_skc_rule_count",
    "excluded_skuid_rule_count",
    "intercepted_order_lines",
    "intercepted_orders",
}

DECISION_FIELDS = {"decision_reason", "order_decision_reason"}

WARNING_FIELDS = {
    "order_low_qty_warning",
    "min_order_ship_qty_exempt_warning",
    "min_order_ship_qty_exempt_applied_warning",
    "order_intercept_warning",
    "small_change_keep_warning",
}

SUMMARY_FIELDS = [
    ("order_lines", "订单行数"),
    ("sales_rows", "销售行数"),
    ("matched_order_lines", "匹配订单行数"),
    ("join_coverage_pct", "匹配覆盖率_百分比"),
    ("total_order_qty", "总下单量"),
    ("total_recommended_qty", "建议发货总量"),
    ("small_change_kept_lines", "触发30%免改行数"),
    ("decision_ship_all", "建议_全发_行数"),
    ("decision_ship_partial", "建议_部分发_行数"),
    ("decision_hold", "建议_暂不发_行数"),
    ("sku_check_exact_match", "SKU校验_完全一致_行数"),
    ("sku_check_normalized_match", "SKU校验_标准化一致_行数"),
    ("sku_check_diff", "SKU校验_不一致_行数"),
    ("sku_check_missing_key", "SKU校验_缺少销售匹配_行数"),
    ("quality_issue_rows", "质量问题行数"),
    ("duplicate_sales_keys", "销售重复键数量"),
    ("global_gap_multiplier", "全局缺口上浮系数"),
    ("sold30_weight", "近30日销量占比"),
    ("sold7_weight", "近7日销量占比"),
    ("zero_sold7_with_sold30_stockout_max_qty", "零7日销量且30日有销量无库存发货上限"),
    ("min_order_ship_qty_threshold", "最小发货阈值"),
    ("low_qty_orders_before_exempt", "阈值前低发货量订单数"),
    ("low_qty_order_lines_before_exempt", "阈值前低发货量订单行数"),
    ("low_qty_orders", "低于阈值订单数_提示"),
    ("low_qty_order_lines", "低于阈值订单行数_提示"),
    ("low_qty_orders_exempted", "低于阈值豁免订单数"),
    ("low_qty_order_lines_exempted", "低于阈值豁免订单行数"),
    ("sku_order_limit_rule_count", "订单内SKU限额规则数"),
    ("sku_order_limit_capped_lines", "触发订单内SKU限额行数"),
    ("excluded_skc_rule_count", "SKC拦截规则数"),
    ("excluded_skuid_rule_count", "SKUID拦截规则数"),
    ("intercepted_order_lines", "命中拦截订单行数"),
    ("intercepted_orders", "拦截导致不发订单数"),
]


def export_reports(
    out_dir: str | Path,
    recommendations: list[dict[str, object]],
    quality_rows: list[dict[str, object]],
    summary: dict[str, object],
) -> dict[str, Path]:
    output_dir = Path(out_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    recommendation_path = output_dir / "发货建议明细.csv"
    quality_path = output_dir / "数据质量报告.csv"
    summary_path = output_dir / "运行摘要.json"

    recommendation_columns = [target for _, target in RECOMMENDATION_FIELDS]
    quality_columns = [target for _, target in QUALITY_FIELDS]

    localized_recommendations = [_localize_recommendation_row(row) for row in recommendations]
    localized_quality_rows = [_localize_quality_row(row) for row in quality_rows]
    localized_summary = _localize_summary(summary)

    _write_csv(recommendation_path, localized_recommendations, recommendation_columns)
    _write_csv(quality_path, localized_quality_rows, quality_columns)
    summary_path.write_text(json.dumps(localized_summary, ensure_ascii=False, indent=2), encoding="utf-8")

    return {
        "final_recommendation": recommendation_path,
        "quality_report": quality_path,
        "run_summary": summary_path,
    }


def _write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        if rows:
            writer.writerows(rows)


def _localize_recommendation_row(row: dict[str, object]) -> dict[str, object]:
    return _localize_row(
        row=row,
        fields=RECOMMENDATION_FIELDS,
        default_value="",
        value_mapper=_localize_recommendation_value,
    )


def _localize_quality_row(row: dict[str, object]) -> dict[str, object]:
    return _localize_row(
        row=row,
        fields=QUALITY_FIELDS,
        default_value="",
        value_mapper=_localize_quality_value,
    )


def _localize_row(
    *,
    row: dict[str, object],
    fields: list[tuple[str, str]],
    default_value: object,
    value_mapper: Callable[[str, object], object],
) -> dict[str, object]:
    localized: dict[str, object] = {}
    for source, target in fields:
        localized[target] = value_mapper(source, row.get(source, default_value))
    return localized


def _localize_summary(summary: dict[str, object]) -> dict[str, object]:
    localized: dict[str, object] = {}
    for source, target in SUMMARY_FIELDS:
        value = summary.get(source, 0)
        if source in SUMMARY_INT_FORMAT_FIELDS:
            localized[target] = _format_int_like(value)
            continue
        localized[target] = value
    return localized


def _format_int_like(value: object) -> object:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _localize_recommendation_value(source: str, value: object) -> object:
    if source in INT_FORMAT_FIELDS:
        return _format_int_like(value)
    if source in DECISION_FIELDS:
        return DECISION_REASON_MAP.get(str(value), value)
    if source == "intercept_reason":
        return INTERCEPT_REASON_MAP.get(str(value), value)
    if source == "sku_code_check":
        return SKU_CHECK_MAP.get(str(value), value)
    if source in WARNING_FIELDS:
        return "是" if str(value) == "yes" else "否"
    return value


def _localize_quality_value(source: str, value: object) -> object:
    if source == "type":
        return QUALITY_TYPE_MAP.get(str(value), value)
    if source == "message":
        return QUALITY_MESSAGE_MAP.get(str(value), value)
    return value
