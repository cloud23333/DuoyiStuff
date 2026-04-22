from __future__ import annotations

import io
import csv
import json
from collections.abc import Callable, Sequence
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XlsxImage
from openpyxl.utils import get_column_letter

from .engine import stockout_mask_cut
from .eval_forecast import EvalRow
from .forecast_curve import build_forecast_curve

RECOMMENDATION_FIELDS = [
    ("internal_order_id", "内部订单号"),
    ("店铺款式编码", "店铺款式编码"),
    ("店铺商品编码", "店铺商品编码"),
    ("原始商品编码", "原始商品编码"),
    ("系统商品编码", "系统商品编码"),
    ("sku_code_check", "SKU编码校验"),
    ("line_order_qty", "订单行数量"),
    ("key_order_qty", "同SKC_SKUID总下单量"),
    ("forecast_strategy", "预测策略"),
    ("forecast_daily_sales", "预测日均销量"),
    ("forecast_stocking_period_sales", "预测备货期销量"),
    ("sku_forecast_abs_error", "SKU预测绝对误差"),
    ("sku_forecast_signed_error", "SKU预测偏差"),
    ("__plot__", "销量与预测曲线"),
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

PLOT_COLUMN_NAME = "销量与预测曲线"
PLOT_COLUMN_WIDTH_CHARS = 46.0
PLOT_ROW_HEIGHT_POINTS = 100.0
PLOT_IMAGE_WIDTH_PX = 320
PLOT_IMAGE_HEIGHT_PX = 128

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
    "missing_daily_sales": "缺少每日销量数据",
}

QUALITY_MESSAGE_MAP = {
    "Order SKU and system SKU differ after normalization": "原始商品编码与系统商品编码在标准化后仍不一致",
    "No sales row found for (SKC, SKUID)": "未找到对应销售键 (SKC, SKUID)",
    "Missing daily sales data for (SKC, SKUID)": "未找到对应每日销量数据 (SKC, SKUID)",
}

INT_FORMAT_FIELDS = {
    "line_order_qty",
    "key_order_qty",
    "gap",
    "key_recommended_total",
    "recommended_ship_before_small_change_rule",
    "recommended_ship",
    "order_recommended_ship_total_before_threshold",
    "min_order_ship_qty_threshold",
}

FORECAST_ERROR_FIELDS = {"sku_forecast_abs_error", "sku_forecast_signed_error"}

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

FORECAST_EVAL_SUMMARY_KEY = "预测回测评估"
FORECAST_EVAL_SUMMARY_FIELDS = [
    ("evaluated_skus", "评估SKU数"),
    ("skipped_insufficient_history", "历史不足跳过SKU数"),
    ("holdout_days", "留出天数"),
    ("mae", "MAE"),
    ("wape", "WAPE"),
    ("bias", "平均偏差"),
]
FORECAST_EVAL_STRATEGY_FIELDS = [
    ("count", "SKU数"),
    ("mae", "MAE"),
    ("wape", "WAPE"),
    ("bias", "平均偏差"),
]
FORECAST_EVAL_INT_FIELDS = {
    "evaluated_skus",
    "skipped_insufficient_history",
    "holdout_days",
    "count",
}


def export_reports(
    out_dir: str | Path,
    recommendations: list[dict[str, object]],
    quality_rows: list[dict[str, object]],
    summary: dict[str, object],
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]],
    eval_rows: Sequence[EvalRow] | None = None,
    eval_summary: dict[str, object] | None = None,
) -> dict[str, Path]:
    output_dir = Path(out_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    recommendation_path = output_dir / "发货建议明细.xlsx"
    quality_path = output_dir / "数据质量报告.csv"
    summary_path = output_dir / "运行摘要.json"
    _remove_stale_recommendation_csv(output_dir / "发货建议明细.csv")

    quality_columns = [target for _, target in QUALITY_FIELDS]

    _annotate_forecast_error(recommendations, eval_rows or ())
    plot_cache = _build_plot_cache(recommendations, daily_sales_by_key)
    localized_recommendations = [_localize_recommendation_row(row) for row in recommendations]
    localized_quality_rows = [_localize_quality_row(row) for row in quality_rows]
    localized_summary = _localize_summary(summary, eval_summary=eval_summary)

    _write_recommendations_xlsx(
        path=recommendation_path,
        rows=localized_recommendations,
        recommendation_rows=recommendations,
        plot_cache=plot_cache,
    )
    _write_csv(quality_path, localized_quality_rows, quality_columns)
    summary_path.write_text(
        json.dumps(localized_summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    return {
        "final_recommendation": recommendation_path,
        "quality_report": quality_path,
        "run_summary": summary_path,
    }


def _remove_stale_recommendation_csv(path: Path) -> None:
    try:
        path.unlink()
    except FileNotFoundError:
        return


def _annotate_forecast_error(
    recommendations: list[dict[str, object]],
    eval_rows: Sequence[EvalRow],
) -> None:
    error_by_key: dict[tuple[str, str], EvalRow] = {
        (row.skc, row.skuid): row for row in eval_rows
    }
    for row in recommendations:
        key = (str(row.get("店铺款式编码", "")), str(row.get("店铺商品编码", "")))
        eval_row = error_by_key.get(key)
        if eval_row is None:
            row["sku_forecast_abs_error"] = ""
            row["sku_forecast_signed_error"] = ""
            continue
        row["sku_forecast_abs_error"] = eval_row.abs_error
        row["sku_forecast_signed_error"] = eval_row.signed_error


def _build_plot_cache(
    recommendations: list[dict[str, object]],
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]],
) -> dict[tuple[str, str], bytes]:
    """Render one plot for every unique SKU in the recommendation rows."""
    from .plots import render_sku_plot

    cache: dict[tuple[str, str], bytes] = {}
    plot_keys = _plot_keys(recommendations)
    for row in recommendations:
        key = (str(row.get("店铺款式编码", "")), str(row.get("店铺商品编码", "")))
        if key in cache or key not in plot_keys:
            continue
        history = daily_sales_by_key.get(key, ())
        forecast = build_forecast_curve(
            strategy=str(row.get("forecast_strategy", "") or ""),
            forecast_daily_sales=_as_float(row.get("forecast_daily_sales", 0.0)),
            stocking_days=_as_float(row.get("stocking_days", 0.0)),
        )
        if not history and not forecast:
            continue
        masked_tail_from = stockout_mask_cut(
            daily_sales=history,
            stock_in_warehouse=_as_float(row.get("wh", 0.0)),
        )
        title = f"SKC:{key[0]} / SKUID:{key[1]}"
        cache[key] = render_sku_plot(
            history=history,
            forecast=forecast,
            title=title,
            masked_tail_from=masked_tail_from,
        )
    return cache


def _plot_keys(
    recommendations: list[dict[str, object]],
) -> set[tuple[str, str]]:
    keys: set[tuple[str, str]] = set()
    for row in recommendations:
        key = (str(row.get("店铺款式编码", "")), str(row.get("店铺商品编码", "")))
        keys.add(key)
    return keys


def _as_float(value: object) -> float:
    if isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value))
    except (TypeError, ValueError):
        return 0.0


def _write_recommendations_xlsx(
    *,
    path: Path,
    rows: list[dict[str, object]],
    recommendation_rows: list[dict[str, object]],
    plot_cache: dict[tuple[str, str], bytes],
) -> None:
    columns = [target for _, target in RECOMMENDATION_FIELDS]
    plot_col_idx = columns.index(PLOT_COLUMN_NAME) + 1
    plot_col_letter = get_column_letter(plot_col_idx)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "发货建议明细"

    for col_idx, name in enumerate(columns, start=1):
        worksheet.cell(row=1, column=col_idx, value=name)

    worksheet.column_dimensions[plot_col_letter].width = PLOT_COLUMN_WIDTH_CHARS

    for row_offset, (localized, source) in enumerate(zip(rows, recommendation_rows)):
        excel_row = row_offset + 2
        for col_idx, name in enumerate(columns, start=1):
            if name == PLOT_COLUMN_NAME:
                continue
            worksheet.cell(row=excel_row, column=col_idx, value=localized.get(name, ""))

        key = (str(source.get("店铺款式编码", "")), str(source.get("店铺商品编码", "")))
        image_bytes = plot_cache.get(key)
        if image_bytes is not None:
            image_stream = io.BytesIO(image_bytes)
            excel_image = XlsxImage(image_stream)
            excel_image.width = PLOT_IMAGE_WIDTH_PX
            excel_image.height = PLOT_IMAGE_HEIGHT_PX
            excel_image.anchor = f"{plot_col_letter}{excel_row}"
            worksheet.add_image(excel_image)
            worksheet.row_dimensions[excel_row].height = PLOT_ROW_HEIGHT_POINTS

    workbook.save(path)


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
    fields: Sequence[tuple[str, str]],
    default_value: object,
    value_mapper: Callable[[str, object], object],
) -> dict[str, object]:
    localized: dict[str, object] = {}
    for source, target in fields:
        if source == "__plot__":
            localized[target] = ""
            continue
        localized[target] = value_mapper(source, row.get(source, default_value))
    return localized


def _localize_summary(
    summary: dict[str, object],
    *,
    eval_summary: dict[str, object] | None = None,
) -> dict[str, object]:
    localized: dict[str, object] = {}
    for source, target in SUMMARY_FIELDS:
        value = summary.get(source, 0)
        if source in SUMMARY_INT_FORMAT_FIELDS:
            localized[target] = _format_int_like(value)
            continue
        localized[target] = value
    if eval_summary is not None:
        localized[FORECAST_EVAL_SUMMARY_KEY] = _localize_eval_summary(eval_summary)
    return localized


def _localize_eval_summary(eval_summary: dict[str, object]) -> dict[str, object]:
    localized: dict[str, object] = {}
    for source, target in FORECAST_EVAL_SUMMARY_FIELDS:
        localized[target] = _localize_eval_summary_value(source, eval_summary.get(source))

    by_strategy = eval_summary.get("by_strategy", {})
    if isinstance(by_strategy, dict):
        localized["按策略"] = {
            str(strategy): _localize_eval_strategy_summary(stats)
            for strategy, stats in by_strategy.items()
            if isinstance(stats, dict)
        }
    else:
        localized["按策略"] = {}
    return localized


def _localize_eval_strategy_summary(stats: dict[object, object]) -> dict[str, object]:
    localized: dict[str, object] = {}
    for source, target in FORECAST_EVAL_STRATEGY_FIELDS:
        localized[target] = _localize_eval_summary_value(source, stats.get(source))
    return localized


def _localize_eval_summary_value(source: str, value: object) -> object:
    if source in FORECAST_EVAL_INT_FIELDS:
        return _format_int_like(value)
    return value


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
