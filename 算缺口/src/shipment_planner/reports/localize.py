from __future__ import annotations

from collections.abc import Callable, Sequence

from ._util import _format_int_like
from .derive import _recommendation_derived_value
from .schema import (
    DECISION_FIELDS,
    DECISION_REASON_MAP,
    FORECAST_EVAL_INT_FIELDS,
    FORECAST_EVAL_STRATEGY_FIELDS,
    FORECAST_EVAL_SUMMARY_FIELDS,
    FORECAST_EVAL_SUMMARY_KEY,
    INT_FORMAT_FIELDS,
    INTERCEPT_REASON_MAP,
    QUALITY_FIELDS,
    QUALITY_MESSAGE_MAP,
    QUALITY_TYPE_MAP,
    RECOMMENDATION_FIELDS,
    SKU_CHECK_MAP,
    SUMMARY_FIELDS,
    SUMMARY_INT_FORMAT_FIELDS,
    WARNING_FIELDS,
)


def _localize_recommendation_row(row: dict[str, object]) -> dict[str, object]:
    return _localize_row(
        row=row,
        fields=RECOMMENDATION_FIELDS,
        default_value="",
        value_mapper=_localize_recommendation_value,
        derived_value_mapper=_recommendation_derived_value,
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
    derived_value_mapper: Callable[[str, dict[str, object]], object] | None = None,
) -> dict[str, object]:
    localized: dict[str, object] = {}
    for source, target in fields:
        if source == "__plot__":
            localized[target] = ""
            continue
        if source.startswith("__") and derived_value_mapper is not None:
            localized[target] = derived_value_mapper(source, row)
            continue
        localized[target] = value_mapper(source, row.get(source, default_value))
    return localized


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
