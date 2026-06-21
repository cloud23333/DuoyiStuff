from __future__ import annotations

from collections.abc import Sequence

from ..engine import stockout_mask_cut
from ..eval_forecast import EvalRow
from ..forecast_curve import build_forecast_curve
from ._util import _as_float


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
    from ..plots import render_sku_plot

    cache: dict[tuple[str, str], bytes] = {}
    for row in recommendations:
        key = (str(row.get("店铺款式编码", "")), str(row.get("店铺商品编码", "")))
        if key in cache:
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
        title = _plot_title(row, key)
        cache[key] = render_sku_plot(
            history=history,
            forecast=forecast,
            title=title,
            masked_tail_from=masked_tail_from,
            forecast_daily_sales=_as_float(row.get("forecast_daily_sales", 0.0)),
        )
    return cache


def _plot_title(
    row: dict[str, object],
    key: tuple[str, str],
) -> str:
    base = f"SKC:{key[0]} / SKUID:{key[1]}"
    model = str(row.get("forecast_model", "") or "").strip()
    if not model:
        return base
    return f"{base} | {model}"
