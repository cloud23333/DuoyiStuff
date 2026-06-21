from __future__ import annotations

from ._util import _as_float, _format_int_like, _is_yes
from .schema import INTERCEPT_REASON_MAP


def _recommendation_derived_value(source: str, row: dict[str, object]) -> object:
    if source == "__item_identity__":
        return _build_item_identity(row)
    if source == "__action_notes__":
        return _build_action_notes(row)
    if source == "__formula__":
        return _build_formula(row)
    if source == "__available_stock__":
        return _format_int_like(
            round(
                _as_float(row.get("wh"))
                + _as_float(row.get("pending_recv"))
                + _as_float(row.get("shipping_in_progress")),
                4,
            )
        )
    return ""


def _build_item_identity(row: dict[str, object]) -> str:
    skc = str(row.get("店铺款式编码", "")).strip()
    skuid = str(row.get("店铺商品编码", "")).strip()
    order_sku = str(row.get("原始商品编码", "")).strip()
    parts = [
        f"SKC {skc}" if skc else "",
        f"SKUID {skuid}" if skuid else "",
        f"SKU {order_sku}" if order_sku else "",
    ]
    return "\n".join(part for part in parts if part)


def _build_action_notes(row: dict[str, object]) -> str:
    notes: list[str] = []

    sku_code_check = str(row.get("sku_code_check", "")).strip()
    if sku_code_check == "diff":
        notes.append("SKU不一致")
    elif sku_code_check == "missing_key":
        notes.append("缺少销售匹配")

    intercept_reason = str(row.get("intercept_reason", "")).strip()
    if intercept_reason:
        notes.append(str(INTERCEPT_REASON_MAP.get(intercept_reason, intercept_reason)))

    if _is_yes(row.get("order_intercept_warning")):
        notes.append("订单拦截导致不发")
    if _is_yes(row.get("order_low_qty_warning")):
        notes.append("订单低于起发量")
    if _is_yes(row.get("min_order_ship_qty_exempt_applied_warning")):
        notes.append("小于10豁免生效")
    elif _is_yes(row.get("min_order_ship_qty_exempt_warning")):
        notes.append("小于10豁免候选")
    if _is_yes(row.get("small_change_keep_warning")):
        notes.append("30%内免改")
    if _is_yes(row.get("base_stock_triggered_warning")):
        notes.append("保底触发")

    anomaly_flags = str(row.get("anomaly_flags", "")).strip()
    if anomaly_flags and anomaly_flags != "无":
        notes.append(f"销量异常：{anomaly_flags}")

    return "；".join(notes)


def _fmt_num(value: object) -> str:
    number = round(_as_float(value), 2)
    if number == int(number):
        return str(int(number))
    return f"{number:.2f}".rstrip("0").rstrip(".")


def _build_formula(row: dict[str, object]) -> str:
    wh = _as_float(row.get("wh"))
    recv = _as_float(row.get("pending_recv"))
    ship = _as_float(row.get("shipping_in_progress"))
    available = wh + recv + ship
    forecast_period = _as_float(row.get("forecast_stocking_period_sales"))
    base_stock = _as_float(row.get("base_stock_qty"))
    gap = _as_float(row.get("gap"))
    key_order_qty = _as_float(row.get("key_order_qty"))
    allocated = _as_float(row.get("recommended_ship_before_small_change_rule"))
    line_qty = _as_float(row.get("line_order_qty"))
    final_ship = _as_float(row.get("recommended_ship"))
    key_recommended = min(key_order_qty, gap)

    steps = [
        f"可用库存 = 仓内{_fmt_num(wh)}+待收{_fmt_num(recv)}"
        f"+发货中{_fmt_num(ship)} = {_fmt_num(available)}"
    ]
    if base_stock > 0:
        steps.append(
            f"缺口 = ⌈max(预测{_fmt_num(forecast_period)}−可用{_fmt_num(available)}, "
            f"保底{_fmt_num(base_stock)}−可用{_fmt_num(available)}, 0)⌉ = {_fmt_num(gap)}"
        )
    else:
        steps.append(
            f"缺口 = ⌈max(预测{_fmt_num(forecast_period)}−可用{_fmt_num(available)}, 0)⌉"
            f" = {_fmt_num(gap)}"
        )
    steps.append(
        f"同款建议 = min(同款下单{_fmt_num(key_order_qty)}, 缺口{_fmt_num(gap)})"
        f" = {_fmt_num(key_recommended)}"
    )
    if allocated != key_recommended:
        steps.append(f"本行分摊 = {_fmt_num(allocated)}")

    intercept_reason = str(row.get("intercept_reason", "")).strip()
    if intercept_reason:
        label = INTERCEPT_REASON_MAP.get(intercept_reason, intercept_reason)
        steps.append(f"{label} → 0")
    if _is_yes(row.get("small_change_keep_warning")):
        steps.append(
            f"30%免改：|下单{_fmt_num(line_qty)}−分摊{_fmt_num(allocated)}|"
            f"/{_fmt_num(line_qty)} ≤ 30% → 保留原量"
        )
    threshold = _fmt_num(row.get("min_order_ship_qty_threshold"))
    order_total = _fmt_num(row.get("order_recommended_ship_total_before_threshold"))
    if _is_yes(row.get("min_order_ship_qty_exempt_applied_warning")):
        steps.append(f"订单合计{order_total} < 起发{threshold}，保底豁免 → 保留")
    elif _is_yes(row.get("order_low_qty_warning")):
        steps.append(f"订单合计{order_total} < 起发{threshold} → 0")

    steps.append(f"∴ 建议发货量 = {_fmt_num(final_ship)}")
    return "\n".join(steps)
