from __future__ import annotations

import pytest

from shipment_planner.parsers import (
    has_target_tag,
    normalize_sku_code,
    parse_hot_style,
    parse_order_time,
    parse_orders,
    parse_quantity_int,
    parse_sales,
    parse_stocking_days,
)


# ---------------------------------------------------------------------------
# parse_quantity_int
# ---------------------------------------------------------------------------


def test_parse_quantity_int_returns_int_for_valid_integer_string():
    assert parse_quantity_int("5", field_name="数量", row_number=2) == 5


def test_parse_quantity_int_returns_int_for_float_string_with_zero_decimal():
    assert parse_quantity_int("3.0", field_name="数量", row_number=2) == 3


def test_parse_quantity_int_returns_zero_for_empty_string():
    assert parse_quantity_int("", field_name="数量", row_number=2) == 0


def test_parse_quantity_int_returns_zero_for_none():
    assert parse_quantity_int(None, field_name="数量", row_number=2) == 0


def test_parse_quantity_int_raises_for_non_numeric_string():
    with pytest.raises(ValueError, match="not a number"):
        parse_quantity_int("abc", field_name="数量", row_number=3)


def test_parse_quantity_int_raises_for_negative_value():
    with pytest.raises(ValueError, match="must be >= 0"):
        parse_quantity_int("-1", field_name="数量", row_number=4)


def test_parse_quantity_int_raises_for_non_integer_float():
    with pytest.raises(ValueError, match="not an integer"):
        parse_quantity_int("2.5", field_name="数量", row_number=5)


def test_parse_quantity_int_handles_comma_separated_number():
    # e.g. "1,000" should be treated as 1000
    assert parse_quantity_int("1,000", field_name="数量", row_number=2) == 1000


# ---------------------------------------------------------------------------
# parse_stocking_days
# ---------------------------------------------------------------------------


def test_parse_stocking_days_returns_float_for_valid_value():
    assert parse_stocking_days("7") == 7.0


def test_parse_stocking_days_returns_zero_for_empty_string():
    assert parse_stocking_days("") == 0.0


def test_parse_stocking_days_returns_zero_for_none():
    assert parse_stocking_days(None) == 0.0


def test_parse_stocking_days_returns_zero_for_string_zero():
    assert parse_stocking_days("0") == 0.0


def test_parse_stocking_days_handles_plus_separated_values():
    # "3+4" should sum to 7.0
    assert parse_stocking_days("3+4") == 7.0


def test_parse_stocking_days_handles_encoded_plus_separator():
    # "_x002B_" is the XML-escaped "+" sometimes found in xlsx exports
    assert parse_stocking_days("3_x002B_4") == 7.0


# ---------------------------------------------------------------------------
# has_target_tag
# ---------------------------------------------------------------------------


def test_has_target_tag_returns_true_when_tag_present():
    assert has_target_tag("今日可发货", "今日可发货") is True


def test_has_target_tag_returns_false_when_tag_absent():
    assert has_target_tag("其他标签", "今日可发货") is False


def test_has_target_tag_returns_false_for_empty_string():
    assert has_target_tag("", "今日可发货") is False


def test_has_target_tag_returns_false_for_none():
    assert has_target_tag(None, "今日可发货") is False


def test_has_target_tag_handles_comma_separated_chinese_tags():
    assert has_target_tag("标签A，今日可发货，标签B", "今日可发货") is True


def test_has_target_tag_handles_ascii_comma_separated_tags():
    assert has_target_tag("tagA,今日可发货,tagB", "今日可发货") is True


def test_has_target_tag_returns_false_when_tag_is_substring_only():
    # "今日可发货Extra" should NOT match target "今日可发货"
    assert has_target_tag("今日可发货Extra", "今日可发货") is False


# ---------------------------------------------------------------------------
# normalize_sku_code
# ---------------------------------------------------------------------------


def test_normalize_sku_code_strips_whitespace():
    assert normalize_sku_code("  SKU-001  ") == "sku-001"


def test_normalize_sku_code_lowercases_result():
    assert normalize_sku_code("SKU-ABC") == "sku-abc"


def test_normalize_sku_code_returns_empty_string_for_none():
    assert normalize_sku_code(None) == ""


def test_normalize_sku_code_returns_empty_string_for_empty_input():
    assert normalize_sku_code("") == ""


def test_normalize_sku_code_removes_internal_spaces():
    assert normalize_sku_code("SKU 001") == "sku001"


# ---------------------------------------------------------------------------
# parse_hot_style
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("value", ["是", "true", "1", "yes"])
def test_parse_hot_style_returns_true_for_truthy_values(value):
    assert parse_hot_style(value) is True


def test_parse_hot_style_returns_false_for_chinese_no():
    assert parse_hot_style("否") is False


def test_parse_hot_style_returns_false_for_empty_string():
    assert parse_hot_style("") is False


def test_parse_hot_style_returns_false_for_none():
    assert parse_hot_style(None) is False


def test_parse_hot_style_is_case_insensitive():
    assert parse_hot_style("YES") is True
    assert parse_hot_style("True") is True


# ---------------------------------------------------------------------------
# parse_orders
# ---------------------------------------------------------------------------

def _make_order_row(**overrides) -> dict[str, str]:
    """Return a minimal valid order row dict that will produce an OrderLine."""
    base = {
        "内部订单号": "ORD-001",
        "下单时间": "2024-01-15 10:00:00",
        "店铺款式编码": "SKC001",
        "店铺商品编码": "SKUID001",
        "商品编码": "PROD001",
        "原始商品编码": "RAW-001",
        "地址": "",
        "数量": "5",
        "状态": "待发货",
        "标签": "今日可发货",
    }
    base.update(overrides)
    return base


def test_parse_orders_returns_order_line_for_valid_row():
    rows = [_make_order_row()]
    lines, in_progress = parse_orders(rows)
    assert len(lines) == 1
    assert lines[0].skc == "SKC001"
    assert lines[0].quantity == 5


def test_parse_orders_skips_row_without_target_tag():
    rows = [_make_order_row(标签="其他标签")]
    lines, in_progress = parse_orders(rows)
    assert lines == []


def test_parse_orders_counts_shipping_in_progress_row():
    rows = [_make_order_row(状态="发货中", 地址="某地址")]
    lines, in_progress = parse_orders(rows)
    assert lines == []
    assert in_progress.get(("SKC001", "SKUID001"), 0) == 5


def test_parse_orders_ignores_extra_columns():
    rows = [_make_order_row(额外列="ignored")]
    lines, _ = parse_orders(rows)
    assert len(lines) == 1


def test_parse_orders_raises_for_invalid_quantity():
    rows = [_make_order_row(数量="not_a_number")]
    with pytest.raises(ValueError, match="not a number"):
        parse_orders(rows)


def test_parse_orders_strips_whitespace_from_skc_and_skuid():
    rows = [_make_order_row(**{"店铺款式编码": "  SKC001  ", "店铺商品编码": "  SKUID001  "})]
    lines, _ = parse_orders(rows)
    assert lines[0].skc == "SKC001"
    assert lines[0].skuid == "SKUID001"


def test_parse_orders_in_progress_with_empty_address_falls_through_to_lines():
    # status == IN_PROGRESS_STATUS but address is empty → should NOT go into
    # in_progress dict; the row should be appended to lines instead.
    rows = [_make_order_row(状态="发货中", 地址="")]
    lines, in_progress = parse_orders(rows)
    assert len(lines) == 1
    assert in_progress == {}


# ---------------------------------------------------------------------------
# parse_order_time
# ---------------------------------------------------------------------------


def test_parse_order_time_parses_valid_datetime_string():
    from datetime import datetime

    result = parse_order_time("2024-01-15 10:30:00", row_number=2)
    assert result == datetime(2024, 1, 15, 10, 30, 0)


def test_parse_order_time_raises_for_empty_string():
    with pytest.raises(ValueError, match="Missing 下单时间"):
        parse_order_time("", row_number=3)


# ---------------------------------------------------------------------------
# parse_sales
# ---------------------------------------------------------------------------

def _make_sales_row(**overrides) -> dict[str, str]:
    """Return a minimal valid sales row dict."""
    base = {
        "平台商品基本信息-skc": "SKC001",
        "平台商品基本信息-是否热销款": "否",
        "平台商品基本信息-平台SKUID": "SKUID001",
        "平台商品基本信息-SKU货号": "SYS-SKU-001",
        "销售数据-近30日销量": "100",
        "销售数据-近7日销量": "30",
        "平台商品基本信息-备货逻辑": "7",
        "平台商品库存信息-平台仓内库存": "50",
        "平台商品库存信息-平台待发货库存": "5",
        "平台商品库存信息-平台待收货库存": "10",
    }
    base.update(overrides)
    return base


def test_parse_sales_returns_sales_record_for_valid_row():
    rows = [_make_sales_row()]
    records = parse_sales(rows)
    assert len(records) == 1
    assert records[0].skc == "SKC001"
    assert records[0].sold30 == 100
    assert records[0].sold7 == 30


def test_parse_sales_parses_hot_style_flag():
    rows = [_make_sales_row(**{"平台商品基本信息-是否热销款": "是"})]
    records = parse_sales(rows)
    assert records[0].is_hot_style is True


def test_parse_sales_returns_zero_for_missing_numeric_columns():
    rows = [_make_sales_row(**{"销售数据-近30日销量": "", "销售数据-近7日销量": ""})]
    records = parse_sales(rows)
    assert records[0].sold30 == 0
    assert records[0].sold7 == 0


def test_parse_sales_ignores_extra_columns():
    rows = [_make_sales_row(额外列="ignored")]
    records = parse_sales(rows)
    assert len(records) == 1


def test_parse_sales_parses_stocking_days_correctly():
    rows = [_make_sales_row(**{"平台商品基本信息-备货逻辑": "3+4"})]
    records = parse_sales(rows)
    assert records[0].stocking_days == 7.0


def test_parse_sales_handles_comma_formatted_numeric_fields():
    # Values like "1,200" (thousands separator) should be parsed correctly.
    rows = [_make_sales_row(**{
        "销售数据-近30日销量": "1,200",
        "销售数据-近7日销量": "2,500",
        "平台商品库存信息-平台仓内库存": "3,000",
    })]
    records = parse_sales(rows)
    assert records[0].sold30 == 1200
    assert records[0].sold7 == 2500
    assert records[0].stock_in_warehouse == 3000.0
