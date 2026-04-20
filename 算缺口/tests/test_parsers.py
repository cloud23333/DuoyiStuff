from __future__ import annotations

import pytest

from shipment_planner.parsers import (
    parse_quantity_int,
    parse_stocking_days,
    parse_temu_daily_sales,
)


def test_parse_stocking_days_handles_excel_encoded_plus() -> None:
    assert parse_stocking_days("3 _x002B_ 4+5") == 12.0


def test_parse_quantity_int_rejects_fractional_quantities() -> None:
    with pytest.raises(ValueError, match="not an integer"):
        parse_quantity_int("1.5", field_name="数量", row_number=2)


def test_parse_temu_daily_sales_sums_duplicate_keys_in_date_order() -> None:
    rows = [
        {
            "平台SKC_ID": "skc-1",
            "平台SKU_ID": "sku-1",
            "4月2日销量": "1",
            "4月1日销量": "2",
        },
        {
            "平台SKC_ID": "skc-1",
            "平台SKU_ID": "sku-1",
            "4月2日销量": "3",
            "4月1日销量": "4",
        },
    ]

    assert parse_temu_daily_sales(rows) == {("skc-1", "sku-1"): (6, 4)}


def test_parse_temu_daily_sales_requires_identity_columns() -> None:
    rows = [{"平台SKC_ID": "skc-1", "4月1日销量": "1"}]

    with pytest.raises(ValueError, match="Missing required columns in Temu daily sales file"):
        parse_temu_daily_sales(rows)


def test_parse_temu_daily_sales_requires_daily_sales_columns() -> None:
    rows = [{"平台SKC_ID": "skc-1", "平台SKU_ID": "sku-1"}]

    with pytest.raises(ValueError, match="Temu daily sales file has no daily sales columns"):
        parse_temu_daily_sales(rows)
