from __future__ import annotations

from collections.abc import Callable
from collections import defaultdict
from datetime import datetime
import functools
import re
from pathlib import Path
from typing import TypeVar

from .models import OrderLine, SalesRecord

ORDER_REQUIRED_COLUMNS = [
    "内部订单号",
    "下单时间",
    "店铺款式编码",
    "店铺商品编码",
    "商品编码",
    "原始商品编码",
    "地址",
    "数量",
    "状态",
    "标签",
]

SALES_REQUIRED_COLUMNS = [
    "平台商品基本信息-skc",
    "平台商品基本信息-是否热销款",
    "平台商品基本信息-平台SKUID",
    "平台商品基本信息-SKU货号",
    "销售数据-近30日销量",
    "销售数据-近7日销量",
    "平台商品基本信息-备货逻辑",
    "平台商品库存信息-平台仓内库存",
    "平台商品库存信息-平台待发货库存",
    "平台商品库存信息-平台待收货库存",
]

TAG_SPLIT_RE = re.compile(r"[，,]")
IN_PROGRESS_STATUS = "发货中"
SHORTAGE_STATUS = "缺货"
ORDER_TIME_FORMAT = "%Y-%m-%d %H:%M:%S"
HOT_STYLE_TRUE_VALUES = {"是", "true", "1", "yes", "y"}
NumberT = TypeVar("NumberT", int, float)


def _clean_text(value: str | None) -> str:
    return (value or "").strip()


def assert_xlsx(path: str | Path) -> None:
    if Path(path).suffix.lower() != ".xlsx":
        raise ValueError(f"Input must be .xlsx: {path}")


def assert_required_columns(header: list[str], required: list[str], file_label: str) -> None:
    missing = missing_required_columns(header, required)
    if missing:
        joined = ", ".join(describe_required_column(column_name) for column_name in missing)
        raise ValueError(f"Missing required columns in {file_label}: {joined}")


def missing_required_columns(header: list[str], required: list[str]) -> list[str]:
    header_set = set(header)
    return [column_name for column_name in required if column_name not in header_set]


def describe_required_column(column_name: str) -> str:
    return column_name


def parse_orders(rows: list[dict[str, str]]) -> tuple[list[OrderLine], dict[tuple[str, str], int]]:
    lines: list[OrderLine] = []
    shipping_in_progress_by_key: dict[tuple[str, str], int] = defaultdict(int)
    for row_number, row in enumerate(rows, start=2):
        row_get = row.get

        if not has_target_tag(row_get("标签"), "今日可发货"):
            continue

        skc = _clean_text(row_get("店铺款式编码"))
        skuid = _clean_text(row_get("店铺商品编码"))
        product_code = _clean_text(row_get("商品编码"))
        qty = parse_quantity_int(
            row_get("数量"),
            field_name="数量",
            row_number=row_number,
        )
        status = _clean_text(row_get("状态"))
        address = _clean_text(row_get("地址"))
        order_time = parse_order_time(row_get("下单时间"), row_number=row_number)

        if status == IN_PROGRESS_STATUS and address:
            shipping_in_progress_by_key[(skc, skuid)] += qty
            continue

        lines.append(
            OrderLine(
                row_number=row_number,
                internal_order_id=_clean_text(row_get("内部订单号")),
                skc=skc,
                skuid=skuid,
                product_code=product_code,
                order_sku=_clean_text(row_get("原始商品编码")),
                status=status,
                order_time=order_time,
                quantity=qty,
            )
        )
    return lines, dict(shipping_in_progress_by_key)


def parse_sales(rows: list[dict[str, str]]) -> list[SalesRecord]:
    records: list[SalesRecord] = []
    for row_number, row in enumerate(rows, start=2):
        row_get = row.get
        skc = _clean_text(row_get("平台商品基本信息-skc"))
        skuid = _clean_text(row_get("平台商品基本信息-平台SKUID"))
        system_sku = _clean_text(row_get("平台商品基本信息-SKU货号"))
        records.append(
            SalesRecord(
                row_number=row_number,
                skc=skc,
                skuid=skuid,
                system_sku=system_sku,
                is_hot_style=parse_hot_style(row_get("平台商品基本信息-是否热销款")),
                sold30=parse_int(row_get("销售数据-近30日销量")),
                sold7=parse_int(row_get("销售数据-近7日销量")),
                stocking_days=parse_stocking_days(row_get("平台商品基本信息-备货逻辑")),
                stock_in_warehouse=parse_float(row_get("平台商品库存信息-平台仓内库存")),
                pending_ship=parse_float(row_get("平台商品库存信息-平台待发货库存")),
                pending_receive=parse_float(row_get("平台商品库存信息-平台待收货库存")),
            )
        )
    return records


def parse_float(value: str | None) -> float:
    return _parse_number_or_default(value, parser=float, default=0.0)


def parse_int(value: str | None) -> int:
    return _parse_number_or_default(
        value,
        parser=_parse_int_from_number_text,
        default=0,
    )


def parse_quantity_int(
    value: str | None,
    *,
    field_name: str,
    row_number: int,
) -> int:
    text = _normalize_number_text(value)
    if not text:
        return 0

    try:
        number = float(text)
    except ValueError as exc:
        raise ValueError(
            f"Invalid {field_name} at orders row {row_number}: {value!r} is not a number"
        ) from exc

    if number < 0:
        raise ValueError(
            f"Invalid {field_name} at orders row {row_number}: value must be >= 0"
        )
    if not number.is_integer():
        raise ValueError(
            f"Invalid {field_name} at orders row {row_number}: {value!r} is not an integer"
        )

    return int(number)


def parse_stocking_days(value: str | None) -> float:
    normalized = _normalize_plus_text(value)
    if not normalized:
        return 0.0

    parts = [part for part in normalized.split("+") if part]
    if not parts:
        return 0.0
    if len(parts) == 1:
        return parse_float(parts[0])
    return sum(parse_float(part) for part in parts)


def parse_order_time(value: str | None, row_number: int) -> datetime:
    text = _clean_text(value)
    if not text:
        raise ValueError(f"Missing 下单时间 at orders row {row_number}")
    try:
        return datetime.strptime(text, ORDER_TIME_FORMAT)
    except ValueError as exc:
        raise ValueError(
            f"Invalid 下单时间 format at orders row {row_number}: {text}. "
            f"Expected format: {ORDER_TIME_FORMAT}"
        ) from exc


def has_target_tag(tags_value: str | None, target_tag: str) -> bool:
    tags_text = _clean_text(tags_value)
    if not tags_text:
        return False
    return any(_clean_text(tag) == target_tag for tag in TAG_SPLIT_RE.split(tags_text))


@functools.lru_cache(maxsize=None)
def normalize_sku_code(value: str | None) -> str:
    return _normalize_plus_text(value).lower()


def _normalize_number_text(value: str | None) -> str:
    return _clean_text(value).replace(",", "")


def _normalize_plus_text(value: str | None) -> str:
    text = _clean_text(value)
    text = text.replace("_x002B_", "+").replace("_x002b_", "+")
    return text.replace(" ", "")


def _parse_int_from_number_text(text: str) -> int:
    return int(float(text))


def _parse_number_or_default(
    value: str | None,
    *,
    parser: Callable[[str], NumberT],
    default: NumberT,
) -> NumberT:
    text = _normalize_number_text(value)
    if not text:
        return default
    try:
        return parser(text)
    except ValueError:
        return default


def parse_hot_style(value: str | None) -> bool:
    return _clean_text(value).lower() in HOT_STYLE_TRUE_VALUES
