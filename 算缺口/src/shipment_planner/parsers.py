from __future__ import annotations

from collections import defaultdict
from datetime import datetime
import functools
import re
from pathlib import Path

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
    "平台商品基本信息-平台SKUID",
    "平台商品基本信息-SKU货号",
    "平台商品基本信息-备货逻辑",
    "平台商品库存信息-平台仓内库存",
    "平台商品库存信息-平台待发货库存",
    "平台商品库存信息-平台待收货库存",
]

TEMU_DAILY_SKC_COLUMN = "平台SKC_ID"
TEMU_DAILY_SKU_COLUMN = "平台SKU_ID"
TEMU_DAILY_REQUIRED_COLUMNS = [TEMU_DAILY_SKC_COLUMN, TEMU_DAILY_SKU_COLUMN]
_TEMU_DAILY_COL_RE = re.compile(r"^(\d+)月(\d+)日销量$")

TAG_SPLIT_RE = re.compile(r"[，,]")
IN_PROGRESS_STATUS = "发货中"
SHORTAGE_STATUS = "缺货"
ORDER_TIME_FORMAT = "%Y-%m-%d %H:%M:%S"


def _clean_text(value: str | None) -> str:
    return (value or "").strip()


def assert_xlsx(path: str | Path) -> None:
    if Path(path).suffix.lower() != ".xlsx":
        raise ValueError(f"Input must be .xlsx: {path}")


def assert_required_columns(header: list[str], required: list[str], file_label: str) -> None:
    missing = missing_required_columns(header, required)
    if missing:
        joined = ", ".join(missing)
        raise ValueError(f"Missing required columns in {file_label}: {joined}")


def missing_required_columns(header: list[str], required: list[str]) -> list[str]:
    header_set = set(header)
    return [column_name for column_name in required if column_name not in header_set]


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
                stocking_days=parse_stocking_days(row_get("平台商品基本信息-备货逻辑")),
                stock_in_warehouse=parse_float(row_get("平台商品库存信息-平台仓内库存")),
                pending_ship=parse_float(row_get("平台商品库存信息-平台待发货库存")),
                pending_receive=parse_float(row_get("平台商品库存信息-平台待收货库存")),
            )
        )
    return records


def parse_float(value: str | None) -> float:
    try:
        return float(_normalize_number_text(value))
    except (TypeError, ValueError):
        return 0.0


def parse_int(value: str | None) -> int:
    try:
        return int(float(_normalize_number_text(value)))
    except (TypeError, ValueError):
        return 0


def parse_quantity_int(
    value: str | None,
    *,
    row_number: int,
) -> int:
    text = _normalize_number_text(value)
    if not text:
        return 0

    try:
        number = float(text)
    except ValueError as exc:
        raise ValueError(
            f"Invalid 数量 at orders row {row_number}: {value!r} is not a number"
        ) from exc

    if number < 0:
        raise ValueError(
            f"Invalid 数量 at orders row {row_number}: value must be >= 0"
        )
    if not number.is_integer():
        raise ValueError(
            f"Invalid 数量 at orders row {row_number}: {value!r} is not an integer"
        )

    return int(number)


def parse_stocking_days(value: str | None) -> float:
    normalized = _normalize_plus_text(value)
    if not normalized:
        return 0.0

    parts = [part for part in normalized.split("+") if part]
    if not parts:
        return 0.0
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


def _temu_date_sort_key(col: str) -> tuple[int, int]:
    m = _TEMU_DAILY_COL_RE.match(col)
    return (int(m.group(1)), int(m.group(2))) if m else (0, 0)


def temu_daily_sales_columns(header: list[str]) -> list[str]:
    return sorted(
        [column_name for column_name in header if _TEMU_DAILY_COL_RE.match(column_name)],
        key=_temu_date_sort_key,
    )


def assert_temu_daily_sales_columns(header: list[str], file_label: str) -> None:
    assert_required_columns(header, TEMU_DAILY_REQUIRED_COLUMNS, file_label)
    if not temu_daily_sales_columns(header):
        raise ValueError(f"{file_label} has no daily sales columns")


def parse_temu_daily_sales(
    rows: list[dict[str, str]],
) -> dict[tuple[str, str], tuple[int, ...]]:
    """Parse Temu daily sales export.

    Returns a mapping of (平台SKC_ID, 平台SKU_ID) → daily sales tuple (oldest →
    newest). Rows sharing the same key (different shops) are summed.
    """
    all_cols = list(dict.fromkeys(column_name for row in rows for column_name in row))
    assert_temu_daily_sales_columns(all_cols, "Temu daily sales file")
    date_cols = temu_daily_sales_columns(all_cols)

    result: dict[tuple[str, str], tuple[int, ...]] = {}
    for row in rows:
        skc_id = _clean_text(row.get(TEMU_DAILY_SKC_COLUMN))
        sku_id = _clean_text(row.get(TEMU_DAILY_SKU_COLUMN))
        if not skc_id or not sku_id:
            continue
        key = (skc_id, sku_id)
        daily = tuple(parse_int(row.get(c)) for c in date_cols)
        if key in result:
            existing = result[key]
            result[key] = tuple(a + b for a, b in zip(existing, daily))
        else:
            result[key] = daily

    return result
