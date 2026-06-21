from __future__ import annotations

import json
import re
from pathlib import Path

from .parsers import normalize_sku_code

DEFAULT_CONSTRAINTS_FILENAME = "shipment_constraints.json"


SPLIT_CODES_RE = re.compile(r"[，,]")


def load_constraints(
    path: str | Path,
    *,
    strict: bool = False,
) -> tuple[dict[str, int], set[str], set[str], bool]:
    constraints_path = Path(path)
    if not constraints_path.exists():
        if strict:
            raise ValueError(f"Constraints file not found: {constraints_path}")
        return {}, set(), set(), False

    data = _read_json_object(constraints_path)
    limits = _parse_sku_order_max_qty(data, constraints_path)
    exclude_skc = _parse_code_set(
        data.get("exclude_skc"),
        label="exclude_skc",
        path=constraints_path,
    )
    exclude_skuid = _parse_code_set(
        data.get("exclude_skuid"),
        label="exclude_skuid",
        path=constraints_path,
    )
    return limits, exclude_skc, exclude_skuid, True


def _parse_sku_order_max_qty(
    data: dict[str, object],
    constraints_path: Path,
) -> dict[str, int]:
    raw_limits = data.get("sku_order_max_qty", {})
    if raw_limits is None:
        return {}
    if not isinstance(raw_limits, dict):
        raise ValueError(
            f"Invalid constraints format in {constraints_path}: "
            "'sku_order_max_qty' must be an object"
        )

    limits: dict[str, int] = {}
    for raw_sku, raw_limit in raw_limits.items():
        if not isinstance(raw_sku, str):
            raise ValueError(
                f"Invalid constraints format in {constraints_path}: "
                "all sku_order_max_qty keys must be strings"
            )
        sku_code = normalize_sku_code(raw_sku)
        if not sku_code:
            raise ValueError(
                f"Invalid constraints format in {constraints_path}: "
                "empty SKU key is not allowed in sku_order_max_qty"
            )
        limits[sku_code] = _parse_non_negative_int(
            raw_limit,
            label=f"sku_order_max_qty[{raw_sku!r}]",
            path=constraints_path,
        )
    return limits


def _parse_code_set(
    value: object,
    *,
    label: str,
    path: Path,
) -> set[str]:
    if value is None:
        return set()
    if not isinstance(value, list):
        raise ValueError(
            f"Invalid constraints format in {path}: "
            f"'{label}' must be an array of strings"
        )

    code_set: set[str] = set()
    for index, item in enumerate(value):
        if not isinstance(item, str):
            raise ValueError(
                f"Invalid {label}[{index}] in {path}: expected string, got {type(item).__name__}"
            )
        for raw_code in SPLIT_CODES_RE.split(item):
            code = raw_code.strip()
            if not code:
                continue
            code_set.add(code)
    return code_set


def _read_json_object(path: Path) -> dict[str, object]:
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise ValueError(f"Failed to read constraints file {path}: {exc}") from exc

    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON in constraints file {path}: {exc}") from exc

    if not isinstance(data, dict):
        raise ValueError(f"Invalid constraints format in {path}: root must be a JSON object")
    return data


def _parse_non_negative_int(value: object, *, label: str, path: Path) -> int:
    if isinstance(value, bool):
        raise ValueError(f"Invalid {label} in {path}: boolean is not allowed")

    number: float
    if isinstance(value, (int, float)):
        number = float(value)
    elif isinstance(value, str):
        text = value.strip()
        if not text:
            raise ValueError(f"Invalid {label} in {path}: value is empty")
        try:
            number = float(text)
        except ValueError as exc:
            raise ValueError(f"Invalid {label} in {path}: {value!r} is not a number") from exc
    else:
        raise ValueError(
            f"Invalid {label} in {path}: expected number or numeric string, got {type(value).__name__}"
        )

    if number < 0:
        raise ValueError(f"Invalid {label} in {path}: value must be >= 0")
    if not number.is_integer():
        raise ValueError(f"Invalid {label} in {path}: value must be an integer")
    return int(number)
