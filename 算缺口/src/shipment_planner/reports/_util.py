from __future__ import annotations


def _as_float(value: object) -> float:
    if isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value))
    except (TypeError, ValueError):
        return 0.0


def _format_int_like(value: object) -> object:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _is_yes(value: object) -> bool:
    return str(value).strip() in {"yes", "是"}
