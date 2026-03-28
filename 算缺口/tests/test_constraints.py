from __future__ import annotations

import json
from pathlib import Path

import pytest

from shipment_planner.constraints import load_constraints, load_sku_order_max_qty


# ---------------------------------------------------------------------------
# load_constraints — missing file
# ---------------------------------------------------------------------------


def test_load_constraints_missing_file_returns_empty_defaults(tmp_path: Path):
    missing = tmp_path / "nonexistent.json"
    limits, exclude_skc, exclude_skuid, loaded = load_constraints(missing)
    assert limits == {}
    assert exclude_skc == set()
    assert exclude_skuid == set()
    assert loaded is False


def test_load_constraints_missing_file_strict_raises(tmp_path: Path):
    missing = tmp_path / "nonexistent.json"
    with pytest.raises(ValueError, match="not found"):
        load_constraints(missing, strict=True)


# ---------------------------------------------------------------------------
# load_constraints — empty JSON object
# ---------------------------------------------------------------------------


def test_load_constraints_empty_json_object_returns_empty_defaults(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text("{}", encoding="utf-8")
    limits, exclude_skc, exclude_skuid, loaded = load_constraints(f)
    assert limits == {}
    assert exclude_skc == set()
    assert exclude_skuid == set()
    assert loaded is True


# ---------------------------------------------------------------------------
# load_constraints — invalid types
# ---------------------------------------------------------------------------


def test_load_constraints_sku_order_max_qty_not_dict_raises(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text(json.dumps({"sku_order_max_qty": "bad_value"}), encoding="utf-8")
    with pytest.raises(ValueError, match="must be an object"):
        load_constraints(f)


def test_load_constraints_max_qty_string_value_raises(tmp_path: Path):
    """A string like 'lots' where a number is expected should raise."""
    f = tmp_path / "constraints.json"
    f.write_text(
        json.dumps({"sku_order_max_qty": {"SKU-001": "lots"}}), encoding="utf-8"
    )
    with pytest.raises(ValueError, match="not a number"):
        load_constraints(f)


def test_load_constraints_max_qty_boolean_value_raises(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text(
        json.dumps({"sku_order_max_qty": {"SKU-001": True}}), encoding="utf-8"
    )
    with pytest.raises(ValueError, match="boolean is not allowed"):
        load_constraints(f)


def test_load_constraints_exclude_skc_not_list_raises(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text(json.dumps({"exclude_skc": "SKC-001"}), encoding="utf-8")
    with pytest.raises(ValueError, match="must be an array"):
        load_constraints(f)


def test_load_constraints_exclude_skuid_element_not_string_raises(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text(json.dumps({"exclude_skuid": [123]}), encoding="utf-8")
    with pytest.raises(ValueError, match="expected string"):
        load_constraints(f)


# ---------------------------------------------------------------------------
# load_constraints — negative max value rejected
# ---------------------------------------------------------------------------


def test_load_constraints_negative_max_qty_raises(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text(
        json.dumps({"sku_order_max_qty": {"SKU-001": -5}}), encoding="utf-8"
    )
    with pytest.raises(ValueError, match=">= 0"):
        load_constraints(f)


# ---------------------------------------------------------------------------
# load_constraints — comma-separated exclude codes
# ---------------------------------------------------------------------------


def test_load_constraints_comma_separated_exclude_skc_parsed_correctly(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text(
        json.dumps({"exclude_skc": ["SKC-001,SKC-002", "SKC-003"]}), encoding="utf-8"
    )
    _, exclude_skc, _, loaded = load_constraints(f)
    assert loaded is True
    assert "SKC-001" in exclude_skc
    assert "SKC-002" in exclude_skc
    assert "SKC-003" in exclude_skc


def test_load_constraints_fullwidth_comma_in_exclude_skuid_parsed_correctly(
    tmp_path: Path,
):
    """Fullwidth Chinese comma '，' should also act as a separator."""
    f = tmp_path / "constraints.json"
    f.write_text(
        json.dumps({"exclude_skuid": ["SKUID-A\uff0cSKUID-B"]}), encoding="utf-8"
    )
    _, _, exclude_skuid, loaded = load_constraints(f)
    assert loaded is True
    assert "SKUID-A" in exclude_skuid
    assert "SKUID-B" in exclude_skuid


# ---------------------------------------------------------------------------
# load_constraints — max quantity cap
# ---------------------------------------------------------------------------


def test_load_constraints_max_qty_cap_stored_correctly(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text(
        json.dumps({"sku_order_max_qty": {"SKU-ABC": 50}}), encoding="utf-8"
    )
    limits, _, _, loaded = load_constraints(f)
    assert loaded is True
    # normalize_sku_code lowercases the key
    assert limits.get("sku-abc") == 50


def test_load_constraints_numeric_string_max_qty_accepted(tmp_path: Path):
    """A numeric string like '100' should be coerced to int 100."""
    f = tmp_path / "constraints.json"
    f.write_text(
        json.dumps({"sku_order_max_qty": {"SKU-XYZ": "100"}}), encoding="utf-8"
    )
    limits, _, _, loaded = load_constraints(f)
    assert loaded is True
    assert limits.get("sku-xyz") == 100


# ---------------------------------------------------------------------------
# load_constraints — zero max qty is valid
# ---------------------------------------------------------------------------


def test_load_constraints_zero_max_qty_is_valid(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text(
        json.dumps({"sku_order_max_qty": {"SKU-ZERO": 0}}), encoding="utf-8"
    )
    limits, _, _, loaded = load_constraints(f)
    assert loaded is True
    assert limits.get("sku-zero") == 0


# ---------------------------------------------------------------------------
# load_constraints — full round-trip (all fields populated)
# ---------------------------------------------------------------------------


def test_load_constraints_full_round_trip(tmp_path: Path):
    payload = {
        "sku_order_max_qty": {"SKU-100": 200, "SKU-200": 50},
        "exclude_skc": ["SKC-X", "SKC-Y"],
        "exclude_skuid": ["SKUID-1,SKUID-2"],
    }
    f = tmp_path / "constraints.json"
    f.write_text(json.dumps(payload), encoding="utf-8")
    limits, exclude_skc, exclude_skuid, loaded = load_constraints(f)
    assert loaded is True
    assert limits["sku-100"] == 200
    assert limits["sku-200"] == 50
    assert "SKC-X" in exclude_skc
    assert "SKC-Y" in exclude_skc
    assert "SKUID-1" in exclude_skuid
    assert "SKUID-2" in exclude_skuid


# ---------------------------------------------------------------------------
# load_sku_order_max_qty — convenience wrapper
# ---------------------------------------------------------------------------


def test_load_sku_order_max_qty_returns_limits_and_loaded_flag(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text(
        json.dumps({"sku_order_max_qty": {"SKU-W": 30}}), encoding="utf-8"
    )
    limits, loaded = load_sku_order_max_qty(f)
    assert loaded is True
    assert limits.get("sku-w") == 30


def test_load_sku_order_max_qty_missing_file_returns_empty(tmp_path: Path):
    missing = tmp_path / "no_file.json"
    limits, loaded = load_sku_order_max_qty(missing)
    assert limits == {}
    assert loaded is False


# ---------------------------------------------------------------------------
# load_constraints — non-integer float rejected
# ---------------------------------------------------------------------------


def test_load_constraints_float_max_qty_raises(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text(
        json.dumps({"sku_order_max_qty": {"SKU-F": 10.5}}), encoding="utf-8"
    )
    with pytest.raises(ValueError, match="must be an integer"):
        load_constraints(f)


# ---------------------------------------------------------------------------
# load_constraints — invalid JSON raises
# ---------------------------------------------------------------------------


def test_load_constraints_invalid_json_raises(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text("not valid json", encoding="utf-8")
    with pytest.raises(ValueError, match="Invalid JSON"):
        load_constraints(f)


# ---------------------------------------------------------------------------
# load_constraints — root must be object
# ---------------------------------------------------------------------------


def test_load_constraints_root_array_raises(tmp_path: Path):
    f = tmp_path / "constraints.json"
    f.write_text("[1, 2, 3]", encoding="utf-8")
    with pytest.raises(ValueError, match="root must be a JSON object"):
        load_constraints(f)
