from __future__ import annotations

from datetime import datetime

import pytest

from shipment_planner.engine import (
    HOT_STYLE_GAP_MULTIPLIER,
    _build_key_states,
    _decision_reason,
    _line_change_ratio,
    _normalize_sales_weights,
    _target_ship_qty,
    build_recommendations,
)
from shipment_planner.models import OrderLine, SalesRecord


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def make_order_line(
    *,
    skc: str = "SKC001",
    skuid: str = "SKUID001",
    quantity: int = 10,
    row_number: int = 2,
    status: str = "待发货",
    internal_order_id: str = "ORD-001",
    order_time: datetime | None = None,
) -> OrderLine:
    return OrderLine(
        row_number=row_number,
        internal_order_id=internal_order_id,
        skc=skc,
        skuid=skuid,
        product_code="PROD001",
        order_sku="ORDER-SKU-001",
        status=status,
        order_time=order_time or datetime(2024, 1, 15, 10, 0, 0),
        quantity=quantity,
    )


def make_sales_record(
    *,
    skc: str = "SKC001",
    skuid: str = "SKUID001",
    sold30: int = 100,
    sold7: int = 30,
    stocking_days: float = 7.0,
    stock_in_warehouse: float = 50.0,
    pending_receive: float = 10.0,
    pending_ship: float = 5.0,
    is_hot_style: bool = False,
    system_sku: str = "SYS-SKU-001",
) -> SalesRecord:
    return SalesRecord(
        row_number=2,
        skc=skc,
        skuid=skuid,
        system_sku=system_sku,
        is_hot_style=is_hot_style,
        sold30=sold30,
        sold7=sold7,
        stocking_days=stocking_days,
        stock_in_warehouse=stock_in_warehouse,
        pending_receive=pending_receive,
        pending_ship=pending_ship,
    )


# ---------------------------------------------------------------------------
# _normalize_sales_weights
# ---------------------------------------------------------------------------


class TestNormalizeSalesWeights:
    def test_already_normalized_weights(self) -> None:
        w30, w7 = _normalize_sales_weights(0.2, 0.8)
        assert abs(w30 + w7 - 1.0) < 1e-9
        assert abs(w30 - 0.2) < 1e-9
        assert abs(w7 - 0.8) < 1e-9

    def test_unnormalized_weights_sum_to_one(self) -> None:
        w30, w7 = _normalize_sales_weights(1.0, 3.0)
        assert abs(w30 + w7 - 1.0) < 1e-9
        assert abs(w30 - 0.25) < 1e-9
        assert abs(w7 - 0.75) < 1e-9

    def test_equal_weights_normalized(self) -> None:
        w30, w7 = _normalize_sales_weights(1.0, 1.0)
        assert abs(w30 - 0.5) < 1e-9
        assert abs(w7 - 0.5) < 1e-9

    def test_zero_sold30_weight_valid(self) -> None:
        w30, w7 = _normalize_sales_weights(0.0, 1.0)
        assert w30 == 0.0
        assert w7 == 1.0

    def test_zero_sold7_weight_valid(self) -> None:
        w30, w7 = _normalize_sales_weights(1.0, 0.0)
        assert w30 == 1.0
        assert w7 == 0.0

    def test_both_zero_raises(self) -> None:
        with pytest.raises(ValueError, match="cannot both be 0"):
            _normalize_sales_weights(0.0, 0.0)

    def test_negative_weight_raises(self) -> None:
        with pytest.raises(ValueError, match="non-negative"):
            _normalize_sales_weights(-1.0, 1.0)


# ---------------------------------------------------------------------------
# _target_ship_qty
# ---------------------------------------------------------------------------


class TestTargetShipQty:
    def test_zero_demand_returns_zero(self) -> None:
        qty = _target_ship_qty(
            sold30=0, sold7=0, stocking_days=7.0, sold30_weight=0.2, sold7_weight=0.8
        )
        assert qty == 0.0

    def test_zero_stocking_days_returns_zero(self) -> None:
        qty = _target_ship_qty(
            sold30=100, sold7=30, stocking_days=0.0, sold30_weight=0.2, sold7_weight=0.8
        )
        assert qty == 0.0

    def test_normal_case(self) -> None:
        # sold30_daily = (0.2 * 100) / 30 ≈ 0.6667
        # sold7_daily  = (0.8 * 30) / 7  ≈ 3.4286
        # target = (0.6667 + 3.4286) * 7 ≈ 28.667
        qty = _target_ship_qty(
            sold30=100, sold7=30, stocking_days=7.0, sold30_weight=0.2, sold7_weight=0.8
        )
        expected = ((0.2 * 100) / 30 + (0.8 * 30) / 7) * 7.0
        assert abs(qty - expected) < 1e-6

    def test_only_sold30_contributes(self) -> None:
        qty = _target_ship_qty(
            sold30=60, sold7=0, stocking_days=30.0, sold30_weight=1.0, sold7_weight=0.0
        )
        # sold30_daily = (1.0 * 60) / 30 = 2.0; target = 2.0 * 30 = 60
        assert abs(qty - 60.0) < 1e-6


# ---------------------------------------------------------------------------
# _build_key_states
# ---------------------------------------------------------------------------


class TestBuildKeyStates:
    def _run(
        self,
        key_demand: dict,
        sales_by_key: dict,
        *,
        global_gap_multiplier: float = 1.0,
        sold30_weight: float = 0.2,
        sold7_weight: float = 0.8,
        zero_sold7_with_sold30_stockout_max_qty: int = 5,
        shipping_in_progress_by_key: dict | None = None,
    ):
        return _build_key_states(
            key_demand=key_demand,
            sales_by_key=sales_by_key,
            shipping_in_progress_by_key=shipping_in_progress_by_key or {},
            global_gap_multiplier=global_gap_multiplier,
            sold30_weight=sold30_weight,
            sold7_weight=sold7_weight,
            zero_sold7_with_sold30_stockout_max_qty=zero_sold7_with_sold30_stockout_max_qty,
        )

    def test_aggregates_order_qty_for_same_key(self) -> None:
        key = ("SKC1", "SKUID1")
        sales = make_sales_record(skc="SKC1", skuid="SKUID1", sold30=0, sold7=0, stock_in_warehouse=100.0)
        states = self._run(
            key_demand={key: 25},
            sales_by_key={key: sales},
        )
        assert states[key].order_qty_total == 25

    def test_different_keys_kept_separate(self) -> None:
        key_a = ("SKC_A", "SKUID_A")
        key_b = ("SKC_B", "SKUID_B")
        sales_a = make_sales_record(skc="SKC_A", skuid="SKUID_A", sold30=0, sold7=0, stock_in_warehouse=100.0)
        sales_b = make_sales_record(skc="SKC_B", skuid="SKUID_B", sold30=0, sold7=0, stock_in_warehouse=100.0)
        states = self._run(
            key_demand={key_a: 10, key_b: 20},
            sales_by_key={key_a: sales_a, key_b: sales_b},
        )
        assert states[key_a].order_qty_total == 10
        assert states[key_b].order_qty_total == 20
        assert states[key_a].skuid == "SKUID_A"
        assert states[key_b].skuid == "SKUID_B"

    def test_no_sales_record_gives_zero_gap(self) -> None:
        key = ("SKC_MISSING", "SKUID_MISSING")
        states = self._run(
            key_demand={key: 15},
            sales_by_key={},
        )
        assert states[key].gap == 0
        assert states[key].recommended_qty_total == 0

    def test_shipping_in_progress_reduces_gap(self) -> None:
        key = ("SKC_IP", "SKUID_IP")
        # With zero stock and positive demand, there should be a gap.
        # Adding shipping_in_progress increases available_stock, reducing the gap.
        sales = make_sales_record(
            skc="SKC_IP", skuid="SKUID_IP",
            sold30=300, sold7=70, stocking_days=7.0,
            stock_in_warehouse=0.0, pending_receive=0.0, pending_ship=0.0,
            is_hot_style=False,
        )
        states_no_progress = self._run(
            key_demand={key: 1000},
            sales_by_key={key: sales},
            shipping_in_progress_by_key={},
        )
        states_with_progress = self._run(
            key_demand={key: 1000},
            sales_by_key={key: sales},
            shipping_in_progress_by_key={key: 50},
        )
        gap_no_progress = states_no_progress[key].gap
        gap_with_progress = states_with_progress[key].gap
        assert gap_with_progress < gap_no_progress

    def test_hot_style_multiplier_applied(self) -> None:
        key = ("SKC_HOT", "SKUID_HOT")
        # Use a simple case where the gap is clearly computable.
        # sold30=0, sold7=0 => target_ship_qty=0, stock=0 => raw_gap=0 — useless.
        # Instead give enough demand but zero stock so gap > 0.
        sales_regular = make_sales_record(
            skc="SKC_HOT", skuid="SKUID_HOT",
            sold30=300, sold7=70, stocking_days=7.0,
            stock_in_warehouse=0.0, pending_receive=0.0, pending_ship=0.0,
            is_hot_style=False,
        )
        sales_hot = make_sales_record(
            skc="SKC_HOT", skuid="SKUID_HOT",
            sold30=300, sold7=70, stocking_days=7.0,
            stock_in_warehouse=0.0, pending_receive=0.0, pending_ship=0.0,
            is_hot_style=True,
        )
        states_regular = self._run(
            key_demand={key: 1000},
            sales_by_key={key: sales_regular},
        )
        states_hot = self._run(
            key_demand={key: 1000},
            sales_by_key={key: sales_hot},
        )
        gap_regular = states_regular[key].gap
        gap_hot = states_hot[key].gap
        # gap_hot should be roughly HOT_STYLE_GAP_MULTIPLIER * gap_regular (both ceil'd)
        assert gap_hot > gap_regular


# ---------------------------------------------------------------------------
# _line_change_ratio
# ---------------------------------------------------------------------------


class TestLineChangeRatio:
    def test_zero_original_returns_zero(self) -> None:
        assert _line_change_ratio(0, 5) == 0.0

    def test_no_change_returns_zero(self) -> None:
        assert _line_change_ratio(10, 10) == 0.0

    def test_full_reduction_to_zero(self) -> None:
        ratio = _line_change_ratio(10, 0)
        assert abs(ratio - 1.0) < 1e-9

    def test_partial_change(self) -> None:
        ratio = _line_change_ratio(10, 7)
        assert abs(ratio - 0.3) < 1e-9

    def test_doubled_quantity_gives_ratio_one(self) -> None:
        # The formula uses abs(original - new) / original, so doubling (5→10)
        # gives abs(5-10)/5 = 1.0, the same ratio as a full reduction (10→0).
        ratio = _line_change_ratio(5, 10)
        assert abs(ratio - 1.0) < 1e-9


# ---------------------------------------------------------------------------
# _decision_reason
# ---------------------------------------------------------------------------


class TestDecisionReason:
    def test_zero_suggested_returns_hold(self) -> None:
        assert _decision_reason(10, 0) == "hold"

    def test_negative_suggested_returns_hold(self) -> None:
        assert _decision_reason(10, -1) == "hold"

    def test_full_ship_returns_ship_all(self) -> None:
        assert _decision_reason(10, 10) == "ship_all"

    def test_over_suggested_returns_ship_all(self) -> None:
        assert _decision_reason(10, 15) == "ship_all"

    def test_partial_returns_ship_partial(self) -> None:
        assert _decision_reason(10, 5) == "ship_partial"


# ---------------------------------------------------------------------------
# Allocation with zero-gap SKU (via build_recommendations)
# ---------------------------------------------------------------------------


class TestZeroGapAllocation:
    def test_zero_gap_sku_gets_zero_recommended_zero_demand(self) -> None:
        # Zero demand (sold30=0, sold7=0) causes target_ship_qty=0,
        # which causes gap=0 regardless of stock level.
        order = make_order_line(quantity=20)
        sales = make_sales_record(
            sold30=0, sold7=0, stocking_days=7.0,
            stock_in_warehouse=1000.0, pending_receive=0.0, pending_ship=0.0,
        )
        recs, _, _ = build_recommendations(
            order_lines=[order],
            sales_records=[sales],
            min_order_ship_qty=1,
        )
        assert len(recs) == 1
        assert recs[0]["recommended_ship"] == 0
        assert recs[0]["gap"] == 0

    def test_zero_gap_sku_gets_zero_recommended_stock_surplus(self) -> None:
        # Positive demand but large stock surplus: available_stock > target_ship_qty,
        # so raw_gap clamps to 0 and nothing is recommended.
        order = make_order_line(quantity=20)
        sales = make_sales_record(
            sold30=30, sold7=10, stocking_days=7.0,
            stock_in_warehouse=10000.0, pending_receive=0.0, pending_ship=0.0,
        )
        recs, _, _ = build_recommendations(
            order_lines=[order],
            sales_records=[sales],
            min_order_ship_qty=1,
        )
        assert len(recs) == 1
        assert recs[0]["recommended_ship"] == 0
        assert recs[0]["gap"] == 0


# ---------------------------------------------------------------------------
# Exempt eligibility bypasses min_order_ship_qty threshold
# ---------------------------------------------------------------------------


class TestExemptEligibility:
    def test_exempt_flag_set_when_sold7_zero_sold30_positive_stockout(self) -> None:
        # sold7=0, sold30>0, and no available stock => exempt eligible
        order = make_order_line(quantity=3)
        sales = make_sales_record(
            sold30=10, sold7=0, stocking_days=7.0,
            stock_in_warehouse=0.0, pending_receive=0.0, pending_ship=0.0,
        )
        recs, _, _ = build_recommendations(
            order_lines=[order],
            sales_records=[sales],
            min_order_ship_qty=10,  # threshold higher than order qty
            zero_sold7_with_sold30_stockout_max_qty=5,
        )
        assert len(recs) == 1
        assert recs[0]["min_order_ship_qty_exempt_eligible"] is True

    def test_no_exempt_flag_when_stock_available(self) -> None:
        # sold7=0, sold30>0, but stock is available => NOT exempt eligible
        order = make_order_line(quantity=3)
        sales = make_sales_record(
            sold30=10, sold7=0, stocking_days=7.0,
            stock_in_warehouse=100.0, pending_receive=0.0, pending_ship=0.0,
        )
        recs, _, _ = build_recommendations(
            order_lines=[order],
            sales_records=[sales],
            min_order_ship_qty=10,
        )
        assert recs[0]["min_order_ship_qty_exempt_eligible"] is False

    def test_no_exempt_flag_when_sold7_nonzero(self) -> None:
        order = make_order_line(quantity=3)
        sales = make_sales_record(
            sold30=10, sold7=5, stocking_days=7.0,
            stock_in_warehouse=0.0, pending_receive=0.0, pending_ship=0.0,
        )
        recs, _, _ = build_recommendations(
            order_lines=[order],
            sales_records=[sales],
            min_order_ship_qty=10,
        )
        assert recs[0]["min_order_ship_qty_exempt_eligible"] is False
