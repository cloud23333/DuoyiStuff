from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from shipment_planner.models import OrderLine, SalesRecord


@pytest.fixture
def sample_order_line() -> OrderLine:
    """A minimal valid OrderLine instance for use in tests."""
    return OrderLine(
        row_number=2,
        internal_order_id="ORD-001",
        skc="SKC001",
        skuid="SKUID001",
        product_code="PROD001",
        order_sku="ORDER-SKU-001",
        status="待发货",
        order_time=datetime(2024, 1, 15, 10, 0, 0),
        quantity=5,
    )


@pytest.fixture
def sample_sales_record() -> SalesRecord:
    """A minimal valid SalesRecord instance for use in tests."""
    return SalesRecord(
        row_number=2,
        skc="SKC001",
        skuid="SKUID001",
        system_sku="SYS-SKU-001",
        is_hot_style=False,
        sold30=100,
        sold7=30,
        stocking_days=7.0,
        stock_in_warehouse=50.0,
        pending_receive=10.0,
        pending_ship=5.0,
    )


@pytest.fixture
def output_dir(tmp_path: Path) -> Path:
    """A temporary directory for test output files."""
    out = tmp_path / "output"
    out.mkdir()
    return out
