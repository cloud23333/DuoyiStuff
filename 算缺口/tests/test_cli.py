from __future__ import annotations

import pytest

from shipment_planner.cli import build_arg_parser, main


def test_trend_recent_days_argument_is_removed() -> None:
    parser = build_arg_parser()

    with pytest.raises(SystemExit):
        parser.parse_args(["--trend-recent-days", "3"])


def test_sales_weight_arguments_are_removed_from_help() -> None:
    help_text = build_arg_parser().format_help()

    assert "--sold30-weight" not in help_text
    assert "--sold7-weight" not in help_text


def test_cli_requires_temu_daily_sales_file_without_fallback() -> None:
    with pytest.raises(ValueError, match="Temu daily sales file is required"):
        main(["--orders", "orders.xlsx", "--sales", "sales.xlsx"])
