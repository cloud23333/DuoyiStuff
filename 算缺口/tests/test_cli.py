from __future__ import annotations

import pytest

from shipment_planner.cli import build_arg_parser


def test_trend_recent_days_argument_is_removed() -> None:
    parser = build_arg_parser()

    with pytest.raises(SystemExit):
        parser.parse_args(["--trend-recent-days", "3"])
