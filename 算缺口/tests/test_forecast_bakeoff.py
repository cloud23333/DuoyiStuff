# tests/test_forecast_bakeoff.py
from __future__ import annotations

from shipment_planner.forecast_bakeoff import CANDIDATE_ESTIMATORS, run_bakeoff


def test_bakeoff_scores_every_candidate():
    daily_sales_by_key = {
        ("skc-1", "sku-1"): tuple([3, 4, 0, 5, 3, 4, 0, 5, 3, 4, 0, 5, 6, 2]),
        ("skc-2", "sku-2"): tuple([0, 0, 2, 0, 0, 3, 0, 0, 2, 0, 0, 3, 0, 4]),
    }
    summary = run_bakeoff(
        daily_sales_by_key=daily_sales_by_key,
        stock_by_key={},
        hot_by_key={},
        holdout_days=3,
        min_train_days=8,
        service_level=0.55,
    )
    assert set(summary["by_estimator"]) == set(CANDIDATE_ESTIMATORS)
    for stats in summary["by_estimator"].values():
        assert "wape" in stats
        assert "pinball" in stats
        assert "death_fill_rate" in stats
        assert "spike_fill_rate" in stats
