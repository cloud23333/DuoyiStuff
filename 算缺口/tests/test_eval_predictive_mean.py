from __future__ import annotations

from shipment_planner import eval_forecast
from shipment_planner.models import ForecastMetrics, SalesRecord


def test_eval_accuracy_uses_predictive_mean(monkeypatch):
    def fake(**kwargs):
        return ForecastMetrics(
            strategy="x", forecast_daily_sales=99.0,
            forecast_stocking_period_sales=99.0 * kwargs["stocking_days"],
            predictive_mean=2.0 * kwargs["stocking_days"],
        )

    monkeypatch.setattr(eval_forecast, "compute_forecast_metrics", fake)
    sales = [SalesRecord(2, "skc-1", "sku-1", "l1", 30, 10, 0, 0)]
    rows, _, _ = eval_forecast.run_holdout_eval(
        sales_records=sales,
        daily_sales_by_key={("skc-1", "sku-1"): (2, 2, 2, 2, 2, 2)},
        holdout_days=3, min_train_days=3,
    )
    # predictive_mean (2/day*3) drives the holdout total, NOT forecast_daily_sales(99)
    assert rows[0].forecast_holdout_total == 6.0
