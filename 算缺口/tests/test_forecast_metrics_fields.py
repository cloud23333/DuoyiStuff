from __future__ import annotations

from shipment_planner.engine import compute_forecast_metrics
from shipment_planner.models import ForecastMetrics


def test_forecast_metrics_has_distribution_fields():
    metrics = ForecastMetrics(
        strategy="正常",
        forecast_daily_sales=1.0,
        forecast_stocking_period_sales=7.0,
        predictive_mean=6.0,
    )
    assert metrics.predictive_mean == 6.0


def test_hurdle_spike_does_not_inflate_forecast():
    flat = (5, 5, 6, 5, 4, 5, 6, 5, 4, 5)
    spiked = flat + (40, 38)
    flat_m = compute_forecast_metrics(daily_sales=flat, stocking_days=7)
    spiked_m = compute_forecast_metrics(daily_sales=spiked, stocking_days=7)
    assert flat_m.forecast_model == "hurdle"
    assert spiked_m.forecast_stocking_period_sales <= flat_m.forecast_stocking_period_sales * 1.1
