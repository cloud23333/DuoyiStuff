from __future__ import annotations

from shipment_planner.models import ForecastMetrics


def test_forecast_metrics_has_distribution_fields():
    metrics = ForecastMetrics(
        strategy="正常",
        forecast_daily_sales=1.0,
        forecast_stocking_period_sales=7.0,
        predictive_mean=6.0,
    )
    assert metrics.predictive_mean == 6.0
