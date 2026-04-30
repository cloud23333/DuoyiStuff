from __future__ import annotations

from shipment_planner import eval_forecast
from shipment_planner.eval_forecast import build_arg_parser
from shipment_planner.forecast_benchmark import (
    DEFAULT_BENCHMARK_MODELS,
    forecast_with_model,
    run_forecast_benchmark,
)
from shipment_planner.models import ForecastMetrics, SalesRecord


def test_intermittent_benchmark_models_do_not_collapse_to_zero():
    daily_sales = (0, 0, 4, 0, 0, 3, 0, 0, 5, 0)

    croston = forecast_with_model("croston_sba", daily_sales, holdout_days=7)
    tsb = forecast_with_model("tsb", daily_sales, holdout_days=7)
    adida = forecast_with_model("adida", daily_sales, holdout_days=7)
    imapa = forecast_with_model("imapa", daily_sales, holdout_days=7)

    assert 0 < croston < 7
    assert 0 < tsb < 7
    assert 0 < adida < 7
    assert 0 < imapa < 7


def test_run_forecast_benchmark_compares_each_model():
    sales_records = [
        SalesRecord(
            row_number=2,
            skc="skc-1",
            skuid="sku-1",
            system_sku="local-1",
            is_hot_style=False,
            stocking_days=7,
            stock_in_warehouse=10,
            pending_receive=0,
            pending_ship=0,
        )
    ]
    daily_sales_by_key = {
        ("skc-1", "sku-1"): (0, 0, 3, 0, 0, 4, 0, 0, 2, 0, 0, 3, 0, 0, 4)
    }

    rows, skipped, summary = run_forecast_benchmark(
        sales_records=sales_records,
        daily_sales_by_key=daily_sales_by_key,
        holdout_days=3,
        min_train_days=8,
    )

    assert skipped == 0
    assert {row.model for row in rows} == set(DEFAULT_BENCHMARK_MODELS)
    assert set(summary["by_model"]) == set(DEFAULT_BENCHMARK_MODELS)
    assert summary["evaluated_skus"] == 1
    assert summary["model_rows"] == len(DEFAULT_BENCHMARK_MODELS)


def test_eval_parser_accepts_benchmark_flag():
    args = build_arg_parser().parse_args(
        [
            "--sales",
            "sales.xlsx",
            "--temu-sales",
            "daily.xlsx",
            "--benchmark",
        ]
    )

    assert args.benchmark is True


def test_holdout_eval_forecasts_the_holdout_horizon(monkeypatch):
    seen_stocking_days: list[float] = []

    def fake_compute_forecast_metrics(**kwargs):
        seen_stocking_days.append(kwargs["stocking_days"])
        return ForecastMetrics(
            strategy="test",
            forecast_daily_sales=1.0,
            forecast_stocking_period_sales=kwargs["stocking_days"],
        )

    monkeypatch.setattr(
        eval_forecast,
        "compute_forecast_metrics",
        fake_compute_forecast_metrics,
    )
    sales_records = [
        SalesRecord(
            row_number=2,
            skc="skc-1",
            skuid="sku-1",
            system_sku="local-1",
            is_hot_style=False,
            stocking_days=30,
            stock_in_warehouse=10,
            pending_receive=0,
            pending_ship=0,
        )
    ]

    eval_forecast.run_holdout_eval(
        sales_records=sales_records,
        daily_sales_by_key={("skc-1", "sku-1"): (1, 2, 3, 4, 5, 6, 7, 8)},
        holdout_days=3,
        min_train_days=3,
    )

    assert seen_stocking_days == [3]
