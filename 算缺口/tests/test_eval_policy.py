from __future__ import annotations

from shipment_planner import eval_policy
from shipment_planner.models import ForecastMetrics, SalesRecord


def _fake_metrics(stocking_period_sales: float):
    def fake(**kwargs):
        return ForecastMetrics(
            strategy="x",
            forecast_daily_sales=0.0,
            forecast_stocking_period_sales=stocking_period_sales,
        )

    return fake


def test_simulate_sku_arithmetic(monkeypatch):
    monkeypatch.setattr(eval_policy, "compute_forecast_metrics", _fake_metrics(10.0))
    target, actual, shipped, lost, leftover = eval_policy.simulate_sku(
        train=(1, 1, 1),
        test=(3, 3, 3),
        stock_in_warehouse=0.0,
        offset=0.0,
        alpha=0.8,
    )
    assert target == 8.0
    assert actual == 9
    assert shipped == 8
    assert lost == 1
    assert leftover == 0


def test_alpha_monotonicity(monkeypatch):
    monkeypatch.setattr(eval_policy, "compute_forecast_metrics", _fake_metrics(10.0))
    _, _, shipped_high, _, _ = eval_policy.simulate_sku(
        train=(1, 1, 1), test=(3, 3, 3), stock_in_warehouse=0.0, offset=0.0, alpha=1.0
    )
    _, _, shipped_low, _, _ = eval_policy.simulate_sku(
        train=(1, 1, 1), test=(3, 3, 3), stock_in_warehouse=0.0, offset=0.0, alpha=0.6
    )
    assert shipped_low <= shipped_high


def test_skip_rule_counts_short_history(monkeypatch):
    monkeypatch.setattr(eval_policy, "compute_forecast_metrics", _fake_metrics(10.0))
    rows, skipped, _ = eval_policy.run_policy_grid(
        daily_sales_by_key={("skc-1", "sku-1"): (2, 2, 2, 2, 2)},
        sales_by_key={},
        offsets=[0.0],
        alphas=[1.0],
        holdout_days=3,
        min_train_days=13,
    )
    assert rows == []
    assert skipped == 1


def test_grid_summary_fill_rate(monkeypatch):
    monkeypatch.setattr(eval_policy, "compute_forecast_metrics", _fake_metrics(10.0))
    sales = SalesRecord(2, "skc-1", "sku-1", "ZX1", 7, 0, 0, 0)
    rows, skipped, grid_summary = eval_policy.run_policy_grid(
        daily_sales_by_key={("skc-1", "sku-1"): (1, 1, 1, 3, 3, 3)},
        sales_by_key={("skc-1", "sku-1"): sales},
        offsets=[0.0],
        alphas=[0.8],
        holdout_days=3,
        min_train_days=3,
    )
    assert skipped == 0
    assert len(rows) == 1
    assert len(grid_summary) == 1
    cell = grid_summary[0]
    assert cell["evaluated_skus"] == 1
    expected_fill = 1 - cell["lost_units"] / rows[0].actual
    assert cell["fill_rate"] == round(expected_fill, 4)
