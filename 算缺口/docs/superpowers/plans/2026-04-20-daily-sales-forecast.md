# Daily Sales Forecast Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace 7/30 weighted shortage forecasting with mandatory Temu daily-sales forecasting and report the selected strategy.

**Architecture:** Keep legacy sales rows as the source for SKU identity, stocking days, inventory, hot-style flag, and reference sold30/sold7 columns. Add daily-sales validation and a small deterministic forecasting helper inside the engine, then pass parsed Temu daily sales from CLI/UI into `build_recommendations()`.

**Tech Stack:** Python 3.10+, pytest, existing csv/json report exporters, existing PyQt6 UI.

---

### Task 1: Parser And CLI Guardrails

**Files:**
- Modify: `src/shipment_planner/parsers.py`
- Modify: `src/shipment_planner/cli.py`
- Test: `tests/test_parsers.py`
- Test: `tests/test_cli.py`

- [ ] **Step 1: Write failing parser tests**

```python
def test_parse_temu_daily_sales_requires_identity_columns() -> None:
    with pytest.raises(ValueError, match="Missing required columns in Temu daily sales file"):
        parse_temu_daily_sales([{"平台SKC_ID": "skc-1", "4月1日销量": "1"}])


def test_parse_temu_daily_sales_requires_daily_columns() -> None:
    with pytest.raises(ValueError, match="no daily sales columns"):
        parse_temu_daily_sales([{"平台SKC_ID": "skc-1", "平台SKU_ID": "sku-1"}])
```

- [ ] **Step 2: Run failing parser tests**

Run: `python3 -m pytest tests/test_parsers.py -v`

Expected: FAIL because `parse_temu_daily_sales()` currently returns empty/zero-date tuples instead of raising.

- [ ] **Step 3: Write failing CLI tests**

```python
def test_cli_help_removes_sales_weight_arguments() -> None:
    help_text = build_arg_parser().format_help()

    assert "--sold30-weight" not in help_text
    assert "--sold7-weight" not in help_text
```

- [ ] **Step 4: Implement parser and CLI guardrails**

Implementation points:
- In `parse_temu_daily_sales()`, derive the union of headers from all rows, require `平台SKC_ID`, `平台SKU_ID`, and at least one `x月x日销量` column.
- In `build_arg_parser()`, remove `--sold30-weight` and `--sold7-weight`; keep `--temu-sales` but update help to state it is required for forecasting.
- In `main()`, raise `ValueError("Temu daily sales file is required.")` when no explicit or auto-detected Temu file is available.

- [ ] **Step 5: Verify task**

Run: `python3 -m pytest tests/test_parsers.py tests/test_cli.py -v`

Expected: PASS.

### Task 2: Daily-Sales Forecast Engine

**Files:**
- Modify: `src/shipment_planner/models.py`
- Modify: `src/shipment_planner/engine.py`
- Test: `tests/test_engine.py`

- [ ] **Step 1: Write failing engine tests**

Add tests for these behaviors:
- `build_recommendations(..., daily_sales_by_key=None)` raises `ValueError("daily sales data is required")`.
- A missing `(SKC, SKUID)` in daily sales raises `ValueError("Missing daily sales data")`.
- Stable daily sales produce `forecast_strategy == "正常"`.
- Rising daily sales produce `forecast_strategy == "激进"`.
- Falling or volatile daily sales produce `forecast_strategy == "保守"`.
- Daily forecast changes `gap` and `recommended_ship` from inventory/stocking-day math.

- [ ] **Step 2: Run failing engine tests**

Run: `python3 -m pytest tests/test_engine.py -v`

Expected: FAIL because the engine has no daily-sales requirement, strategy, or forecast fields.

- [ ] **Step 3: Implement engine model**

Implementation points:
- Add `ForecastMetrics` dataclass with `strategy`, `forecast_daily_sales`, and `forecast_stocking_period_sales`.
- Add a `forecast_metrics` field to `KeyState`.
- Add `daily_sales_by_key` as a required keyword argument to `build_recommendations()`.
- Validate all demanded keys have non-empty daily sales.
- Replace `_target_ship_qty()` with `_forecast_daily_sales()` and compute `raw_gap = forecast_stocking_period_sales - available_stock`.
- Keep hot-style and global gap multipliers.
- Remove weight constants and `_normalize_sales_weights()`.

- [ ] **Step 4: Verify engine tests**

Run: `python3 -m pytest tests/test_engine.py -v`

Expected: PASS.

### Task 3: Reports, Summary, UI, And Workflow

**Files:**
- Modify: `src/shipment_planner/reports.py`
- Modify: `src/shipment_planner/summary.py`
- Modify: `src/planner_ui/workflow.py`
- Modify: `src/planner_ui/app.py`
- Test: `tests/test_engine.py`
- Test: `tests/test_cli.py`

- [ ] **Step 1: Write failing report assertion**

Assert recommendation rows include `forecast_strategy`, `forecast_daily_sales`, and `forecast_stocking_period_sales`.

- [ ] **Step 2: Implement report and summary updates**

Implementation points:
- Add `forecast_strategy` after `sold7`.
- Add `forecast_daily_sales` and `forecast_stocking_period_sales` as readable debug columns.
- Remove summary fields and console output for sold30/sold7 weights.
- Keep sold30/sold7 in recommendations as reference fields only.

- [ ] **Step 3: Implement UI and workflow updates**

Implementation points:
- Remove sold7/sold30 spin boxes, weight sync handlers, and validation.
- Require a Temu file selection in UI input readiness and validation.
- Pass `temu_sales_path` through `run_planner()` and CLI without weight arguments.
- Update Temu placeholder and log text to say the file is required.

- [ ] **Step 4: Verify all tests and syntax**

Run:
```bash
python3 -m compileall src
python3 -m pytest
```

Expected: both commands pass.
