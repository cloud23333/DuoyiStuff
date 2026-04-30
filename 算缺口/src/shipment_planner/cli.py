from __future__ import annotations

import argparse
import re
from datetime import datetime
from pathlib import Path
from typing import TypeAlias

from .constraints import DEFAULT_CONSTRAINTS_FILENAME, load_constraints
from .engine import build_recommendations
from .eval_forecast import (
    DEFAULT_HOLDOUT_DAYS,
    DEFAULT_MIN_TRAIN_DAYS,
    run_holdout_eval,
)
from .parsers import (
    ORDER_REQUIRED_COLUMNS,
    SALES_REQUIRED_COLUMNS,
    TEMU_DAILY_REQUIRED_COLUMNS,
    assert_required_columns,
    assert_temu_daily_sales_columns,
    assert_xlsx,
    missing_required_columns,
    parse_orders,
    parse_sales,
    parse_temu_daily_sales,
    temu_daily_sales_columns,
)
from .reports import export_reports
from .xlsx_reader import read_xlsx_table

_XlsxData: TypeAlias = tuple[list[str], list[dict[str, str]]]
HeaderCache: TypeAlias = dict[Path, _XlsxData | None]
_FILENAME_EXPORT_TIMESTAMP_RE = re.compile(r"(?<!\d)(?:19|20)\d{12}(?!\d)")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Shipment suggestion app (xlsx-only).")
    parser.add_argument("--orders", help="Orders xlsx path (optional with auto-detect)")
    parser.add_argument("--sales", help="Sales xlsx path (optional with auto-detect)")
    parser.add_argument(
        "--temu-sales",
        help="Temu daily sales xlsx path (required for daily-sales forecasting)",
    )
    parser.add_argument(
        "--input-dir",
        default="data/input",
        help="Directory for auto-detecting input xlsx files (default: data/input)",
    )
    parser.add_argument(
        "--out-dir",
        default="data/output",
        help="Output directory for reports (default: data/output)",
    )
    parser.add_argument(
        "--min-order-ship-qty",
        type=int,
        default=10,
        help="Low-qty blocking threshold by order total suggested qty (default: 10)",
    )
    parser.add_argument(
        "--constraints",
        help=(
            "Constraints JSON path. If omitted, app auto-loads "
            f"{DEFAULT_CONSTRAINTS_FILENAME} from --input-dir when present."
        ),
    )
    parser.add_argument(
        "--global-gap-multiplier",
        type=float,
        default=1.0,
        help="Global multiplier applied to key gaps after hot-style rule (default: 1.0)",
    )
    parser.add_argument(
        "--eval-holdout-days",
        type=int,
        default=DEFAULT_HOLDOUT_DAYS,
        help=(
            "Trailing days held out for forecast-quality eval appended to the "
            f"summary and detail rows (default: {DEFAULT_HOLDOUT_DAYS})."
        ),
    )
    parser.add_argument(
        "--eval-min-train-days",
        type=int,
        default=DEFAULT_MIN_TRAIN_DAYS,
        help=(
            "Skip forecast eval for SKUs with fewer training days than this "
            f"(default: {DEFAULT_MIN_TRAIN_DAYS})."
        ),
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    input_dir = Path(args.input_dir)
    needs_auto_detect = args.orders is None or args.sales is None
    candidates = _list_xlsx_candidates(input_dir) if needs_auto_detect else []
    header_cache: HeaderCache = {}

    orders_path = _resolve_input_path(
        explicit_path=args.orders,
        candidates=candidates,
        header_cache=header_cache,
        required_columns=ORDER_REQUIRED_COLUMNS,
        label="orders file",
    )
    sales_path = _resolve_input_path(
        explicit_path=args.sales,
        candidates=candidates,
        header_cache=header_cache,
        required_columns=SALES_REQUIRED_COLUMNS,
        label="sales file",
    )

    assert_xlsx(orders_path)
    assert_xlsx(sales_path)

    if args.orders is None:
        print(f"Auto-selected orders file: {orders_path}")
    if args.sales is None:
        print(f"Auto-selected sales file: {sales_path}")

    # Resolve required Temu daily sales file.
    temu_sales_path: Path | None = None
    if args.temu_sales is not None:
        temu_sales_path = Path(args.temu_sales)
        assert_xlsx(temu_sales_path)
    else:
        temu_sales_path = _try_auto_detect_temu(candidates, header_cache)
        if temu_sales_path is not None:
            print(f"Auto-selected Temu daily sales file: {temu_sales_path}")
    if temu_sales_path is None:
        raise ValueError("Temu daily sales file is required.")

    cached_orders = header_cache.get(orders_path)
    order_header, order_rows = cached_orders if cached_orders is not None else read_xlsx_table(orders_path)
    cached_sales = header_cache.get(sales_path)
    sales_header, sales_rows = cached_sales if cached_sales is not None else read_xlsx_table(sales_path)

    assert_required_columns(order_header, ORDER_REQUIRED_COLUMNS, "orders file")
    assert_required_columns(sales_header, SALES_REQUIRED_COLUMNS, "sales file")

    order_lines, shipping_in_progress_by_key = parse_orders(order_rows)
    sales_records = parse_sales(sales_rows)

    cached_temu = header_cache.get(temu_sales_path)
    temu_header, temu_rows = (
        cached_temu if cached_temu is not None else read_xlsx_table(temu_sales_path)
    )
    assert_temu_daily_sales_columns(temu_header, "Temu daily sales file")
    daily_sales_by_key = parse_temu_daily_sales(temu_rows)
    print(f"Loaded Temu daily sales: {len(daily_sales_by_key)} SKUs")
    constraints_path = (
        Path(args.constraints)
        if args.constraints is not None
        else input_dir / DEFAULT_CONSTRAINTS_FILENAME
    )
    sku_order_max_qty, exclude_skc, exclude_skuid, constraints_loaded = load_constraints(
        constraints_path,
        strict=(args.constraints is not None),
    )
    global_gap_multiplier = args.global_gap_multiplier

    recommendations, quality_rows, summary = build_recommendations(
        order_lines,
        sales_records,
        min_order_ship_qty=args.min_order_ship_qty,
        sku_order_max_qty=sku_order_max_qty,
        exclude_skc=exclude_skc,
        exclude_skuid=exclude_skuid,
        shipping_in_progress_by_key=shipping_in_progress_by_key,
        global_gap_multiplier=global_gap_multiplier,
        daily_sales_by_key=daily_sales_by_key,
    )
    eval_rows, _eval_skipped, eval_summary = run_holdout_eval(
        sales_records=sales_records,
        daily_sales_by_key=daily_sales_by_key,
        holdout_days=args.eval_holdout_days,
        min_train_days=args.eval_min_train_days,
    )
    outputs = export_reports(
        args.out_dir,
        recommendations,
        quality_rows,
        summary,
        daily_sales_by_key=daily_sales_by_key,
        eval_rows=eval_rows,
        eval_summary=eval_summary,
    )

    _print_run_summary(
        summary=summary,
        outputs=outputs,
        constraints_loaded=constraints_loaded,
        constraints_path=constraints_path,
    )

    return 0


def _list_xlsx_candidates(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        raise ValueError(f"Input directory not found: {input_dir}")

    candidates = [path for path in input_dir.iterdir() if path.is_file() and path.suffix.lower() == ".xlsx"]
    candidates.sort(key=_xlsx_candidate_recency_key, reverse=True)

    if not candidates:
        raise ValueError(f"No .xlsx files found in input directory: {input_dir}")
    return candidates


def _xlsx_candidate_recency_key(path: Path) -> tuple[str, float]:
    match = _FILENAME_EXPORT_TIMESTAMP_RE.search(path.name)
    timestamp = (
        match.group(0)
        if match is not None
        else datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y%m%d%H%M%S")
    )
    return timestamp, path.stat().st_mtime


def _auto_select_xlsx(
    candidates: list[Path],
    header_cache: HeaderCache,
    required_columns: list[str],
    label: str,
) -> Path:
    for candidate in candidates:
        header = _cached_header(candidate, header_cache)
        if header is None:
            continue
        if _contains_all_required_columns(header, required_columns):
            return candidate

    searched_files = ", ".join(path.name for path in candidates)
    raise ValueError(
        f"Could not auto-detect {label}. "
        "Please pass it explicitly via CLI arguments. "
        f"Searched files: {searched_files}"
    )


def _resolve_input_path(
    *,
    explicit_path: str | None,
    candidates: list[Path],
    header_cache: HeaderCache,
    required_columns: list[str],
    label: str,
) -> Path:
    if explicit_path is not None:
        return Path(explicit_path)
    return _auto_select_xlsx(candidates, header_cache, required_columns, label)


def _cached_header(path: Path, header_cache: HeaderCache) -> list[str] | None:
    if path in header_cache:
        cached = header_cache[path]
        return cached[0] if cached is not None else None

    try:
        header, rows = read_xlsx_table(path)
    except Exception:
        header_cache[path] = None
        return None

    header_cache[path] = (header, rows)
    return header


def _contains_all_required_columns(header: list[str], required_columns: list[str]) -> bool:
    return not missing_required_columns(header, required_columns)


def _try_auto_detect_temu(
    candidates: list[Path],
    header_cache: HeaderCache,
) -> Path | None:
    """Return the first candidate that looks like a Temu daily sales file, or None."""
    for candidate in candidates:
        header = _cached_header(candidate, header_cache)
        if header is None:
            continue
        if (
            _contains_all_required_columns(header, TEMU_DAILY_REQUIRED_COLUMNS)
            and temu_daily_sales_columns(header)
        ):
            return candidate
    return None


def _print_run_summary(
    *,
    summary: dict[str, object],
    outputs: dict[str, Path],
    constraints_loaded: bool,
    constraints_path: Path,
) -> None:
    print("Run completed.")
    print(f"Orders rows: {summary['order_lines']}")
    print(f"Sales rows: {summary['sales_rows']}")
    print(
        "Join coverage: "
        f"{summary['matched_order_lines']}/{summary['order_lines']} "
        f"({summary['join_coverage_pct']}%)"
    )
    print(f"Quality issue rows: {summary['quality_issue_rows']}")
    print(f"Recommended qty: {summary['total_recommended_qty']}")
    print(f"30%-change keep lines: {summary['small_change_kept_lines']}")
    print(f"Demand profiles: {summary['demand_profile_summary']}")
    print(f"Anomaly flags: {summary['anomaly_flag_summary']}")
    print(f"Service levels: {summary['service_level_summary']}")
    print(f"Forecast models: {summary['forecast_model_summary']}")
    print(f"Global gap multiplier: {summary['global_gap_multiplier']}")
    print(f"Min order ship qty threshold: {summary['min_order_ship_qty_threshold']}")
    print(f"Low-qty orders before exemption: {summary['low_qty_orders_before_exempt']}")
    print(f"Low-qty orders exempted: {summary['low_qty_orders_exempted']}")
    print(f"Low-qty orders blocked: {summary['low_qty_orders']}")

    if constraints_loaded:
        print(
            "SKU-per-order limits: "
            f"{summary['sku_order_limit_rule_count']} rules "
            f"(capped lines: {summary['sku_order_limit_capped_lines']}) "
            f"from {constraints_path}"
        )
        print(
            "SKC/SKUID exclusions: "
            f"{summary['excluded_skc_rule_count']} SKC, "
            f"{summary['excluded_skuid_rule_count']} SKUID "
            f"(intercepted lines: {summary['intercepted_order_lines']}, "
            f"intercepted orders: {summary['intercepted_orders']})"
        )
    else:
        print("SKU-per-order limits: none")
        print("SKC/SKUID exclusions: none")

    print(f"Outputs: {outputs['final_recommendation']}")
    print(f"         {outputs['quality_report']}")
    print(f"         {outputs['run_summary']}")


if __name__ == "__main__":
    raise SystemExit(main())
