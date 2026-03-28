from __future__ import annotations

import argparse
from pathlib import Path
from typing import TypeAlias

from .constraints import DEFAULT_CONSTRAINTS_FILENAME, load_constraints
from .engine import (
    DEFAULT_ZERO_SOLD7_WITH_SOLD30_STOCKOUT_MAX_QTY,
    DEFAULT_SOLD30_WEIGHT,
    DEFAULT_SOLD7_WEIGHT,
    build_recommendations,
)
from .parsers import (
    ORDER_REQUIRED_COLUMNS,
    SALES_REQUIRED_COLUMNS,
    assert_required_columns,
    assert_xlsx,
    missing_required_columns,
    parse_orders,
    parse_sales,
)
from .reports import export_reports
from .xlsx_reader import read_xlsx_table

_XlsxData: TypeAlias = tuple[list[str], list[dict[str, str]]]
HeaderCache: TypeAlias = dict[Path, _XlsxData | None]


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Shipment suggestion app (xlsx-only).")
    parser.add_argument("--orders", help="Orders xlsx path (optional with auto-detect)")
    parser.add_argument("--sales", help="Sales xlsx path (optional with auto-detect)")
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
        "--sold30-weight",
        type=float,
        default=DEFAULT_SOLD30_WEIGHT,
        help=f"Weight for sold30 in blended demand (default: {DEFAULT_SOLD30_WEIGHT})",
    )
    parser.add_argument(
        "--sold7-weight",
        type=float,
        default=DEFAULT_SOLD7_WEIGHT,
        help=f"Weight for sold7 in blended demand (default: {DEFAULT_SOLD7_WEIGHT})",
    )
    parser.add_argument(
        "--zero-sold7-with-sold30-stockout-max-qty",
        type=int,
        default=DEFAULT_ZERO_SOLD7_WITH_SOLD30_STOCKOUT_MAX_QTY,
        help=(
            "Max suggested qty when sold7=0, sold30>0, and available stock is 0 "
            f"(default: {DEFAULT_ZERO_SOLD7_WITH_SOLD30_STOCKOUT_MAX_QTY})"
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

    cached_orders = header_cache.get(orders_path)
    order_header, order_rows = cached_orders if cached_orders is not None else read_xlsx_table(orders_path)
    cached_sales = header_cache.get(sales_path)
    sales_header, sales_rows = cached_sales if cached_sales is not None else read_xlsx_table(sales_path)

    assert_required_columns(order_header, ORDER_REQUIRED_COLUMNS, "orders file")
    assert_required_columns(sales_header, SALES_REQUIRED_COLUMNS, "sales file")

    order_lines, shipping_in_progress_by_key = parse_orders(order_rows)
    sales_records = parse_sales(sales_rows)
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
        zero_sold7_with_sold30_stockout_max_qty=(
            args.zero_sold7_with_sold30_stockout_max_qty
        ),
        sku_order_max_qty=sku_order_max_qty,
        exclude_skc=exclude_skc,
        exclude_skuid=exclude_skuid,
        shipping_in_progress_by_key=shipping_in_progress_by_key,
        global_gap_multiplier=global_gap_multiplier,
        sold30_weight=args.sold30_weight,
        sold7_weight=args.sold7_weight,
    )
    outputs = export_reports(args.out_dir, recommendations, quality_rows, summary)

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
    candidates.sort(key=lambda path: path.stat().st_mtime, reverse=True)

    if not candidates:
        raise ValueError(f"No .xlsx files found in input directory: {input_dir}")
    return candidates


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
    print(f"Recommended qty: {summary['total_recommended_qty']}")
    print(f"30%-change keep lines: {summary['small_change_kept_lines']}")
    print(f"Global gap multiplier: {summary['global_gap_multiplier']}")
    print(
        "Sales weights: "
        f"sold30={summary['sold30_weight']}, "
        f"sold7={summary['sold7_weight']}"
    )
    print(
        "Zero-sold7 stockout cap: "
        f"{summary['zero_sold7_with_sold30_stockout_max_qty']}"
    )
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
