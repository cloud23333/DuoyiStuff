from __future__ import annotations

import contextlib
from dataclasses import dataclass
from datetime import datetime
import io
import json
from pathlib import Path
import sys

from shipment_planner.cli import main as planner_cli_main
from shipment_planner.constraints import DEFAULT_CONSTRAINTS_FILENAME
from shipment_planner.engine import DEFAULT_ZERO_SOLD7_WITH_SOLD30_STOCKOUT_MAX_QTY
from shipment_planner.parsers import (
    ORDER_REQUIRED_COLUMNS,
    SALES_REQUIRED_COLUMNS,
    describe_required_column,
    missing_required_columns,
)
from shipment_planner.xlsx_reader import read_xlsx_table


@dataclass(slots=True)
class PlannerRunResult:
    output_dir: Path
    recommendation_path: Path
    quality_path: Path
    summary_path: Path
    console_output: str
    constraints_path: Path
    constraints_template_created: bool


def get_constraints_config_dir() -> Path:
    base_dir = _resolve_app_base_dir()
    if getattr(sys, "frozen", False):
        return base_dir
    return base_dir / "data" / "input"


def get_constraints_path() -> Path:
    return get_constraints_config_dir() / DEFAULT_CONSTRAINTS_FILENAME


def ensure_constraints_template() -> tuple[Path, bool]:
    constraints_path = get_constraints_path()
    constraints_path.parent.mkdir(parents=True, exist_ok=True)
    if constraints_path.exists():
        return constraints_path, False

    template = {
        "sku_order_max_qty": {},
        "exclude_skc": [],
        "exclude_skuid": [],
    }
    constraints_path.write_text(
        json.dumps(template, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return constraints_path, True


def extract_unique_skc(order_path: str | Path) -> list[str]:
    path = Path(order_path)
    _assert_existing_file(path, "订单文件")
    _assert_xlsx(path, "订单文件")

    header, rows = read_xlsx_table(path)
    _assert_required_columns(header, ORDER_REQUIRED_COLUMNS, "订单文件")

    seen: set[str] = set()
    unique_codes: list[str] = []
    for row in rows:
        skc = (row.get("店铺款式编码") or "").strip()
        if not skc or skc in seen:
            continue
        seen.add(skc)
        unique_codes.append(skc)
    return unique_codes


def run_planner(
    *,
    orders_path: str | Path,
    sales_path: str | Path,
    output_dir: str | Path,
    sold30_weight: float,
    sold7_weight: float,
    global_gap_multiplier: float,
    zero_sold7_with_sold30_stockout_max_qty: int = DEFAULT_ZERO_SOLD7_WITH_SOLD30_STOCKOUT_MAX_QTY,
) -> PlannerRunResult:
    orders = Path(orders_path)
    sales = Path(sales_path)
    base_output_dir = Path(output_dir)
    input_specs = (
        (orders, ORDER_REQUIRED_COLUMNS, "订单文件"),
        (sales, SALES_REQUIRED_COLUMNS, "销售文件"),
    )

    for file_path, _, label in input_specs:
        _assert_existing_file(file_path, label)
        _assert_xlsx(file_path, label)

    headers: list[list[str]] = []
    for file_path, _, _ in input_specs:
        header, _ = read_xlsx_table(file_path)
        headers.append(header)
    for header, (_, required_columns, label) in zip(headers, input_specs):
        _assert_required_columns(header, required_columns, label)

    out_dir = _prepare_run_output_dir(base_output_dir)
    constraints_path, constraints_template_created = ensure_constraints_template()

    args = [
        "--orders",
        str(orders),
        "--sales",
        str(sales),
        "--out-dir",
        str(out_dir),
        "--constraints",
        str(constraints_path),
        "--sold30-weight",
        str(sold30_weight),
        "--sold7-weight",
        str(sold7_weight),
        "--global-gap-multiplier",
        str(global_gap_multiplier),
        "--zero-sold7-with-sold30-stockout-max-qty",
        str(zero_sold7_with_sold30_stockout_max_qty),
    ]

    stdout_buffer = io.StringIO()
    with contextlib.redirect_stdout(stdout_buffer):
        exit_code = planner_cli_main(args)
    if exit_code != 0:
        raise RuntimeError(f"发货建议程序异常退出，返回码：{exit_code}")

    recommendation = out_dir / "发货建议明细.csv"
    quality = out_dir / "数据质量报告.csv"
    summary_path = out_dir / "运行摘要.json"
    missing_files = [
        str(path)
        for path in (recommendation, quality, summary_path)
        if not path.exists()
    ]
    if missing_files:
        missing = ", ".join(missing_files)
        raise RuntimeError(f"运行完成但缺少输出文件：{missing}")

    localized_console_output = _build_localized_console_output(
        summary_path=summary_path,
        constraints_path=constraints_path,
        fallback_output=stdout_buffer.getvalue().strip(),
    )
    return PlannerRunResult(
        output_dir=out_dir,
        recommendation_path=recommendation,
        quality_path=quality,
        summary_path=summary_path,
        console_output=localized_console_output,
        constraints_path=constraints_path,
        constraints_template_created=constraints_template_created,
    )


def _assert_existing_file(path: Path, label: str) -> None:
    if not path.exists():
        raise FileNotFoundError(f"{label}不存在：{path}")
    if not path.is_file():
        raise ValueError(f"{label}不是文件：{path}")


def _assert_xlsx(path: Path, label: str) -> None:
    if path.suffix.lower() != ".xlsx":
        raise ValueError(f"{label}必须是 .xlsx 文件：{path}")


def _assert_required_columns(header: list[str], required: list[str], label: str) -> None:
    missing = missing_required_columns(header, required)
    if missing:
        details = ", ".join(describe_required_column(column_name) for column_name in missing)
        raise ValueError(f"{label}缺少必填列：{details}")


def _build_localized_console_output(
    *,
    summary_path: Path,
    constraints_path: Path,
    fallback_output: str,
) -> str:
    summary_data = _read_summary_json(summary_path)
    if summary_data is None:
        return fallback_output

    def summary_value(key: str) -> object:
        return summary_data.get(key, "-")

    order_lines = summary_value("订单行数")
    sales_rows = summary_value("销售行数")
    matched_lines = summary_value("匹配订单行数")
    join_coverage_pct = summary_value("匹配覆盖率_百分比")
    total_recommended = summary_value("建议发货总量")
    small_change_kept_lines = summary_value("触发30%免改行数")
    global_gap_multiplier = summary_value("全局缺口上浮系数")
    sold30_weight = summary_value("近30日销量占比")
    sold7_weight = summary_value("近7日销量占比")
    zero_sold7_stockout_cap = summary_value("零7日销量且30日有销量无库存发货上限")
    min_order_ship_qty = summary_value("最小发货阈值")
    low_qty_orders_before_exempt = summary_value("阈值前低发货量订单数")
    low_qty_orders = summary_value("低于阈值订单数_提示")
    low_qty_orders_exempted = summary_value("低于阈值豁免订单数")

    sku_limit_rule_count = summary_value("订单内SKU限额规则数")
    sku_limit_capped_lines = summary_value("触发订单内SKU限额行数")
    excluded_skc_rule_count = summary_value("SKC拦截规则数")
    excluded_skuid_rule_count = summary_value("SKUID拦截规则数")
    intercepted_order_lines = summary_value("命中拦截订单行数")
    intercepted_orders = summary_value("拦截导致不发订单数")

    lines = [
        "运行完成。",
        f"订单行数：{order_lines}",
        f"销售行数：{sales_rows}",
        f"匹配覆盖率：{matched_lines}/{order_lines} ({join_coverage_pct}%)",
        f"建议发货数量：{total_recommended}",
        f"30%变动保留行数：{small_change_kept_lines}",
        f"全局缺口倍率：{global_gap_multiplier}",
        f"销量权重：sold30={sold30_weight}, sold7={sold7_weight}",
        f"零7日销量且30日有销量无库存发货上限：{zero_sold7_stockout_cap}",
        f"最小发货阈值：{min_order_ship_qty}",
        f"阈值前低发货量订单数：{low_qty_orders_before_exempt}",
        f"低发货量豁免订单数：{low_qty_orders_exempted}",
        f"低发货量拦截订单数：{low_qty_orders}",
        (
            "单单SKU上限规则："
            f"{sku_limit_rule_count} 条（触发行数：{sku_limit_capped_lines}）"
        ),
        (
            "SKC/SKUID 排除规则："
            f"{excluded_skc_rule_count} SKC, "
            f"{excluded_skuid_rule_count} SKUID "
            f"(命中订单行：{intercepted_order_lines}, "
            f"拦截订单：{intercepted_orders})"
        ),
        f"约束配置文件：{constraints_path}",
    ]
    return "\n".join(lines)


def _read_summary_json(path: Path) -> dict[str, object] | None:
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return None

    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return None

    if not isinstance(data, dict):
        return None
    return data


def _prepare_run_output_dir(base_dir: Path) -> Path:
    base_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = base_dir / f"output_{timestamp}"
    sequence = 1
    while candidate.exists():
        candidate = base_dir / f"output_{timestamp}_{sequence:02d}"
        sequence += 1
    candidate.mkdir(parents=False, exist_ok=False)
    return candidate


def _resolve_app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]
