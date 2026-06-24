from __future__ import annotations

import contextlib
import io
import json
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from shipment_planner.cli import main as planner_cli_main
from shipment_planner.constraints import DEFAULT_CONSTRAINTS_FILENAME
from shipment_planner.parsers import (
    ORDER_REQUIRED_COLUMNS,
    SALES_REQUIRED_COLUMNS,
    TEMU_DAILY_REQUIRED_COLUMNS,
    missing_required_columns,
    temu_daily_sales_columns,
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
    service_level_offset: float,
    base_stock_qty: int,
    protection_interval_factor: float,
    temu_sales_path: str | Path,
) -> PlannerRunResult:
    orders = Path(orders_path)
    sales = Path(sales_path)
    temu_sales = Path(temu_sales_path)
    base_output_dir = Path(output_dir)
    input_specs = (
        (orders, ORDER_REQUIRED_COLUMNS, "订单文件"),
        (sales, SALES_REQUIRED_COLUMNS, "销售文件"),
        (temu_sales, TEMU_DAILY_REQUIRED_COLUMNS, "Temu每日销量文件"),
    )

    for file_path, _, label in input_specs:
        _assert_existing_file(file_path, label)
        _assert_xlsx(file_path, label)

    headers: list[list[str]] = []
    for file_path, _, _ in input_specs:
        header, _ = read_xlsx_table(file_path)
        headers.append(header)
    for header, (_, required_columns, label) in zip(headers, input_specs, strict=True):
        if label == "Temu每日销量文件":
            _assert_temu_daily_columns(header)
            continue
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
        "--service-level-offset",
        str(service_level_offset),
        "--base-stock-qty",
        str(base_stock_qty),
        "--protection-interval-factor",
        str(protection_interval_factor),
        "--temu-sales",
        str(temu_sales),
    ]

    stdout_buffer = io.StringIO()
    with contextlib.redirect_stdout(stdout_buffer):
        exit_code = planner_cli_main(args)
    if exit_code != 0:
        raise RuntimeError(f"发货建议程序异常退出，返回码：{exit_code}")

    recommendation = out_dir / "发货建议明细.xlsx"
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
        details = ", ".join(missing)
        raise ValueError(f"{label}缺少必填列：{details}")


def _assert_temu_daily_columns(header: list[str]) -> None:
    _assert_required_columns(header, TEMU_DAILY_REQUIRED_COLUMNS, "Temu每日销量文件")
    if not temu_daily_sales_columns(header):
        raise ValueError("Temu每日销量文件缺少日期销量列。")


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
    lines = [
        "运行完成。",
        f"订单行数：{order_lines}",
        f"销售行数：{summary_value('销售行数')}",
        (
            f"匹配覆盖率：{summary_value('匹配订单行数')}/{order_lines} "
            f"({summary_value('匹配覆盖率_百分比')}%)"
        ),
        f"质量问题行数：{summary_value('质量问题行数')}",
        f"建议发货数量：{summary_value('建议发货总量')}",
        f"30%变动保留行数：{summary_value('触发30%免改行数')}",
        f"需求类型分布：{summary_value('需求类型分布')}",
        f"异常标记分布：{summary_value('异常标记分布')}",
        f"服务水平分布：{summary_value('服务水平分布')}",
        f"预测模型分布：{summary_value('预测模型分布')}",
        f"全局服务水平偏移：{summary_value('全局服务水平偏移')}",
        f"保底库存目标：{summary_value('保底库存目标')}",
        (
            f"保底触发：{summary_value('保底触发SKU数')} 个SKU，"
            f"{summary_value('保底触发订单行数')} 行"
        ),
        f"最小发货阈值：{summary_value('最小发货阈值')}",
        f"阈值前低发货量订单数：{summary_value('阈值前低发货量订单数')}",
        f"低发货量豁免订单数：{summary_value('低于阈值豁免订单数')}",
        f"低发货量拦截订单数：{summary_value('低于阈值订单数_提示')}",
        (
            "单单SKU上限规则："
            f"{summary_value('订单内SKU限额规则数')} 条"
            f"（触发行数：{summary_value('触发订单内SKU限额行数')}）"
        ),
        (
            "SKC/SKUID 排除规则："
            f"{summary_value('SKC拦截规则数')} SKC, "
            f"{summary_value('SKUID拦截规则数')} SKUID "
            f"(命中订单行：{summary_value('命中拦截订单行数')}, "
            f"拦截订单：{summary_value('拦截导致不发订单数')})"
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
    created = tempfile.mkdtemp(prefix=f"output_{timestamp}_", dir=base_dir)
    return Path(created)


def _resolve_app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]
