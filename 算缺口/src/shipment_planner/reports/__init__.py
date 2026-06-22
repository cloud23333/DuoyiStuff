from __future__ import annotations

import csv
import json
from collections.abc import Sequence
from pathlib import Path

from ..eval_forecast import EvalRow
from .localize import (
    _localize_quality_row,
    _localize_recommendation_row,
    _localize_summary,
)
from .plots_cache import _annotate_forecast_error, _build_plot_cache
from .schema import QUALITY_FIELDS, RECOMMENDATION_FIELDS
from .xlsx_writer import _write_recommendations_xlsx

__all__ = [
    "export_reports",
    "RECOMMENDATION_FIELDS",
    "_annotate_forecast_error",
    "_build_plot_cache",
    "_localize_recommendation_row",
    "_localize_quality_row",
    "_localize_summary",
    "_write_recommendations_xlsx",
]


def export_reports(
    out_dir: str | Path,
    recommendations: list[dict[str, object]],
    quality_rows: list[dict[str, object]],
    summary: dict[str, object],
    daily_sales_by_key: dict[tuple[str, str], tuple[int, ...]],
    eval_rows: Sequence[EvalRow] | None = None,
    eval_summary: dict[str, object] | None = None,
) -> dict[str, Path]:
    output_dir = Path(out_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    recommendation_path = output_dir / "发货建议明细.xlsx"
    quality_path = output_dir / "数据质量报告.csv"
    summary_path = output_dir / "运行摘要.json"
    (output_dir / "发货建议明细.csv").unlink(missing_ok=True)

    quality_columns = [target for _, target in QUALITY_FIELDS]

    _annotate_forecast_error(recommendations, eval_rows or ())
    plot_cache = _build_plot_cache(recommendations, daily_sales_by_key)
    localized_recommendations = [_localize_recommendation_row(row) for row in recommendations]
    localized_quality_rows = [_localize_quality_row(row) for row in quality_rows]
    localized_summary = _localize_summary(summary, eval_summary=eval_summary)

    _write_recommendations_xlsx(
        path=recommendation_path,
        rows=localized_recommendations,
        recommendation_rows=recommendations,
        plot_cache=plot_cache,
    )
    _write_csv(quality_path, localized_quality_rows, quality_columns)
    summary_path.write_text(
        json.dumps(localized_summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    return {
        "final_recommendation": recommendation_path,
        "quality_report": quality_path,
        "run_summary": summary_path,
    }


def _write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        if rows:
            writer.writerows(rows)
