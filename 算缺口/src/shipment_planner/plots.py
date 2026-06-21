from __future__ import annotations

import io
import math
from collections.abc import Sequence

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.ticker import MaxNLocator  # noqa: E402

_PLOT_WIDTH_IN = 4.0
_PLOT_HEIGHT_IN = 1.6
_PLOT_DPI = 80
_TITLE_MAX_CHARS = 34
_TITLE_ID_MAX_CHARS = 14
_Y_AXIS_HEADROOM_RATIO = 0.08


def render_sku_plot(
    history: Sequence[float],
    forecast: Sequence[float],
    title: str,
    masked_tail_from: int | None = None,
    forecast_daily_sales: float | None = None,
) -> bytes:
    """Render a small PNG showing history and forecast side by side.

    X axis: day offset where 0 is the last historical day; positive is future.
    History: past days (x = -(n-1)..0). Forecast: future days (x = 1..m).
    When ``masked_tail_from`` is set, the history at index >= that value is
    rendered as greyed-out "OOS (masked)" so operators can see what the
    forecast pipeline ignored.
    """
    fig, ax = plt.subplots(figsize=(_PLOT_WIDTH_IN, _PLOT_HEIGHT_IN), dpi=_PLOT_DPI)
    try:
        _draw_history(ax, history, masked_tail_from)
        _draw_forecast(ax, forecast)
        if len(history) > 0 and len(forecast) > 0:
            ax.axvline(0.5, color="gray", linestyle="--", linewidth=0.7)
        _draw_forecast_label(ax, forecast_daily_sales)
        _style_axes(ax, title=title, history=history, forecast=forecast)
        fig.tight_layout(pad=0.3)

        buffer = io.BytesIO()
        fig.savefig(buffer, format="png")
    finally:
        plt.close(fig)

    return buffer.getvalue()


def _split_history(
    history: Sequence[float],
    masked_tail_from: int | None,
) -> tuple[list[int], list[float], list[int], list[float]]:
    n = len(history)
    cut = n
    if masked_tail_from is not None and 0 <= masked_tail_from < n:
        cut = masked_tail_from
    x = list(range(-(n - 1), 1))
    return x[:cut], list(history[:cut]), x[cut:], list(history[cut:])


def _draw_history(
    ax,
    history: Sequence[float],
    masked_tail_from: int | None,
) -> None:
    used_x, used_y, masked_x, masked_y = _split_history(history, masked_tail_from)

    if used_x:
        ax.axhline(
            _mean(used_y),
            color="darkgray",
            linestyle=":",
            linewidth=0.8,
            label="avg",
        )
        ax.plot(
            used_x,
            used_y,
            color="steelblue",
            linewidth=1.2,
            marker="o",
            markersize=2.5,
            label="history",
        )

    if masked_x:
        if used_x:
            ax.plot(
                [used_x[-1], masked_x[0]],
                [used_y[-1], masked_y[0]],
                color="lightgrey",
                linewidth=0.8,
                linestyle="--",
            )
        ax.plot(
            masked_x,
            masked_y,
            color="lightgrey",
            linewidth=0.8,
            linestyle="--",
            marker="o",
            markersize=2.5,
            markerfacecolor="lightgrey",
            markeredgecolor="lightgrey",
            label="OOS (masked)",
        )


def _draw_forecast(ax, forecast: Sequence[float]) -> None:
    if len(forecast) == 0:
        return
    fc_x = list(range(1, len(forecast) + 1))
    ax.plot(
        fc_x,
        list(forecast),
        color="darkorange",
        linewidth=1.2,
        marker="o",
        markersize=2.5,
        label="forecast",
    )


def _draw_forecast_label(ax, forecast_daily_sales: float | None) -> None:
    label = _forecast_daily_sales_label(forecast_daily_sales)
    if not label:
        return
    ax.text(
        0.98,
        0.96,
        label,
        transform=ax.transAxes,
        ha="right",
        va="top",
        fontsize=6,
        color="darkorange",
    )


def _style_axes(
    ax,
    *,
    title: str,
    history: Sequence[float],
    forecast: Sequence[float],
) -> None:
    ax.set_title(_compact_title(title), fontsize=6, pad=1)
    ax.tick_params(axis="both", labelsize=6)
    _format_y_axis(ax, history=history, forecast=forecast)
    ax.grid(True, alpha=0.25, linewidth=0.5)
    if len(history) > 0 or len(forecast) > 0:
        ax.legend(fontsize=6, loc="upper left", frameon=False)
    ax.margins(x=0.02, y=0)


def _compact_title(title: str) -> str:
    title = title.strip()
    if title.startswith("SKC:") and " / SKUID:" in title:
        skc, skuid = title.removeprefix("SKC:").split(" / SKUID:", maxsplit=1)
        return (
            f"SKC:{_shorten_middle(skc, _TITLE_ID_MAX_CHARS)} / "
            f"SKU:{_shorten_middle(skuid, _TITLE_ID_MAX_CHARS)}"
        )
    return _shorten_middle(title, _TITLE_MAX_CHARS)


def _shorten_middle(value: str, max_chars: int) -> str:
    if len(value) <= max_chars:
        return value
    if max_chars <= 3:
        return value[:max_chars]
    left_chars = (max_chars - 3) // 2
    right_chars = max_chars - 3 - left_chars
    return f"{value[:left_chars]}...{value[-right_chars:]}"


def _mean(values: Sequence[float]) -> float:
    return sum(float(value) for value in values) / len(values)


def _forecast_daily_sales_label(value: float | None) -> str:
    if value is None:
        return ""
    value = float(value)
    if not math.isfinite(value):
        return ""
    formatted = f"{value:.2f}".rstrip("0").rstrip(".")
    return f"fc {formatted}/d"


def _format_y_axis(
    ax,
    *,
    history: Sequence[float],
    forecast: Sequence[float],
) -> None:
    values = [float(value) for value in (*history, *forecast)]
    max_value = max(values, default=0.0)
    y_max = max(1.0, max_value * (1.0 + _Y_AXIS_HEADROOM_RATIO))
    ax.set_ylim(0, y_max)
    ax.yaxis.set_major_locator(MaxNLocator(nbins=4, integer=True, min_n_ticks=2))
