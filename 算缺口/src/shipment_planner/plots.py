from __future__ import annotations

import io
from collections.abc import Sequence

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt  # noqa: E402

_PLOT_WIDTH_IN = 4.0
_PLOT_HEIGHT_IN = 1.6
_PLOT_DPI = 80


def render_sku_plot(
    history: Sequence[float],
    forecast: Sequence[float],
    title: str,
    masked_tail_from: int | None = None,
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
        hist_len = len(history)
        fc_len = len(forecast)

        cut = hist_len
        if masked_tail_from is not None and 0 <= masked_tail_from < hist_len:
            cut = masked_tail_from

        if hist_len > 0:
            hist_x = list(range(-(hist_len - 1), 1))
            used_x = hist_x[:cut]
            used_y = list(history[:cut])
            masked_x = hist_x[cut:]
            masked_y = list(history[cut:])

            if used_x:
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
                # Bridge from last used point to the masked segment so the
                # viewer can see it's contiguous in time.
                if used_x:
                    bridge_x = [used_x[-1], masked_x[0]]
                    bridge_y = [used_y[-1], masked_y[0]]
                    ax.plot(
                        bridge_x,
                        bridge_y,
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

        if fc_len > 0:
            fc_x = list(range(1, fc_len + 1))
            ax.plot(
                fc_x,
                list(forecast),
                color="darkorange",
                linewidth=1.2,
                marker="o",
                markersize=2.5,
                label="forecast",
            )

        if hist_len > 0 and fc_len > 0:
            ax.axvline(0.5, color="gray", linestyle="--", linewidth=0.7)

        ax.set_title(title, fontsize=7)
        ax.tick_params(axis="both", labelsize=6)
        ax.grid(True, alpha=0.25, linewidth=0.5)
        if hist_len > 0 or fc_len > 0:
            ax.legend(fontsize=6, loc="upper left", frameon=False)
        ax.margins(x=0.02)

        fig.tight_layout(pad=0.3)

        buffer = io.BytesIO()
        fig.savefig(buffer, format="png")
    finally:
        plt.close(fig)

    return buffer.getvalue()
