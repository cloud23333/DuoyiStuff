from __future__ import annotations

import os
import tempfile

os.environ.setdefault("MPLCONFIGDIR", f"{tempfile.gettempdir()}/matplotlib-cache")

from shipment_planner import plots  # noqa: E402


def test_render_sku_plot_draws_history_average_reference_line(monkeypatch):
    ax = _FakeAxis()
    fig = _FakeFigure()

    monkeypatch.setattr(plots.plt, "subplots", lambda **_: (fig, ax))
    monkeypatch.setattr(plots.plt, "close", lambda _: None)

    plots.render_sku_plot(
        history=[2, 4, 6],
        forecast=[5, 5],
        title="sku-1",
    )

    assert ax.horizontal_lines == [
        {
            "y": 4.0,
            "color": "darkgray",
            "linestyle": ":",
            "linewidth": 0.8,
            "label": "avg",
        }
    ]


def test_render_sku_plot_average_ignores_masked_stockout_tail(monkeypatch):
    ax = _FakeAxis()
    fig = _FakeFigure()

    monkeypatch.setattr(plots.plt, "subplots", lambda **_: (fig, ax))
    monkeypatch.setattr(plots.plt, "close", lambda _: None)

    plots.render_sku_plot(
        history=[10, 10, 0, 0],
        forecast=[8, 8],
        title="sku-1",
        masked_tail_from=2,
    )

    assert ax.horizontal_lines[0]["y"] == 10.0


def test_render_sku_plot_labels_forecast_daily_sales(monkeypatch):
    ax = _FakeAxis()
    fig = _FakeFigure()

    monkeypatch.setattr(plots.plt, "subplots", lambda **_: (fig, ax))
    monkeypatch.setattr(plots.plt, "close", lambda _: None)

    plots.render_sku_plot(
        history=[2, 4, 6],
        forecast=[5, 5],
        title="sku-1",
        forecast_daily_sales=3.25,
    )

    assert ax.texts == [
        {
            "x": 0.98,
            "y": 0.96,
            "text": "fc 3.25/d",
            "transform": ax.transAxes,
            "ha": "right",
            "va": "top",
            "fontsize": 6,
            "color": "darkorange",
        }
    ]


class _FakeFigure:
    def tight_layout(self, **_kwargs):
        pass

    def savefig(self, buffer, **_kwargs):
        buffer.write(b"png")


class _FakeAxis:
    def __init__(self):
        self.horizontal_lines: list[dict[str, object]] = []
        self.texts: list[dict[str, object]] = []
        self.transAxes = object()
        self.yaxis = _FakeYAxis()

    def plot(self, *_args, **_kwargs):
        pass

    def axhline(self, y, **kwargs):
        self.horizontal_lines.append({"y": y, **kwargs})

    def axvline(self, *_args, **_kwargs):
        pass

    def set_title(self, *_args, **_kwargs):
        pass

    def tick_params(self, *_args, **_kwargs):
        pass

    def text(self, x, y, text, **kwargs):
        self.texts.append({"x": x, "y": y, "text": text, **kwargs})

    def set_ylim(self, *_args, **_kwargs):
        pass

    def grid(self, *_args, **_kwargs):
        pass

    def legend(self, *_args, **_kwargs):
        pass

    def margins(self, *_args, **_kwargs):
        pass


class _FakeYAxis:
    def set_major_locator(self, *_args, **_kwargs):
        pass
