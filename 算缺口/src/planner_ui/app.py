from __future__ import annotations

import io
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QObject, Qt, QThread, QUrl, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QDesktopServices, QFont, QGuiApplication, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QDoubleSpinBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSlider,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from planner_ui.workflow import (
    PlannerRunResult,
    ensure_constraints_template,
    extract_unique_skc,
    get_constraints_config_dir,
    get_constraints_path,
    run_planner,
)
from shipment_planner.alpha_preview import AlphaCurve
from shipment_planner.engine import DEFAULT_BASE_STOCK_QTY


class PlannerRunWorker(QObject):
    finished = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, run_kwargs: dict[str, Any]) -> None:
        super().__init__()
        self._run_kwargs = run_kwargs

    @pyqtSlot()
    def run(self) -> None:
        try:
            result = run_planner(**self._run_kwargs)
        except Exception as exc:  # pragma: no cover - UI error channel
            self.failed.emit(str(exc))
            return
        self.finished.emit(result)


class PreviewWorker(QObject):
    finished = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, orders_path: Path, sales_path: Path, temu_path: Path) -> None:
        super().__init__()
        self._orders_path = orders_path
        self._sales_path = sales_path
        self._temu_path = temu_path

    @pyqtSlot()
    def run(self) -> None:
        try:
            from shipment_planner.alpha_preview import compute_alpha_curve

            curve = compute_alpha_curve(
                orders_path=self._orders_path,
                sales_path=self._sales_path,
                temu_path=self._temu_path,
            )
        except Exception as exc:  # pragma: no cover - UI error channel
            self.failed.emit(str(exc))
            return
        self.finished.emit(curve)


class PlannerWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self._run_thread: QThread | None = None
        self._run_worker: PlannerRunWorker | None = None
        self._preview_thread: QThread | None = None
        self._preview_worker: PreviewWorker | None = None
        self._alpha_curve: AlphaCurve | None = None
        self._alpha_syncing = False
        self._constraints_ready = False
        self._last_dialog_dir: Path | None = None

        self.setWindowTitle("发货建议工具")

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        root_layout = QVBoxLayout(central_widget)
        root_layout.setContentsMargins(10, 10, 10, 10)
        root_layout.setSpacing(8)

        root_layout.addWidget(self._build_orders_group())
        root_layout.addWidget(self._build_run_group())
        root_layout.addWidget(self._build_skc_group(), stretch=1)
        root_layout.addWidget(self._build_log_group(), stretch=1)

        self.setStyleSheet(_app_stylesheet())
        self._set_default_window_size()
        self._set_status("请选择订单文件开始。")
        self._init_constraints_template()
        self._refresh_run_button_state()

    def _set_default_window_size(self) -> None:
        target_width = max(self.sizeHint().width(), self.minimumSizeHint().width())
        self.resize(target_width, 760)

    def _build_orders_group(self) -> QGroupBox:
        group = QGroupBox("文件")
        layout = QGridLayout(group)
        layout.setHorizontalSpacing(8)
        layout.setVerticalSpacing(6)
        layout.setColumnStretch(1, 1)

        self.order_path_edit = QLineEdit()
        self.order_path_edit.setReadOnly(True)
        self.order_path_edit.setPlaceholderText(".xlsx 订单文件")
        self.order_browse_button = QPushButton("订单")
        self.order_browse_button.clicked.connect(self._on_pick_orders)
        self.order_browse_button.setMinimumWidth(92)

        self.sales_path_edit = QLineEdit()
        self.sales_path_edit.setReadOnly(True)
        self.sales_path_edit.setPlaceholderText(".xlsx 销售文件")
        self.sales_browse_button = QPushButton("销售")
        self.sales_browse_button.clicked.connect(self._on_pick_sales)
        self.sales_browse_button.setMinimumWidth(92)

        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setReadOnly(True)
        self.output_dir_edit.setPlaceholderText("输出目录")
        self.output_browse_button = QPushButton("输出")
        self.output_browse_button.clicked.connect(self._on_pick_output_dir)
        self.output_browse_button.setMinimumWidth(92)

        self.temu_path_edit = QLineEdit()
        self.temu_path_edit.setReadOnly(True)
        self.temu_path_edit.setPlaceholderText(".xlsx Temu每日销量文件（必选）")
        self.temu_browse_button = QPushButton("Temu明细")
        self.temu_browse_button.clicked.connect(self._on_pick_temu_sales)
        self.temu_browse_button.setMinimumWidth(92)
        self.temu_clear_button = QPushButton("✕")
        self.temu_clear_button.setFixedWidth(28)
        self.temu_clear_button.setToolTip("清除 Temu 明细文件")
        self.temu_clear_button.clicked.connect(self._on_clear_temu_sales)
        self.temu_clear_button.setEnabled(False)

        self.constraints_path_edit = QLineEdit()
        self.constraints_path_edit.setReadOnly(True)
        self.constraints_path_edit.setText(str(get_constraints_path()))
        self.constraints_path_edit.setToolTip(self.constraints_path_edit.text())

        self.open_config_dir_button = QPushButton("配置目录")
        self.open_config_dir_button.clicked.connect(self._on_open_config_dir)
        self.open_config_dir_button.setMinimumWidth(92)

        temu_row = QHBoxLayout()
        temu_row.setSpacing(4)
        temu_row.addWidget(self.temu_path_edit)
        temu_row.addWidget(self.temu_clear_button)

        layout.addWidget(QLabel("订单"), 0, 0)
        layout.addWidget(self.order_path_edit, 0, 1)
        layout.addWidget(self.order_browse_button, 0, 2)
        layout.addWidget(QLabel("销售"), 1, 0)
        layout.addWidget(self.sales_path_edit, 1, 1)
        layout.addWidget(self.sales_browse_button, 1, 2)
        layout.addWidget(QLabel("输出"), 2, 0)
        layout.addWidget(self.output_dir_edit, 2, 1)
        layout.addWidget(self.output_browse_button, 2, 2)
        layout.addWidget(QLabel("Temu"), 3, 0)
        layout.addLayout(temu_row, 3, 1)
        layout.addWidget(self.temu_browse_button, 3, 2)
        layout.addWidget(QLabel("配置"), 4, 0)
        layout.addWidget(self.constraints_path_edit, 4, 1)
        layout.addWidget(self.open_config_dir_button, 4, 2)
        return group

    def _build_run_group(self) -> QGroupBox:
        group = QGroupBox("参数")
        layout = QVBoxLayout(group)
        layout.setSpacing(6)

        params_grid = QGridLayout()
        params_grid.setHorizontalSpacing(8)
        params_grid.setVerticalSpacing(6)
        params_grid.setColumnStretch(4, 1)

        self.service_level_offset_spin = QDoubleSpinBox()
        self.service_level_offset_spin.setDecimals(2)
        self.service_level_offset_spin.setRange(-0.3, 0.3)
        self.service_level_offset_spin.setSingleStep(0.05)
        self.service_level_offset_spin.setValue(0.0)
        self.service_level_offset_spin.setFixedWidth(108)

        self.base_stock_qty_spin = QSpinBox()
        self.base_stock_qty_spin.setRange(0, 9999)
        self.base_stock_qty_spin.setSingleStep(1)
        self.base_stock_qty_spin.setValue(DEFAULT_BASE_STOCK_QTY)
        self.base_stock_qty_spin.setFixedWidth(108)

        self.alpha_slider = QSlider(Qt.Orientation.Horizontal)
        self.alpha_slider.setRange(50, 100)
        self.alpha_slider.setSingleStep(5)
        self.alpha_slider.setPageStep(5)
        self.alpha_slider.setValue(100)
        self.alpha_slider.valueChanged.connect(self._on_alpha_slider_changed)
        self.alpha_slider.valueChanged.connect(self._update_alpha_estimate)
        self.alpha_slider.sliderReleased.connect(self._render_alpha_plot)

        self.alpha_value_spin = QDoubleSpinBox()
        self.alpha_value_spin.setDecimals(2)
        self.alpha_value_spin.setRange(0.50, 1.00)
        self.alpha_value_spin.setSingleStep(0.05)
        self.alpha_value_spin.setValue(1.0)
        self.alpha_value_spin.setFixedWidth(90)
        self.alpha_value_spin.valueChanged.connect(self._on_alpha_spin_changed)
        self.alpha_value_spin.editingFinished.connect(self._render_alpha_plot)

        self.alpha_suggested_label = QLabel("建议值：—")
        self.alpha_adopt_button = QPushButton("采用建议")
        self.alpha_adopt_button.setEnabled(False)
        self.alpha_adopt_button.clicked.connect(self._on_adopt_suggested_alpha)

        self.alpha_estimate_label = QLabel("预计今日发货：先点“预览发货量”")
        self.alpha_estimate_label.setWordWrap(True)
        estimate_font = QFont()
        estimate_font.setBold(True)
        estimate_font.setPointSize(13)
        self.alpha_estimate_label.setFont(estimate_font)
        self.alpha_estimate_label.setStyleSheet("color: #0f766e;")

        self.preview_button = QPushButton("预览发货量")
        self.preview_button.clicked.connect(self._on_preview_clicked)

        self.alpha_plot_label = QLabel()
        self.alpha_plot_label.setFixedSize(440, 220)
        self.alpha_plot_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.status_label = QLabel("")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        self.status_label.setMinimumHeight(28)

        self.run_button = QPushButton("开始运行")
        self.run_button.setObjectName("runButton")
        self.run_button.clicked.connect(self._on_run_clicked)
        self.run_button.setMinimumWidth(110)

        params_grid.addWidget(QLabel("全局服务水平偏移"), 0, 0)
        params_grid.addWidget(self.service_level_offset_spin, 0, 1)
        params_grid.addWidget(QLabel("保底库存"), 0, 2)
        params_grid.addWidget(self.base_stock_qty_spin, 0, 3)

        alpha_card = self._build_alpha_card()

        action_row = QHBoxLayout()
        action_row.setSpacing(8)
        action_row.addWidget(self.status_label, stretch=1)
        action_row.addWidget(self.run_button)

        layout.addLayout(params_grid)
        layout.addWidget(alpha_card)
        layout.addLayout(action_row)
        return group

    def _build_alpha_card(self) -> QGroupBox:
        card = QGroupBox("发货保守度 α · 越小越保守，少发降积压")
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(8)
        card_layout.setContentsMargins(10, 6, 10, 8)

        action_row = QHBoxLayout()
        action_row.setSpacing(8)
        action_row.addWidget(self.preview_button)
        action_row.addStretch(1)
        action_row.addWidget(self.alpha_suggested_label)
        action_row.addWidget(self.alpha_adopt_button)

        anchor_font = QFont()
        anchor_font.setPointSize(8)
        left_anchor = QLabel("正常 1.0")
        left_anchor.setFont(anchor_font)
        left_anchor.setStyleSheet("color: #888;")
        right_anchor = QLabel("0.5 保守")
        right_anchor.setFont(anchor_font)
        right_anchor.setStyleSheet("color: #888;")

        slider_row = QHBoxLayout()
        slider_row.setSpacing(8)
        slider_row.addWidget(left_anchor)
        slider_row.addWidget(self.alpha_slider, stretch=1)
        slider_row.addWidget(right_anchor)
        slider_row.addWidget(QLabel("α"))
        slider_row.addWidget(self.alpha_value_spin)

        plot_row = QHBoxLayout()
        plot_row.addStretch(1)
        plot_row.addWidget(self.alpha_plot_label)
        plot_row.addStretch(1)

        card_layout.addLayout(action_row)
        card_layout.addLayout(slider_row)
        card_layout.addWidget(self.alpha_estimate_label)
        card_layout.addLayout(plot_row)
        return card

    def _build_skc_group(self) -> QGroupBox:
        group = QGroupBox("SKC")
        layout = QVBoxLayout(group)
        layout.setSpacing(6)

        toolbar = QHBoxLayout()
        self.skc_count_label = QLabel("唯一 SKC：0")
        self.skc_count_label.setObjectName("skcCountLabel")
        self.copy_skc_button = QPushButton("复制")
        self.copy_skc_button.clicked.connect(self._on_copy_skc)
        self.copy_skc_button.setEnabled(False)

        toolbar.addWidget(self.skc_count_label)
        toolbar.addStretch(1)
        toolbar.addWidget(self.copy_skc_button)

        self.skc_text_edit = QPlainTextEdit()
        self.skc_text_edit.setReadOnly(True)
        self.skc_text_edit.setPlaceholderText("唯一 SKC（每行一个）")
        self.skc_text_edit.setFont(_monospace_font())
        self.skc_text_edit.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.skc_text_edit.setObjectName("skcTextEdit")
        self.skc_text_edit.setMinimumHeight(140)

        layout.addLayout(toolbar)
        layout.addWidget(self.skc_text_edit)
        return group

    def _build_log_group(self) -> QGroupBox:
        group = QGroupBox("运行日志")
        layout = QVBoxLayout(group)

        header_row = QHBoxLayout()
        header_row.addStretch(1)

        self.log_text_edit = QPlainTextEdit()
        self.log_text_edit.setReadOnly(True)
        self.log_text_edit.setFont(_monospace_font())
        self.log_text_edit.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.log_text_edit.setObjectName("logTextEdit")
        self.log_text_edit.setMinimumHeight(180)

        self.clear_log_button = QPushButton("清空日志")
        self.clear_log_button.clicked.connect(self.log_text_edit.clear)
        header_row.addWidget(self.clear_log_button)

        layout.addLayout(header_row)
        layout.addWidget(self.log_text_edit)
        return group

    @pyqtSlot()
    def _on_pick_orders(self) -> None:
        selected_path = self._pick_xlsx_file(
            "选择订单文件",
            self.order_path_edit.text(),
            self.sales_path_edit.text(),
            self.output_dir_edit.text(),
        )
        if not selected_path:
            return
        self._set_path_edit(self.order_path_edit, selected_path)
        self._remember_dialog_dir(Path(selected_path))
        self._append_log(f"已选择订单文件：{selected_path}")
        self._load_unique_skc(Path(selected_path))
        self._refresh_run_button_state()

    @pyqtSlot()
    def _on_pick_sales(self) -> None:
        selected_path = self._pick_xlsx_file(
            "选择销售文件",
            self.sales_path_edit.text(),
            self.order_path_edit.text(),
            self.output_dir_edit.text(),
        )
        if not selected_path:
            return
        self._set_path_edit(self.sales_path_edit, selected_path)
        self._remember_dialog_dir(Path(selected_path))
        self._append_log(f"已选择销售文件：{selected_path}")
        self._set_status("销售文件已准备好。")
        self._refresh_run_button_state()

    @pyqtSlot()
    def _on_pick_temu_sales(self) -> None:
        selected_path = self._pick_xlsx_file(
            "选择 Temu 销售明细文件",
            self.temu_path_edit.text(),
            self.order_path_edit.text(),
            self.sales_path_edit.text(),
        )
        if not selected_path:
            return
        self._set_path_edit(self.temu_path_edit, selected_path)
        self._remember_dialog_dir(Path(selected_path))
        self._append_log(f"已选择 Temu 明细文件：{selected_path}")
        self.temu_clear_button.setEnabled(True)

    @pyqtSlot()
    def _on_clear_temu_sales(self) -> None:
        self.temu_path_edit.clear()
        self.temu_path_edit.setToolTip("")
        self.temu_clear_button.setEnabled(False)
        self._append_log("已清除 Temu 明细文件。")

    @pyqtSlot()
    def _on_pick_output_dir(self) -> None:
        selected_dir = self._pick_directory(
            "选择输出目录",
            self.output_dir_edit.text(),
            self.order_path_edit.text(),
            self.sales_path_edit.text(),
        )
        if not selected_dir:
            return
        self._set_path_edit(self.output_dir_edit, selected_dir)
        self._remember_dialog_dir(Path(selected_dir))
        self._append_log(f"已选择输出目录：{selected_dir}")
        self._set_status("输出目录已设置。")
        self._refresh_run_button_state()

    @pyqtSlot()
    def _on_open_config_dir(self) -> None:
        try:
            constraints_path, created = ensure_constraints_template()
        except Exception as exc:
            self._constraints_ready = False
            self._append_log(f"【错误】打开配置目录失败：{exc}")
            self._set_status(f"打开配置目录失败：{exc}", error=True)
            QMessageBox.critical(self, "打开配置目录失败", str(exc))
            self._refresh_run_button_state()
            return

        self._apply_constraints_path(constraints_path)
        if created:
            self._append_log(f"已创建约束配置模板：{constraints_path}")

        config_dir = constraints_path.parent
        self._remember_dialog_dir(config_dir)
        opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(config_dir)))
        if not opened:
            message = f"无法打开目录，请手动前往：{config_dir}"
            self._append_log(f"【错误】{message}")
            self._set_status(message, error=True)
            QMessageBox.warning(self, "无法打开目录", message)
            self._refresh_run_button_state()
            return

        self._append_log(f"已打开配置目录：{config_dir}")
        self._set_status("已打开约束配置目录。")
        self._refresh_run_button_state()

    @pyqtSlot()
    def _on_copy_skc(self) -> None:
        skc_text = self.skc_text_edit.toPlainText().strip()
        if not skc_text:
            QMessageBox.warning(self, "无 SKC 可复制", "请先导入有效的订单文件。")
            return
        clipboard = QGuiApplication.clipboard()
        clipboard.setText(skc_text)
        self._set_status("SKC 已复制到剪贴板。")
        self._append_log("已复制 SKC 到剪贴板。")

    def _load_unique_skc(self, order_path: Path) -> None:
        try:
            skc_codes = extract_unique_skc(order_path)
        except Exception as exc:
            self._reset_order_selection()
            self._set_status(f"读取订单文件失败：{exc}", error=True)
            self._append_log(f"【错误】读取订单文件失败：{exc}")
            QMessageBox.critical(self, "订单文件读取失败", str(exc))
            self._refresh_run_button_state()
            return

        self.skc_text_edit.setPlainText("\n".join(skc_codes))
        self.skc_count_label.setText(f"唯一 SKC：{len(skc_codes)}")
        self.copy_skc_button.setEnabled(bool(skc_codes))
        self._set_status("已提取唯一 SKC，可继续导入销售文件。")
        self._append_log(f"提取唯一 SKC 数量：{len(skc_codes)}")

    @pyqtSlot()
    def _on_run_clicked(self) -> None:
        run_request = self._collect_run_request()
        if run_request is None:
            return
        self._start_run(run_request)

    @pyqtSlot(int)
    def _on_alpha_slider_changed(self, value: int) -> None:
        if self._alpha_syncing:
            return
        self._alpha_syncing = True
        try:
            self.alpha_value_spin.setValue(value / 100)
        finally:
            self._alpha_syncing = False

    @pyqtSlot(float)
    def _on_alpha_spin_changed(self, value: float) -> None:
        if self._alpha_syncing:
            return
        self._alpha_syncing = True
        try:
            self.alpha_slider.setValue(round(value * 100))
        finally:
            self._alpha_syncing = False
        self._update_alpha_estimate()

    @pyqtSlot()
    def _on_adopt_suggested_alpha(self) -> None:
        if self._alpha_curve is None:
            return
        self.alpha_slider.setValue(round(self._alpha_curve.suggested_alpha * 100))
        self._render_alpha_plot()

    @pyqtSlot()
    def _on_preview_clicked(self) -> None:
        orders_text = self.order_path_edit.text().strip()
        sales_text = self.sales_path_edit.text().strip()
        temu_text = self.temu_path_edit.text().strip()
        if not orders_text or not sales_text or not temu_text:
            QMessageBox.warning(
                self,
                "信息不完整",
                "请先选择订单文件、销售文件和 Temu每日销量文件。",
            )
            return
        self._start_preview(
            Path(orders_text), Path(sales_text), Path(temu_text)
        )

    def _start_preview(
        self, orders_path: Path, sales_path: Path, temu_path: Path
    ) -> None:
        self.preview_button.setEnabled(False)
        self._set_status("正在预览，请稍候...")

        self._preview_thread = QThread(self)
        self._preview_worker = PreviewWorker(orders_path, sales_path, temu_path)
        self._preview_worker.moveToThread(self._preview_thread)

        self._preview_thread.started.connect(self._preview_worker.run)
        self._preview_worker.finished.connect(self._on_preview_finished)
        self._preview_worker.failed.connect(self._on_preview_failed)

        self._preview_worker.finished.connect(self._preview_thread.quit)
        self._preview_worker.failed.connect(self._preview_thread.quit)
        self._preview_worker.finished.connect(self._preview_worker.deleteLater)
        self._preview_worker.failed.connect(self._preview_worker.deleteLater)

        self._preview_thread.finished.connect(self._on_preview_thread_finished)
        self._preview_thread.finished.connect(self._preview_thread.deleteLater)
        self._preview_thread.start()

    @pyqtSlot(object)
    def _on_preview_finished(self, curve: object) -> None:
        if not isinstance(curve, AlphaCurve):
            self._on_preview_failed("预览结果数据格式异常。")
            return

        self._alpha_curve = curve
        self.alpha_slider.setEnabled(True)
        self.alpha_value_spin.setEnabled(True)
        self.alpha_adopt_button.setEnabled(True)
        self.alpha_suggested_label.setText(f"建议值：{curve.suggested_alpha:.2f}")
        self._render_alpha_plot()
        self._update_alpha_estimate()
        self.preview_button.setEnabled(True)
        self._set_status("预览完成，可拖动滑块查看预计发货量。")

    @pyqtSlot(str)
    def _on_preview_failed(self, message: str) -> None:
        self._append_log(f"【错误】预览失败：{message}")
        self._set_status(f"预览失败：{message}", error=True)
        self.preview_button.setEnabled(True)
        QMessageBox.critical(self, "预览失败", message)

    @pyqtSlot()
    def _on_preview_thread_finished(self) -> None:
        self._preview_worker = None
        self._preview_thread = None

    def _update_alpha_estimate(self) -> None:
        curve = self._alpha_curve
        if curve is None:
            self.alpha_estimate_label.setText("预计今日发货：先点“预览发货量”")
            return

        current_alpha = self.alpha_slider.value() / 100
        point = min(
            curve.points, key=lambda p: abs(p.alpha - current_alpha)
        )
        default_ship = curve.default_ship_units
        if default_ship > 0:
            pct = 100 * (point.ship_units - default_ship) / default_ship
        else:
            pct = 0.0
        self.alpha_estimate_label.setText(
            f"预计今日发货 ≈ {point.ship_units} 件（比正常 {pct:+.0f}%）"
        )

    def _render_alpha_plot(self) -> None:
        curve = self._alpha_curve
        if curve is None or not curve.points:
            return

        import matplotlib
        from matplotlib.backends.backend_agg import FigureCanvasAgg
        from matplotlib.figure import Figure

        cjk_font = _resolve_cjk_font()
        previous_sans = matplotlib.rcParams["font.sans-serif"]
        previous_minus = matplotlib.rcParams["axes.unicode_minus"]
        if cjk_font is not None:
            matplotlib.rcParams["font.sans-serif"] = [cjk_font, *previous_sans]
            matplotlib.rcParams["axes.unicode_minus"] = False

        figure = Figure(figsize=(4.6, 2.4), dpi=120)
        FigureCanvasAgg(figure)
        try:
            self._draw_alpha_axes(figure, curve, use_cjk=cjk_font is not None)
            buffer = io.BytesIO()
            figure.savefig(buffer, format="png")
        finally:
            figure.clear()
            matplotlib.rcParams["font.sans-serif"] = previous_sans
            matplotlib.rcParams["axes.unicode_minus"] = previous_minus

        pixmap = QPixmap()
        pixmap.loadFromData(buffer.getvalue())
        self.alpha_plot_label.setPixmap(pixmap)

    def _draw_alpha_axes(self, figure, curve: AlphaCurve, *, use_cjk: bool) -> None:
        labels = _PLOT_LABELS_CJK if use_cjk else _PLOT_LABELS_EN
        ordered = sorted(curve.points, key=lambda p: p.alpha)
        alphas = [p.alpha for p in ordered]
        ship_units = [p.ship_units for p in ordered]
        lost_units = [p.lost_units for p in ordered]

        ax_ship = figure.add_subplot(111)
        ship_line, = ax_ship.plot(
            alphas, ship_units, color="#1f77b4", marker="o", markersize=3,
            label=labels["ship"],
        )
        ax_ship.set_xlabel("α", fontsize=7)
        ax_ship.set_ylabel(labels["ship"], color="#1f77b4", fontsize=7)
        ax_ship.tick_params(axis="both", labelsize=6)
        ax_ship.tick_params(axis="y", labelcolor="#1f77b4")

        ax_lost = ax_ship.twinx()
        lost_line, = ax_lost.plot(
            alphas, lost_units, color="#d62728", marker="s", markersize=3,
            label=labels["risk"],
        )
        ax_lost.set_ylabel(labels["risk"], color="#d62728", fontsize=7)
        ax_lost.tick_params(axis="y", labelsize=6, labelcolor="#d62728")

        current_alpha = self.alpha_slider.value() / 100
        current_line = ax_ship.axvline(
            current_alpha, color="#333333", linewidth=1.0, label=labels["current"]
        )
        suggested_line = ax_ship.axvline(
            curve.suggested_alpha, color="#2ca02c", linewidth=1.0, linestyle="--",
            label=labels["suggested"],
        )

        handles = [ship_line, lost_line, current_line, suggested_line]
        ax_ship.legend(
            handles=handles,
            labels=[handle.get_label() for handle in handles],
            fontsize=6,
            loc="upper center",
            bbox_to_anchor=(0.5, -0.18),
            ncol=2,
            frameon=False,
            columnspacing=1.2,
            handlelength=1.4,
            handletextpad=0.4,
        )
        figure.tight_layout(pad=0.6)

    @pyqtSlot(object)
    def _on_run_finished(self, result_obj: object) -> None:
        if not isinstance(result_obj, PlannerRunResult):
            self._on_run_failed("运行结果数据格式异常。")
            return

        if result_obj.console_output:
            self._append_log(result_obj.console_output)
        if result_obj.constraints_template_created:
            self._append_log(f"已自动创建约束配置模板：{result_obj.constraints_path}")
        self._append_log(f"输出目录：{result_obj.output_dir}")
        self._append_log(f"输出文件：{result_obj.recommendation_path}")
        self._append_log(f"输出文件：{result_obj.quality_path}")
        self._append_log(f"输出文件：{result_obj.summary_path}")
        self._set_status("运行完成，结果已写入输出子目录。")

        dialog = QMessageBox(self)
        dialog.setIcon(QMessageBox.Icon.Information)
        dialog.setWindowTitle("运行完成")
        dialog.setText("输出文件已生成：")
        dialog.setInformativeText(
            "\n".join(
                [
                    f"目录：{result_obj.output_dir}",
                    str(result_obj.recommendation_path),
                    str(result_obj.quality_path),
                    str(result_obj.summary_path),
                ]
            )
        )
        open_recommendation_button = dialog.addButton(
            "打开明细表", QMessageBox.ButtonRole.ActionRole
        )
        open_output_dir_button = dialog.addButton(
            "打开输出目录", QMessageBox.ButtonRole.ActionRole
        )
        dialog.addButton(QMessageBox.StandardButton.Ok)
        dialog.exec()

        clicked_button = dialog.clickedButton()
        if clicked_button == open_recommendation_button:
            self._open_generated_path(result_obj.recommendation_path, "明细表")
        elif clicked_button == open_output_dir_button:
            self._open_generated_path(result_obj.output_dir, "输出目录")

    def _open_generated_path(self, path: Path, label: str) -> None:
        opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
        if opened:
            self._append_log(f"已打开{label}：{path}")
            self._set_status(f"已打开{label}。")
            return

        message = f"无法打开{label}，请手动前往：{path}"
        self._append_log(f"【错误】{message}")
        self._set_status(message, error=True)
        QMessageBox.warning(self, f"无法打开{label}", message)

    @pyqtSlot(str)
    def _on_run_failed(self, message: str) -> None:
        self._append_log(f"【错误】运行失败：{message}")
        self._set_status(f"运行失败：{message}", error=True)
        QMessageBox.critical(self, "运行失败", message)

    @pyqtSlot()
    def _on_run_thread_finished(self) -> None:
        self._run_worker = None
        self._run_thread = None
        self._set_running_state(False)

    def _refresh_run_button_state(self) -> None:
        self.run_button.setEnabled(
            self._inputs_ready_for_run()
            and self._constraints_ready
            and self._run_thread is None
        )

    def _set_running_state(self, running: bool) -> None:
        temu_selected = bool(self.temu_path_edit.text().strip())
        for control in (
            self.order_browse_button,
            self.sales_browse_button,
            self.output_browse_button,
            self.temu_browse_button,
            self.open_config_dir_button,
            self.service_level_offset_spin,
            self.base_stock_qty_spin,
            self.alpha_slider,
            self.alpha_value_spin,
            self.alpha_adopt_button,
            self.preview_button,
            self.clear_log_button,
        ):
            control.setEnabled(not running)
        self.temu_clear_button.setEnabled(not running and temu_selected)

        if running:
            self.copy_skc_button.setEnabled(False)
            self.run_button.setEnabled(False)
            return

        self.copy_skc_button.setEnabled(bool(self.skc_text_edit.toPlainText().strip()))
        self.alpha_adopt_button.setEnabled(self._alpha_curve is not None)
        self._refresh_run_button_state()

    def _set_status(self, message: str, *, error: bool = False) -> None:
        self.status_label.setText(message)
        color = "#b42318" if error else "#166534"
        self.status_label.setStyleSheet(f"color: {color};")

    def _append_log(self, message: str) -> None:
        if not message:
            return
        self.log_text_edit.appendPlainText(message)
        cursor = self.log_text_edit.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.log_text_edit.setTextCursor(cursor)

    def _collect_run_request(self) -> dict[str, Any] | None:
        if not self._constraints_ready:
            message = "约束配置未就绪，请点击“打开配置目录”修复后再运行。"
            self._set_status(message, error=True)
            QMessageBox.warning(self, "配置未就绪", message)
            return None

        orders_text = self.order_path_edit.text().strip()
        sales_text = self.sales_path_edit.text().strip()
        output_text = self.output_dir_edit.text().strip()
        temu_text = self.temu_path_edit.text().strip()
        if not orders_text or not sales_text or not temu_text or not output_text:
            QMessageBox.warning(
                self,
                "信息不完整",
                "请先选择订单文件、销售文件、Temu每日销量文件和输出目录。",
            )
            return None

        run_kwargs: dict[str, Any] = {
            "orders_path": Path(orders_text),
            "sales_path": Path(sales_text),
            "output_dir": Path(output_text),
            "service_level_offset": float(self.service_level_offset_spin.value()),
            "base_stock_qty": int(self.base_stock_qty_spin.value()),
            "protection_interval_factor": float(self.alpha_value_spin.value()),
            "temu_sales_path": Path(temu_text),
        }

        validation_error = self._validate_run_inputs(
            orders_path=run_kwargs["orders_path"],
            sales_path=run_kwargs["sales_path"],
            temu_sales_path=run_kwargs["temu_sales_path"],
            output_dir=run_kwargs["output_dir"],
            service_level_offset=run_kwargs["service_level_offset"],
            base_stock_qty=run_kwargs["base_stock_qty"],
            protection_interval_factor=run_kwargs["protection_interval_factor"],
        )
        if validation_error is not None:
            self._set_status(validation_error, error=True)
            QMessageBox.warning(self, "输入无效", validation_error)
            return None

        return run_kwargs

    def _start_run(self, run_kwargs: dict[str, Any]) -> None:
        self._set_running_state(True)
        self._set_status("正在运行，请稍候...")
        self._append_log("开始运行发货建议计算...")

        self._run_thread = QThread(self)
        self._run_worker = PlannerRunWorker(run_kwargs)
        self._run_worker.moveToThread(self._run_thread)

        self._run_thread.started.connect(self._run_worker.run)
        self._run_worker.finished.connect(self._on_run_finished)
        self._run_worker.failed.connect(self._on_run_failed)

        self._run_worker.finished.connect(self._run_thread.quit)
        self._run_worker.failed.connect(self._run_thread.quit)
        self._run_worker.finished.connect(self._run_worker.deleteLater)
        self._run_worker.failed.connect(self._run_worker.deleteLater)

        self._run_thread.finished.connect(self._on_run_thread_finished)
        self._run_thread.finished.connect(self._run_thread.deleteLater)
        self._run_thread.start()

    def _init_constraints_template(self) -> None:
        try:
            constraints_path, created = ensure_constraints_template()
        except Exception as exc:
            self._constraints_ready = False
            message = f"初始化约束配置失败：{exc}"
            self._append_log(f"【错误】{message}")
            self._set_status(message, error=True)
            self._refresh_run_button_state()
            return

        self._apply_constraints_path(constraints_path)
        if created:
            self._append_log(f"首次启动已创建约束配置模板：{constraints_path}")
        else:
            self._append_log(f"约束配置文件：{constraints_path}")
        self._refresh_run_button_state()

    def _preferred_dialog_dir(self, *raw_paths: str) -> str:
        for raw_path in raw_paths:
            path_text = raw_path.strip()
            if not path_text:
                continue

            candidate = Path(path_text)
            if candidate.is_file():
                return str(candidate.parent)
            if candidate.is_dir():
                return str(candidate)
            parent = candidate.parent
            if parent.exists() and parent.is_dir():
                return str(parent)

        if self._last_dialog_dir is not None and self._last_dialog_dir.exists():
            return str(self._last_dialog_dir)

        constraints_dir = get_constraints_config_dir()
        if constraints_dir.exists() and constraints_dir.is_dir():
            return str(constraints_dir)

        return str(Path.home())

    def _pick_xlsx_file(self, title: str, *raw_paths: str) -> str:
        start_dir = self._preferred_dialog_dir(*raw_paths)
        selected_path, _ = QFileDialog.getOpenFileName(
            self,
            title,
            start_dir,
            "Excel 文件 (*.xlsx)",
        )
        return selected_path

    def _pick_directory(self, title: str, *raw_paths: str) -> str:
        start_dir = self._preferred_dialog_dir(*raw_paths)
        return QFileDialog.getExistingDirectory(self, title, start_dir)

    def _set_path_edit(self, edit: QLineEdit, path_text: str) -> None:
        edit.setText(path_text)
        edit.setToolTip(path_text)

    def _apply_constraints_path(self, constraints_path: Path) -> None:
        self._constraints_ready = True
        self._set_path_edit(self.constraints_path_edit, str(constraints_path))
        self._remember_dialog_dir(constraints_path.parent)

    def _reset_order_selection(self) -> None:
        self.order_path_edit.clear()
        self.order_path_edit.setToolTip("")
        self.skc_text_edit.clear()
        self.skc_count_label.setText("唯一 SKC：0")
        self.copy_skc_button.setEnabled(False)

    def _inputs_ready_for_run(self) -> bool:
        orders_text = self.order_path_edit.text().strip()
        sales_text = self.sales_path_edit.text().strip()
        temu_text = self.temu_path_edit.text().strip()
        output_text = self.output_dir_edit.text().strip()
        if not orders_text or not sales_text or not temu_text or not output_text:
            return False

        validation_error = self._validate_run_inputs(
            orders_path=Path(orders_text),
            sales_path=Path(sales_text),
            temu_sales_path=Path(temu_text),
            output_dir=Path(output_text),
            service_level_offset=float(self.service_level_offset_spin.value()),
            base_stock_qty=int(self.base_stock_qty_spin.value()),
            protection_interval_factor=float(self.alpha_value_spin.value()),
        )
        return validation_error is None

    def _remember_dialog_dir(self, path: Path) -> None:
        candidate = path if path.is_dir() else path.parent
        if candidate.exists() and candidate.is_dir():
            self._last_dialog_dir = candidate.resolve()

    def _validate_run_inputs(
        self,
        *,
        orders_path: Path,
        sales_path: Path,
        temu_sales_path: Path,
        output_dir: Path,
        service_level_offset: float,
        base_stock_qty: int,
        protection_interval_factor: float,
    ) -> str | None:
        if not orders_path.exists():
            return f"订单文件不存在：{orders_path}"
        if not orders_path.is_file():
            return f"订单路径不是文件：{orders_path}"
        if orders_path.suffix.lower() != ".xlsx":
            return f"订单文件不是 xlsx 格式：{orders_path}"

        if not sales_path.exists():
            return f"销售文件不存在：{sales_path}"
        if not sales_path.is_file():
            return f"销售路径不是文件：{sales_path}"
        if sales_path.suffix.lower() != ".xlsx":
            return f"销售文件不是 xlsx 格式：{sales_path}"

        if not temu_sales_path.exists():
            return f"Temu每日销量文件不存在：{temu_sales_path}"
        if not temu_sales_path.is_file():
            return f"Temu每日销量路径不是文件：{temu_sales_path}"
        if temu_sales_path.suffix.lower() != ".xlsx":
            return f"Temu每日销量文件不是 xlsx 格式：{temu_sales_path}"

        if output_dir.exists() and not output_dir.is_dir():
            return f"输出路径不是目录：{output_dir}"
        if not -0.3 <= service_level_offset <= 0.3:
            return "全局服务水平偏移需在 [-0.3, 0.3] 之间。"
        if base_stock_qty < 0:
            return "保底库存不能小于 0。"
        if not 0.0 < protection_interval_factor <= 1.0:
            return "备货期覆盖系数 α 需在 (0, 1] 之间。"
        return None


_CJK_FONT_CANDIDATES = (
    "PingFang SC",
    "Heiti SC",
    "STHeiti",
    "Arial Unicode MS",
    "Hiragino Sans GB",
)
_PLOT_LABELS_CJK = {
    "ship": "发货量",
    "risk": "缺货风险",
    "current": "当前",
    "suggested": "建议",
}
_PLOT_LABELS_EN = {
    "ship": "ship",
    "risk": "stockout risk",
    "current": "current",
    "suggested": "suggested",
}
_resolved_cjk_font: list[str | None] = []


def _resolve_cjk_font() -> str | None:
    if _resolved_cjk_font:
        return _resolved_cjk_font[0]

    from matplotlib import font_manager

    available = {font.name for font in font_manager.fontManager.ttflist}
    chosen = next(
        (family for family in _CJK_FONT_CANDIDATES if family in available), None
    )
    _resolved_cjk_font.append(chosen)
    return chosen


def _monospace_font() -> QFont:
    font = QFont("Consolas")
    font.setStyleHint(QFont.StyleHint.Monospace)
    return font


def _app_stylesheet() -> str:
    return """
    QWidget {
        font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
        font-size: 13px;
    }
    QGroupBox {
        margin-top: 8px;
        padding-top: 6px;
        font-weight: 600;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 8px;
        padding: 0 2px;
    }
    QPushButton {
        min-height: 28px;
    }
    QPushButton#runButton {
        font-weight: 700;
    }
    QLabel#skcCountLabel {
        font-weight: 600;
    }
    QPlainTextEdit#skcTextEdit,
    QPlainTextEdit#logTextEdit {
        font-size: 12px;
    }
    """


def main() -> int:
    app = QApplication([])
    window = PlannerWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
