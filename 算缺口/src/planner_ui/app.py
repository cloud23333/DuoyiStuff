from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from PyQt6.QtCore import QObject, Qt, QThread, QUrl, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QDesktopServices, QFont, QGuiApplication
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
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from shipment_planner.engine import DEFAULT_ZERO_SOLD7_WITH_SOLD30_STOCKOUT_MAX_QTY

from planner_ui.workflow import (
    PlannerRunResult,
    ensure_constraints_template,
    extract_unique_skc,
    get_constraints_config_dir,
    get_constraints_path,
    run_planner,
)


@dataclass(slots=True)
class RunRequest:
    orders_path: Path
    sales_path: Path
    output_dir: Path
    sold30_weight: float
    sold7_weight: float
    global_gap_multiplier: float
    zero_sold7_with_sold30_stockout_max_qty: int


class PlannerRunWorker(QObject):
    finished = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(
        self,
        *,
        orders_path: Path,
        sales_path: Path,
        output_dir: Path,
        sold30_weight: float,
        sold7_weight: float,
        global_gap_multiplier: float,
        zero_sold7_with_sold30_stockout_max_qty: int,
    ) -> None:
        super().__init__()
        self._orders_path = orders_path
        self._sales_path = sales_path
        self._output_dir = output_dir
        self._sold30_weight = sold30_weight
        self._sold7_weight = sold7_weight
        self._global_gap_multiplier = global_gap_multiplier
        self._zero_sold7_with_sold30_stockout_max_qty = (
            zero_sold7_with_sold30_stockout_max_qty
        )

    @pyqtSlot()
    def run(self) -> None:
        try:
            result = run_planner(
                orders_path=self._orders_path,
                sales_path=self._sales_path,
                output_dir=self._output_dir,
                sold30_weight=self._sold30_weight,
                sold7_weight=self._sold7_weight,
                global_gap_multiplier=self._global_gap_multiplier,
                zero_sold7_with_sold30_stockout_max_qty=(
                    self._zero_sold7_with_sold30_stockout_max_qty
                ),
            )
        except Exception as exc:  # pragma: no cover - UI error channel
            self.failed.emit(str(exc))
            return
        self.finished.emit(result)


class PlannerWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self._run_thread: QThread | None = None
        self._run_worker: PlannerRunWorker | None = None
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

        self.constraints_path_edit = QLineEdit()
        self.constraints_path_edit.setReadOnly(True)
        self.constraints_path_edit.setText(str(get_constraints_path()))
        self.constraints_path_edit.setToolTip(self.constraints_path_edit.text())

        self.open_config_dir_button = QPushButton("配置目录")
        self.open_config_dir_button.clicked.connect(self._on_open_config_dir)
        self.open_config_dir_button.setMinimumWidth(92)

        layout.addWidget(QLabel("订单"), 0, 0)
        layout.addWidget(self.order_path_edit, 0, 1)
        layout.addWidget(self.order_browse_button, 0, 2)
        layout.addWidget(QLabel("销售"), 1, 0)
        layout.addWidget(self.sales_path_edit, 1, 1)
        layout.addWidget(self.sales_browse_button, 1, 2)
        layout.addWidget(QLabel("输出"), 2, 0)
        layout.addWidget(self.output_dir_edit, 2, 1)
        layout.addWidget(self.output_browse_button, 2, 2)
        layout.addWidget(QLabel("配置"), 3, 0)
        layout.addWidget(self.constraints_path_edit, 3, 1)
        layout.addWidget(self.open_config_dir_button, 3, 2)
        return group

    def _build_run_group(self) -> QGroupBox:
        group = QGroupBox("参数")
        layout = QVBoxLayout(group)
        layout.setSpacing(6)

        params_grid = QGridLayout()
        params_grid.setHorizontalSpacing(8)
        params_grid.setVerticalSpacing(6)
        params_grid.setColumnStretch(5, 1)

        self.sold7_weight_spin = QDoubleSpinBox()
        self.sold7_weight_spin.setDecimals(2)
        self.sold7_weight_spin.setRange(0.0, 1.0)
        self.sold7_weight_spin.setSingleStep(0.01)
        self.sold7_weight_spin.setValue(0.8)
        self.sold7_weight_spin.setFixedWidth(108)

        self.sold30_weight_spin = QDoubleSpinBox()
        self.sold30_weight_spin.setDecimals(2)
        self.sold30_weight_spin.setRange(0.0, 1.0)
        self.sold30_weight_spin.setSingleStep(0.01)
        self.sold30_weight_spin.setValue(0.2)
        self.sold30_weight_spin.setFixedWidth(108)

        self.sold7_weight_spin.valueChanged.connect(self._on_sold7_weight_changed)
        self.sold30_weight_spin.valueChanged.connect(self._on_sold30_weight_changed)

        self.global_gap_multiplier_spin = QDoubleSpinBox()
        self.global_gap_multiplier_spin.setDecimals(2)
        self.global_gap_multiplier_spin.setRange(0.0001, 9999.0)
        self.global_gap_multiplier_spin.setSingleStep(0.01)
        self.global_gap_multiplier_spin.setValue(1.0)
        self.global_gap_multiplier_spin.setFixedWidth(108)

        self.zero_sold7_stockout_cap_spin = QSpinBox()
        self.zero_sold7_stockout_cap_spin.setRange(0, 999999)
        self.zero_sold7_stockout_cap_spin.setSingleStep(1)
        self.zero_sold7_stockout_cap_spin.setValue(
            DEFAULT_ZERO_SOLD7_WITH_SOLD30_STOCKOUT_MAX_QTY
        )
        self.zero_sold7_stockout_cap_spin.setFixedWidth(108)

        self.status_label = QLabel("")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        self.status_label.setMinimumHeight(28)

        self.run_button = QPushButton("开始运行")
        self.run_button.setObjectName("runButton")
        self.run_button.clicked.connect(self._on_run_clicked)
        self.run_button.setMinimumWidth(110)

        weight_link_label = QLabel("↔")
        weight_link_label.setObjectName("weightLinkLabel")
        weight_link_label.setToolTip("7日与30日联动，二者之和始终为 1。")
        weight_link_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        weight_link_label.setFixedWidth(18)

        params_grid.addWidget(QLabel("7日"), 0, 0)
        params_grid.addWidget(self.sold7_weight_spin, 0, 1)
        params_grid.addWidget(weight_link_label, 0, 2)
        params_grid.addWidget(QLabel("30日"), 0, 3)
        params_grid.addWidget(self.sold30_weight_spin, 0, 4)

        params_grid.addWidget(QLabel("系数"), 1, 0)
        params_grid.addWidget(self.global_gap_multiplier_spin, 1, 1)
        params_grid.addWidget(QLabel("保底"), 1, 3)
        params_grid.addWidget(self.zero_sold7_stockout_cap_spin, 1, 4)

        action_row = QHBoxLayout()
        action_row.setSpacing(8)
        action_row.addWidget(self.status_label, stretch=1)
        action_row.addWidget(self.run_button)

        layout.addLayout(params_grid)
        layout.addLayout(action_row)
        return group

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

        QMessageBox.information(
            self,
            "运行完成",
            "\n".join(
                [
                    "输出文件已生成：",
                    f"目录：{result_obj.output_dir}",
                    str(result_obj.recommendation_path),
                    str(result_obj.quality_path),
                    str(result_obj.summary_path),
                ]
            ),
        )

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
        for control in (
            self.order_browse_button,
            self.sales_browse_button,
            self.output_browse_button,
            self.open_config_dir_button,
            self.sold7_weight_spin,
            self.sold30_weight_spin,
            self.global_gap_multiplier_spin,
            self.zero_sold7_stockout_cap_spin,
            self.clear_log_button,
        ):
            control.setEnabled(not running)

        if running:
            self.copy_skc_button.setEnabled(False)
            self.run_button.setEnabled(False)
            return

        self.copy_skc_button.setEnabled(bool(self.skc_text_edit.toPlainText().strip()))
        self._refresh_run_button_state()

    def _set_status(self, message: str, *, error: bool = False) -> None:
        self.status_label.setText(message)
        color = "#b42318" if error else "#166534"
        self.status_label.setStyleSheet(f"color: {color};")

    @pyqtSlot(float)
    def _on_sold7_weight_changed(self, value: float) -> None:
        self._sync_weight_pair(paired_spin=self.sold30_weight_spin, value=value)

    @pyqtSlot(float)
    def _on_sold30_weight_changed(self, value: float) -> None:
        self._sync_weight_pair(paired_spin=self.sold7_weight_spin, value=value)

    def _sync_weight_pair(self, *, paired_spin: QDoubleSpinBox, value: float) -> None:
        normalized = min(max(value, 0.0), 1.0)
        paired = round(1.0 - normalized, 2)
        self._set_spin_without_signal(paired_spin, paired)

    def _set_spin_without_signal(self, spin: QDoubleSpinBox, value: float) -> None:
        spin.blockSignals(True)
        spin.setValue(value)
        spin.blockSignals(False)

    def _append_log(self, message: str) -> None:
        if not message:
            return
        self.log_text_edit.appendPlainText(message)
        cursor = self.log_text_edit.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.log_text_edit.setTextCursor(cursor)

    def _collect_run_request(self) -> RunRequest | None:
        if not self._constraints_ready:
            message = "约束配置未就绪，请点击“打开配置目录”修复后再运行。"
            self._set_status(message, error=True)
            QMessageBox.warning(self, "配置未就绪", message)
            return None

        orders_text = self.order_path_edit.text().strip()
        sales_text = self.sales_path_edit.text().strip()
        output_text = self.output_dir_edit.text().strip()
        if not orders_text or not sales_text or not output_text:
            QMessageBox.warning(self, "信息不完整", "请先选择订单文件、销售文件和输出目录。")
            return None

        run_request = RunRequest(
            orders_path=Path(orders_text),
            sales_path=Path(sales_text),
            output_dir=Path(output_text),
            sold30_weight=float(self.sold30_weight_spin.value()),
            sold7_weight=float(self.sold7_weight_spin.value()),
            global_gap_multiplier=float(self.global_gap_multiplier_spin.value()),
            zero_sold7_with_sold30_stockout_max_qty=int(
                self.zero_sold7_stockout_cap_spin.value()
            ),
        )

        validation_error = self._validate_run_inputs(
            orders_path=run_request.orders_path,
            sales_path=run_request.sales_path,
            output_dir=run_request.output_dir,
            sold30_weight=run_request.sold30_weight,
            sold7_weight=run_request.sold7_weight,
            global_gap_multiplier=run_request.global_gap_multiplier,
            zero_sold7_with_sold30_stockout_max_qty=(
                run_request.zero_sold7_with_sold30_stockout_max_qty
            ),
        )
        if validation_error is not None:
            self._set_status(validation_error, error=True)
            QMessageBox.warning(self, "输入无效", validation_error)
            return None

        return run_request

    def _start_run(self, run_request: RunRequest) -> None:
        self._set_running_state(True)
        self._set_status("正在运行，请稍候...")
        self._append_log("开始运行发货建议计算...")

        self._run_thread = QThread(self)
        self._run_worker = PlannerRunWorker(
            orders_path=run_request.orders_path,
            sales_path=run_request.sales_path,
            output_dir=run_request.output_dir,
            sold30_weight=run_request.sold30_weight,
            sold7_weight=run_request.sold7_weight,
            global_gap_multiplier=run_request.global_gap_multiplier,
            zero_sold7_with_sold30_stockout_max_qty=(
                run_request.zero_sold7_with_sold30_stockout_max_qty
            ),
        )
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
        output_text = self.output_dir_edit.text().strip()
        if not orders_text or not sales_text or not output_text:
            return False

        orders_path = Path(orders_text)
        sales_path = Path(sales_text)
        output_path = Path(output_text)
        return (
            orders_path.is_file()
            and sales_path.is_file()
            and orders_path.suffix.lower() == ".xlsx"
            and sales_path.suffix.lower() == ".xlsx"
            and (output_path.is_dir() or not output_path.exists())
        )

    def _remember_dialog_dir(self, path: Path) -> None:
        candidate = path if path.is_dir() else path.parent
        if candidate.exists() and candidate.is_dir():
            self._last_dialog_dir = candidate.resolve()

    def _validate_run_inputs(
        self,
        *,
        orders_path: Path,
        sales_path: Path,
        output_dir: Path,
        sold30_weight: float,
        sold7_weight: float,
        global_gap_multiplier: float,
        zero_sold7_with_sold30_stockout_max_qty: int,
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

        if output_dir.exists() and not output_dir.is_dir():
            return f"输出路径不是目录：{output_dir}"
        if sold30_weight < 0 or sold7_weight < 0:
            return "7天和30天销量占比不能为负数。"
        if sold30_weight > 1 or sold7_weight > 1:
            return "7天和30天销量占比不能大于 1。"
        if abs((sold30_weight + sold7_weight) - 1.0) > 0.0001:
            return "7天和30天销量占比之和必须为 1。"
        if global_gap_multiplier <= 0:
            return "全局缺口上浮系数必须大于 0。"
        if zero_sold7_with_sold30_stockout_max_qty < 0:
            return "保底不能为负数。"
        return None


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
    QLabel#weightLinkLabel {
        font-weight: 700;
        color: #64748b;
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
