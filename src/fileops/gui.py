from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import QThread, Signal, Qt
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QProgressBar,
    QRadioButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from .document_split import split_documents_by_structure
from .models import OperationResult, RunReport
from .operations import CommonOptions, copy_items, delete_items, move_items, rename_items, split_items
from .reporting import write_report


class OperationWorker(QThread):
    progress_changed = Signal(int, int, str)
    log_message = Signal(str)
    finished_status = Signal(str, bool)

    def __init__(self, params: dict[str, object], operation_value_to_label: dict[str, str]) -> None:
        super().__init__()
        self.params = params
        self.operation_value_to_label = operation_value_to_label

    def run(self) -> None:
        operation = str(self.params["operation"])
        sources = list(self.params["sources"])
        report = RunReport(
            command=operation,
            dry_run_mode=bool(self.params["dry_run"]),
            workspace=str(self.params["workspace"]),
        )

        status_map = {
            "success": "成功",
            "failed": "失败",
            "skipped": "跳过",
            "dry_run": "预演",
        }

        try:
            total = len(sources)
            for idx, source in enumerate(sources, start=1):
                source_path = Path(source)
                self.progress_changed.emit(idx - 1, total, f"处理中：{source_path.name}")
                self.log_message.emit(f"[{idx}/{total}] 开始处理：{source_path}")

                rename_index = int(self.params.get("start_index", 1)) + idx - 1
                results = self._run_single(operation, source_path, rename_index)
                for item in results:
                    report.add(item)
                    op_label = self.operation_value_to_label.get(item.operation, item.operation)
                    status_text = status_map.get(item.status.value, item.status.value)
                    self.log_message.emit(
                        f"[{status_text}] {op_label} | {item.source} -> {item.destination} | {item.message}"
                    )

                self.progress_changed.emit(idx, total, f"已完成：{source_path.name}")

            report_path_text = str(self.params["report_path"]).strip()
            output_path = write_report(report, Path(report_path_text).resolve(strict=False)) if report_path_text else None

            summary = report.summary()
            self.log_message.emit("")
            self.log_message.emit(
                "汇总: "
                f"总数={summary['total']} 成功={summary['success']} 预演={summary['dry_run']} "
                f"跳过={summary['skipped']} 失败={summary['failed']}"
            )
            if output_path is not None:
                self.log_message.emit(f"报告输出: {output_path}")

            has_failure = summary["failed"] > 0
            final_status = "执行完成（存在失败）" if has_failure else "执行完成"
            self.finished_status.emit(final_status, has_failure)

        except Exception as exc:  # noqa: BLE001
            self.finished_status.emit(f"执行失败：{exc}", True)

    def _run_single(self, operation: str, source: Path, rename_index: int) -> list[OperationResult]:
        workspace = Path(self.params["workspace"])
        dry_run = bool(self.params["dry_run"])

        if operation in {"copy", "move", "rename", "split"}:
            common = CommonOptions(
                workspace=workspace,
                dry_run=dry_run,
                overwrite=str(self.params["overwrite"]),
            )

            if operation == "copy":
                return copy_items([source], Path(self.params["destination"]), common)
            if operation == "move":
                return move_items([source], Path(self.params["destination"]), common)
            if operation == "rename":
                return rename_items([source], str(self.params["pattern"]), rename_index, common)
            return split_items([source], Path(self.params["destination"]), float(self.params["split_size_mb"]), common)

        if operation == "doc_split":
            return split_documents_by_structure(
                sources=[source],
                destination=Path(self.params["destination"]),
                workspace=workspace,
                dry_run=dry_run,
                heading_mode=str(self.params["heading_mode"]),
                include_image_text=bool(self.params["include_image_text"]),
            )

        return delete_items(
            sources=[source],
            workspace=workspace,
            dry_run=dry_run,
            use_trash=bool(self.params["use_trash"]),
        )


class FileOpsWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("FileOps 文件操作工具")
        self.resize(1160, 820)

        self.operation_label_to_value = {
            "复制": "copy",
            "移动": "move",
            "重命名": "rename",
            "删除": "delete",
            "按大小拆分": "split",
            "文档拆分": "doc_split",
        }
        self.operation_value_to_label = {value: key for key, value in self.operation_label_to_value.items()}
        self.doc_mode_label_to_value = {
            "按一级标题": "h1",
            "按二级标题": "h2",
            "按一级+二级标题": "h1_h2",
        }

        self.worker: OperationWorker | None = None
        self._build_ui()
        self._apply_styles()
        self._sync_operation_fields()

    def _build_ui(self) -> None:
        central = QWidget(self)
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(10)

        header = QFrame()
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        title = QLabel("FileOps")
        title.setObjectName("titleLabel")
        subtitle = QLabel("支持复制/移动/重命名/删除/按大小拆分/文档拆分（标题分段 + 图片文字提取）")
        subtitle.setObjectName("subTitleLabel")
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        root_layout.addWidget(header)

        config_group = QGroupBox("基础配置")
        config_layout = QHBoxLayout(config_group)
        config_layout.addWidget(QLabel("操作类型"))
        self.operation_combo = QComboBox()
        self.operation_combo.addItems(list(self.operation_label_to_value.keys()))
        self.operation_combo.currentTextChanged.connect(self._sync_operation_fields)
        config_layout.addWidget(self.operation_combo)

        config_layout.addWidget(QLabel("工作区"))
        self.workspace_edit = QLineEdit(str(Path.cwd()))
        config_layout.addWidget(self.workspace_edit, 1)
        browse_workspace_button = QPushButton("浏览")
        browse_workspace_button.clicked.connect(self._select_workspace)
        config_layout.addWidget(browse_workspace_button)
        root_layout.addWidget(config_group)

        source_group = QGroupBox("源文件列表")
        source_layout = QHBoxLayout(source_group)
        self.source_list = QListWidget()
        source_layout.addWidget(self.source_list, 1)

        source_button_layout = QVBoxLayout()
        add_file_button = QPushButton("添加文件")
        add_file_button.clicked.connect(self._add_files)
        add_folder_button = QPushButton("添加文件夹")
        add_folder_button.clicked.connect(self._add_folder)
        remove_button = QPushButton("移除选中")
        remove_button.clicked.connect(self._remove_selected_sources)
        clear_button = QPushButton("清空列表")
        clear_button.clicked.connect(self._clear_sources)

        source_button_layout.addWidget(add_file_button)
        source_button_layout.addWidget(add_folder_button)
        source_button_layout.addWidget(remove_button)
        source_button_layout.addWidget(clear_button)
        source_button_layout.addStretch(1)
        source_layout.addLayout(source_button_layout)
        root_layout.addWidget(source_group)

        options_group = QGroupBox("操作参数")
        options_layout = QVBoxLayout(options_group)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("输出目录/目标路径"))
        self.destination_edit = QLineEdit()
        row1.addWidget(self.destination_edit, 1)
        browse_dest_button = QPushButton("浏览")
        browse_dest_button.clicked.connect(self._select_destination)
        row1.addWidget(browse_dest_button)
        options_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("覆盖策略"))
        self.overwrite_combo = QComboBox()
        self.overwrite_combo.addItems(["never", "always", "rename"])
        row2.addWidget(self.overwrite_combo)

        row2.addWidget(QLabel("重命名模板"))
        self.rename_pattern_edit = QLineEdit("{stem}_{index}{ext}")
        row2.addWidget(self.rename_pattern_edit, 1)

        row2.addWidget(QLabel("起始序号"))
        self.start_index_spin = QSpinBox()
        self.start_index_spin.setMinimum(1)
        self.start_index_spin.setMaximum(999999)
        self.start_index_spin.setValue(1)
        row2.addWidget(self.start_index_spin)
        options_layout.addLayout(row2)

        row3 = QHBoxLayout()
        self.trash_radio = QRadioButton("删除到回收站")
        self.hard_delete_radio = QRadioButton("永久删除")
        self.trash_radio.setChecked(True)
        row3.addWidget(self.trash_radio)
        row3.addWidget(self.hard_delete_radio)

        row3.addSpacing(20)
        row3.addWidget(QLabel("分片大小(MB)"))
        self.split_size_spin = QDoubleSpinBox()
        self.split_size_spin.setMinimum(0.01)
        self.split_size_spin.setMaximum(20480.0)
        self.split_size_spin.setDecimals(2)
        self.split_size_spin.setValue(20.0)
        row3.addWidget(self.split_size_spin)

        row3.addSpacing(20)
        row3.addWidget(QLabel("标题拆分规则"))
        self.doc_mode_combo = QComboBox()
        self.doc_mode_combo.addItems(list(self.doc_mode_label_to_value.keys()))
        row3.addWidget(self.doc_mode_combo)

        self.include_ocr_check = QCheckBox("提取图片文字（OCR）")
        self.include_ocr_check.setChecked(True)
        row3.addWidget(self.include_ocr_check)
        row3.addStretch(1)
        options_layout.addLayout(row3)

        root_layout.addWidget(options_group)

        run_group = QGroupBox("执行")
        run_layout = QVBoxLayout(run_group)

        run_row = QHBoxLayout()
        self.dry_run_check = QCheckBox("预演模式（不写入）")
        run_row.addWidget(self.dry_run_check)
        run_row.addWidget(QLabel("报告文件"))
        self.report_edit = QLineEdit()
        run_row.addWidget(self.report_edit, 1)
        save_report_button = QPushButton("另存为")
        save_report_button.clicked.connect(self._select_report_file)
        run_row.addWidget(save_report_button)

        self.run_button = QPushButton("开始执行")
        self.run_button.clicked.connect(self._execute_operation)
        run_row.addWidget(self.run_button)
        run_layout.addLayout(run_row)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        run_layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("进度：未开始")
        run_layout.addWidget(self.progress_label)
        root_layout.addWidget(run_group)

        log_group = QGroupBox("执行日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        root_layout.addWidget(log_group, 1)

        self.status_label = QLabel("就绪")
        root_layout.addWidget(self.status_label)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                font-family: 'Microsoft YaHei UI';
                font-size: 10pt;
            }
            QMainWindow {
                background: #f3f6fb;
            }
            QGroupBox {
                border: 1px solid #dbe2ef;
                border-radius: 8px;
                margin-top: 10px;
                background: #ffffff;
                font-weight: 600;
                color: #1e293b;
                padding-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 6px;
            }
            QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QListWidget, QPlainTextEdit {
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 4px 6px;
                background: #ffffff;
                color: #0f172a;
            }
            QPushButton {
                background: #e2e8f0;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 6px 12px;
            }
            QPushButton:hover {
                background: #d5deea;
            }
            QPushButton:disabled {
                color: #94a3b8;
                background: #f1f5f9;
            }
            QProgressBar {
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                text-align: center;
                background: #e2e8f0;
            }
            QProgressBar::chunk {
                background: #3b82f6;
                border-radius: 5px;
            }
            #titleLabel {
                font-size: 22px;
                font-weight: 700;
                color: #0f172a;
            }
            #subTitleLabel {
                color: #475569;
            }
            """
        )

    def _current_operation(self) -> str:
        label = self.operation_combo.currentText().strip()
        return self.operation_label_to_value.get(label, "copy")

    def _set_widget_enabled(self, widget: QWidget, enabled: bool) -> None:
        widget.setEnabled(enabled)

    def _sync_operation_fields(self) -> None:
        operation = self._current_operation()

        show_destination = operation in {"copy", "move", "split", "doc_split"}
        show_overwrite = operation in {"copy", "move", "rename", "split"}
        show_rename = operation == "rename"
        show_delete = operation == "delete"
        show_split = operation == "split"
        show_doc_split = operation == "doc_split"

        self._set_widget_enabled(self.destination_edit, show_destination)
        self._set_widget_enabled(self.overwrite_combo, show_overwrite)
        self._set_widget_enabled(self.rename_pattern_edit, show_rename)
        self._set_widget_enabled(self.start_index_spin, show_rename)
        self._set_widget_enabled(self.trash_radio, show_delete)
        self._set_widget_enabled(self.hard_delete_radio, show_delete)
        self._set_widget_enabled(self.split_size_spin, show_split)
        self._set_widget_enabled(self.doc_mode_combo, show_doc_split)
        self._set_widget_enabled(self.include_ocr_check, show_doc_split)

    def _set_running(self, running: bool) -> None:
        self.run_button.setEnabled(not running)
        self.operation_combo.setEnabled(not running)

    def _append_log(self, text: str) -> None:
        self.log_text.appendPlainText(text)

    def _on_worker_progress(self, done: int, total: int, detail: str) -> None:
        percent = 100 if total == 0 else int((done / total) * 100)
        self.progress_bar.setValue(percent)
        self.progress_label.setText(f"进度：{done}/{total}（{percent}%）  {detail}")

    def _on_worker_log(self, text: str) -> None:
        self._append_log(text)

    def _on_worker_finished(self, status: str, is_error: bool) -> None:
        self._set_running(False)
        self.status_label.setText(status)
        if is_error:
            QMessageBox.critical(self, "执行结果", status)
        else:
            QMessageBox.information(self, "执行结果", status)

        if self.worker is not None:
            self.worker.deleteLater()
            self.worker = None

    def _select_workspace(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择工作区", self.workspace_edit.text().strip() or str(Path.cwd()))
        if selected:
            self.workspace_edit.setText(selected)

    def _select_destination(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择输出目录", self.workspace_edit.text().strip() or str(Path.cwd()))
        if selected:
            self.destination_edit.setText(selected)

    def _select_report_file(self) -> None:
        selected, _ = QFileDialog.getSaveFileName(
            self,
            "选择报告文件",
            self.workspace_edit.text().strip() or str(Path.cwd()),
            "JSON 文件 (*.json);;全部文件 (*.*)",
        )
        if selected:
            self.report_edit.setText(selected)

    def _add_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "选择文件", self.workspace_edit.text().strip() or str(Path.cwd()))
        for file_path in files:
            self._append_source(file_path)

    def _add_folder(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择文件夹", self.workspace_edit.text().strip() or str(Path.cwd()))
        if selected:
            self._append_source(selected)

    def _append_source(self, value: str) -> None:
        for idx in range(self.source_list.count()):
            if self.source_list.item(idx).text() == value:
                return
        self.source_list.addItem(QListWidgetItem(value))

    def _remove_selected_sources(self) -> None:
        for item in self.source_list.selectedItems():
            self.source_list.takeItem(self.source_list.row(item))

    def _clear_sources(self) -> None:
        self.source_list.clear()

    def _collect_parameters(self) -> dict[str, object]:
        operation = self._current_operation()
        workspace = Path(self.workspace_edit.text().strip() or ".").resolve(strict=False)

        sources: list[Path] = []
        for idx in range(self.source_list.count()):
            sources.append(Path(self.source_list.item(idx).text()).resolve(strict=False))
        if not sources:
            raise ValueError("请先添加至少一个源文件或目录。")

        params: dict[str, object] = {
            "operation": operation,
            "workspace": workspace,
            "sources": sources,
            "dry_run": self.dry_run_check.isChecked(),
            "overwrite": self.overwrite_combo.currentText().strip() or "never",
            "report_path": self.report_edit.text().strip(),
        }

        if operation in {"copy", "move", "split", "doc_split"}:
            dest_text = self.destination_edit.text().strip()
            if not dest_text:
                raise ValueError("该操作需要指定输出目录/目标路径。")
            params["destination"] = Path(dest_text).resolve(strict=False)

        if operation == "rename":
            pattern = self.rename_pattern_edit.text().strip()
            if not pattern:
                raise ValueError("请填写重命名模板。")
            params["pattern"] = pattern
            params["start_index"] = int(self.start_index_spin.value())

        if operation == "delete":
            params["use_trash"] = self.trash_radio.isChecked()

        if operation == "split":
            params["split_size_mb"] = float(self.split_size_spin.value())

        if operation == "doc_split":
            params["heading_mode"] = self.doc_mode_label_to_value[self.doc_mode_combo.currentText()]
            params["include_image_text"] = self.include_ocr_check.isChecked()

        return params

    def _execute_operation(self) -> None:
        if self.worker is not None:
            return

        try:
            params = self._collect_parameters()
        except ValueError as exc:
            QMessageBox.critical(self, "参数错误", str(exc))
            return

        if self._current_operation() == "delete" and not bool(params["dry_run"]) and not bool(params.get("use_trash", True)):
            confirmed = QMessageBox.question(
                self,
                "确认永久删除",
                "你选择了“永久删除”，该操作不可恢复，是否继续？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if confirmed != QMessageBox.Yes:
                return

        self._set_running(True)
        self.progress_bar.setValue(0)
        self.progress_label.setText("进度：0/0（0%）  准备中...")
        self.status_label.setText("执行中...")
        self._append_log("----------------------------------------")
        self._append_log(f"开始执行：{self.operation_combo.currentText()}")

        self.worker = OperationWorker(params, self.operation_value_to_label)
        self.worker.progress_changed.connect(self._on_worker_progress)
        self.worker.log_message.connect(self._on_worker_log)
        self.worker.finished_status.connect(self._on_worker_finished)
        self.worker.start()


def launch_gui() -> None:
    app = QApplication.instance() or QApplication(sys.argv)
    window = FileOpsWindow()
    window.show()
    app.exec()


if __name__ == "__main__":
    launch_gui()
