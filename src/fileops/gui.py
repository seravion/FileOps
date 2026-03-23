from __future__ import annotations

import os
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


TRANSLATIONS: dict[str, dict[str, str]] = {
    "zh": {
        "window_title": "FileOps 文件操作工具",
        "subtitle": "支持复制/移动/重命名/删除/按大小拆分/文档拆分（标题分段 + 图片文字提取）",
        "group_basic": "基础配置",
        "label_operation": "操作类型",
        "label_language": "语言",
        "label_workspace": "工作区（安全范围）",
        "button_browse": "浏览",
        "group_sources": "源文件列表",
        "button_add_file": "添加文件",
        "button_add_folder": "添加文件夹",
        "button_remove_selected": "移除选中",
        "button_clear_list": "清空列表",
        "group_options": "操作参数",
        "label_destination": "输出目录/目标路径",
        "label_overwrite": "覆盖策略",
        "label_rename_pattern": "重命名模板",
        "label_start_index": "起始序号",
        "radio_trash": "删除到回收站",
        "radio_hard_delete": "永久删除",
        "label_split_size": "分片大小(MB)",
        "label_doc_mode": "标题拆分规则",
        "check_include_ocr": "提取图片文字（OCR）",
        "group_run": "执行",
        "check_dry_run": "预演模式（不写入）",
        "label_report_file": "报告文件",
        "button_save_as": "另存为",
        "button_run": "开始执行",
        "group_log": "执行日志",
        "status_ready": "就绪",
        "status_running": "执行中...",
        "progress_not_started": "进度：未开始",
        "progress_preparing": "进度：0/0（0%）  准备中...",
        "progress_runtime": "进度：{done}/{total}（{percent}%）  {detail}",
        "dialog_result_title": "执行结果",
        "dialog_select_workspace": "选择工作区",
        "dialog_select_destination": "选择输出目录",
        "dialog_select_report_file": "选择报告文件",
        "dialog_select_file": "选择文件",
        "dialog_select_folder": "选择文件夹",
        "dialog_json_filter": "JSON 文件 (*.json);;全部文件 (*.*)",
        "dialog_param_error_title": "参数错误",
        "dialog_confirm_delete_title": "确认永久删除",
        "dialog_confirm_delete_text": "你选择了“永久删除”，该操作不可恢复，是否继续？",
        "error_workspace_diff_disk": "工作区与源路径不在同一磁盘，请调整为同一盘后重试。",
        "error_workspace_infer": "无法自动推导工作区，请手动设置。",
        "workspace_auto_adjusted": "自动调整工作区为：{workspace}",
        "error_no_sources": "请先添加至少一个源文件或目录。",
        "error_missing_destination": "该操作需要指定输出目录/目标路径。",
        "error_missing_pattern": "请填写重命名模板。",
        "log_start_execution": "开始执行：{operation}",
        "op_copy": "复制",
        "op_move": "移动",
        "op_rename": "重命名",
        "op_delete": "删除",
        "op_split": "按大小拆分",
        "op_doc_split": "文档拆分",
        "doc_mode_h1": "按一级标题",
        "doc_mode_h2": "按二级标题",
        "doc_mode_h1_h2": "按一级+二级标题",
        "status_success": "成功",
        "status_failed": "失败",
        "status_skipped": "跳过",
        "status_dry_run": "预演",
        "worker_processing": "处理中：{name}",
        "worker_start_item": "[{idx}/{total}] 开始处理：{source}",
        "worker_done_item": "已完成：{name}",
        "worker_summary": "汇总: 总数={total} 成功={success} 预演={dry_run} 跳过={skipped} 失败={failed}",
        "worker_report_output": "报告输出: {path}",
        "worker_failure_details": "失败详情:",
        "worker_remaining_failures": "- 其余 {count} 条请查看报告文件。",
        "worker_check_log_or_report": "请查看执行日志或报告文件。",
        "worker_finished_with_failures": "执行完成（存在失败）",
        "worker_finished": "执行完成",
        "worker_exception": "[异常] {error}",
    },
    "en": {
        "window_title": "FileOps File Operations Tool",
        "subtitle": "Supports copy/move/rename/delete/split-by-size/document split (heading split + OCR)",
        "group_basic": "Basic Settings",
        "label_operation": "Operation",
        "label_language": "Language",
        "label_workspace": "Workspace (safe scope)",
        "button_browse": "Browse",
        "group_sources": "Source List",
        "button_add_file": "Add File",
        "button_add_folder": "Add Folder",
        "button_remove_selected": "Remove Selected",
        "button_clear_list": "Clear List",
        "group_options": "Operation Parameters",
        "label_destination": "Output Directory / Target Path",
        "label_overwrite": "Overwrite Policy",
        "label_rename_pattern": "Rename Pattern",
        "label_start_index": "Start Index",
        "radio_trash": "Move to Recycle Bin",
        "radio_hard_delete": "Delete Permanently",
        "label_split_size": "Chunk Size (MB)",
        "label_doc_mode": "Heading Split Rule",
        "check_include_ocr": "Extract image text (OCR)",
        "group_run": "Run",
        "check_dry_run": "Dry run mode (no writes)",
        "label_report_file": "Report File",
        "button_save_as": "Save As",
        "button_run": "Start",
        "group_log": "Execution Log",
        "status_ready": "Ready",
        "status_running": "Running...",
        "progress_not_started": "Progress: not started",
        "progress_preparing": "Progress: 0/0 (0%)  Preparing...",
        "progress_runtime": "Progress: {done}/{total} ({percent}%)  {detail}",
        "dialog_result_title": "Execution Result",
        "dialog_select_workspace": "Select Workspace",
        "dialog_select_destination": "Select Output Directory",
        "dialog_select_report_file": "Select Report File",
        "dialog_select_file": "Select File",
        "dialog_select_folder": "Select Folder",
        "dialog_json_filter": "JSON Files (*.json);;All Files (*.*)",
        "dialog_param_error_title": "Parameter Error",
        "dialog_confirm_delete_title": "Confirm Permanent Delete",
        "dialog_confirm_delete_text": "You selected permanent delete. This action cannot be undone. Continue?",
        "error_workspace_diff_disk": "Workspace and source paths are on different drives. Please use the same drive.",
        "error_workspace_infer": "Unable to infer workspace automatically. Please set it manually.",
        "workspace_auto_adjusted": "Workspace auto-adjusted to: {workspace}",
        "error_no_sources": "Add at least one source file or folder first.",
        "error_missing_destination": "This operation requires an output directory or target path.",
        "error_missing_pattern": "Please provide a rename pattern.",
        "log_start_execution": "Start operation: {operation}",
        "op_copy": "Copy",
        "op_move": "Move",
        "op_rename": "Rename",
        "op_delete": "Delete",
        "op_split": "Split by Size",
        "op_doc_split": "Document Split",
        "doc_mode_h1": "By H1",
        "doc_mode_h2": "By H2",
        "doc_mode_h1_h2": "By H1 + H2",
        "status_success": "Success",
        "status_failed": "Failed",
        "status_skipped": "Skipped",
        "status_dry_run": "Dry Run",
        "worker_processing": "Processing: {name}",
        "worker_start_item": "[{idx}/{total}] Start: {source}",
        "worker_done_item": "Completed: {name}",
        "worker_summary": "Summary: total={total} success={success} dry_run={dry_run} skipped={skipped} failed={failed}",
        "worker_report_output": "Report written: {path}",
        "worker_failure_details": "Failure details:",
        "worker_remaining_failures": "- {count} more item(s). See the report file.",
        "worker_check_log_or_report": "Check execution logs or report file for details.",
        "worker_finished_with_failures": "Completed (with failures)",
        "worker_finished": "Completed",
        "worker_exception": "[Exception] {error}",
    },
}

LANGUAGE_OPTIONS: list[tuple[str, str]] = [("zh", "中文"), ("en", "English")]
OPERATION_VALUES: list[str] = ["copy", "move", "rename", "delete", "split", "doc_split"]
DOC_MODE_VALUES: list[str] = ["h1", "h2", "h1_h2"]


def _translate(language: str, key: str, **kwargs: object) -> str:
    fallback_table = TRANSLATIONS["zh"]
    table = TRANSLATIONS.get(language, fallback_table)
    template = table.get(key, fallback_table.get(key, key))
    return template.format(**kwargs)


class OperationWorker(QThread):
    progress_changed = Signal(int, int, str)
    log_message = Signal(str)
    finished_status = Signal(str, bool, str)

    def __init__(self, params: dict[str, object], operation_value_to_label: dict[str, str], language: str) -> None:
        super().__init__()
        self.params = params
        self.operation_value_to_label = operation_value_to_label
        self.language = language if language in TRANSLATIONS else "zh"

    def _tr(self, key: str, **kwargs: object) -> str:
        return _translate(self.language, key, **kwargs)

    def run(self) -> None:
        operation = str(self.params["operation"])
        sources = list(self.params["sources"])
        report = RunReport(
            command=operation,
            dry_run_mode=bool(self.params["dry_run"]),
            workspace=str(self.params["workspace"]),
        )

        status_map = {
            "success": self._tr("status_success"),
            "failed": self._tr("status_failed"),
            "skipped": self._tr("status_skipped"),
            "dry_run": self._tr("status_dry_run"),
        }

        try:
            total = len(sources)
            failure_details: list[str] = []
            for idx, source in enumerate(sources, start=1):
                source_path = Path(source)
                self.progress_changed.emit(idx - 1, total, self._tr("worker_processing", name=source_path.name))
                self.log_message.emit(self._tr("worker_start_item", idx=idx, total=total, source=source_path))

                rename_index = int(self.params.get("start_index", 1)) + idx - 1
                results = self._run_single(operation, source_path, rename_index)
                for item in results:
                    report.add(item)
                    op_label = self.operation_value_to_label.get(item.operation, item.operation)
                    status_text = status_map.get(item.status.value, item.status.value)
                    self.log_message.emit(
                        f"[{status_text}] {op_label} | {item.source} -> {item.destination} | {item.message}"
                    )
                    if item.status.value == "failed":
                        failure_details.append(f"{Path(item.source).name}: {item.message}")

                self.progress_changed.emit(idx, total, self._tr("worker_done_item", name=source_path.name))

            report_path_text = str(self.params["report_path"]).strip()
            output_path = write_report(report, Path(report_path_text).resolve(strict=False)) if report_path_text else None

            summary = report.summary()
            self.log_message.emit("")
            self.log_message.emit(
                self._tr(
                    "worker_summary",
                    total=summary["total"],
                    success=summary["success"],
                    dry_run=summary["dry_run"],
                    skipped=summary["skipped"],
                    failed=summary["failed"],
                )
            )
            if output_path is not None:
                self.log_message.emit(self._tr("worker_report_output", path=output_path))

            has_failure = summary["failed"] > 0
            if has_failure:
                detail_lines = failure_details[:3]
                if detail_lines:
                    self.log_message.emit(self._tr("worker_failure_details"))
                    for line in detail_lines:
                        self.log_message.emit(f"- {line}")
                    remain_count = len(failure_details) - len(detail_lines)
                    if remain_count > 0:
                        self.log_message.emit(self._tr("worker_remaining_failures", count=remain_count))

                detail_text = "\n".join(detail_lines) if detail_lines else self._tr("worker_check_log_or_report")
                self.finished_status.emit(self._tr("worker_finished_with_failures"), True, detail_text)
            else:
                self.finished_status.emit(self._tr("worker_finished"), False, "")

        except Exception as exc:  # noqa: BLE001
            self.log_message.emit(self._tr("worker_exception", error=exc))
            self.finished_status.emit(self._tr("status_failed"), True, str(exc))

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
        self.resize(1160, 820)
        cwd = Path.cwd()
        self.default_workspace = str(Path(cwd.anchor) if cwd.anchor else cwd)

        self.language = "zh"
        self.operation_values = OPERATION_VALUES[:]
        self.doc_mode_values = DOC_MODE_VALUES[:]

        self.worker: OperationWorker | None = None
        self._build_ui()
        self._apply_styles()
        self._apply_language(initial=True)
        self._sync_operation_fields()

    def _tr(self, key: str, **kwargs: object) -> str:
        return _translate(self.language, key, **kwargs)

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
        self.subtitle_label = QLabel("")
        self.subtitle_label.setObjectName("subTitleLabel")
        header_layout.addWidget(title)
        header_layout.addWidget(self.subtitle_label)
        root_layout.addWidget(header)

        self.config_group = QGroupBox("")
        config_layout = QHBoxLayout(self.config_group)
        self.operation_label = QLabel("")
        config_layout.addWidget(self.operation_label)
        self.operation_combo = QComboBox()
        self.operation_combo.currentIndexChanged.connect(lambda _idx: self._sync_operation_fields())
        config_layout.addWidget(self.operation_combo)

        self.language_label = QLabel("")
        config_layout.addWidget(self.language_label)
        self.language_combo = QComboBox()
        for code, label in LANGUAGE_OPTIONS:
            self.language_combo.addItem(label, code)
        self.language_combo.currentIndexChanged.connect(lambda _idx: self._on_language_changed())
        config_layout.addWidget(self.language_combo)

        self.workspace_label = QLabel("")
        config_layout.addWidget(self.workspace_label)
        self.workspace_edit = QLineEdit(self.default_workspace)
        config_layout.addWidget(self.workspace_edit, 1)
        self.browse_workspace_button = QPushButton("")
        self.browse_workspace_button.clicked.connect(self._select_workspace)
        config_layout.addWidget(self.browse_workspace_button)
        root_layout.addWidget(self.config_group)

        self.source_group = QGroupBox("")
        source_layout = QHBoxLayout(self.source_group)
        self.source_list = QListWidget()
        source_layout.addWidget(self.source_list, 1)

        source_button_layout = QVBoxLayout()
        self.add_file_button = QPushButton("")
        self.add_file_button.clicked.connect(self._add_files)
        self.add_folder_button = QPushButton("")
        self.add_folder_button.clicked.connect(self._add_folder)
        self.remove_button = QPushButton("")
        self.remove_button.clicked.connect(self._remove_selected_sources)
        self.clear_button = QPushButton("")
        self.clear_button.clicked.connect(self._clear_sources)

        source_button_layout.addWidget(self.add_file_button)
        source_button_layout.addWidget(self.add_folder_button)
        source_button_layout.addWidget(self.remove_button)
        source_button_layout.addWidget(self.clear_button)
        source_button_layout.addStretch(1)
        source_layout.addLayout(source_button_layout)
        root_layout.addWidget(self.source_group)

        self.options_group = QGroupBox("")
        options_layout = QVBoxLayout(self.options_group)

        row1 = QHBoxLayout()
        self.destination_label = QLabel("")
        row1.addWidget(self.destination_label)
        self.destination_edit = QLineEdit()
        row1.addWidget(self.destination_edit, 1)
        self.browse_dest_button = QPushButton("")
        self.browse_dest_button.clicked.connect(self._select_destination)
        row1.addWidget(self.browse_dest_button)
        options_layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.overwrite_label = QLabel("")
        row2.addWidget(self.overwrite_label)
        self.overwrite_combo = QComboBox()
        self.overwrite_combo.addItems(["never", "always", "rename"])
        row2.addWidget(self.overwrite_combo)

        self.rename_pattern_label = QLabel("")
        row2.addWidget(self.rename_pattern_label)
        self.rename_pattern_edit = QLineEdit("{stem}_{index}{ext}")
        row2.addWidget(self.rename_pattern_edit, 1)

        self.start_index_label = QLabel("")
        row2.addWidget(self.start_index_label)
        self.start_index_spin = QSpinBox()
        self.start_index_spin.setMinimum(1)
        self.start_index_spin.setMaximum(999999)
        self.start_index_spin.setValue(1)
        row2.addWidget(self.start_index_spin)
        options_layout.addLayout(row2)

        row3 = QHBoxLayout()
        self.trash_radio = QRadioButton("")
        self.hard_delete_radio = QRadioButton("")
        self.trash_radio.setChecked(True)
        row3.addWidget(self.trash_radio)
        row3.addWidget(self.hard_delete_radio)

        row3.addSpacing(20)
        self.split_size_label = QLabel("")
        row3.addWidget(self.split_size_label)
        self.split_size_spin = QDoubleSpinBox()
        self.split_size_spin.setMinimum(0.01)
        self.split_size_spin.setMaximum(20480.0)
        self.split_size_spin.setDecimals(2)
        self.split_size_spin.setValue(20.0)
        row3.addWidget(self.split_size_spin)

        row3.addSpacing(20)
        self.doc_mode_label = QLabel("")
        row3.addWidget(self.doc_mode_label)
        self.doc_mode_combo = QComboBox()
        row3.addWidget(self.doc_mode_combo)

        self.include_ocr_check = QCheckBox("")
        self.include_ocr_check.setChecked(True)
        row3.addWidget(self.include_ocr_check)
        row3.addStretch(1)
        options_layout.addLayout(row3)

        root_layout.addWidget(self.options_group)

        self.run_group = QGroupBox("")
        run_layout = QVBoxLayout(self.run_group)

        run_row = QHBoxLayout()
        self.dry_run_check = QCheckBox("")
        run_row.addWidget(self.dry_run_check)
        self.report_label = QLabel("")
        run_row.addWidget(self.report_label)
        self.report_edit = QLineEdit()
        run_row.addWidget(self.report_edit, 1)
        self.save_report_button = QPushButton("")
        self.save_report_button.clicked.connect(self._select_report_file)
        run_row.addWidget(self.save_report_button)

        self.run_button = QPushButton("")
        self.run_button.clicked.connect(self._execute_operation)
        run_row.addWidget(self.run_button)
        run_layout.addLayout(run_row)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        run_layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("")
        run_layout.addWidget(self.progress_label)
        root_layout.addWidget(self.run_group)

        self.log_group = QGroupBox("")
        log_layout = QVBoxLayout(self.log_group)
        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        root_layout.addWidget(self.log_group, 1)

        self.status_label = QLabel("")
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

    def _on_language_changed(self) -> None:
        code = str(self.language_combo.currentData() or "zh")
        if code == self.language:
            return
        self.language = code
        self._apply_language(initial=False)

    def _apply_language(self, initial: bool) -> None:
        self.setWindowTitle(self._tr("window_title"))
        self.subtitle_label.setText(self._tr("subtitle"))
        self.config_group.setTitle(self._tr("group_basic"))
        self.operation_label.setText(self._tr("label_operation"))
        self.language_label.setText(self._tr("label_language"))
        self.workspace_label.setText(self._tr("label_workspace"))
        self.browse_workspace_button.setText(self._tr("button_browse"))
        self.source_group.setTitle(self._tr("group_sources"))
        self.add_file_button.setText(self._tr("button_add_file"))
        self.add_folder_button.setText(self._tr("button_add_folder"))
        self.remove_button.setText(self._tr("button_remove_selected"))
        self.clear_button.setText(self._tr("button_clear_list"))
        self.options_group.setTitle(self._tr("group_options"))
        self.destination_label.setText(self._tr("label_destination"))
        self.browse_dest_button.setText(self._tr("button_browse"))
        self.overwrite_label.setText(self._tr("label_overwrite"))
        self.rename_pattern_label.setText(self._tr("label_rename_pattern"))
        self.start_index_label.setText(self._tr("label_start_index"))
        self.trash_radio.setText(self._tr("radio_trash"))
        self.hard_delete_radio.setText(self._tr("radio_hard_delete"))
        self.split_size_label.setText(self._tr("label_split_size"))
        self.doc_mode_label.setText(self._tr("label_doc_mode"))
        self.include_ocr_check.setText(self._tr("check_include_ocr"))
        self.run_group.setTitle(self._tr("group_run"))
        self.dry_run_check.setText(self._tr("check_dry_run"))
        self.report_label.setText(self._tr("label_report_file"))
        self.save_report_button.setText(self._tr("button_save_as"))
        self.run_button.setText(self._tr("button_run"))
        self.log_group.setTitle(self._tr("group_log"))

        self._rebuild_operation_combo()
        self._rebuild_doc_mode_combo()
        self._sync_operation_fields()

        if initial:
            self.status_label.setText(self._tr("status_ready"))
            self.progress_label.setText(self._tr("progress_not_started"))
            return

        if self.worker is None:
            ready_values = {_translate(code, "status_ready") for code, _label in LANGUAGE_OPTIONS}
            progress_values = {_translate(code, "progress_not_started") for code, _label in LANGUAGE_OPTIONS}
            if self.status_label.text() in ready_values:
                self.status_label.setText(self._tr("status_ready"))
            if self.progress_label.text() in progress_values:
                self.progress_label.setText(self._tr("progress_not_started"))

    def _rebuild_operation_combo(self) -> None:
        current_value = self._current_operation() if self.operation_combo.count() > 0 else "copy"
        self.operation_combo.blockSignals(True)
        self.operation_combo.clear()
        for operation in self.operation_values:
            self.operation_combo.addItem(self._tr(f"op_{operation}"), operation)
        target_index = self.operation_combo.findData(current_value)
        self.operation_combo.setCurrentIndex(target_index if target_index >= 0 else 0)
        self.operation_combo.blockSignals(False)

    def _rebuild_doc_mode_combo(self) -> None:
        current_value = str(self.doc_mode_combo.currentData() or "h1") if self.doc_mode_combo.count() > 0 else "h1"
        self.doc_mode_combo.blockSignals(True)
        self.doc_mode_combo.clear()
        for mode in self.doc_mode_values:
            self.doc_mode_combo.addItem(self._tr(f"doc_mode_{mode}"), mode)
        target_index = self.doc_mode_combo.findData(current_value)
        self.doc_mode_combo.setCurrentIndex(target_index if target_index >= 0 else 0)
        self.doc_mode_combo.blockSignals(False)

    def _operation_value_to_label(self) -> dict[str, str]:
        return {operation: self._tr(f"op_{operation}") for operation in self.operation_values}

    def _current_operation(self) -> str:
        value = self.operation_combo.currentData()
        return str(value) if value else "copy"

    def _set_widget_enabled(self, widget: QWidget, enabled: bool) -> None:
        widget.setEnabled(enabled)

    @staticmethod
    def _is_within_workspace(path: Path, workspace: Path) -> bool:
        try:
            path.resolve(strict=False).relative_to(workspace.resolve(strict=False))
            return True
        except ValueError:
            return False

    def _normalize_workspace(self, workspace: Path, paths: list[Path]) -> tuple[Path, str | None]:
        outside = [path for path in paths if not self._is_within_workspace(path, workspace)]
        if not outside:
            return workspace, None

        candidates: list[str] = [str(workspace.resolve(strict=False))]
        for path in paths:
            resolved = path.resolve(strict=False)
            if resolved.exists() and resolved.is_file():
                candidates.append(str(resolved.parent))
            else:
                candidates.append(str(resolved))

        try:
            common_text = os.path.commonpath(candidates)
        except ValueError as exc:
            raise ValueError(self._tr("error_workspace_diff_disk")) from exc

        common = Path(common_text).resolve(strict=False)
        if common.is_file():
            common = common.parent
        if not common.parts:
            raise ValueError(self._tr("error_workspace_infer"))

        return common, self._tr("workspace_auto_adjusted", workspace=common)

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
        self.language_combo.setEnabled(not running)

    def _append_log(self, text: str) -> None:
        self.log_text.appendPlainText(text)

    def _on_worker_progress(self, done: int, total: int, detail: str) -> None:
        percent = 100 if total == 0 else int((done / total) * 100)
        self.progress_bar.setValue(percent)
        self.progress_label.setText(self._tr("progress_runtime", done=done, total=total, percent=percent, detail=detail))

    def _on_worker_log(self, text: str) -> None:
        self._append_log(text)

    def _on_worker_finished(self, status: str, is_error: bool, detail: str) -> None:
        self._set_running(False)
        self.status_label.setText(status)
        if is_error:
            message = status if not detail else f"{status}\n\n{detail}"
            QMessageBox.critical(self, self._tr("dialog_result_title"), message)
        else:
            QMessageBox.information(self, self._tr("dialog_result_title"), status)

        if self.worker is not None:
            self.worker.deleteLater()
            self.worker = None

    def _select_workspace(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            self._tr("dialog_select_workspace"),
            self.workspace_edit.text().strip() or str(Path.cwd()),
        )
        if selected:
            self.workspace_edit.setText(selected)

    def _select_destination(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            self._tr("dialog_select_destination"),
            self.workspace_edit.text().strip() or str(Path.cwd()),
        )
        if selected:
            self.destination_edit.setText(selected)

    def _select_report_file(self) -> None:
        selected, _ = QFileDialog.getSaveFileName(
            self,
            self._tr("dialog_select_report_file"),
            self.workspace_edit.text().strip() or str(Path.cwd()),
            self._tr("dialog_json_filter"),
        )
        if selected:
            self.report_edit.setText(selected)

    def _add_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self,
            self._tr("dialog_select_file"),
            self.workspace_edit.text().strip() or str(Path.cwd()),
        )
        for file_path in files:
            self._append_source(file_path)

    def _add_folder(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            self._tr("dialog_select_folder"),
            self.workspace_edit.text().strip() or str(Path.cwd()),
        )
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
            raise ValueError(self._tr("error_no_sources"))

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
                raise ValueError(self._tr("error_missing_destination"))
            params["destination"] = Path(dest_text).resolve(strict=False)

        if operation == "rename":
            pattern = self.rename_pattern_edit.text().strip()
            if not pattern:
                raise ValueError(self._tr("error_missing_pattern"))
            params["pattern"] = pattern
            params["start_index"] = int(self.start_index_spin.value())

        if operation == "delete":
            params["use_trash"] = self.trash_radio.isChecked()

        if operation == "split":
            params["split_size_mb"] = float(self.split_size_spin.value())

        if operation == "doc_split":
            params["heading_mode"] = str(self.doc_mode_combo.currentData() or "h1")
            params["include_image_text"] = self.include_ocr_check.isChecked()

        candidate_paths = list(sources)
        if "destination" in params:
            candidate_paths.append(Path(params["destination"]))

        workspace, workspace_note = self._normalize_workspace(workspace, candidate_paths)
        params["workspace"] = workspace
        params["workspace_note"] = workspace_note

        return params

    def _execute_operation(self) -> None:
        if self.worker is not None:
            return

        try:
            params = self._collect_parameters()
        except ValueError as exc:
            QMessageBox.critical(self, self._tr("dialog_param_error_title"), str(exc))
            return

        if self._current_operation() == "delete" and not bool(params["dry_run"]) and not bool(params.get("use_trash", True)):
            confirmed = QMessageBox.question(
                self,
                self._tr("dialog_confirm_delete_title"),
                self._tr("dialog_confirm_delete_text"),
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if confirmed != QMessageBox.Yes:
                return

        self._set_running(True)
        self.progress_bar.setValue(0)
        self.progress_label.setText(self._tr("progress_preparing"))
        self.status_label.setText(self._tr("status_running"))
        self._append_log("----------------------------------------")
        self._append_log(self._tr("log_start_execution", operation=self.operation_combo.currentText()))

        self.workspace_edit.setText(str(params["workspace"]))
        workspace_note = params.get("workspace_note")
        if workspace_note:
            self._append_log(str(workspace_note))

        self.worker = OperationWorker(
            params=params,
            operation_value_to_label=self._operation_value_to_label(),
            language=self.language,
        )
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
