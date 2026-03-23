from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .document_split import split_documents_by_structure
from .models import OperationResult, RunReport
from .operations import CommonOptions, copy_items, delete_items, move_items, rename_items, split_items
from .reporting import write_report


class FileOpsGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("FileOps 文件操作工具")
        self.root.geometry("1120x780")
        self.root.minsize(980, 700)

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

        self.operation_label_var = tk.StringVar(value="复制")
        self.workspace_var = tk.StringVar(value=str(Path.cwd()))
        self.destination_var = tk.StringVar(value="")
        self.rename_pattern_var = tk.StringVar(value="{stem}_{index}{ext}")
        self.start_index_var = tk.IntVar(value=1)
        self.overwrite_var = tk.StringVar(value="never")
        self.file_split_size_var = tk.StringVar(value="20")
        self.doc_mode_label_var = tk.StringVar(value="按一级+二级标题")
        self.include_image_text_var = tk.BooleanVar(value=True)
        self.report_path_var = tk.StringVar(value="")
        self.dry_run_var = tk.BooleanVar(value=False)
        self.use_trash_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="就绪")
        self.progress_text_var = tk.StringVar(value="进度：未开始")
        self.progress_var = tk.DoubleVar(value=0.0)

        self._build_styles()
        self._build_ui()
        self._sync_operation_fields()

    def _build_styles(self) -> None:
        self.root.configure(bg="#f3f6fb")
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", font=("Microsoft YaHei UI", 10))
        style.configure("Page.TFrame", background="#f3f6fb")
        style.configure("Header.TFrame", background="#f3f6fb")
        style.configure("HeaderTitle.TLabel", background="#f3f6fb", foreground="#0f172a", font=("Microsoft YaHei UI", 16, "bold"))
        style.configure("HeaderSub.TLabel", background="#f3f6fb", foreground="#475569", font=("Microsoft YaHei UI", 10))
        style.configure("Card.TLabelframe", background="#ffffff", bordercolor="#dbe2ef", relief="solid", borderwidth=1)
        style.configure("Card.TLabelframe.Label", background="#ffffff", foreground="#1e293b", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(14, 7))
        style.configure("Progress.Horizontal.TProgressbar", troughcolor="#e2e8f0", background="#3b82f6", thickness=14)

    def _build_ui(self) -> None:
        page = ttk.Frame(self.root, style="Page.TFrame", padding=12)
        page.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(page, style="Header.TFrame")
        header.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(header, text="FileOps", style="HeaderTitle.TLabel").pack(anchor=tk.W)
        ttk.Label(
            header,
            text="支持复制/移动/重命名/删除/按大小拆分/文档拆分（标题分段 + 图片文字提取）",
            style="HeaderSub.TLabel",
        ).pack(anchor=tk.W, pady=(2, 0))

        top = ttk.LabelFrame(page, text="基础配置", style="Card.TLabelframe", padding=10)
        top.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(top, text="操作类型").grid(row=0, column=0, sticky=tk.W)
        self.operation_box = ttk.Combobox(
            top,
            textvariable=self.operation_label_var,
            state="readonly",
            values=list(self.operation_label_to_value.keys()),
            width=18,
        )
        self.operation_box.grid(row=0, column=1, padx=(8, 18), sticky=tk.W)
        self.operation_box.bind("<<ComboboxSelected>>", lambda _event: self._sync_operation_fields())

        ttk.Label(top, text="工作区").grid(row=0, column=2, sticky=tk.W)
        ttk.Entry(top, textvariable=self.workspace_var, width=74).grid(row=0, column=3, padx=8, sticky=tk.EW)
        ttk.Button(top, text="浏览", command=self._select_workspace).grid(row=0, column=4, padx=(6, 0), sticky=tk.W)
        top.grid_columnconfigure(3, weight=1)

        source_frame = ttk.LabelFrame(page, text="源文件列表", style="Card.TLabelframe", padding=10)
        source_frame.pack(fill=tk.BOTH, expand=False, pady=(0, 10))

        self.source_list = tk.Listbox(
            source_frame,
            height=8,
            selectmode=tk.EXTENDED,
            bg="#f8fafc",
            fg="#0f172a",
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground="#cbd5e1",
            highlightcolor="#93c5fd",
            font=("Microsoft YaHei UI", 10),
        )
        self.source_list.grid(row=0, column=0, rowspan=4, sticky=tk.NSEW)

        scrollbar = ttk.Scrollbar(source_frame, orient=tk.VERTICAL, command=self.source_list.yview)
        scrollbar.grid(row=0, column=1, rowspan=4, sticky=tk.NS)
        self.source_list.configure(yscrollcommand=scrollbar.set)

        ttk.Button(source_frame, text="添加文件", command=self._add_files).grid(row=0, column=2, padx=(10, 0), pady=(0, 6), sticky=tk.EW)
        ttk.Button(source_frame, text="添加文件夹", command=self._add_folder).grid(row=1, column=2, padx=(10, 0), pady=6, sticky=tk.EW)
        ttk.Button(source_frame, text="移除选中", command=self._remove_selected_sources).grid(row=2, column=2, padx=(10, 0), pady=6, sticky=tk.EW)
        ttk.Button(source_frame, text="清空列表", command=self._clear_sources).grid(row=3, column=2, padx=(10, 0), pady=(6, 0), sticky=tk.EW)

        source_frame.grid_columnconfigure(0, weight=1)
        source_frame.grid_rowconfigure(0, weight=1)

        options_frame = ttk.LabelFrame(page, text="操作参数", style="Card.TLabelframe", padding=10)
        options_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(options_frame, text="输出目录/目标路径").grid(row=0, column=0, sticky=tk.W)
        self.destination_entry = ttk.Entry(options_frame, textvariable=self.destination_var, width=74)
        self.destination_entry.grid(row=0, column=1, padx=8, sticky=tk.EW)
        self.destination_button = ttk.Button(options_frame, text="浏览", command=self._select_destination)
        self.destination_button.grid(row=0, column=2, padx=(6, 0), sticky=tk.W)

        ttk.Label(options_frame, text="覆盖策略").grid(row=1, column=0, sticky=tk.W, pady=(8, 0))
        self.overwrite_box = ttk.Combobox(
            options_frame,
            textvariable=self.overwrite_var,
            state="readonly",
            values=["never", "always", "rename"],
            width=14,
        )
        self.overwrite_box.grid(row=1, column=1, sticky=tk.W, pady=(8, 0))

        ttk.Label(options_frame, text="重命名模板").grid(row=2, column=0, sticky=tk.W, pady=(8, 0))
        self.rename_entry = ttk.Entry(options_frame, textvariable=self.rename_pattern_var, width=45)
        self.rename_entry.grid(row=2, column=1, padx=8, sticky=tk.W, pady=(8, 0))

        ttk.Label(options_frame, text="起始序号").grid(row=2, column=2, sticky=tk.E, pady=(8, 0))
        self.start_index_spin = ttk.Spinbox(options_frame, from_=1, to=999999, textvariable=self.start_index_var, width=8)
        self.start_index_spin.grid(row=2, column=3, sticky=tk.W, pady=(8, 0), padx=(8, 0))

        self.delete_trash_radio = ttk.Radiobutton(options_frame, text="删除到回收站", variable=self.use_trash_var, value=True)
        self.delete_trash_radio.grid(row=3, column=0, sticky=tk.W, pady=(8, 0))
        self.delete_hard_radio = ttk.Radiobutton(options_frame, text="永久删除", variable=self.use_trash_var, value=False)
        self.delete_hard_radio.grid(row=3, column=1, sticky=tk.W, pady=(8, 0))

        ttk.Label(options_frame, text="分片大小(MB)").grid(row=4, column=0, sticky=tk.W, pady=(8, 0))
        self.file_split_size_entry = ttk.Entry(options_frame, textvariable=self.file_split_size_var, width=12)
        self.file_split_size_entry.grid(row=4, column=1, sticky=tk.W, pady=(8, 0))

        ttk.Label(options_frame, text="标题拆分规则").grid(row=5, column=0, sticky=tk.W, pady=(8, 0))
        self.doc_mode_box = ttk.Combobox(
            options_frame,
            textvariable=self.doc_mode_label_var,
            state="readonly",
            values=list(self.doc_mode_label_to_value.keys()),
            width=22,
        )
        self.doc_mode_box.grid(row=5, column=1, sticky=tk.W, pady=(8, 0))

        self.include_image_text_check = ttk.Checkbutton(options_frame, text="提取图片文字（OCR）", variable=self.include_image_text_var)
        self.include_image_text_check.grid(row=5, column=2, columnspan=2, sticky=tk.W, pady=(8, 0), padx=(16, 0))

        options_frame.grid_columnconfigure(1, weight=1)

        run_frame = ttk.LabelFrame(page, text="执行", style="Card.TLabelframe", padding=10)
        run_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(run_frame, text="预演模式（不写入）", variable=self.dry_run_var).grid(row=0, column=0, sticky=tk.W)
        ttk.Label(run_frame, text="报告文件").grid(row=0, column=1, padx=(20, 0), sticky=tk.W)
        ttk.Entry(run_frame, textvariable=self.report_path_var, width=58).grid(row=0, column=2, padx=8, sticky=tk.EW)
        ttk.Button(run_frame, text="另存为", command=self._select_report_file).grid(row=0, column=3, padx=(6, 0), sticky=tk.W)

        self.run_button = ttk.Button(run_frame, text="开始执行", style="Primary.TButton", command=self._execute_operation)
        self.run_button.grid(row=0, column=4, padx=(16, 0), sticky=tk.E)

        self.progress = ttk.Progressbar(
            run_frame,
            style="Progress.Horizontal.TProgressbar",
            mode="determinate",
            maximum=100,
            variable=self.progress_var,
        )
        self.progress.grid(row=1, column=0, columnspan=5, sticky=tk.EW, pady=(10, 4))

        ttk.Label(run_frame, textvariable=self.progress_text_var, foreground="#334155").grid(row=2, column=0, columnspan=5, sticky=tk.W)

        run_frame.grid_columnconfigure(2, weight=1)

        log_frame = ttk.LabelFrame(page, text="执行日志", style="Card.TLabelframe", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(
            log_frame,
            height=14,
            wrap=tk.WORD,
            bg="#0f172a",
            fg="#e2e8f0",
            insertbackground="#e2e8f0",
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground="#334155",
            font=("Consolas", 10),
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        status_bar = ttk.Frame(page, style="Page.TFrame")
        status_bar.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(status_bar, textvariable=self.status_var, foreground="#334155").pack(anchor=tk.W)

    def _current_operation(self) -> str:
        return self.operation_label_to_value[self.operation_label_var.get()]

    def _sync_operation_fields(self) -> None:
        op = self._current_operation()

        show_destination = op in {"copy", "move", "split", "doc_split"}
        show_overwrite = op in {"copy", "move", "rename", "split"}
        show_rename = op == "rename"
        show_delete = op == "delete"
        show_file_split = op == "split"
        show_doc_split = op == "doc_split"

        self._set_state(self.destination_entry, show_destination)
        self._set_state(self.destination_button, show_destination)

        self._set_state(self.overwrite_box, show_overwrite, readonly=True)
        self._set_state(self.rename_entry, show_rename)
        self._set_state(self.start_index_spin, show_rename)

        self._set_state(self.delete_trash_radio, show_delete)
        self._set_state(self.delete_hard_radio, show_delete)

        self._set_state(self.file_split_size_entry, show_file_split)

        self._set_state(self.doc_mode_box, show_doc_split, readonly=True)
        self._set_state(self.include_image_text_check, show_doc_split)

    def _set_state(self, widget: ttk.Widget, enabled: bool, readonly: bool = False) -> None:
        if not enabled:
            widget.configure(state="disabled")
            return
        widget.configure(state="readonly" if readonly else "normal")

    def _set_running(self, running: bool) -> None:
        if running:
            self.run_button.configure(state="disabled")
            self.operation_box.configure(state="disabled")
        else:
            self.run_button.configure(state="normal")
            self.operation_box.configure(state="readonly")

    def _append_log(self, text: str) -> None:
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)

    def _thread_log(self, text: str) -> None:
        self.root.after(0, lambda: self._append_log(text))

    def _thread_progress(self, done: int, total: int, detail: str) -> None:
        def apply() -> None:
            percent = 100.0 if total == 0 else (done / total) * 100.0
            self.progress_var.set(percent)
            self.progress_text_var.set(f"进度：{done}/{total}（{percent:.0f}%）  {detail}")

        self.root.after(0, apply)

    def _select_workspace(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.workspace_var.get() or str(Path.cwd()))
        if selected:
            self.workspace_var.set(selected)

    def _select_destination(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.workspace_var.get() or str(Path.cwd()))
        if selected:
            self.destination_var.set(selected)

    def _select_report_file(self) -> None:
        selected = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON 文件", "*.json"), ("全部文件", "*.*")],
            initialdir=self.workspace_var.get() or str(Path.cwd()),
        )
        if selected:
            self.report_path_var.set(selected)

    def _add_files(self) -> None:
        files = filedialog.askopenfilenames(initialdir=self.workspace_var.get() or str(Path.cwd()))
        for item in files:
            self._append_source(item)

    def _add_folder(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.workspace_var.get() or str(Path.cwd()))
        if selected:
            self._append_source(selected)

    def _append_source(self, value: str) -> None:
        existing = set(self.source_list.get(0, tk.END))
        if value not in existing:
            self.source_list.insert(tk.END, value)

    def _remove_selected_sources(self) -> None:
        selected_indexes = list(self.source_list.curselection())
        for idx in reversed(selected_indexes):
            self.source_list.delete(idx)

    def _clear_sources(self) -> None:
        self.source_list.delete(0, tk.END)

    def _execute_operation(self) -> None:
        if self.run_button.instate(["disabled"]):
            return

        try:
            params = self._collect_parameters()
        except ValueError as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        if self._current_operation() == "delete" and not params["dry_run"] and not bool(params.get("use_trash", True)):
            confirmed = messagebox.askyesno("确认永久删除", "你选择了“永久删除”，该操作不可恢复，是否继续？")
            if not confirmed:
                return

        self.progress_var.set(0.0)
        self.progress_text_var.set("进度：0/0（0%）  准备中...")
        self.status_var.set("执行中...")
        self._set_running(True)
        self._append_log("----------------------------------------")
        self._append_log(f"开始执行：{self.operation_label_var.get()}")

        worker = threading.Thread(target=self._run_operation, args=(params,), daemon=True)
        worker.start()

    def _collect_parameters(self) -> dict[str, object]:
        operation = self._current_operation()
        workspace = Path(self.workspace_var.get().strip() or ".").resolve(strict=False)

        source_values = list(self.source_list.get(0, tk.END))
        sources = [Path(item).resolve(strict=False) for item in source_values]
        if not sources:
            raise ValueError("请先添加至少一个源文件或目录。")

        dry_run = bool(self.dry_run_var.get())
        overwrite = self.overwrite_var.get().strip() or "never"

        payload: dict[str, object] = {
            "operation": operation,
            "workspace": workspace,
            "sources": sources,
            "dry_run": dry_run,
            "overwrite": overwrite,
            "report_path": self.report_path_var.get().strip(),
        }

        if operation in {"copy", "move", "split", "doc_split"}:
            destination_text = self.destination_var.get().strip()
            if not destination_text:
                raise ValueError("该操作需要指定输出目录/目标路径。")
            payload["destination"] = Path(destination_text).resolve(strict=False)

        if operation == "rename":
            pattern = self.rename_pattern_var.get().strip()
            if not pattern:
                raise ValueError("请填写重命名模板。")
            payload["pattern"] = pattern
            payload["start_index"] = int(self.start_index_var.get())

        if operation == "delete":
            payload["use_trash"] = bool(self.use_trash_var.get())

        if operation == "split":
            try:
                split_size_mb = float(self.file_split_size_var.get().strip())
            except ValueError as exc:
                raise ValueError("分片大小必须是数字（单位 MB）。") from exc
            if split_size_mb <= 0:
                raise ValueError("分片大小必须大于 0。")
            payload["split_size_mb"] = split_size_mb

        if operation == "doc_split":
            payload["heading_mode"] = self.doc_mode_label_to_value[self.doc_mode_label_var.get()]
            payload["include_image_text"] = bool(self.include_image_text_var.get())

        return payload

    def _run_single(
        self,
        operation: str,
        source: Path,
        params: dict[str, object],
        rename_index: int,
    ) -> list[OperationResult]:
        workspace = Path(params["workspace"])
        dry_run = bool(params["dry_run"])

        if operation in {"copy", "move", "rename", "split"}:
            common = CommonOptions(
                workspace=workspace,
                dry_run=dry_run,
                overwrite=str(params["overwrite"]),
            )

            if operation == "copy":
                return copy_items([source], Path(params["destination"]), common)
            if operation == "move":
                return move_items([source], Path(params["destination"]), common)
            if operation == "rename":
                return rename_items([source], str(params["pattern"]), rename_index, common)
            return split_items([source], Path(params["destination"]), float(params["split_size_mb"]), common)

        if operation == "doc_split":
            return split_documents_by_structure(
                sources=[source],
                destination=Path(params["destination"]),
                workspace=workspace,
                dry_run=dry_run,
                heading_mode=str(params["heading_mode"]),
                include_image_text=bool(params["include_image_text"]),
            )

        return delete_items(
            sources=[source],
            workspace=workspace,
            dry_run=dry_run,
            use_trash=bool(params["use_trash"]),
        )

    def _run_operation(self, params: dict[str, object]) -> None:
        operation = str(params["operation"])
        sources = list(params["sources"])
        report = RunReport(command=operation, dry_run_mode=bool(params["dry_run"]), workspace=str(params["workspace"]))
        total = len(sources)

        status_map = {
            "success": "成功",
            "failed": "失败",
            "skipped": "跳过",
            "dry_run": "预演",
        }

        try:
            for idx, source in enumerate(sources, start=1):
                self._thread_progress(idx - 1, total, f"处理中：{source.name}")
                self._thread_log(f"[{idx}/{total}] 开始处理：{source}")

                results = self._run_single(operation, source, params, rename_index=int(params.get("start_index", 1)) + idx - 1)
                for item in results:
                    report.add(item)
                    op_label = self.operation_value_to_label.get(item.operation, item.operation)
                    status_text = status_map.get(item.status.value, item.status.value)
                    self._thread_log(f"[{status_text}] {op_label} | {item.source} -> {item.destination} | {item.message}")

                self._thread_progress(idx, total, f"已完成：{source.name}")

            report_path_text = str(params["report_path"]).strip()
            output_path = write_report(report, Path(report_path_text).resolve(strict=False)) if report_path_text else None

            summary = report.summary()
            self._thread_log("")
            self._thread_log(
                "汇总: "
                f"总数={summary['total']} 成功={summary['success']} 预演={summary['dry_run']} "
                f"跳过={summary['skipped']} 失败={summary['failed']}"
            )
            if output_path is not None:
                self._thread_log(f"报告输出: {output_path}")

            has_failure = summary["failed"] > 0
            final_status = "执行完成（存在失败）" if has_failure else "执行完成"
            self.root.after(0, lambda: self._finish_run(final_status, has_failure))

        except Exception as exc:  # noqa: BLE001
            self.root.after(0, lambda: self._finish_run(f"执行失败：{exc}", True))

    def _finish_run(self, status: str, is_error: bool) -> None:
        self.status_var.set(status)
        self._set_running(False)

        if is_error:
            messagebox.showerror("执行结果", status)
        else:
            messagebox.showinfo("执行结果", status)


def launch_gui() -> None:
    root = tk.Tk()
    FileOpsGUI(root)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()
