from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .models import RunReport
from .operations import CommonOptions, copy_items, delete_items, move_items, rename_items
from .reporting import write_report


class FileOpsGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("FileOps")
        self.root.geometry("980x700")
        self.root.minsize(920, 640)

        self.operation_var = tk.StringVar(value="copy")
        self.workspace_var = tk.StringVar(value=str(Path.cwd()))
        self.destination_var = tk.StringVar(value="")
        self.rename_pattern_var = tk.StringVar(value="{stem}_{index}{ext}")
        self.start_index_var = tk.IntVar(value=1)
        self.overwrite_var = tk.StringVar(value="never")
        self.report_path_var = tk.StringVar(value="")
        self.dry_run_var = tk.BooleanVar(value=False)
        self.use_trash_var = tk.BooleanVar(value=True)

        self._build_ui()
        self._sync_operation_fields()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Operation").grid(row=0, column=0, sticky=tk.W)
        operation_box = ttk.Combobox(
            top,
            textvariable=self.operation_var,
            state="readonly",
            values=["copy", "move", "rename", "delete"],
            width=16,
        )
        operation_box.grid(row=0, column=1, padx=(8, 16), sticky=tk.W)
        operation_box.bind("<<ComboboxSelected>>", lambda _event: self._sync_operation_fields())

        ttk.Label(top, text="Workspace").grid(row=0, column=2, sticky=tk.W)
        ttk.Entry(top, textvariable=self.workspace_var, width=58).grid(row=0, column=3, padx=8, sticky=tk.EW)
        ttk.Button(top, text="Browse", command=self._select_workspace).grid(row=0, column=4, padx=(6, 0), sticky=tk.W)

        top.grid_columnconfigure(3, weight=1)

        source_frame = ttk.LabelFrame(self.root, text="Sources", padding=10)
        source_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=(0, 10))

        self.source_list = tk.Listbox(source_frame, height=8, selectmode=tk.EXTENDED)
        self.source_list.grid(row=0, column=0, rowspan=4, sticky=tk.NSEW)

        scrollbar = ttk.Scrollbar(source_frame, orient=tk.VERTICAL, command=self.source_list.yview)
        scrollbar.grid(row=0, column=1, rowspan=4, sticky=tk.NS)
        self.source_list.configure(yscrollcommand=scrollbar.set)

        ttk.Button(source_frame, text="Add Files", command=self._add_files).grid(row=0, column=2, padx=(10, 0), pady=(0, 6), sticky=tk.EW)
        ttk.Button(source_frame, text="Add Folder", command=self._add_folder).grid(row=1, column=2, padx=(10, 0), pady=6, sticky=tk.EW)
        ttk.Button(source_frame, text="Remove Selected", command=self._remove_selected_sources).grid(row=2, column=2, padx=(10, 0), pady=6, sticky=tk.EW)
        ttk.Button(source_frame, text="Clear", command=self._clear_sources).grid(row=3, column=2, padx=(10, 0), pady=(6, 0), sticky=tk.EW)

        source_frame.grid_columnconfigure(0, weight=1)
        source_frame.grid_rowconfigure(0, weight=1)

        options_frame = ttk.LabelFrame(self.root, text="Operation Options", padding=10)
        options_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Label(options_frame, text="Destination").grid(row=0, column=0, sticky=tk.W)
        self.destination_entry = ttk.Entry(options_frame, textvariable=self.destination_var, width=70)
        self.destination_entry.grid(row=0, column=1, padx=8, sticky=tk.EW)
        self.destination_button = ttk.Button(options_frame, text="Browse", command=self._select_destination)
        self.destination_button.grid(row=0, column=2, padx=(6, 0), sticky=tk.W)

        ttk.Label(options_frame, text="Overwrite").grid(row=1, column=0, sticky=tk.W, pady=(8, 0))
        self.overwrite_box = ttk.Combobox(
            options_frame,
            textvariable=self.overwrite_var,
            state="readonly",
            values=["never", "always", "rename"],
            width=12,
        )
        self.overwrite_box.grid(row=1, column=1, sticky=tk.W, pady=(8, 0))

        ttk.Label(options_frame, text="Rename Pattern").grid(row=2, column=0, sticky=tk.W, pady=(8, 0))
        self.rename_entry = ttk.Entry(options_frame, textvariable=self.rename_pattern_var, width=45)
        self.rename_entry.grid(row=2, column=1, padx=8, sticky=tk.W, pady=(8, 0))

        ttk.Label(options_frame, text="Start Index").grid(row=2, column=2, sticky=tk.E, pady=(8, 0))
        self.start_index_spin = ttk.Spinbox(options_frame, from_=1, to=999999, textvariable=self.start_index_var, width=8)
        self.start_index_spin.grid(row=2, column=3, sticky=tk.W, pady=(8, 0), padx=(8, 0))

        self.delete_trash_radio = ttk.Radiobutton(options_frame, text="Delete to Trash", variable=self.use_trash_var, value=True)
        self.delete_trash_radio.grid(row=3, column=0, sticky=tk.W, pady=(8, 0))
        self.delete_hard_radio = ttk.Radiobutton(options_frame, text="Hard Delete", variable=self.use_trash_var, value=False)
        self.delete_hard_radio.grid(row=3, column=1, sticky=tk.W, pady=(8, 0))

        options_frame.grid_columnconfigure(1, weight=1)

        run_frame = ttk.LabelFrame(self.root, text="Run", padding=10)
        run_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Checkbutton(run_frame, text="Dry Run", variable=self.dry_run_var).grid(row=0, column=0, sticky=tk.W)
        ttk.Label(run_frame, text="Report File").grid(row=0, column=1, padx=(20, 0), sticky=tk.W)
        ttk.Entry(run_frame, textvariable=self.report_path_var, width=58).grid(row=0, column=2, padx=8, sticky=tk.EW)
        ttk.Button(run_frame, text="Save As", command=self._select_report_file).grid(row=0, column=3, padx=(6, 0), sticky=tk.W)

        self.run_button = ttk.Button(run_frame, text="Execute", command=self._execute_operation)
        self.run_button.grid(row=0, column=4, padx=(16, 0), sticky=tk.E)

        run_frame.grid_columnconfigure(2, weight=1)

        log_frame = ttk.LabelFrame(self.root, text="Result Log", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.log_text = tk.Text(log_frame, height=14, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self.root, textvariable=self.status_var, anchor=tk.W).pack(fill=tk.X, padx=10, pady=(0, 10))

    def _sync_operation_fields(self) -> None:
        op = self.operation_var.get()

        show_destination = op in {"copy", "move"}
        show_overwrite = op in {"copy", "move", "rename"}
        show_rename = op == "rename"
        show_delete = op == "delete"

        self._set_widget_state(self.destination_entry, show_destination)
        self._set_widget_state(self.destination_button, show_destination)

        self._set_widget_state(self.overwrite_box, show_overwrite)
        self._set_widget_state(self.rename_entry, show_rename)
        self._set_widget_state(self.start_index_spin, show_rename)

        self._set_widget_state(self.delete_trash_radio, show_delete)
        self._set_widget_state(self.delete_hard_radio, show_delete)

    def _set_widget_state(self, widget: ttk.Widget, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        widget.configure(state=state)

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
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
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
            messagebox.showerror("Invalid Input", str(exc))
            return

        if not params["dry_run"]:
            confirmed = messagebox.askyesno("Confirm", f"Execute '{params['operation']}' on {len(params['sources'])} item(s)?")
            if not confirmed:
                return

        self.run_button.configure(state="disabled")
        self.status_var.set("Running...")

        worker = threading.Thread(target=self._run_operation, args=(params,), daemon=True)
        worker.start()

    def _collect_parameters(self) -> dict[str, object]:
        operation = self.operation_var.get()
        workspace = Path(self.workspace_var.get().strip() or ".").resolve(strict=False)

        source_values = list(self.source_list.get(0, tk.END))
        sources = [Path(item).resolve(strict=False) for item in source_values]

        if not sources:
            raise ValueError("Please add at least one source path.")

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

        if operation in {"copy", "move"}:
            dest = self.destination_var.get().strip()
            if not dest:
                raise ValueError("Destination is required for copy and move.")
            payload["destination"] = Path(dest).resolve(strict=False)

        if operation == "rename":
            pattern = self.rename_pattern_var.get().strip()
            if not pattern:
                raise ValueError("Rename pattern is required.")
            payload["pattern"] = pattern
            payload["start_index"] = int(self.start_index_var.get())

        if operation == "delete":
            payload["use_trash"] = bool(self.use_trash_var.get())

        return payload

    def _run_operation(self, params: dict[str, object]) -> None:
        operation = params["operation"]
        workspace = params["workspace"]
        sources = params["sources"]
        dry_run = params["dry_run"]

        report = RunReport(command=str(operation), dry_run_mode=bool(dry_run), workspace=str(workspace))

        try:
            if operation in {"copy", "move", "rename"}:
                common = CommonOptions(
                    workspace=workspace,
                    dry_run=bool(dry_run),
                    overwrite=str(params["overwrite"]),
                )

                if operation == "copy":
                    results = copy_items(sources, params["destination"], common)
                elif operation == "move":
                    results = move_items(sources, params["destination"], common)
                else:
                    results = rename_items(
                        sources=sources,
                        pattern=str(params["pattern"]),
                        start_index=int(params["start_index"]),
                        options=common,
                    )
            else:
                results = delete_items(
                    sources=sources,
                    workspace=workspace,
                    dry_run=bool(dry_run),
                    use_trash=bool(params["use_trash"]),
                )

            for item in results:
                report.add(item)

            report_path_text = str(params["report_path"]).strip()
            output_path = write_report(report, Path(report_path_text).resolve(strict=False)) if report_path_text else None

            summary = report.summary()
            log_lines = []
            for item in report.results:
                log_lines.append(f"[{item.status.value.upper()}] {item.operation}: {item.source} -> {item.destination} | {item.message}")
            log_lines.append("")
            log_lines.append(
                "Summary: "
                f"total={summary['total']} success={summary['success']} dry_run={summary['dry_run']} "
                f"skipped={summary['skipped']} failed={summary['failed']}"
            )
            if output_path is not None:
                log_lines.append(f"Report: {output_path}")

            has_failure = summary["failed"] > 0
            status = "Finished with errors" if has_failure else "Completed"
            self.root.after(0, lambda: self._finish_run("\n".join(log_lines), status, has_failure))

        except Exception as exc:  # noqa: BLE001
            self.root.after(0, lambda: self._finish_run(f"Error: {exc}", "Failed", True))

    def _finish_run(self, text: str, status: str, is_error: bool) -> None:
        self.log_text.insert(tk.END, text + "\n\n")
        self.log_text.see(tk.END)
        self.status_var.set(status)
        self.run_button.configure(state="normal")

        if is_error:
            messagebox.showerror("Execution Result", status)
        else:
            messagebox.showinfo("Execution Result", status)


def launch_gui() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    FileOpsGUI(root)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()
