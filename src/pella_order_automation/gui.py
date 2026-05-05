from __future__ import annotations

import queue
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .extractor import build_job
from .workbook_writer import write_workbook


class PellaOrderApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Pella Order Automation")
        self.geometry("980x720")
        self.minsize(860, 640)
        self.configure(bg=COLORS["bg"])

        self.po_path = tk.StringVar()
        self.work_order_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status = tk.StringVar(value="Ready")
        self.status_detail = tk.StringVar(value="Select a purchase order and work order to begin.")
        self.result_queue: queue.Queue[tuple[str, object]] = queue.Queue()

        self._configure_style()
        self._build_ui()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("App.TFrame", background=COLORS["bg"])
        style.configure("Panel.TFrame", background=COLORS["panel"], relief="flat")
        style.configure("Subtle.TFrame", background=COLORS["subtle"])
        style.configure("Title.TLabel", background=COLORS["bg"], foreground=COLORS["ink"], font=("Helvetica", 22, "bold"))
        style.configure("Subtitle.TLabel", background=COLORS["bg"], foreground=COLORS["muted"], font=("Helvetica", 12))
        style.configure("PanelTitle.TLabel", background=COLORS["panel"], foreground=COLORS["ink"], font=("Helvetica", 13, "bold"))
        style.configure("PanelText.TLabel", background=COLORS["panel"], foreground=COLORS["muted"], font=("Helvetica", 10))
        style.configure("Small.TLabel", background=COLORS["panel"], foreground=COLORS["muted"], font=("Helvetica", 9))
        style.configure("Status.TLabel", background=COLORS["panel"], foreground=COLORS["accent"], font=("Helvetica", 12, "bold"))
        style.configure("StatusDetail.TLabel", background=COLORS["panel"], foreground=COLORS["muted"], font=("Helvetica", 10))
        style.configure("CardValue.TLabel", background=COLORS["subtle"], foreground=COLORS["ink"], font=("Helvetica", 14, "bold"))
        style.configure("CardLabel.TLabel", background=COLORS["subtle"], foreground=COLORS["muted"], font=("Helvetica", 9))
        style.configure("TEntry", fieldbackground=COLORS["input"], bordercolor=COLORS["border"], lightcolor=COLORS["border"], darkcolor=COLORS["border"], padding=8)
        style.configure("TButton", font=("Helvetica", 10), padding=(12, 8))
        style.configure("Accent.TButton", font=("Helvetica", 12, "bold"), padding=(18, 11), background=COLORS["accent"], foreground="white")
        style.map(
            "Accent.TButton",
            background=[("active", COLORS["accent_dark"]), ("disabled", COLORS["disabled"])],
            foreground=[("disabled", "#f1f5f9")],
        )
        style.configure("Horizontal.TProgressbar", troughcolor=COLORS["subtle"], background=COLORS["accent"], bordercolor=COLORS["subtle"])

    def _build_ui(self) -> None:
        root = ttk.Frame(self, style="App.TFrame", padding=22)
        root.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(root, style="App.TFrame")
        header.pack(fill=tk.X)
        ttk.Label(header, text="Pella Order Automation", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            header,
            text="Generate completed workbooks from a purchase order CSV and work order XLSX.",
            style="Subtitle.TLabel",
        ).pack(anchor=tk.W, pady=(4, 0))

        body = ttk.Frame(root, style="App.TFrame")
        body.pack(fill=tk.BOTH, expand=True, pady=(18, 0))
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        left = ttk.Frame(body, style="App.TFrame")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        right = ttk.Frame(body, style="App.TFrame")
        right.grid(row=0, column=1, sticky="nsew")

        self._build_input_panel(left)
        self._build_summary_panel(left)
        self._build_status_panel(right)
        self._build_messages_panel(right)

    def _build_input_panel(self, parent: ttk.Frame) -> None:
        panel = self._panel(parent)
        panel.pack(fill=tk.X)
        ttk.Label(panel, text="Files", style="PanelTitle.TLabel").pack(anchor=tk.W)
        ttk.Label(panel, text="Choose the two source files and where to save the completed workbook.", style="PanelText.TLabel").pack(anchor=tk.W, pady=(2, 12))

        self._file_row(panel, "Purchase Order", "CSV source file", self.po_path, self._choose_po)
        self._file_row(panel, "Work Order", "XLSX source file", self.work_order_path, self._choose_work_order)
        self._file_row(panel, "Completed Workbook", "Output XLSX file", self.output_path, self._choose_output)

        actions = ttk.Frame(panel, style="Panel.TFrame")
        actions.pack(fill=tk.X, pady=(16, 0))
        self.generate_button = ttk.Button(actions, text="Generate Workbook", style="Accent.TButton", command=self._generate)
        self.generate_button.pack(side=tk.LEFT)
        ttk.Button(actions, text="Clear", command=self._clear).pack(side=tk.LEFT, padx=(10, 0))

    def _build_summary_panel(self, parent: ttk.Frame) -> None:
        panel = self._panel(parent)
        panel.pack(fill=tk.BOTH, expand=True, pady=(14, 0))
        ttk.Label(panel, text="Generation Summary", style="PanelTitle.TLabel").pack(anchor=tk.W)

        cards = ttk.Frame(panel, style="Panel.TFrame")
        cards.pack(fill=tk.X, pady=(12, 10))
        cards.columnconfigure((0, 1, 2), weight=1)
        self.line_count_value = tk.StringVar(value="-")
        self.row_count_value = tk.StringVar(value="-")
        self.warning_count_value = tk.StringVar(value="-")
        self._metric_card(cards, "Jamb lines", self.line_count_value, 0)
        self._metric_card(cards, "BOM rows", self.row_count_value, 1)
        self._metric_card(cards, "Warnings", self.warning_count_value, 2)

        self.summary_box = self._text_box(panel, "Summary information will appear here after generation.", height=11)

    def _build_status_panel(self, parent: ttk.Frame) -> None:
        panel = self._panel(parent)
        panel.pack(fill=tk.X)
        ttk.Label(panel, text="Status", style="PanelTitle.TLabel").pack(anchor=tk.W)
        ttk.Label(panel, textvariable=self.status, style="Status.TLabel").pack(anchor=tk.W, pady=(12, 0))
        ttk.Label(panel, textvariable=self.status_detail, style="StatusDetail.TLabel", wraplength=300).pack(anchor=tk.W, pady=(4, 8))
        self.progress = ttk.Progressbar(panel, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=(8, 0))

    def _build_messages_panel(self, parent: ttk.Frame) -> None:
        panel = self._panel(parent)
        panel.pack(fill=tk.BOTH, expand=True, pady=(14, 0))

        notebook = ttk.Notebook(panel)
        notebook.pack(fill=tk.BOTH, expand=True)

        warning_frame = ttk.Frame(notebook, style="Panel.TFrame", padding=10)
        error_frame = ttk.Frame(notebook, style="Panel.TFrame", padding=10)
        notebook.add(warning_frame, text="Warnings")
        notebook.add(error_frame, text="Errors")

        self.warning_box = self._text_box(warning_frame, "Warnings will appear here after generation.", height=16)
        self.error_box = self._text_box(error_frame, "Errors will appear here if generation fails.", height=16)

    def _panel(self, parent: ttk.Frame) -> ttk.Frame:
        panel = ttk.Frame(parent, style="Panel.TFrame", padding=18)
        return panel

    def _file_row(self, parent: ttk.Frame, title: str, subtitle: str, variable: tk.StringVar, command) -> None:
        row = ttk.Frame(parent, style="Subtle.TFrame", padding=12)
        row.pack(fill=tk.X, pady=6)
        row.columnconfigure(1, weight=1)

        labels = ttk.Frame(row, style="Subtle.TFrame")
        labels.grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Label(labels, text=title, background=COLORS["subtle"], foreground=COLORS["ink"], font=("Helvetica", 11, "bold")).pack(anchor=tk.W)
        ttk.Label(labels, text=subtitle, background=COLORS["subtle"], foreground=COLORS["muted"], font=("Helvetica", 9)).pack(anchor=tk.W)

        entry = ttk.Entry(row, textvariable=variable)
        entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        ttk.Button(row, text="Choose", command=command).grid(row=0, column=2, sticky="e")

    def _metric_card(self, parent: ttk.Frame, label: str, value: tk.StringVar, column: int) -> None:
        card = ttk.Frame(parent, style="Subtle.TFrame", padding=12)
        card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
        ttk.Label(card, textvariable=value, style="CardValue.TLabel").pack(anchor=tk.W)
        ttk.Label(card, text=label, style="CardLabel.TLabel").pack(anchor=tk.W, pady=(2, 0))

    def _text_box(self, parent: ttk.Frame, initial: str, height: int) -> tk.Text:
        frame = ttk.Frame(parent, style="Panel.TFrame")
        frame.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text = tk.Text(
            frame,
            height=height,
            wrap=tk.WORD,
            relief=tk.FLAT,
            borderwidth=0,
            bg=COLORS["input"],
            fg=COLORS["ink"],
            insertbackground=COLORS["ink"],
            selectbackground="#bfdbfe",
            font=("Helvetica", 10),
            padx=12,
            pady=10,
            yscrollcommand=scrollbar.set,
        )
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.configure(command=text.yview)
        text.insert("1.0", initial)
        text.configure(state=tk.DISABLED)
        return text

    def _choose_po(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose Purchase Order CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self.po_path.set(path)
            self._suggest_output()
            self._ready_detail()

    def _choose_work_order(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose Work Order XLSX",
            filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.work_order_path.set(path)
            self._suggest_output()
            self._ready_detail()

    def _choose_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save Completed Workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)
            self._ready_detail()

    def _suggest_output(self) -> None:
        if self.output_path.get() or not self.work_order_path.get():
            return
        work_order = Path(self.work_order_path.get())
        self.output_path.set(str(work_order.with_name(f"{work_order.stem} - Completed.xlsx")))

    def _ready_detail(self) -> None:
        missing = []
        if not self.po_path.get():
            missing.append("purchase order")
        if not self.work_order_path.get():
            missing.append("work order")
        if not self.output_path.get():
            missing.append("output path")
        if missing:
            self.status.set("Ready")
            self.status_detail.set(f"Select {', '.join(missing)}.")
        else:
            self.status.set("Ready to generate")
            self.status_detail.set("All files selected. Click Generate Workbook.")

    def _clear(self) -> None:
        self.po_path.set("")
        self.work_order_path.set("")
        self.output_path.set("")
        self.status.set("Ready")
        self.status_detail.set("Select a purchase order and work order to begin.")
        self.line_count_value.set("-")
        self.row_count_value.set("-")
        self.warning_count_value.set("-")
        self._set_summary("Summary information will appear here after generation.")
        self._set_warnings("Warnings will appear here after generation.")
        self._set_errors("Errors will appear here if generation fails.")

    def _generate(self) -> None:
        po_path = self.po_path.get()
        work_order_path = self.work_order_path.get()
        output_path = self.output_path.get()

        if not po_path or not work_order_path or not output_path:
            messagebox.showwarning("Missing files", "Choose a PO, work order, and output path first.")
            return

        self.generate_button.configure(state=tk.DISABLED)
        self.status.set("Generating")
        self.status_detail.set("Parsing files and building the completed workbook.")
        self.progress.start(12)
        self.line_count_value.set("-")
        self.row_count_value.set("-")
        self.warning_count_value.set("-")
        self._set_summary("Working...")
        self._set_warnings("Working...")
        self._set_errors("No errors.")

        thread = threading.Thread(
            target=self._generate_worker,
            args=(po_path, work_order_path, output_path),
            daemon=True,
        )
        thread.start()
        self.after(100, self._poll_result_queue)

    def _generate_worker(self, po_path: str, work_order_path: str, output_path: str) -> None:
        try:
            job = build_job(po_path, work_order_path)
            output = write_workbook(job, output_path)
            warnings = self._dedupe_warnings(
                [*job.warnings, *(warning for line in job.jamb_lines for warning in line.warnings)]
            )
            summary_text = self._format_summary(job, output, len(warnings))
            warning_text = self._format_warnings(warnings)
            self.result_queue.put(("success", (output, summary_text, warning_text, len(job.jamb_lines), len(warnings))))
        except Exception as exc:  # noqa: BLE001 - user-facing GUI boundary
            self.result_queue.put(("error", (exc, traceback.format_exc())))

    def _poll_result_queue(self) -> None:
        try:
            status, payload = self.result_queue.get_nowait()
        except queue.Empty:
            self.after(100, self._poll_result_queue)
            return

        if status == "success":
            output, summary_text, warning_text, line_count, warning_count = payload
            self._generation_done(output, summary_text, warning_text, line_count, warning_count)
        else:
            exc, details = payload
            self._generation_failed(exc, details)

    def _generation_done(self, output: Path, summary: str, warnings: str, line_count: int, warning_count: int) -> None:
        self.progress.stop()
        self.generate_button.configure(state=tk.NORMAL)
        self.status.set("Completed")
        self.status_detail.set(f"Generated {output.name}")
        self.line_count_value.set(str(line_count))
        self.row_count_value.set(str(line_count * 4))
        self.warning_count_value.set(str(warning_count))
        self._set_summary(summary)
        self._set_warnings(warnings)
        self._set_errors("No errors.")

    def _generation_failed(self, exc: Exception, details: str) -> None:
        self.progress.stop()
        self.generate_button.configure(state=tk.NORMAL)
        self.status.set("Failed")
        self.status_detail.set("Generation failed. Review the Errors tab for details.")
        self.line_count_value.set("-")
        self.row_count_value.set("-")
        self.warning_count_value.set("-")
        self._set_summary("Generation failed.")
        self._set_warnings("Generation failed before warnings could be calculated.")
        self._set_errors(f"{exc}\n\nDetails:\n{details}")
        messagebox.showerror("Generation failed", str(exc))

    def _set_summary(self, text: str) -> None:
        self._set_text(self.summary_box, text)

    def _set_warnings(self, text: str) -> None:
        self._set_text(self.warning_box, text)

    def _set_errors(self, text: str) -> None:
        self._set_text(self.error_box, text)

    def _set_text(self, widget: tk.Text, text: str) -> None:
        widget.configure(state=tk.NORMAL)
        widget.delete("1.0", tk.END)
        widget.insert("1.0", text)
        widget.configure(state=tk.DISABLED)

    def _format_summary(self, job, output: Path, warning_count: int) -> str:
        request_dates = sorted(
            {
                line.po_line.requested_date.isoformat()
                for line in job.jamb_lines
                if line.po_line.requested_date
            }
        )
        return "\n".join(
            [
                f"Output: {output}",
                f"Purchase order: {Path(job.po_path).name}",
                f"Work order: {Path(job.work_order_path).name}",
                "",
                f"Order number: {job.order_number or 'Unknown'}",
                f"Pella PO number: {job.pella_po_number or 'Unknown'}",
                f"Customer: {job.customer_name or 'Unknown'}",
                f"Request date(s): {', '.join(request_dates) if request_dates else 'Blank'}",
                "",
                f"Jamb input lines: {len(job.jamb_lines)}",
                f"Generated BOM rows: {len(job.jamb_lines) * 4}",
                f"Warnings: {warning_count}",
            ]
        )

    def _format_warnings(self, warnings) -> str:
        if not warnings:
            return "No warnings."
        return "\n\n".join(f"[{warning.code}]\n{warning.message}" for warning in warnings)

    def _dedupe_warnings(self, warnings):
        unique = {}
        for warning in warnings:
            unique[(warning.code, warning.message)] = warning
        return list(unique.values())


COLORS = {
    "bg": "#eef3f7",
    "panel": "#ffffff",
    "subtle": "#f6f8fb",
    "input": "#fbfdff",
    "ink": "#1f2937",
    "muted": "#64748b",
    "border": "#d8e2ea",
    "accent": "#2563eb",
    "accent_dark": "#1d4ed8",
    "disabled": "#94a3b8",
}


def main() -> None:
    app = PellaOrderApp()
    app.mainloop()


if __name__ == "__main__":
    main()
