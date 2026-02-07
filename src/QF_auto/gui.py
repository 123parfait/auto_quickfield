from __future__ import annotations

import argparse
import csv
import os
import queue
import threading
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .connection import pythoncom, win32com, dispatch_qf_app, ensure_model_loaded, open_problem
from .geometry import list_block_labels
from .workflow import (
    _cases_from_mapping_all,
    _cases_from_mapping_pair,
    _parse_label_values,
    _positions_line,
    run_batch_force_plan,
)


def _parse_xy(text: str) -> tuple[float, float]:
    raw = text.strip()
    raw = raw.replace("\uff0c", ",").replace(";", ",")
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    if len(parts) != 2:
        raise ValueError("Expected x,y")
    return float(parts[0]), float(parts[1])


def launch_gui() -> int:
    root = tk.Tk()
    root.title("QF Auto")
    root.geometry("720x600+100+80")

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    BG = "#FFFFFF"
    PANEL = "#FFFFFF"
    TEXT = "#1B2A3A"
    BORDER = "#E1E1E1"
    ACCENT = "#BFC5CC"

    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=PANEL, foreground=TEXT)
    style.configure("TLabelframe", background=PANEL)
    style.configure("TLabelframe.Label", background=PANEL, foreground=TEXT)
    style.configure("App.TFrame", background=BG)
    style.configure("Card.TFrame", background=PANEL)
    style.configure("Card.TLabelframe", background=PANEL, bordercolor=BORDER, relief="solid")
    style.configure("Card.TLabelframe.Label", background=PANEL, foreground=TEXT, font=("Segoe UI", 9, "bold"))
    style.configure("App.TLabel", background=BG, foreground=TEXT)
    style.configure("Card.TLabel", background=PANEL, foreground=TEXT)
    style.configure("Gray.TButton", padding=(10, 5), foreground=TEXT, background=ACCENT)
    style.configure("App.TRadiobutton", background=BG, foreground=TEXT)
    style.configure("Card.TRadiobutton", background=PANEL, foreground=TEXT)
    style.configure("App.TEntry", fieldbackground=PANEL, foreground=TEXT, font=("Segoe UI", 9))
    style.configure("App.TCombobox", fieldbackground=PANEL, foreground=TEXT, font=("Segoe UI", 9))
    style.map("Gray.TButton", background=[("active", "#B3B8BE"), ("disabled", "#E4E6E9")])
    style.map(
        "App.TCombobox",
        fieldbackground=[("readonly", PANEL)],
        background=[("readonly", PANEL)],
        foreground=[("readonly", TEXT)],
    )

    root.configure(background=BG)

    main = ttk.Frame(root, padding=(12, 10), style="App.TFrame")
    main.pack(fill="both", expand=True)
    main.columnconfigure(0, weight=1)
    main.rowconfigure(0, weight=0)
    main.rowconfigure(3, weight=1)

    INTEGRAL_CHOICES = [
        ("MaxwellForce", 15),
        ("MaxwellTorque", 16),
        ("MagneticCoenergy", 14),
        ("MagneticEnergy", 13),
        ("FluxLinkage", 17),
        ("LorentzForce", 23),
    ]

    # Variables
    out_var = tk.StringVar(value="outputs\\force_table.csv")
    mode_var = tk.StringVar(value="pair")
    start_var = tk.StringVar(value="-1,0")
    end_var = tk.StringVar(value="1,0")
    step_var = tk.StringVar(value="1")
    status_var = tk.StringVar(value="Idle")

    # Content split
    content = ttk.Frame(main, style="App.TFrame")
    content.grid(row=0, column=0, sticky="nsew", pady=6)
    content.columnconfigure(0, weight=1)
    content.columnconfigure(1, weight=1)

    # Values
    values = ttk.LabelFrame(content, text="Values", padding=10, style="Card.TLabelframe")
    values.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
    values.columnconfigure(0, weight=1)
    values.rowconfigure(5, weight=0)

    top_row = ttk.Frame(values, style="Card.TFrame")
    top_row.grid(row=0, column=0, sticky="ew")
    ttk.Button(
        top_row,
        text="Load labels",
        command=lambda: _load_label_choices(labels_combo, value_label_combo),
        style="Gray.TButton",
    ).pack(side="left")
    case_row = ttk.Frame(top_row, style="Card.TFrame")
    case_row.pack(side="right")
    ttk.Label(case_row, text="Case mode", style="Card.TLabel").pack(side="left", padx=(0, 6))
    ttk.Radiobutton(case_row, text="all", value="all", variable=mode_var, style="Card.TRadiobutton").pack(side="left")
    ttk.Radiobutton(case_row, text="pair", value="pair", variable=mode_var, style="Card.TRadiobutton").pack(
        side="left", padx=8
    )

    ttk.Label(values, text="Input", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=(6, 0))
    values_row = ttk.Frame(values, style="Card.TFrame")
    values_row.grid(row=3, column=0, sticky="ew", pady=(0, 6))
    values_row.columnconfigure(1, weight=0)

    value_label_combo = ttk.Combobox(values_row, state="readonly", values=[], style="App.TCombobox", width=9)
    value_label_combo.grid(row=0, column=0, sticky="ew", padx=(0, 6))
    value_entry = ttk.Entry(values_row, style="App.TEntry", width=8)
    value_entry.grid(row=0, column=1, sticky="w")
    value_entry.bind(
        "<Return>", lambda _e: _add_value_pair(value_label_combo, value_entry, values_list)
    )

    values_btn_row = ttk.Frame(values, style="App.TFrame")
    values_btn_row.grid(row=4, column=0, sticky="ew", pady=4)
    ttk.Button(
        values_btn_row,
        text="Confirm",
        command=lambda: _add_value_pair(value_label_combo, value_entry, values_list),
        style="Gray.TButton",
    ).pack(side="left", padx=6)
    ttk.Button(
        values_btn_row,
        text="Clear selected",
        command=lambda: _clear_selected_multi(values_list, labels_list, outputs_list),
        style="Gray.TButton",
    ).pack(side="left")

    values_list_frame = ttk.Frame(values, style="Card.TFrame", padding=(0, 0, 0, 8))
    values_list_frame.grid(row=5, column=0, sticky="nsew")
    values_list = tk.Listbox(
        values_list_frame,
        height=4,
        width=26,
        selectmode="extended",
        exportselection=False,
        bg=PANEL,
        fg=TEXT,
        highlightthickness=1,
        highlightbackground=BORDER,
        highlightcolor=ACCENT,
        relief="solid",
        borderwidth=2,
        selectbackground="#EFF1F3",
        selectforeground=TEXT,
    )
    values_list.pack(fill="both", expand=True)

    # Labels + outputs
    movement = ttk.LabelFrame(content, text="Select", padding=10, style="Card.TLabelframe")
    movement.grid(row=0, column=1, sticky="nsew")
    movement.columnconfigure(0, weight=1)
    movement.rowconfigure(1, weight=0)
    movement.rowconfigure(4, weight=0)

    ttk.Label(movement, text="Move labels", style="Card.TLabel").grid(row=0, column=0, sticky="w")
    labels_row = ttk.Frame(movement, style="Card.TFrame")
    labels_row.grid(row=1, column=0, sticky="nsew")
    labels_row.columnconfigure(1, weight=1)

    labels_combo = ttk.Combobox(labels_row, state="readonly", values=[], style="App.TCombobox", width=10)
    labels_combo.grid(row=0, column=0, sticky="ew", padx=(0, 6))
    labels_combo.bind("<<ComboboxSelected>>", lambda _e: _add_choice(labels_combo, labels_list))
    labels_list_frame = ttk.Frame(labels_row, style="Card.TFrame", padding=(0, 0, 0, 8))
    labels_list_frame.grid(row=0, column=1, sticky="nsew")
    labels_list_frame.columnconfigure(0, weight=1)
    labels_list_frame.rowconfigure(0, weight=1)
    labels_list = tk.Listbox(
        labels_list_frame,
        height=4,
        width=26,
        selectmode="extended",
        exportselection=False,
        bg=PANEL,
        fg=TEXT,
        highlightthickness=1,
        highlightbackground=BORDER,
        highlightcolor=ACCENT,
        relief="solid",
        borderwidth=2,
        selectbackground="#EFF1F3",
        selectforeground=TEXT,
    )
    labels_list.grid(row=0, column=0, sticky="nsew")
    labels_scroll = ttk.Scrollbar(labels_list_frame, orient="vertical", command=labels_list.yview)
    labels_scroll.grid(row=0, column=1, sticky="ns")
    labels_list.configure(yscrollcommand=labels_scroll.set)

    ttk.Label(movement, text="Outputs", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=(6, 0))
    out_row = ttk.Frame(movement, style="Card.TFrame")
    out_row.grid(row=4, column=0, sticky="nsew")
    out_row.columnconfigure(1, weight=1)

    outputs_combo = ttk.Combobox(
        out_row, state="readonly", values=[n for n, _ in INTEGRAL_CHOICES], style="App.TCombobox", width=10
    )
    outputs_combo.grid(row=0, column=0, sticky="ew", padx=(0, 6))
    outputs_combo.bind("<<ComboboxSelected>>", lambda _e: _add_choice(outputs_combo, outputs_list))
    outputs_list_frame = ttk.Frame(out_row, style="Card.TFrame", padding=(0, 0, 0, 8))
    outputs_list_frame.grid(row=0, column=1, sticky="nsew")
    outputs_list_frame.columnconfigure(0, weight=1)
    outputs_list_frame.rowconfigure(0, weight=1)
    outputs_list = tk.Listbox(
        outputs_list_frame,
        height=4,
        width=26,
        selectmode="extended",
        exportselection=False,
        bg=PANEL,
        fg=TEXT,
        highlightthickness=1,
        highlightbackground=BORDER,
        highlightcolor=ACCENT,
        relief="solid",
        borderwidth=2,
        selectbackground="#EFF1F3",
        selectforeground=TEXT,
    )
    outputs_list.grid(row=0, column=0, sticky="nsew")
    out_scroll = ttk.Scrollbar(outputs_list_frame, orient="vertical", command=outputs_list.yview)
    out_scroll.grid(row=0, column=1, sticky="ns")
    outputs_list.configure(yscrollcommand=out_scroll.set)

    out_btn_row = ttk.Frame(movement, style="Card.TFrame")
    out_btn_row.grid(row=5, column=0, sticky="ew", pady=4)
    ttk.Label(out_btn_row, text=" ", style="Card.TLabel").pack(side="left")

    display_to_integral: dict[str, tuple[str, int]] = {name: (name, int_id) for name, int_id in INTEGRAL_CHOICES}
    if INTEGRAL_CHOICES:
        outputs_combo.current(0)
        outputs_list.insert("end", INTEGRAL_CHOICES[0][0])

    # Motion
    motion = ttk.LabelFrame(main, text="Motion", padding=10, style="Card.TLabelframe")
    motion.grid(row=1, column=0, sticky="ew")
    motion.columnconfigure(0, weight=1, uniform="motion")
    motion.columnconfigure(1, weight=1, uniform="motion")
    motion.columnconfigure(2, weight=1, uniform="motion")

    start_frame = ttk.Frame(motion, style="Card.TFrame")
    start_frame.grid(row=0, column=0, sticky="ew")
    start_frame.columnconfigure(1, weight=1)
    ttk.Label(start_frame, text="Start x0,y0", style="Card.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Entry(start_frame, textvariable=start_var, width=8, style="App.TEntry").grid(
        row=0, column=1, sticky="ew", padx=(6, 0)
    )

    end_frame = ttk.Frame(motion, style="Card.TFrame")
    end_frame.grid(row=0, column=1, sticky="ew", padx=(6, 6))
    end_frame.columnconfigure(1, weight=1)
    ttk.Label(end_frame, text="End x1,y1", style="Card.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Entry(end_frame, textvariable=end_var, width=8, style="App.TEntry").grid(
        row=0, column=1, sticky="ew", padx=(6, 0)
    )

    step_frame = ttk.Frame(motion, style="Card.TFrame")
    step_frame.grid(row=0, column=2, sticky="ew")
    step_frame.columnconfigure(1, weight=1)
    ttk.Label(step_frame, text="Step", style="Card.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Entry(step_frame, textvariable=step_var, width=5, style="App.TEntry").grid(
        row=0, column=1, sticky="ew", padx=(6, 0)
    )

    # Output
    output = ttk.LabelFrame(main, text="Output", padding=10, style="Card.TLabelframe")
    output.grid(row=2, column=0, sticky="ew")
    output.columnconfigure(1, weight=1)

    ttk.Label(output, text="CSV", style="Card.TLabel").grid(row=0, column=0, sticky="w")
    out_entry = ttk.Entry(output, textvariable=out_var, style="App.TEntry", width=36)
    out_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6))
    ttk.Button(
        output,
        text="Browse...",
        command=lambda: out_var.set(
            filedialog.asksaveasfilename(
                title="Save CSV",
                defaultextension=".csv",
                filetypes=[("CSV", "*.csv"), ("All files", "*.*")],
            )
        ),
        style="Gray.TButton",
    ).grid(row=0, column=2, sticky="e")

    # Log
    log_frame = ttk.LabelFrame(main, text="Log", padding=10, style="Card.TLabelframe")
    log_frame.grid(row=3, column=0, sticky="nsew", pady=(6, 0))
    log_frame.rowconfigure(0, weight=1)
    log_frame.columnconfigure(0, weight=1)
    log_text = tk.Text(
        log_frame,
        height=9,
        width=70,
        wrap="word",
        state="disabled",
        bg=PANEL,
        fg=TEXT,
        highlightthickness=1,
        highlightbackground=BORDER,
        highlightcolor=ACCENT,
        relief="solid",
        borderwidth=2,
    )
    log_text.grid(row=0, column=0, sticky="nsew")
    log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
    log_scroll.grid(row=0, column=1, sticky="ns")
    log_text.configure(yscrollcommand=log_scroll.set)

    # Footer
    footer = ttk.Frame(main, style="App.TFrame")
    footer.grid(row=4, column=0, sticky="ew", pady=(6, 0))
    footer.columnconfigure(0, weight=1)
    status = ttk.Label(footer, textvariable=status_var, style="App.TLabel")
    status.grid(row=0, column=0, sticky="w")

    run_btn = ttk.Button(footer, text="Run", style="Gray.TButton")
    stop_btn = ttk.Button(footer, text="Stop", state="disabled", style="Gray.TButton")
    show_btn = ttk.Button(footer, text="Show table", style="Gray.TButton")
    clear_btn = ttk.Button(footer, text="Clear all", style="Gray.TButton")
    run_btn.grid(row=0, column=1, padx=4)
    stop_btn.grid(row=0, column=2, padx=4)
    show_btn.grid(row=0, column=3, padx=4)
    clear_btn.grid(row=0, column=4, padx=4)

    log_q: queue.Queue[str] = queue.Queue()
    cancel_event = threading.Event()
    worker: dict[str, Optional[threading.Thread]] = {"thread": None}

    def append_log(msg: str) -> None:
        log_q.put(msg)

    def flush_log() -> None:
        updated = False
        while True:
            try:
                msg = log_q.get_nowait()
            except queue.Empty:
                break
            log_text.configure(state="normal")
            log_text.insert("end", msg + "\n")
            log_text.see("end")
            log_text.configure(state="disabled")
            updated = True
        if updated:
            log_text.see("end")
        root.after(200, flush_log)

    def set_running(running: bool) -> None:
        run_btn.configure(state="disabled" if running else "normal")
        stop_btn.configure(state="normal" if running else "disabled")
        status_var.set("Running" if running else "Idle")

    def clear_log() -> None:
        log_text.configure(state="normal")
        log_text.delete("1.0", "end")
        log_text.configure(state="disabled")

    def clear_all() -> None:
        clear_log()
        values_list.delete(0, "end")
        labels_list.delete(0, "end")
        outputs_list.delete(0, "end")
        value_entry.delete(0, "end")
        value_label_combo.set("")
        labels_combo.set("")
        out_var.set("outputs\\force_table.csv")
        mode_var.set("pair")
        start_var.set("-1,0")
        end_var.set("1,0")
        step_var.set("1")
        value_label_combo["values"] = []
        labels_combo["values"] = []
        if INTEGRAL_CHOICES:
            outputs_combo.current(0)
            outputs_list.insert("end", INTEGRAL_CHOICES[0][0])

    def show_table() -> None:
        path = out_var.get().strip()
        if not path:
            messagebox.showerror("Output", "Output path is empty.")
            return
        if not os.path.exists(path):
            messagebox.showerror("Output", f"File not found:\n{path}")
            return

        try:
            with open(path, newline="", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                header = next(reader, [])
                rows = reader
        except UnicodeDecodeError:
            with open(path, newline="", encoding="mbcs") as f:
                reader = csv.reader(f)
                header = next(reader, [])
                rows = reader
        except Exception as exc:
            messagebox.showerror("Output", f"Failed to read file:\n{exc}")
            return

        if not header:
            messagebox.showerror("Output", "CSV file is empty.")
            return

        preview = tk.Toplevel(root)
        preview.title("Output Table")
        preview.geometry("740x420+140+120")
        preview.configure(background=BG)

        preview_main = ttk.Frame(preview, padding=10, style="App.TFrame")
        preview_main.pack(fill="both", expand=True)
        preview_main.rowconfigure(0, weight=1)
        preview_main.columnconfigure(0, weight=1)

        tree = ttk.Treeview(preview_main, columns=header, show="headings")
        for col in header:
            tree.heading(col, text=col)
            tree.column(col, width=110, anchor="w")

        y_scroll = ttk.Scrollbar(preview_main, orient="vertical", command=tree.yview)
        x_scroll = ttk.Scrollbar(preview_main, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        max_rows = 500
        shown = 0
        for row in rows:
            if shown >= max_rows:
                break
            if len(row) < len(header):
                row = row + [""] * (len(header) - len(row))
            tree.insert("", "end", values=row[: len(header)])
            shown += 1

        if shown == max_rows:
            ttk.Label(
                preview_main,
                text=f"Showing first {max_rows} rows.",
                style="App.TLabel",
            ).grid(row=2, column=0, sticky="w", pady=(6, 0))

    def collect_inputs() -> dict:
        if win32com is None:
            raise RuntimeError("pywin32 is not available. Install with: pip install pywin32")

        mode = mode_var.get().strip().lower()
        value_items = list(values_list.get(0, "end"))
        if not value_items:
            raise ValueError("Add at least one label value.")
        mapping = _parse_label_values("; ".join(value_items))
        if mode == "pair":
            cases = _cases_from_mapping_pair(mapping)
        else:
            cases = _cases_from_mapping_all(mapping)
        if not cases:
            raise ValueError("No cases provided.")

        field_name = "Loading"

        selected_labels = list(labels_list.get(0, "end"))
        if not selected_labels:
            raise ValueError("Select at least one moving label.")
        move_labels = list(selected_labels)

        selected_outputs = list(outputs_list.get(0, "end"))
        if not selected_outputs:
            raise ValueError("Select at least one output value.")
        integrals = [display_to_integral[name] for name in selected_outputs if name in display_to_integral]

        x0, y0 = _parse_xy(start_var.get())
        x1, y1 = _parse_xy(end_var.get())
        step = float(step_var.get().strip())
        positions = _positions_line(x0, y0, x1, y1, step)

        return {
            "pbm": "",
            "model_path": "",
            "move_labels": move_labels,
            "cases": cases,
            "field_name": field_name,
            "positions": positions,
            "integrals": integrals,
            "mesh": True,
            "remesh": True,
            "mesh_once": False,
            "sleep_s": 0.5,
            "out_path": out_var.get().strip(),
        }

    def run_clicked() -> None:
        if worker["thread"] is not None and worker["thread"].is_alive():
            return
        try:
            params = collect_inputs()
        except Exception as exc:
            messagebox.showerror("Input error", str(exc))
            return

        cancel_event.clear()
        set_running(True)
        append_log("Starting...")

        def _worker() -> None:
            if pythoncom is not None:
                try:
                    pythoncom.CoInitialize()
                except Exception:
                    pass
            try:
                rc = run_batch_force_plan(
                    log=append_log,
                    cancel=cancel_event,
                    **params,
                )
                append_log(f"Done. Exit code: {rc}")
            except Exception as exc:
                append_log(f"Error: {exc}")
            finally:
                if pythoncom is not None:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
                root.after(0, lambda: set_running(False))

        t = threading.Thread(target=_worker, daemon=True)
        worker["thread"] = t
        t.start()

    def stop_clicked() -> None:
        if worker["thread"] is not None and worker["thread"].is_alive():
            cancel_event.set()
            append_log("Cancel requested.")

    run_btn.configure(command=run_clicked)
    stop_btn.configure(command=stop_clicked)
    show_btn.configure(command=show_table)
    clear_btn.configure(command=clear_all)

    flush_log()
    root.mainloop()
    return 0


def _add_choice(combo: ttk.Combobox, listbox: tk.Listbox) -> None:
    choice = combo.get().strip()
    if not choice:
        return
    existing = set(listbox.get(0, "end"))
    if choice in existing:
        return
    listbox.insert("end", choice)


def _add_value_pair(combo: ttk.Combobox, entry: ttk.Entry, listbox: tk.Listbox) -> None:
    label = combo.get().strip()
    values = entry.get().strip()
    if not label or not values:
        return
    text = f"{label}={values}"
    existing = set(listbox.get(0, "end"))
    if text in existing:
        return
    listbox.insert("end", text)
    entry.delete(0, "end")


def _clear_selected(listbox: tk.Listbox) -> None:
    selected = list(listbox.curselection())
    for idx in reversed(selected):
        listbox.delete(idx)


def _clear_selected_multi(*listboxes: tk.Listbox) -> None:
    for lb in listboxes:
        _clear_selected(lb)


def _load_label_choices(*combos: ttk.Combobox) -> None:
    if win32com is None:
        messagebox.showerror("Missing dependency", "pywin32 is not available.")
        return

    qf = dispatch_qf_app()
    problem = open_problem(qf, "")
    if problem is None:
        messagebox.showerror("QuickField", "No active problem found or PBM not found.")
        return

    model = ensure_model_loaded(problem, None)
    if model is None:
        messagebox.showerror("QuickField", "Failed to load model.")
        return

    labels = list_block_labels(model)
    for combo in combos:
        combo["values"] = labels
        if labels:
            combo.current(0)


def cmd_gui(args: argparse.Namespace) -> int:
    _ = args
    return launch_gui()


