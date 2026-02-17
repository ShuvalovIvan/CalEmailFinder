"""
Data Mapper Application (Production Build).
"""

import os
import time
import json
import threading
import queue
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont
from typing import List, Optional, Any, Dict

# Third-party imports
from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
import pandas as pd
import numpy as np


# --- 1. ROBUST BROWSER CHECK ---
def ensure_browser_installed():
    """Checks for Playwright browsers and installs Chromium if missing."""
    try:
        # Try to launch a dummy browser to see if it exists
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            p.chromium.launch(headless=True).close()
    except Exception:
        # If launch fails, try to install
        try:
            print("Installing browser engine (first run only)...")
            # FIX: Use sys.executable to ensure we use the correct python environment
            subprocess.check_call(
                [sys.executable, "-m", "playwright", "install", "chromium"]
            )
        except Exception as e:
            messagebox.showerror(
                "Setup Error", f"Failed to install browser engine:\n{e}"
            )


# --- 2. SCRAPER IMPORT / MOCK ---
try:
    import scraper as sc  # type: ignore
except ImportError:

    class MockScraper:
        def __init__(self):
            self.current_url = "N/A"

        def find_principal_data(self, text: str) -> Dict[str, str]:
            self.current_url = f"https://www.cde.ca.gov/mock_search?q={text}"
            time.sleep(0.5)
            if "fail" in text.lower():
                raise TimeoutError("Simulated Network Timeout")
            return {
                "First Name": "MockFirst",
                "Last Name": "MockLast",
                "Job Title": "Principal",
                "Email": "mock@school.edu",
                "Phone": "555-0199",
            }

        def close(self) -> None:
            pass

    class ScraperWrapper:
        CDEScraper = MockScraper

    sc = ScraperWrapper()


# --- 3. CONSTANTS ---
RECOVERY_DATA_FILE = "_recovery_data.csv"
RECOVERY_META_FILE = "_recovery_meta.json"
WINDOW_SIZE = "1300x700"
SIDEBAR_WIDTH = 250
AVAILABLE_FIELDS = ["First Name", "Last Name", "Job Title", "Email", "Phone"]


# --- 4. GUI CLASSES ---


class FieldMappingDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, current_columns: List[str]) -> None:
        super().__init__(parent)
        self.title("Map Extracted Data")
        self.geometry("500x500")
        self.transient(parent)
        self.grab_set()
        self.column_map: Dict[str, str] = {}
        self.cancelled = True
        self._setup_ui(current_columns)

    def _setup_ui(self, current_columns: List[str]) -> None:
        main_frame = tk.Frame(self, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            main_frame,
            text="Match Scraper Fields to Columns",
            font=("Arial", 12, "bold"),
        ).pack(pady=(0, 10))
        tk.Label(
            main_frame, text="Select a destination for each field found.", fg="gray"
        ).pack(pady=(0, 15))

        grid_frame = tk.Frame(main_frame)
        grid_frame.pack(fill=tk.X)

        self.selections = {}
        for i, field in enumerate(AVAILABLE_FIELDS):
            tk.Label(grid_frame, text=f"{field} :", font=("Arial", 10, "bold")).grid(
                row=i, column=0, sticky="e", padx=10, pady=8
            )
            options = ["--- Skip ---", f"[NEW] {field}"] + current_columns
            var = tk.StringVar(value=options[1])
            dropdown = ttk.Combobox(
                grid_frame, textvariable=var, values=options, state="readonly", width=30
            )
            dropdown.grid(row=i, column=1, sticky="w", padx=10, pady=8)
            self.selections[field] = var

        tk.Button(
            main_frame,
            text="Start Extraction",
            command=self.on_submit,
            bg="#e6ffe6",
            height=2,
        ).pack(side=tk.BOTTOM, fill=tk.X, pady=10)

    def on_submit(self) -> None:
        for field, var in self.selections.items():
            choice = var.get()
            if choice == "--- Skip ---":
                continue
            elif choice.startswith("[NEW]"):
                self.column_map[field] = field
            else:
                self.column_map[field] = choice

        if not self.column_map:
            messagebox.showwarning("Warning", "Please map at least one field.")
            return
        self.cancelled = False
        self.destroy()


class ProgressWindow(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        total: int,
        start_val: int = 0,
        title: str = "Processing...",
    ) -> None:
        super().__init__(parent)
        self.title(title)
        self.geometry("400x200")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        self.cancelled = threading.Event()
        self.paused = threading.Event()
        self.save_and_quit = threading.Event()

        self._setup_ui(total, start_val)
        self.protocol("WM_DELETE_WINDOW", self.cancel_process)

    def _setup_ui(self, total: int, start_val: int) -> None:
        frame = tk.Frame(self, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        self.lbl_status = tk.Label(frame, text="Starting...", anchor="w")
        self.lbl_status.pack(fill=tk.X, pady=(0, 5))
        self.progress = ttk.Progressbar(
            frame, orient="horizontal", length=300, mode="determinate"
        )
        self.progress.pack(fill=tk.X, pady=5)
        self.progress["maximum"] = total
        self.progress["value"] = start_val
        self.lbl_count = tk.Label(frame, text=f"{start_val} / {total}", fg="gray")
        self.lbl_count.pack(pady=2)

        btn_frame = tk.Frame(frame)
        btn_frame.pack(pady=10, fill=tk.X)

        self.btn_pause = tk.Button(
            btn_frame, text="Pause", command=self.toggle_pause, width=10
        )
        self.btn_pause.pack(side=tk.LEFT, padx=5)
        self.btn_save = tk.Button(
            btn_frame, text="Save & Quit", command=self.trigger_save_quit, width=15
        )
        self.btn_save.pack(side=tk.LEFT, padx=5)
        tk.Button(
            btn_frame, text="Cancel", command=self.cancel_process, width=10, fg="red"
        ).pack(side=tk.RIGHT, padx=5)

    def update_progress(self, value: int, status_text: Optional[str] = None) -> None:
        if self.cancelled.is_set():
            return
        self.progress["value"] = value
        self.lbl_count.config(text=f"{value} / {self.progress['maximum']}")
        if status_text:
            self.lbl_status.config(text=status_text)
        self.update()

    def toggle_pause(self) -> None:
        if self.paused.is_set():
            self.paused.clear()
            self.btn_pause.config(text="Pause", bg="SystemButtonFace")
            self.lbl_status.config(text="Resuming...")
        else:
            self.paused.set()
            self.btn_pause.config(text="Resume", bg="#e6ffe6")
            self.lbl_status.config(text="PAUSED. Click Resume to continue.")

    def trigger_save_quit(self) -> None:
        self.paused.set()
        self.save_and_quit.set()
        self.lbl_status.config(text="Saving progress and closing...")

    def cancel_process(self) -> None:
        if messagebox.askyesno(
            "Cancel", "Are you sure? Unsaved progress will be lost."
        ):
            self.cancelled.set()
            self.destroy()


class DataViewer(TkinterDnD.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Data Mapper App")
        self.geometry(WINDOW_SIZE)
        self.df: Optional[pd.DataFrame] = None
        self.msg_queue = queue.Queue()
        self.thread_decision = {}

        self._setup_layout()
        self._setup_bindings()
        self.after(500, self.check_for_recovery)

    def _setup_layout(self) -> None:
        top_frame = tk.Frame(self)
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        main_body = tk.Frame(self)
        main_body.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        sidebar = tk.Frame(main_body, width=SIDEBAR_WIDTH)
        sidebar.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        sidebar.pack_propagate(False)
        data_frame = tk.Frame(main_body)
        data_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        data_frame.rowconfigure(0, weight=1)
        data_frame.columnconfigure(0, weight=1)

        tk.Button(top_frame, text="Open File", command=self.open_file_dialog).pack(
            side=tk.LEFT
        )
        tk.Button(top_frame, text="Save / Export", command=self.save_file_dialog).pack(
            side=tk.LEFT, padx=10
        )
        tk.Button(
            top_frame,
            text="Export Failed Rows",
            command=self.export_failed_rows,
            bg="#fff0f0",
        ).pack(side=tk.LEFT, padx=10)
        tk.Button(
            top_frame,
            text="Merge Fixed Data",
            command=self.merge_back_data,
            bg="#e6f3ff",
        ).pack(side=tk.LEFT, padx=10)
        self.lbl_status = tk.Label(
            top_frame, text="Drag & Drop a CSV or XLSX file here", fg="gray"
        )
        self.lbl_status.pack(side=tk.LEFT, padx=10)

        tk.Label(sidebar, text="Column Selector", font=("Arial", 10, "bold")).pack(
            pady=(0, 5)
        )
        self.col_listbox = tk.Listbox(sidebar, selectmode=tk.EXTENDED, height=15)
        self.col_listbox.pack(side=tk.TOP, fill=tk.X)
        tk.Label(
            sidebar, text="Ctrl+Click to select multiple", fg="gray", font=("Arial", 8)
        ).pack()

        manage_frame = tk.Frame(sidebar)
        manage_frame.pack(side=tk.TOP, fill=tk.X, pady=5)
        tk.Button(
            manage_frame, text="▲ Move Up", command=lambda: self.move_column(-1)
        ).pack(side=tk.LEFT, expand=True, fill=tk.X)
        tk.Button(
            manage_frame, text="▼ Move Down", command=lambda: self.move_column(1)
        ).pack(side=tk.LEFT, expand=True, fill=tk.X)
        tk.Button(
            sidebar,
            text="Delete Selected Column(s)",
            bg="#ffcccc",
            command=self.delete_columns,
        ).pack(side=tk.TOP, fill=tk.X, pady=(2, 10))

        action_frame = tk.Frame(sidebar)
        action_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=10)
        tk.Label(action_frame, text="--- Step 1: Merge ---", fg="#555").pack(
            pady=(10, 2)
        )
        tk.Button(
            action_frame, text="Merge Selected Cols", command=self.merge_columns
        ).pack(fill=tk.X, pady=2)
        tk.Label(action_frame, text="--- Step 2: Transform ---", fg="#555").pack(
            pady=(15, 2)
        )
        tk.Button(
            action_frame,
            text="Extract Principal Info",
            command=self.extract_info,
            bg="#e6f3ff",
        ).pack(fill=tk.X, pady=2)
        tk.Label(
            action_frame,
            text="(Uses selected col as input)",
            font=("Arial", 8),
            fg="gray",
        ).pack()

        self.tree = ttk.Treeview(data_frame, show="headings")
        vsb = ttk.Scrollbar(data_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(data_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    def _setup_bindings(self) -> None:
        self.tree.bind("<ButtonRelease-1>", self.copy_cell_content)
        self.tree.bind("<Shift-MouseWheel>", self.on_horizontal_scroll)
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self.handle_drop)

    # --- RECOVERY & I/O ---
    def check_for_recovery(self) -> None:
        if os.path.exists(RECOVERY_META_FILE) and os.path.exists(RECOVERY_DATA_FILE):
            if messagebox.askyesno("Resume?", "Unfinished process found.\nResume?"):
                self.resume_process()
            else:
                self.clear_recovery_files()

    def save_recovery_state(
        self, current_index: int, source_col: str, column_map: Dict[str, str]
    ) -> None:
        if self.df is None:
            return
        try:
            self.df.to_csv(RECOVERY_DATA_FILE, index=False)
            with open(RECOVERY_META_FILE, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "current_index": current_index,
                        "source_col": source_col,
                        "column_map": column_map,
                    },
                    f,
                )
        except Exception:
            pass

    def clear_recovery_files(self) -> None:
        try:
            if os.path.exists(RECOVERY_DATA_FILE):
                os.remove(RECOVERY_DATA_FILE)
            if os.path.exists(RECOVERY_META_FILE):
                os.remove(RECOVERY_META_FILE)
        except OSError:
            pass

    def resume_process(self) -> None:
        try:
            self.df = pd.read_csv(RECOVERY_DATA_FILE)
            with open(RECOVERY_META_FILE, "r", encoding="utf-8") as f:
                meta = json.load(f)
            self.refresh_display()
            self.lbl_status.config(text=f"Resumed from row {meta['current_index']}")
            self.run_extraction_thread(
                meta["source_col"], meta["column_map"], meta["current_index"]
            )
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def ask_sheet_selection(self, sheet_names: List[str]) -> Optional[str]:
        popup = tk.Toplevel(self)
        popup.title("Select Sheet")
        popup.geometry("300x150")
        popup.transient(self)
        popup.grab_set()
        tk.Label(popup, text="Multiple sheets found.").pack(pady=10)
        selected = tk.StringVar()
        combo = ttk.Combobox(
            popup, textvariable=selected, values=sheet_names, state="readonly"
        )
        combo.pack(pady=5)
        combo.current(0)
        res = [None]

        def on_confirm():
            res[0] = selected.get()
            popup.destroy()

        tk.Button(popup, text="Load Sheet", command=on_confirm).pack(pady=10)
        self.wait_window(popup)
        return res[0]

    def ask_column_selection(
        self, title: str, prompt: str, df_to_use=None
    ) -> Optional[str]:
        popup = tk.Toplevel(self)
        popup.title(title)
        popup.geometry("300x150")
        popup.transient(self)
        popup.grab_set()
        tk.Label(popup, text=prompt).pack(pady=10)
        selected = tk.StringVar()
        target_df = df_to_use if df_to_use is not None else self.df
        cols = list(target_df.columns) if target_df is not None else []
        combo = ttk.Combobox(
            popup, textvariable=selected, values=cols, state="readonly"
        )
        combo.pack(pady=5)
        if cols:
            combo.current(0)
        res = [None]

        def on_confirm():
            res[0] = selected.get()
            popup.destroy()

        tk.Button(popup, text="Select", command=on_confirm).pack(pady=10)
        self.wait_window(popup)
        return res[0]

    def open_file_dialog(self) -> None:
        fp = filedialog.askopenfilename(
            filetypes=[("Data Files", "*.csv;*.xlsx;*.xls")]
        )
        if fp:
            self.load_data(fp)

    def handle_drop(self, event: Any) -> None:
        fp = event.data
        if fp.startswith("{") and fp.endswith("}"):
            fp = fp[1:-1]
        self.load_data(fp)

    def load_data(self, file_path: str) -> None:
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".csv":
                self.df = pd.read_csv(file_path)
            elif ext in [".xlsx", ".xls"]:
                xls = pd.ExcelFile(file_path)
                sheet_to_load = xls.sheet_names[0]
                if len(xls.sheet_names) > 1:
                    sel = self.ask_sheet_selection(xls.sheet_names)
                    if not sel:
                        return
                    sheet_to_load = sel
                self.df = pd.read_excel(xls, sheet_name=sheet_to_load)
            else:
                return
            self.refresh_display()
            self.lbl_status.config(text=f"Loaded: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def save_file_dialog(self) -> None:
        if self.df is None:
            messagebox.showwarning("Warning", "No data to save.")
            return
        fp = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx")],
        )
        if not fp:
            return
        try:
            ext = os.path.splitext(fp)[1].lower()
            if ext == ".csv":
                self.df.to_csv(fp, index=False)
            elif ext == ".xlsx":
                with pd.ExcelWriter(fp, engine="xlsxwriter") as writer:
                    self.df.to_excel(writer, index=False, sheet_name="Sheet1")
                    fmt = writer.book.add_format({"text_wrap": True})  # type: ignore
                    writer.sheets["Sheet1"].set_column("A:AZ", 20, fmt)  # type: ignore
            messagebox.showinfo("Success", "File saved!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def export_failed_rows(self) -> None:
        if self.df is None:
            return
        target_col = self.ask_column_selection(
            "Export Failed", "Which column to check for errors?"
        )
        if not target_col:
            return
        col_str = self.df[target_col].fillna("").astype(str).str.strip()
        failed_df = self.df[
            col_str.isin(["", "Error", "no_email_found", "nan", "None", "{}"])
        ]
        if failed_df.empty:
            messagebox.showinfo("Info", "No failed rows found!")
            return
        fp = filedialog.asksaveasfilename(
            defaultextension=".csv", initialfile="failed_rows_retry.csv"
        )
        if fp:
            failed_df.to_csv(fp, index=False)
            messagebox.showinfo("Success", f"Exported {len(failed_df)} rows.")

    # --- MERGE BACK ---
    def merge_back_data(self) -> None:
        if self.df is None:
            messagebox.showwarning("Error", "Please load your MAIN/MASTER file first.")
            return
        fp = filedialog.askopenfilename(
            title="Select the Fixed/Retried CSV file",
            filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx;*.xls")],
        )
        if not fp:
            return
        try:
            if fp.endswith(".csv"):
                fixed_df = pd.read_csv(fp)
            else:
                fixed_df = pd.read_excel(fp)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read fixed file:\n{e}")
            return

        key_col = self.ask_column_selection(
            "Select Key Column", "Which column is the Unique ID (e.g., School Name)?"
        )
        if not key_col:
            return
        if key_col not in fixed_df.columns:
            messagebox.showerror(
                "Error", f"Key column '{key_col}' not found in the fixed file."
            )
            return
        shared_columns = [
            col for col in fixed_df.columns if col in self.df.columns and col != key_col
        ]
        if not shared_columns:
            messagebox.showwarning("Merge Error", "No matching data columns found!")
            return
        confirm_msg = (
            f"Found {len(shared_columns)} columns to update:\n\n"
            + ", ".join(shared_columns)
            + "\n\nProceed with merge?"
        )
        if not messagebox.askyesno("Confirm Merge", confirm_msg):
            return

        try:
            self.df[key_col] = self.df[key_col].astype(str)
            fixed_df[key_col] = fixed_df[key_col].astype(str)
            for col in shared_columns:
                self.df[col] = self.df[col].astype(object)
            rows_updated = 0
            for index, fixed_row in fixed_df.iterrows():
                key_val = fixed_row[key_col]
                mask = self.df[key_col] == key_val
                if mask.any():
                    row_has_updates = False
                    for col in shared_columns:
                        new_val = fixed_row[col]
                        if not pd.isna(new_val) and str(new_val).strip() not in [
                            "",
                            "nan",
                            "None",
                            "Error",
                            "{}",
                        ]:
                            self.df.loc[mask, col] = new_val
                            row_has_updates = True
                    if row_has_updates:
                        rows_updated += 1
            self.refresh_display()
            messagebox.showinfo(
                "Merge Complete", f"Successfully updated {rows_updated} rows."
            )
        except Exception as e:
            messagebox.showerror("Merge Error", f"An error occurred during merge:\n{e}")

    # --- ERROR DIALOG ---
    def ask_error_resolution(
        self, row_num: int, error_url: str, current_term: str
    ) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Connection Error Resolution")
        dialog.geometry("500x350")
        dialog.transient(self)
        dialog.grab_set()
        tk.Label(
            dialog, text=f"Error at Row {row_num}", font=("Arial", 12, "bold"), fg="red"
        ).pack(pady=10)
        tk.Label(
            dialog, text="Last Visited URL (Timed Out):", font=("Arial", 9, "bold")
        ).pack(anchor="w", padx=10)
        err_frame = tk.Frame(dialog, padx=10)
        err_frame.pack(fill=tk.BOTH, expand=True)
        text_w = tk.Text(err_frame, height=4, width=50, bg="#f0f0f0")
        text_w.insert("1.0", error_url)
        text_w.configure(state="disabled")
        text_w.pack()
        tk.Label(dialog, text="Search Term Used:", font=("Arial", 10, "bold")).pack(
            pady=(10, 5)
        )
        term_var = tk.StringVar(value=current_term)
        entry = tk.Entry(dialog, textvariable=term_var, width=50, font=("Arial", 11))
        entry.pack(pady=5)
        btn_frame = tk.Frame(dialog, pady=20)
        btn_frame.pack(fill=tk.X)

        def on_retry():
            self.thread_decision = {"action": "retry", "new_term": term_var.get()}
            dialog.destroy()

        def on_skip():
            self.thread_decision = {"action": "skip", "new_term": None}
            dialog.destroy()

        def on_cancel():
            self.thread_decision = {"action": "cancel", "new_term": None}
            dialog.destroy()

        tk.Button(btn_frame, text="Skip Row", command=on_skip, width=12).pack(
            side=tk.LEFT, padx=20
        )
        tk.Button(
            btn_frame, text="Retry", command=on_retry, bg="#e6ffe6", width=12
        ).pack(side=tk.LEFT, padx=10)
        tk.Button(
            btn_frame, text="Stop Process", command=on_cancel, bg="#ffe6e6", width=12
        ).pack(side=tk.RIGHT, padx=20)
        dialog.protocol("WM_DELETE_WINDOW", on_cancel)
        self.wait_window(dialog)

    # --- COLUMN OPERATIONS (INDENTED CORRECTLY) ---
    def delete_columns(self) -> None:
        if self.df is None:
            return
        sel = [self.col_listbox.get(i) for i in self.col_listbox.curselection()]
        if sel and messagebox.askyesno("Delete", f"Delete {len(sel)} cols?"):
            self.df.drop(columns=sel, inplace=True)
            self.refresh_display()

    def move_column(self, direction: int) -> None:
        if self.df is None:
            return
        sel = list(self.col_listbox.curselection())
        if len(sel) != 1:
            return
        idx = sel[0]
        cols = list(self.df.columns)
        if (direction == -1 and idx == 0) or (direction == 1 and idx == len(cols) - 1):
            return
        nidx = idx + direction
        cols[idx], cols[nidx] = cols[nidx], cols[idx]
        self.df = self.df[cols]
        self.refresh_display()
        self.col_listbox.selection_set(nidx)
        self.col_listbox.see(nidx)

    def merge_columns(self) -> None:
        if self.df is None:
            return
        sel = [self.col_listbox.get(i) for i in self.col_listbox.curselection()]
        if not sel:
            return
        name = base = "+".join(sel)
        c = 1
        while name in self.df.columns:
            name = f"{base}_{c}"
            c += 1
        try:
            self.df[name] = self.df[sel].fillna("").astype(str).agg(" ".join, axis=1)
            self.refresh_display()
            messagebox.showinfo("Success", f"Created: {name}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def copy_cell_content(self, event: Any) -> None:
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id, col_id = self.tree.identify_row(event.y), self.tree.identify_column(
            event.x
        )
        if not row_id or not col_id:
            return
        col_index = int(col_id.replace("#", "")) - 1
        vals = self.tree.item(row_id, "values")
        if col_index < len(vals):
            self.clipboard_clear()
            self.clipboard_append(vals[col_index])
            self.update()
            self.lbl_status.config(text=f"Copied: '{str(vals[col_index])[:30]}...'")

    def copy_column(self, col_name: str) -> None:
        if self.df is None:
            return
        try:
            self.df[col_name].to_clipboard(index=False, header=False)
            self.lbl_status.config(text=f"Copied: {col_name}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def on_horizontal_scroll(self, event: Any) -> str:
        if event.delta:
            self.tree.xview_scroll(int(-1 * (event.delta / 120) * 5), "units")
        return "break"

    def refresh_display(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.col_listbox.delete(0, tk.END)
        self.tree["columns"] = []
        if self.df is None:
            return
        cols = list(self.df.columns)
        self.tree["columns"] = cols
        font = tkfont.Font()
        for i, col in enumerate(cols):
            self.tree.heading(col, text=col, command=lambda c=col: self.copy_column(c))
            mw = font.measure(str(col)) + 20
            for val in self.df.iloc[:, i].fillna("").astype(str).head(50):
                for line in str(val).split("\n"):
                    w = font.measure(line) + 20
                    mw = max(mw, w)
            self.tree.column(col, width=min(mw, 600), stretch=False)
            self.col_listbox.insert(tk.END, col)
        try:
            mx = self.df.astype(str).apply(lambda x: x.str.count("\n").max()).max()
            if pd.isna(mx):
                mx = 0
            ttk.Style().configure(
                "Treeview", rowheight=25 + (min(int(mx + 1), 15) - 1) * 18
            )
        except:
            pass
        for row in self.df.fillna("").to_numpy().tolist():
            self.tree.insert("", "end", values=row)

    # --- EXTRACTION ---
    def extract_info(self) -> None:
        if self.df is None:
            return
        sel = self.col_listbox.curselection()
        if len(sel) != 1:
            messagebox.showwarning("Error", "Select exactly ONE source column.")
            return
        source = self.col_listbox.get(sel[0])
        dialog = FieldMappingDialog(self, list(self.df.columns))
        self.wait_window(dialog)
        if dialog.cancelled:
            return
        column_map = dialog.column_map
        for target in column_map.values():
            if target not in self.df.columns:
                self.df[target] = ""
            self.df[target] = self.df[target].astype(object)
        self.refresh_display()
        self.run_extraction_thread(source, column_map, 0)

    def run_extraction_thread(
        self, source_col: str, column_map: Dict[str, str], start_index: int
    ) -> None:
        input_data = self.df[source_col].fillna("").astype(str).tolist()
        total = len(input_data)
        self.pwin = ProgressWindow(
            self, total, start_index, "Scraping Principal Info..."
        )

        def worker():
            try:
                scraper = sc.CDEScraper()
            except Exception as e:
                self.msg_queue.put(("error", f"Scraper Start Error:\n{e}"))
                return

            for i in range(start_index, total):
                if self.pwin.cancelled.is_set():
                    self.msg_queue.put(("cancelled",))
                    break
                if self.pwin.save_and_quit.is_set():
                    self.msg_queue.put(("save_quit", i, source_col, column_map))
                    break

                while self.pwin.paused.is_set():
                    time.sleep(0.1)
                    if self.pwin.cancelled.is_set():
                        self.msg_queue.put(("cancelled",))
                        return
                    if self.pwin.save_and_quit.is_set():
                        self.msg_queue.put(("save_quit", i, source_col, column_map))
                        return

                current_search_term = input_data[i]
                while True:
                    try:
                        result_dict = scraper.find_principal_data(current_search_term)
                        self.msg_queue.put(("result", i, result_dict))
                        break
                    except Exception as e:
                        last_url = getattr(scraper, "current_url", "URL Unknown")
                        self.pwin.paused.set()
                        self.msg_queue.put(
                            ("network_error", i + 1, last_url, current_search_term)
                        )
                        while self.pwin.paused.is_set():
                            time.sleep(0.1)
                            if self.pwin.cancelled.is_set():
                                break
                        if self.pwin.cancelled.is_set():
                            break
                        decision = self.thread_decision.copy()
                        if decision.get("action") == "skip":
                            self.msg_queue.put(("result", i, {}))
                            break
                        elif decision.get("action") == "retry":
                            current_search_term = decision.get(
                                "new_term", current_search_term
                            )
                            continue
                        elif decision.get("action") == "cancel":
                            self.pwin.cancelled.set()
                            break

                if self.pwin.cancelled.is_set():
                    self.msg_queue.put(("cancelled",))
                    break
                if i % 10 == 0:
                    self.msg_queue.put(("autosave", i + 1, source_col, column_map))

            try:
                scraper.close()
            except:
                pass
            if (
                not self.pwin.cancelled.is_set()
                and not self.pwin.save_and_quit.is_set()
            ):
                self.msg_queue.put(("done",))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self.monitor_queue, column_map)

    def monitor_queue(self, column_map: Dict[str, str]) -> None:
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                mtype = msg[0]
                if mtype == "result":
                    idx, res = msg[1], msg[2]
                    for field, target in column_map.items():
                        self.df.at[idx, target] = res.get(field, "")
                    self.pwin.update_progress(idx + 1, f"Row {idx + 1}...")
                    if idx % 5 == 0:
                        children = self.tree.get_children()
                        if idx < len(children):
                            vals = list(self.tree.item(children[idx], "values"))
                            for field, target in column_map.items():
                                vals[self.df.columns.get_loc(target)] = res.get(
                                    field, ""
                                )
                            self.tree.item(children[idx], values=vals)
                elif mtype == "network_error":
                    row_num, err_text, term = msg[1], msg[2], msg[3]
                    self.pwin.lbl_status.config(text="⚠ Error - Waiting for input...")
                    self.pwin.btn_pause.config(text="Resume", bg="#e6ffe6")
                    self.ask_error_resolution(row_num, err_text, term)
                    self.pwin.toggle_pause()
                elif mtype == "error":
                    self.pwin.destroy()
                    messagebox.showerror("Error", msg[1])
                    return
                elif mtype == "autosave":
                    self.save_recovery_state(msg[1], msg[2], msg[3])
                elif mtype == "cancelled":
                    self.clear_recovery_files()
                    self.pwin.destroy()
                    return
                elif mtype == "save_quit":
                    self.save_recovery_state(msg[1], msg[2], msg[3])
                    self.pwin.destroy()
                    self.destroy()
                    return
                elif mtype == "done":
                    self.clear_recovery_files()
                    self.pwin.destroy()
                    self.refresh_display()
                    messagebox.showinfo("Done", "Complete!")
                    return
        except queue.Empty:
            pass
        if hasattr(self, "pwin") and self.pwin.winfo_exists():
            self.after(100, self.monitor_queue, column_map)


if __name__ == "__main__":
    ensure_browser_installed()
    app = DataViewer()
    app.mainloop()
