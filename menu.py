"""
Data Mapper Application (Multi-threaded).

Features:
- Threaded scraping (prevents GUI freezing).
- Pause/Resume/Cancel/Save-and-Quit.
- Crash recovery (auto-save).
- Excel formatting (text wrap).
- Export failed rows.
"""

import os
import time
import json
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import tkinter.font as tkfont
from typing import List, Optional, Any, Union

# Third-party imports
# pylint: disable=import-error
from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
import pandas as pd
import numpy as np

# Try to import scraper, provide dummy if missing
try:
    import scraper as sc  # type: ignore
except ImportError:

    class MockScraper:
        def find_emails(self, text: str) -> List[str]:
            # Simulate network delay without blocking GUI
            time.sleep(0.2)
            if "fail" in text.lower():
                return []
            return ["mock@email.com"]

        def close(self) -> None:
            pass

    class ScraperWrapper:
        CDEScraper = MockScraper

    sc = ScraperWrapper()  # type: ignore


# Constants
RECOVERY_DATA_FILE = "_recovery_data.csv"
RECOVERY_META_FILE = "_recovery_meta.json"
WINDOW_SIZE = "1200x700"
SIDEBAR_WIDTH = 250


class DestinationDialog(tk.Toplevel):
    """Modal dialog to ask the user where to save the transformation results."""

    def __init__(self, parent: tk.Misc, current_columns: List[str]) -> None:
        super().__init__(parent)
        self.title("Select Destination")
        self.geometry("350x250")
        self.transient(parent)
        self.grab_set()

        self.result_action: Optional[str] = None
        self.result_col_name: Optional[str] = None
        self._setup_ui(current_columns)

    def _setup_ui(self, current_columns: List[str]) -> None:
        padding = tk.Frame(self, padx=20, pady=20)
        padding.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            padding, text="Where should the result go?", font=("Arial", 11, "bold")
        ).pack(pady=(0, 15))
        self.var_choice = tk.StringVar(value="new")

        # Option 1: New Column
        frame_new = tk.Frame(padding)
        frame_new.pack(fill=tk.X, pady=5)
        tk.Radiobutton(
            frame_new, text="Create New Column:", variable=self.var_choice, value="new"
        ).pack(side=tk.LEFT)
        self.entry_new = tk.Entry(frame_new)
        self.entry_new.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=(10, 0))
        self.entry_new.insert(0, "Principal_Email")

        # Option 2: Existing Column
        frame_exist = tk.Frame(padding)
        frame_exist.pack(fill=tk.X, pady=5)
        tk.Radiobutton(
            frame_exist,
            text="Overwrite Column:",
            variable=self.var_choice,
            value="existing",
        ).pack(side=tk.LEFT)
        self.combo_exist = ttk.Combobox(
            frame_exist, values=current_columns, state="readonly"
        )
        self.combo_exist.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=(10, 0))
        if current_columns:
            self.combo_exist.current(0)

        tk.Button(
            padding,
            text="Run Transformation",
            command=self.on_submit,
            bg="#dddddd",
            height=2,
        ).pack(side=tk.BOTTOM, fill=tk.X, pady=10)

    def on_submit(self) -> None:
        choice = self.var_choice.get()
        if choice == "new":
            name = self.entry_new.get().strip()
            if not name:
                messagebox.showerror("Error", "Please enter a name.")
                return
            self.result_action = "new"
            self.result_col_name = name
        else:
            name = self.combo_exist.get()
            if not name:
                messagebox.showerror("Error", "Please select a column.")
                return
            self.result_action = "existing"
            self.result_col_name = name
        self.destroy()


class ProgressWindow(tk.Toplevel):
    """Popup window displaying a progress bar with Pause and Save/Quit capabilities."""

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

        # Thread-safe flags
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
        self.paused.set()  # Pause to stop processing
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

        self.msg_queue = queue.Queue()  # Queue for thread communication

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

        # Top Widgets
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
        self.lbl_status = tk.Label(
            top_frame, text="Drag & Drop a CSV or XLSX file here", fg="gray"
        )
        self.lbl_status.pack(side=tk.LEFT, padx=10)

        # Sidebar Widgets
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
            text="Find Principal Email",
            command=self.extract_emails,
            bg="#e6f3ff",
        ).pack(fill=tk.X, pady=2)
        tk.Label(
            action_frame,
            text="(Uses selected col as input)",
            font=("Arial", 8),
            fg="gray",
        ).pack()

        # Treeview
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

    # --- RECOVERY ---
    def check_for_recovery(self) -> None:
        if os.path.exists(RECOVERY_META_FILE) and os.path.exists(RECOVERY_DATA_FILE):
            if messagebox.askyesno(
                "Resume?", "Unfinished process found.\nDo you want to resume?"
            ):
                self.resume_process()
            else:
                self.clear_recovery_files()

    def save_recovery_state(
        self, current_index: int, source_col: str, target_col: str
    ) -> None:
        if self.df is None:
            return
        try:
            self.df.to_csv(RECOVERY_DATA_FILE, index=False)
            meta = {
                "current_index": current_index,
                "source_col": source_col,
                "target_col": target_col,
            }
            with open(RECOVERY_META_FILE, "w", encoding="utf-8") as f:
                json.dump(meta, f)
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
                meta["source_col"], meta["target_col"], meta["current_index"]
            )
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # --- UI HELPERS ---
    def ask_sheet_selection(self, sheet_names: List[str]) -> Optional[str]:
        popup = tk.Toplevel(self)
        popup.title("Select Sheet")
        popup.geometry("300x150")
        popup.transient(self)
        popup.grab_set()
        tk.Label(popup, text="Multiple sheets found.\nPlease select one:").pack(pady=10)
        selected_sheet = tk.StringVar()
        combo = ttk.Combobox(
            popup, textvariable=selected_sheet, values=sheet_names, state="readonly"
        )
        combo.pack(pady=5)
        combo.current(0)
        result: List[Optional[str]] = [None]

        def on_confirm() -> None:
            result[0] = selected_sheet.get()
            popup.destroy()

        tk.Button(popup, text="Load Sheet", command=on_confirm).pack(pady=10)
        self.wait_window(popup)
        return result[0]

    def ask_column_selection(self, title: str, prompt: str) -> Optional[str]:
        popup = tk.Toplevel(self)
        popup.title(title)
        popup.geometry("300x150")
        popup.transient(self)
        popup.grab_set()
        tk.Label(popup, text=prompt).pack(pady=10)
        selected_col = tk.StringVar()
        cols = list(self.df.columns) if self.df is not None else []
        combo = ttk.Combobox(
            popup, textvariable=selected_col, values=cols, state="readonly"
        )
        combo.pack(pady=5)
        if cols:
            combo.current(0)
        result: List[Optional[str]] = [None]

        def on_confirm() -> None:
            result[0] = selected_col.get()
            popup.destroy()

        tk.Button(popup, text="Select", command=on_confirm).pack(pady=10)
        self.wait_window(popup)
        return result[0]

    def on_horizontal_scroll(self, event: Any) -> str:
        if event.delta:
            self.tree.xview_scroll(int(-1 * (event.delta / 120) * 5), "units")
        return "break"

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
            self.lbl_status.config(text=f"Copied column: '{col_name}'")
        except Exception as e:
            messagebox.showerror("Clipboard Error", str(e))

    # --- FILE I/O ---
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
                sheets = xls.sheet_names
                sheet_to_load = sheets[0]
                if len(sheets) > 1:
                    sel = self.ask_sheet_selection(sheets)
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
                    workbook = writer.book
                    worksheet = writer.sheets["Sheet1"]
                    wrap_format = workbook.add_format({"text_wrap": True})
                    worksheet.set_column("A:AZ", 20, wrap_format)
            messagebox.showinfo("Success", "File saved successfully!")
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save file.\n{e}")

    def export_failed_rows(self) -> None:
        if self.df is None:
            return
        target_col = self.ask_column_selection(
            "Export Failed", "Which column contains the results/emails?"
        )
        if not target_col:
            return
        col_str = self.df[target_col].fillna("").astype(str).str.strip()
        fail_markers = ["", "Error", "no_email_found", "nan", "None"]
        failed_df = self.df[col_str.isin(fail_markers)]
        if failed_df.empty:
            messagebox.showinfo("Info", "No failed rows found based on that column!")
            return
        fp = filedialog.asksaveasfilename(
            defaultextension=".csv", initialfile="failed_rows_retry.csv"
        )
        if fp:
            failed_df.to_csv(fp, index=False)
            messagebox.showinfo("Success", f"Exported {len(failed_df)} failed rows.")

    # --- DISPLAY LOGIC ---
    def adjust_row_height(self) -> None:
        if self.df is None or self.df.empty:
            return
        try:
            max_newlines = (
                self.df.astype(str).apply(lambda x: x.str.count("\n").max()).max()
            )
            if pd.isna(max_newlines):
                max_newlines = 0
            lines = min(int(max_newlines + 1), 15)
            ttk.Style().configure("Treeview", rowheight=25 + ((lines - 1) * 18))
        except Exception:
            pass

    def autosize_columns(self) -> None:
        if self.df is None:
            return
        font = tkfont.Font()
        for i, col_name in enumerate(self.df.columns):
            mw = font.measure(str(col_name)) + 20
            sample = self.df.iloc[:, i].fillna("").astype(str).head(50)
            for val in sample:
                for line in str(val).split("\n"):
                    w = font.measure(line) + 20
                    if w > mw:
                        mw = w
            self.tree.column(col_name, width=min(mw, 600), stretch=False)

    def refresh_display(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.col_listbox.delete(0, tk.END)
        self.tree["columns"] = []
        if self.df is None:
            return
        cols = list(self.df.columns)
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col, command=lambda c=col: self.copy_column(c))
            self.tree.column(col, width=100, stretch=False)
            self.col_listbox.insert(tk.END, col)
        for row in self.df.fillna("").to_numpy().tolist():
            self.tree.insert("", "end", values=row)
        self.autosize_columns()
        self.adjust_row_height()

    # --- COLUMN OPERATIONS ---
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
        base = "+".join(sel)
        name = base
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

    # --- THREADED SCRAPER LOGIC ---
    def extract_emails(self) -> None:
        if self.df is None:
            return
        sel = self.col_listbox.curselection()
        if len(sel) != 1:
            messagebox.showwarning("Error", "Select exactly ONE source column.")
            return
        source = self.col_listbox.get(sel[0])

        dialog = DestinationDialog(self, list(self.df.columns))
        self.wait_window(dialog)
        if not dialog.result_action:
            return
        target = dialog.result_col_name
        if target not in self.df.columns:
            self.df[target] = ""
            self.refresh_display()

        # Start Thread
        self.run_extraction_thread(source, target, 0)

    def run_extraction_thread(
        self, source_col: str, target_col: str, start_index: int
    ) -> None:
        """Starts the worker thread and the monitoring loop."""
        input_data = self.df[source_col].fillna("").astype(str).tolist()
        total = len(input_data)

        # Create UI Window
        self.pwin = ProgressWindow(self, total, start_index, "Scraping Emails...")

        # Define Worker Function
        def worker():
            scraper = sc.CDEScraper()
            for i in range(start_index, total):
                # Check for kill signals from main thread
                if self.pwin.cancelled.is_set():
                    self.msg_queue.put(("cancelled",))
                    break

                if self.pwin.save_and_quit.is_set():
                    self.msg_queue.put(("save_quit", i, source_col, target_col))
                    break

                # Handle Pause (Sleep inside thread)
                while self.pwin.paused.is_set():
                    time.sleep(0.1)
                    if self.pwin.cancelled.is_set():
                        self.msg_queue.put(("cancelled",))
                        return
                    if self.pwin.save_and_quit.is_set():
                        self.msg_queue.put(("save_quit", i, source_col, target_col))
                        return

                # Scrape
                try:
                    res_str = "\n".join(scraper.find_emails(input_data[i]))
                except Exception:
                    res_str = "Error"

                # Send Result
                self.msg_queue.put(("result", i, res_str))

                # Auto-save request
                if i % 10 == 0:
                    self.msg_queue.put(("autosave", i + 1, source_col, target_col))

            scraper.close()
            if (
                not self.pwin.cancelled.is_set()
                and not self.pwin.save_and_quit.is_set()
            ):
                self.msg_queue.put(("done",))

        # Start Thread
        threading.Thread(target=worker, daemon=True).start()

        # Start Polling
        self.after(100, self.monitor_queue, target_col)

    def monitor_queue(self, target_col: str) -> None:
        """Polls the queue for updates from the thread."""
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                msg_type = msg[0]

                if msg_type == "result":
                    idx, res_str = msg[1], msg[2]
                    # Update Dataframe safely in main thread
                    self.df.at[idx, target_col] = res_str
                    # Update GUI
                    self.pwin.update_progress(idx + 1, f"Row {idx + 1}...")
                    # Update Treeview (every 5 rows visual update)
                    if idx % 5 == 0:
                        children = self.tree.get_children()
                        if idx < len(children):
                            vals = list(self.tree.item(children[idx], "values"))
                            cidx = self.df.columns.get_loc(target_col)
                            vals[cidx] = res_str
                            self.tree.item(children[idx], values=vals)

                elif msg_type == "autosave":
                    self.save_recovery_state(msg[1], msg[2], msg[3])

                elif msg_type == "cancelled":
                    self.clear_recovery_files()
                    self.pwin.destroy()
                    return  # Stop monitoring

                elif msg_type == "save_quit":
                    self.save_recovery_state(msg[1], msg[2], msg[3])
                    self.pwin.destroy()
                    self.destroy()
                    return

                elif msg_type == "done":
                    self.clear_recovery_files()
                    self.pwin.destroy()
                    self.refresh_display()
                    messagebox.showinfo("Done", "Process Complete!")
                    return

        except queue.Empty:
            # Continue monitoring
            pass

        # Schedule next check if window is still open
        if hasattr(self, "pwin") and self.pwin.winfo_exists():
            self.after(100, self.monitor_queue, target_col)


if __name__ == "__main__":
    app = DataViewer()
    app.mainloop()
