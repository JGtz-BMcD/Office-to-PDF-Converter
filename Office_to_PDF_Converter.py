import os
import sys
import re
import shutil
import tempfile
import time
import hashlib
from datetime import datetime
import ctypes

import win32com.client
import pywintypes
import winerror
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import pymupdf

VERSION = "7.0"
VERSION_DATE = "August 24, 2026"
APP_TITLE = f"BMcD Office to PDF Converter v{VERSION}"

# --- Word/Excel COM constants (named for readability; values confirmed live
# against the installed Word/Excel object model before writing this file) ---
WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_ITEM_DOCUMENT_WITH_MARKUP = 7
WD_WINDOW_VIEW_TYPE_PRINT = 3
WD_REVISIONS_VIEW_FINAL = 0
WD_REVISIONS_MODE_INLINE = 1
WD_REVISED_LINES_MARK_NONE = 0
WD_REVISED_LINES_MARK_RIGHT_BORDER = 2
WD_HEADER_FOOTER_TYPES = (1, 2, 3)
WD_STORY_TYPES_HEADER_FOOTER = (6, 7, 8, 9, 10, 11)
WD_REVISION_INSERT = 1
WD_REVISION_DELETE = 2
WD_REVISION_FORMAT_CHANGE = 3
XL_FIXED_FORMAT_PDF = 0
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# File-ledger disposition values used by the completeness reconciliation pass.
STATUS_PENDING = "Pending"
STATUS_CONVERTED = "Converted"
STATUS_INCLUDED_EXISTING_PDF = "Included (already a PDF)"
STATUS_EXCLUDED_USER = "Excluded by user"
STATUS_EXCLUDED_UNREADABLE = "Excluded (could not be read)"
STATUS_EXCLUDED_NOT_APPLICABLE = "Not applicable"
STATUS_FAILED = "Failed"

# COM error codes known to be transient during bulk Office automation (the callee
# is momentarily busy/showing UI) rather than a real, permanent failure.
TRANSIENT_COM_HRESULTS = {
    winerror.RPC_E_CALL_REJECTED,
    winerror.RPC_E_SERVERCALL_RETRYLATER,
    winerror.RPC_E_SERVERCALL_REJECTED,
    winerror.RPC_E_CANTCALLOUT_AGAIN,
    winerror.RPC_E_CANTCALLOUT_ININPUTSYNCCALL,
    winerror.RPC_E_DISCONNECTED,
}


def standard_sort_key(s):
    return os.path.basename(s).lower()


def natural_sort_key(s):
    s_base = os.path.basename(s)
    s_stem, s_ext = os.path.splitext(s_base)
    stem_key = [int(t) if t.isdigit() else t.lower() for t in re.split(r"([0-9]+)", s_stem)]
    return (stem_key, s_ext.lower())


def revision_aware_sort_key(s):
    s_base = os.path.basename(s)
    s_stem, s_ext = os.path.splitext(s_base)
    match = re.match(r"^(.*?)((?:[-_][a-zA-Z0-9]+)*)$", s_stem)
    if match:
        base_part, rev_part = match.groups()
    else:
        base_part, rev_part = s_stem, ""
    base_key = [int(t) if t.isdigit() else t.lower() for t in re.split(r"([0-9]+)", base_part)]
    return (base_key, rev_part.lower())


def compress_page_ranges(page_numbers):
    if not page_numbers:
        return "None"
    nums = sorted(set(int(p) for p in page_numbers))
    ranges = []
    start = prev = nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
        else:
            ranges.append(str(start) if start == prev else f"{start}-{prev}")
            start = prev = n
    ranges.append(str(start) if start == prev else f"{start}-{prev}")
    return ", ".join(ranges)


def retry_com_call(action, description, max_attempts=4, initial_delay=0.5, backoff=2.0):
    attempt = 0
    while True:
        try:
            return action()
        except pywintypes.com_error as e:
            hresult = e.args[0] if e.args else None
            transient = hresult in TRANSIENT_COM_HRESULTS or "rejected by callee" in str(e).lower()
            attempt += 1
            if not transient or attempt >= max_attempts:
                raise
            print(f"Transient COM error on {description} (attempt {attempt}/{max_attempts}), retrying...")
            time.sleep(initial_delay * (backoff ** (attempt - 1)))


def compute_sha256(path, chunk_size=1048576):
    try:
        digest = hashlib.sha256()
        with open(path, "rb") as f:
            while True:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                digest.update(chunk)
        return digest.hexdigest()
    except Exception:
        return None


def bring_window_to_front(window):
    try:
        window.lift()
        window.attributes("-topmost", True)
        window.after(250, lambda: window.attributes("-topmost", False))
        window.focus_force()
    except Exception:
        pass


def flash_window(window, count=5):
    try:
        hwnd = window.winfo_id()

        class FLASHWINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.c_uint),
                ("hwnd", ctypes.c_void_p),
                ("dwFlags", ctypes.c_uint),
                ("uCount", ctypes.c_uint),
                ("dwTimeout", ctypes.c_uint),
            ]

        FLASHW_TRAY = 0x00000002
        FLASHW_TIMERNOFG = 0x0000000C
        info = FLASHWINFO(ctypes.sizeof(FLASHWINFO), hwnd, FLASHW_TRAY | FLASHW_TIMERNOFG, count, 0)
        ctypes.windll.user32.FlashWindowEx(ctypes.byref(info))
    except Exception:
        pass


class ScrollableFrame(tk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class ExcelTabSelectionDialog(tk.Toplevel):
    def __init__(self, parent, excel_file_info):
        super().__init__(parent)
        self.title("Select Excel Tabs for Conversion")
        self.geometry("450x480")
        self.resizable(False, False)
        self.result = {}
        self.excluded_by_user = []
        self.vars = {}
        self.transient(parent)
        self.grab_set()
        bring_window_to_front(self)
        tk.Label(self, text="Select the tabs you want to convert to PDF:", font=("Arial", 10, "bold")).pack(pady=10, padx=10, anchor="w")
        tk.Label(self, text="All tabs are selected by default. Uncheck any tab you do not want converted.", font=("Arial", 8, "italic"), fg="gray", wraplength=420, justify=tk.LEFT).pack(padx=10, anchor="w")
        sf = ScrollableFrame(self)
        sf.pack(fill="both", expand=True, padx=10, pady=5)
        inner = sf.scrollable_frame
        for file, sheets in excel_file_info.items():
            file_var = tk.BooleanVar(value=True)
            self.vars[file] = {"var": file_var, "sheets": {}}
            tk.Checkbutton(inner, text=file, variable=file_var, font=("Arial", 9, "bold"), command=lambda f=file: self.toggle_file(f)).pack(anchor="w", padx=5, pady=(5, 0))
            for sheet in sheets:
                sheet_var = tk.BooleanVar(value=True)
                self.vars[file]["sheets"][sheet] = sheet_var
                tk.Checkbutton(inner, text=sheet, variable=sheet_var, command=lambda f=file: self.check_file_state(f)).pack(anchor="w", padx=25)
        tk.Label(self, text="Scroll down for more files/tabs if needed", font=("Arial", 8, "italic"), fg="gray").pack(pady=(5, 0))
        tk.Button(self, text="Confirm Selection", width=20, command=self.on_confirm).pack(pady=10)

    def toggle_file(self, file):
        state = self.vars[file]["var"].get()
        for sheet_var in self.vars[file]["sheets"].values():
            sheet_var.set(state)

    def check_file_state(self, file):
        self.vars[file]["var"].set(all(var.get() for var in self.vars[file]["sheets"].values()))

    def on_confirm(self):
        pending_result = {}
        pending_excluded = []
        for file, data in self.vars.items():
            selected = [sheet for sheet, var in data["sheets"].items() if var.get()]
            if selected:
                pending_result[file] = selected
            else:
                pending_excluded.append(file)
        if pending_excluded:
            names = "\n".join(f"- {f}" for f in pending_excluded)
            proceed = messagebox.askyesno(
                "Confirm Excluded Files",
                "No tabs are selected for the following file(s), so they will be completely "
                f"EXCLUDED from this conversion:\n\n{names}\n\nContinue?",
                parent=self,
            )
            if not proceed:
                return
        self.result = pending_result
        self.excluded_by_user = pending_excluded
        self.destroy()


class WordOptionsDialog(tk.Toplevel):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title("Global Word Options")
        self.geometry("600x700")
        self.minsize(600, 700)
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        bring_window_to_front(self)
        self.temp_export_mode_var = tk.StringVar(value=self.app.word_export_mode_var.get())
        self.temp_append_index_var = tk.BooleanVar(value=self.app.word_append_change_index_var.get())
        self.temp_ignore_hf_revisions_var = tk.BooleanVar(value=self.app.word_ignore_hf_revisions_var.get())
        tk.Label(self, text="Global Word Options", font=("Arial", 11, "bold")).pack(pady=(15, 5), padx=15, anchor="w")
        tk.Label(self, text="These settings apply to every Word file converted in the selected folder. Excel files and existing PDF files are not modified by these Word options.", font=("Arial", 9), wraplength=555, justify=tk.LEFT).pack(padx=15, pady=(0, 10), anchor="w")
        options_frame = tk.Frame(self)
        options_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(options_frame, text="Word Export Mode", font=("Arial", 10, "bold")).pack(anchor="w", pady=(0, 5))
        tk.Radiobutton(options_frame, text="Clean PDF (Current Behavior)", variable=self.temp_export_mode_var, value="clean", command=self._sync_states).pack(anchor="w")
        tk.Label(options_frame, text="Default. Comments and tracked changes are hidden.", fg="gray").pack(anchor="w", padx=25)
        tk.Radiobutton(options_frame, text="Revision Bars Only", variable=self.temp_export_mode_var, value="revision", command=self._sync_states).pack(anchor="w", pady=(8, 0))
        tk.Label(options_frame, text="Plot simple right-side change bars to track future versions better.", fg="gray").pack(anchor="w", padx=25)
        tk.Radiobutton(options_frame, text="Custom Markup Detail", variable=self.temp_export_mode_var, value="custom", command=self._sync_states).pack(anchor="w", pady=(8, 0))
        tk.Label(options_frame, text="Choose exactly what shows, like Word's own Track Changes display menu.", fg="gray").pack(anchor="w", padx=25)
        self.configure_markup_btn = tk.Button(options_frame, text="Configure Markup Detail...", command=self.open_markup_detail_dialog)
        self.configure_markup_btn.pack(anchor="w", padx=25, pady=(4, 0))
        self.index_cb = tk.Checkbutton(options_frame, text="If combining PDFs, append a Change Identification Report at the end", variable=self.temp_append_index_var)
        self.index_cb.pack(anchor="w", pady=(12, 3))
        self.ignore_hf_cb = tk.Checkbutton(options_frame, text="Ignore tracked changes in Word headers and footers during revision-bar export", variable=self.temp_ignore_hf_revisions_var)
        self.ignore_hf_cb.pack(anchor="w", pady=(3, 3))
        note = (
            "Note: Revision Bars Only and Custom Markup Detail both export from a temporary copy of each Word "
            "file -- hidden comments/changes are removed or accepted in that temporary copy only, and your "
            "original file is never modified or saved over. Clean PDF exports directly from the original file "
            "(read-only), with no temporary copy needed. The Change Identification Report above only applies "
            "to Revision Bars Only."
        )
        tk.Label(self, text=note, font=("Arial", 8, "italic"), fg="firebrick", wraplength=555, justify=tk.LEFT).pack(padx=20, pady=(10, 5), anchor="w")
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="Apply Options", width=18, command=self.apply_options).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="Cancel", width=12, command=self.destroy).pack(side=tk.LEFT, padx=8)
        self._sync_states()

    def _sync_states(self):
        mode = self.temp_export_mode_var.get()
        revision_enabled = mode == "revision"
        self.index_cb.config(state=tk.NORMAL if revision_enabled else tk.DISABLED)
        self.ignore_hf_cb.config(state=tk.NORMAL if revision_enabled else tk.DISABLED)
        self.configure_markup_btn.config(state=tk.NORMAL if mode == "custom" else tk.DISABLED)

    def open_markup_detail_dialog(self):
        MarkupDetailDialog(self, self.app)

    def apply_options(self):
        self.app.word_export_mode_var.set(self.temp_export_mode_var.get())
        self.app.word_append_change_index_var.set(self.temp_append_index_var.get())
        self.app.word_ignore_hf_revisions_var.set(self.temp_ignore_hf_revisions_var.get())
        mode_value = self.app.word_export_mode_var.get()
        mode = {"revision": "Revision Bars Only", "custom": "Custom Markup Detail"}.get(mode_value, "Clean PDF")
        index = "with Change Identification Report" if self.app.word_append_change_index_var.get() else "without Change Identification Report"
        hf = "header/footer revisions ignored" if self.app.word_ignore_hf_revisions_var.get() else "header/footer revisions included"
        if mode_value == "custom":
            shown = []
            if self.app.custom_show_comments_var.get():
                shown.append("Comments")
            if self.app.custom_show_ins_del_var.get():
                shown.append("Insertions/Deletions")
            if self.app.custom_show_format_changes_var.get():
                shown.append("Formatting")
            summary = ", ".join(shown) if shown else "nothing (fully clean)"
            self.app.status_label.config(text=f"Word Options Applied: {mode} (showing: {summary}).")
        else:
            self.app.status_label.config(text=f"Word Options Applied: {mode} ({index}; {hf}).")
        self.destroy()


class MarkupDetailDialog(tk.Toplevel):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title("Configure Custom Markup Detail")
        self.geometry("480x420")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        bring_window_to_front(self)
        self.temp_show_comments_var = tk.BooleanVar(value=self.app.custom_show_comments_var.get())
        self.temp_show_ins_del_var = tk.BooleanVar(value=self.app.custom_show_ins_del_var.get())
        self.temp_show_format_var = tk.BooleanVar(value=self.app.custom_show_format_changes_var.get())

        tk.Label(self, text="Show Markup", font=("Arial", 10, "bold")).pack(pady=(15, 5), padx=15, anchor="w")
        tk.Label(self, text="Choose which categories of tracked changes appear in the PDF. Anything unchecked is treated as accepted (or removed, for comments) before export.", font=("Arial", 8, "italic"), fg="gray", wraplength=440, justify=tk.LEFT).pack(padx=15, anchor="w")
        show_frame = tk.Frame(self)
        show_frame.pack(fill=tk.X, padx=25, pady=10)
        self.comments_cb = tk.Checkbutton(show_frame, text="Comments", variable=self.temp_show_comments_var)
        self.comments_cb.pack(anchor="w")
        self.ins_del_cb = tk.Checkbutton(show_frame, text="Insertions and Deletions", variable=self.temp_show_ins_del_var)
        self.ins_del_cb.pack(anchor="w")
        self.format_cb = tk.Checkbutton(show_frame, text="Formatting Changes", variable=self.temp_show_format_var)
        self.format_cb.pack(anchor="w")

        note = (
            "Note: like Revision Bars Only, this mode exports from a temporary copy of each Word file. "
            "Hidden comments are removed from the temporary copy only; hidden insertions, deletions, and "
            "formatting changes are accepted in the temporary copy only. Your original Word file is never "
            "modified or saved over."
        )
        tk.Label(self, text=note, font=("Arial", 8, "italic"), fg="firebrick", wraplength=440, justify=tk.LEFT).pack(padx=20, pady=(10, 5), anchor="w")

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="Apply", width=14, command=self.apply_and_close).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="Cancel", width=12, command=self.destroy).pack(side=tk.LEFT, padx=8)

    def apply_and_close(self):
        self.app.custom_show_comments_var.set(self.temp_show_comments_var.get())
        self.app.custom_show_ins_del_var.set(self.temp_show_ins_del_var.get())
        self.app.custom_show_format_changes_var.set(self.temp_show_format_var.get())
        self.destroy()


class ReorderDialog(simpledialog.Dialog):
    def __init__(self, parent, file_list):
        self.file_list = file_list
        self.result_list = None
        self._move_job = None
        flash_window(parent)
        super().__init__(parent, title="Reorder Files for Combined PDF")

    def body(self, master):
        controls = tk.Frame(master)
        controls.pack(pady=5, padx=5, fill=tk.X)
        tk.Label(controls, text="Sort Method:").grid(row=0, column=0, padx=(0, 5), sticky="w")
        self.sort_method_var = tk.StringVar()
        self.sort_combo = ttk.Combobox(controls, textvariable=self.sort_method_var, values=["Revision Sort", "Natural Sort (Default)", "Standard Sort (A-Z)"], state="readonly", width=25)
        self.sort_combo.grid(row=0, column=1, sticky="ew")
        self.sort_combo.current(1)
        self.sort_combo.bind("<<ComboboxSelected>>", self.apply_view_rules)
        self.move_covers_var = tk.BooleanVar()
        tk.Checkbutton(controls, text="Auto-move Covers to Ends", variable=self.move_covers_var, command=self.apply_view_rules).grid(row=1, column=0, columnspan=3, pady=(5, 0), sticky="w")
        controls.columnconfigure(1, weight=1)
        list_frame = tk.Frame(master)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, selectmode=tk.SINGLE, width=80, height=25)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)
        self.apply_view_rules()
        btn_frame = tk.Frame(master)
        btn_frame.pack(pady=10)
        up = tk.Button(btn_frame, text="Move Up", width=15)
        up.pack(side=tk.LEFT, padx=5)
        up.bind("<ButtonPress-1>", lambda e: self.start_move("up"))
        up.bind("<ButtonRelease-1>", lambda e: self.stop_move())
        down = tk.Button(btn_frame, text="Move Down", width=15)
        down.pack(side=tk.LEFT, padx=5)
        down.bind("<ButtonPress-1>", lambda e: self.start_move("down"))
        down.bind("<ButtonRelease-1>", lambda e: self.stop_move())
        return self.listbox

    def apply_view_rules(self, event=None):
        method = self.sort_method_var.get()
        sort_key = revision_aware_sort_key if "Revision" in method else natural_sort_key if "Natural" in method else standard_sort_key
        self.file_list.sort(key=sort_key)
        if self.move_covers_var.get():
            fronts, backs, mids = [], [], []
            for f in self.file_list:
                name = os.path.basename(f).lower()
                if "front cover" in name:
                    fronts.append(f)
                elif "back cover" in name:
                    backs.append(f)
                else:
                    mids.append(f)
            self.file_list = fronts + mids + backs
        self._update_listbox_ui()

    def _update_listbox_ui(self):
        sel = self.listbox.curselection()
        self.listbox.delete(0, tk.END)
        for path in self.file_list:
            self.listbox.insert(tk.END, os.path.basename(path))
        if sel and sel[0] < self.listbox.size():
            self.listbox.selection_set(sel[0])

    def start_move(self, direction):
        if self._move_job:
            self.after_cancel(self._move_job)
        if direction == "up":
            self.move_up()
        else:
            self.move_down()
        self._move_job = self.after(100, lambda: self.start_move(direction))

    def stop_move(self):
        if self._move_job:
            self.after_cancel(self._move_job)
        self._move_job = None

    def move_up(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] == 0:
            return
        idx = sel[0]
        self.file_list.insert(idx - 1, self.file_list.pop(idx))
        self._update_listbox_ui()
        self.listbox.selection_set(idx - 1)
        self.listbox.see(idx - 1)

    def move_down(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] == self.listbox.size() - 1:
            return
        idx = sel[0]
        self.file_list.insert(idx + 1, self.file_list.pop(idx))
        self._update_listbox_ui()
        self.listbox.selection_set(idx + 1)
        self.listbox.see(idx + 1)

    def apply(self):
        self.result_list = self.file_list


class BatchConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("680x760")
        self.root.minsize(680, 680)
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        self.source_folder = ""
        self.all_files = []
        self.excel_tabs_to_convert = {}
        self.file_ledger = {}
        self.word_export_mode_var = tk.StringVar(value="clean")
        self.word_append_change_index_var = tk.BooleanVar(value=True)
        self.word_ignore_hf_revisions_var = tk.BooleanVar(value=False)
        self.custom_show_comments_var = tk.BooleanVar(value=True)
        self.custom_show_ins_del_var = tk.BooleanVar(value=True)
        self.custom_show_format_changes_var = tk.BooleanVar(value=True)
        self._word_original_options = {}
        self.start_time = None
        self.created_pdf_count = 0
        self.expected_pdf_count = 0
        self.current_file_index = 0
        self.total_files_to_convert = 0
        top = tk.Frame(self.root)
        top.pack(pady=10, padx=10, fill=tk.X)
        self.select_btn = tk.Button(top, text="Select Folder with Word/Excel/PDF files", command=self.select_folder)
        self.select_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))
        self.word_options_btn = tk.Button(top, text="Word Options", width=14, command=self.open_word_options_dialog)
        self.word_options_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.help_btn = tk.Button(top, text="Help", width=10, command=self.show_help)
        self.help_btn.pack(side=tk.LEFT)
        list_frame = tk.Frame(self.root)
        list_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        tk.Label(list_frame, text="Files to be Processed:", anchor="w").pack(fill=tk.X)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
        self.file_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        status_box = tk.LabelFrame(self.root, text="Conversion Status", padx=8, pady=6)
        status_box.pack(padx=10, pady=(0, 8), fill=tk.X)
        self.status_current_file = tk.Label(status_box, text="Current File: None", anchor="w")
        self.status_current_file.pack(fill=tk.X)
        self.status_progress = tk.Label(status_box, text="Progress: 0 / 0 files", anchor="w")
        self.status_progress.pack(fill=tk.X)
        self.status_mode = tk.Label(status_box, text="Export Mode: Clean PDF", anchor="w")
        self.status_mode.pack(fill=tk.X)
        self.status_pdf_count = tk.Label(status_box, text="Expected PDFs: 0 | Created PDFs: 0", anchor="w")
        self.status_pdf_count.pack(fill=tk.X)
        self.status_elapsed = tk.Label(status_box, text="Elapsed Time: 00:00", anchor="w")
        self.status_elapsed.pack(fill=tk.X)
        bottom = tk.Frame(self.root)
        bottom.pack(pady=10, padx=10, fill=tk.X)
        self.status_label = tk.Label(bottom, text="Please select a folder to begin.", wraplength=640)
        self.status_label.pack(fill=tk.X)
        self.progress_bar = ttk.Progressbar(bottom, length=300, mode="determinate")
        self.progress_bar.pack(pady=5, fill=tk.X)
        self.precheck_btn = tk.Button(bottom, text="Optional checks for contracts", state=tk.DISABLED, command=self.open_checks_dialog)
        self.precheck_btn.pack(pady=(5, 15), fill=tk.X)
        self.start_btn = tk.Button(bottom, text="Start Conversion & Combine", state=tk.DISABLED, command=self.run_conversion_and_combine, height=2, font=("Arial", 9, "bold"))
        self.start_btn.pack(pady=(0, 5), fill=tk.X)

    def update_elapsed_label(self):
        if self.start_time is None:
            return
        elapsed = int(time.time() - self.start_time)
        mins, secs = divmod(elapsed, 60)
        hours, mins = divmod(mins, 60)
        text = f"Elapsed Time: {hours:02d}:{mins:02d}:{secs:02d}" if hours else f"Elapsed Time: {mins:02d}:{secs:02d}"
        self.status_elapsed.config(text=text)
        self.root.update_idletasks()

    def update_status_panel(self, current_file=None):
        mode = {"revision": "Revision Bars Only", "custom": "Custom Markup Detail"}.get(self.word_export_mode_var.get(), "Clean PDF")
        self.status_mode.config(text=f"Export Mode: {mode}")
        self.status_pdf_count.config(text=f"Expected PDFs: {self.expected_pdf_count} | Created PDFs: {self.created_pdf_count}")
        self.status_progress.config(text=f"Progress: {self.current_file_index} / {self.total_files_to_convert} files")
        if current_file is not None:
            self.status_current_file.config(text=f"Current File: {current_file}")
        self.update_elapsed_label()

    def open_word_options_dialog(self):
        WordOptionsDialog(self.root, self)

    def open_checks_dialog(self):
        messagebox.showinfo("Optional Checks", "Optional contract checks are not changed by the revision-bar workflow. Use your last full internal build if you need the legacy pre-check engine.")

    def show_help(self):
        help_text = (
            f"BMcD Office to PDF Converter\nVersion {VERSION}, {VERSION_DATE}\nAuthor: jgutierrez2@burnsmcd.com\n\n"
            "Converts Word and Excel files to PDF and combines them.\n\n"
            "Default Word Behavior:\n- Comments are hidden.\n- Tracked changes are hidden.\n- Clean PDFs are generated by default.\n\n"
            "Optional Revision-Bar Mode:\n- User may select Revision Bars Only in Word Options.\n- This plots simple right-side change bars to help track future versions.\n- Comments are removed from temporary export copies only.\n- A Change Identification Report may be appended when combining PDFs.\n\n"
            "Optional Custom Markup Detail Mode:\n- User may select Custom Markup Detail in Word Options, then Configure Markup Detail.\n- Independently show or hide Comments, Insertions and Deletions, and Formatting Changes in the PDF output.\n- Like Revision Bars Only, this exports from a temporary copy of each Word file. Hidden comments are removed and hidden changes are accepted in that temporary copy only -- the original file is never modified or saved over.\n- Does not participate in the Change Identification Report (that report is specific to Revision Bars Only).\n\n"
            "Header/Footer Revision Option:\n- User may ignore tracked changes in headers and footers during revision-bar export.\n- This helps prevent title block or footer-only edits from creating change bars.\n\n"
            "Revision Baseline Workflow:\n- After a successful Revision-Bar export, the tool can establish a new revision baseline.\n- It creates backup copies first.\n- It then accepts tracked changes in the local Word files and saves them.\n- Remember to re-upload the new saved Word files to SharePoint.\n\n"
            "Conversion Completeness Check:\n- Every file found in the source folder is tracked from selection through completion.\n- If anything does not end up converted, included, or explicitly excluded, a detailed report is shown before PDFs are combined.\n- A SHA-256 checksum of every source file and every PDF it produced is recorded in the summary for audit purposes.\n\n"
            "Conversion Summary:\n- A Conversion_Summary.txt report is saved in the Converted_PDFs folder."
        )
        messagebox.showinfo("Help", help_text)

    def select_folder(self):
        folder = filedialog.askdirectory(title="Select Folder")
        if not folder:
            return
        self.source_folder = folder
        self.status_label.config(text=f"Folder Selected: {self.source_folder}")
        self.file_listbox.delete(0, tk.END)
        self.start_btn.config(state=tk.DISABLED)
        self.precheck_btn.config(state=tk.DISABLED)
        self.excel_tabs_to_convert = {}
        self.file_ledger = {}
        valid_extensions = (".docx", ".doc", ".docm", ".xlsx", ".xls", ".xlsm", ".pdf")
        raw_listing = sorted(os.listdir(self.source_folder), key=natural_sort_key)
        self.build_initial_file_ledger(raw_listing, valid_extensions)
        self.all_files = sorted([f for f in raw_listing if f.lower().endswith(valid_extensions) and not f.startswith("~$")], key=natural_sort_key)
        if not self.all_files:
            messagebox.showwarning("No Files", "No compatible Word, Excel, or PDF files found.")
            return
        excel_files = [f for f in self.all_files if f.lower().endswith((".xlsx", ".xls", ".xlsm"))]
        if excel_files:
            self.analyze_and_prompt_excel_tabs(excel_files)
        for filename in self.all_files:
            self.file_listbox.insert(tk.END, filename)
            if filename in self.excel_tabs_to_convert:
                for tab in self.excel_tabs_to_convert[filename]:
                    self.file_listbox.insert(tk.END, f"      -> {tab}")
        self.start_btn.config(state=tk.NORMAL)
        self.precheck_btn.config(state=tk.NORMAL)

    def build_initial_file_ledger(self, raw_listing, valid_extensions):
        for name in raw_listing:
            full_path = os.path.join(self.source_folder, name)
            if os.path.isdir(full_path):
                continue
            lower = name.lower()
            entry = {
                "full_path": full_path,
                "status": STATUS_PENDING,
                "reason": "",
                "output_paths": [],
                "sha256_source": None,
                "sha256_outputs": {},
            }
            if name.startswith("~$"):
                entry["status"] = STATUS_EXCLUDED_NOT_APPLICABLE
                entry["reason"] = "Lock file"
            elif not lower.endswith(valid_extensions):
                entry["status"] = STATUS_EXCLUDED_NOT_APPLICABLE
                entry["reason"] = "Unsupported file type"
            else:
                entry["sha256_source"] = compute_sha256(full_path)
            self.file_ledger[name] = entry

    def analyze_and_prompt_excel_tabs(self, excel_files):
        self.status_label.config(text="Analyzing Excel tabs...")
        self.root.update_idletasks()
        excel_app = None
        info = {}
        unreadable = []
        try:
            excel_app = self.create_excel_application()
            for file in excel_files:
                path = os.path.abspath(os.path.join(self.source_folder, file))
                try:
                    wb = retry_com_call(lambda: excel_app.Workbooks.Open(path, ReadOnly=True), f"open {file}")
                    info[file] = [sheet.Name for sheet in wb.Sheets]
                    wb.Close(False)
                except Exception as e:
                    unreadable.append((file, str(e)))
                    entry = self.file_ledger.get(file)
                    if entry is not None:
                        entry["status"] = STATUS_EXCLUDED_UNREADABLE
                        entry["reason"] = str(e)
        finally:
            if excel_app:
                excel_app.Quit()
        if unreadable:
            names = "\n".join(f"- {f}: {reason}" for f, reason in unreadable)
            messagebox.showwarning(
                "Excel Files Could Not Be Read",
                "The following Excel file(s) could not be opened and will be EXCLUDED from this "
                f"conversion:\n\n{names}",
            )
        if info:
            dialog = ExcelTabSelectionDialog(self.root, info)
            self.root.wait_window(dialog)
            self.excel_tabs_to_convert = dialog.result
            for file in dialog.excluded_by_user:
                entry = self.file_ledger.get(file)
                if entry is not None:
                    entry["status"] = STATUS_EXCLUDED_USER
                    entry["reason"] = "No tabs selected"
            self.all_files = [f for f in self.all_files if f not in excel_files or f in self.excel_tabs_to_convert]
        else:
            self.all_files = [f for f in self.all_files if f not in excel_files]
        self.status_label.config(text=f"Folder Selected: {self.source_folder}")

    def calculate_expected_converted_pdf_count(self):
        count = 0
        for filename in self.all_files:
            lower = filename.lower()
            if lower.endswith((".docx", ".doc", ".docm")):
                count += 1
            elif lower.endswith((".xlsx", ".xls", ".xlsm")):
                count += len(self.excel_tabs_to_convert.get(filename, []))
        return count

    def record_ledger_output(self, filename, output_path):
        entry = self.file_ledger.get(filename)
        if entry is None:
            return
        entry["output_paths"].append(output_path)
        entry["sha256_outputs"][os.path.basename(output_path)] = compute_sha256(output_path)

    def reconcile_file_ledger(self):
        buckets = {}
        for filename, entry in self.file_ledger.items():
            buckets.setdefault(entry["status"], []).append(filename)
        return buckets

    def format_ledger_report(self, buckets):
        order = [
            (STATUS_CONVERTED, "Converted"),
            (STATUS_INCLUDED_EXISTING_PDF, "Included (already a PDF)"),
            (STATUS_EXCLUDED_USER, "Excluded by user"),
            (STATUS_EXCLUDED_UNREADABLE, "Excluded (could not be read)"),
            (STATUS_FAILED, "Failed"),
            (STATUS_PENDING, "UNACCOUNTED FOR"),
        ]
        lines = []
        for status, label in order:
            names = sorted(buckets.get(status, []))
            if not names:
                continue
            lines.append(f"{label} ({len(names)}):")
            for name in names:
                reason = self.file_ledger.get(name, {}).get("reason", "")
                suffix = f" - {reason}" if reason else ""
                lines.append(f"  - {name}{suffix}")
        unaccounted = buckets.get(STATUS_PENDING, [])
        lines.append("")
        lines.append("RECONCILED: OK" if not unaccounted else f"RECONCILED: MISMATCH -- {len(unaccounted)} file(s) unaccounted for")
        return "\n".join(lines)

    def show_reconciliation_dialog(self, ledger_buckets):
        dialog = tk.Toplevel(self.root)
        dialog.title("Conversion Completeness Report")
        dialog.geometry("560x480")
        dialog.transient(self.root)
        dialog.grab_set()
        bring_window_to_front(dialog)
        tk.Label(dialog, text="Every file found in the source folder, and what happened to it:", font=("Arial", 10, "bold"), wraplength=530, justify=tk.LEFT).pack(padx=10, pady=(10, 5), anchor="w")
        sf = ScrollableFrame(dialog)
        sf.pack(fill="both", expand=True, padx=10, pady=5)
        text_widget = tk.Text(sf.scrollable_frame, wrap="word", width=68, height=24, borderwidth=0)
        text_widget.insert("1.0", self.format_ledger_report(ledger_buckets))
        text_widget.config(state=tk.DISABLED)
        text_widget.pack(fill="both", expand=True)
        tk.Button(dialog, text="Close", width=14, command=dialog.destroy).pack(pady=10)
        dialog.wait_window()

    def wait_for_file_ready(self, file_path, attempts=20, delay=0.25):
        last_size = -1
        stable = 0
        for _ in range(attempts):
            if os.path.exists(file_path):
                try:
                    size = os.path.getsize(file_path)
                    if size > 0 and size == last_size:
                        stable += 1
                    else:
                        stable = 0
                    if stable >= 2:
                        return True
                    last_size = size
                except Exception:
                    pass
            time.sleep(delay)
        return os.path.exists(file_path) and os.path.getsize(file_path) > 0

    def verify_pdf_created(self, output_path, source_label):
        if not self.wait_for_file_ready(output_path):
            raise RuntimeError(f"PDF was not created or was empty for {source_label}")

    def create_word_application(self):
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        self._safe_set_property(word, "AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE)
        return word

    def create_excel_application(self):
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        self._safe_set_property(excel, "AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE)
        return excel

    def _safe_get_property(self, obj, prop_name, default=None):
        try:
            return getattr(obj, prop_name)
        except Exception:
            return default

    def _safe_set_property(self, obj, prop_name, value):
        try:
            setattr(obj, prop_name, value)
            return True
        except Exception:
            return False

    def _safe_get_word_option(self, word, prop_name):
        return self._safe_get_property(word.Options, prop_name)

    def _safe_set_word_option(self, word, prop_name, value):
        self._safe_set_property(word.Options, prop_name, value)

    def capture_word_original_options(self, word):
        if self._word_original_options:
            return
        for prop in ["RevisedLinesMark", "InsertedTextMark", "InsertedTextColor", "DeletedTextMark", "DeletedTextColor", "RevisedPropertiesMark", "RevisedPropertiesColor", "CommentsColor", "RevisionsMode"]:
            self._word_original_options[prop] = self._safe_get_word_option(word, prop)

    def restore_word_original_options(self, word):
        for prop, val in self._word_original_options.items():
            if val is not None:
                self._safe_set_word_option(word, prop, val)
        self._word_original_options = {}

    def make_temp_word_copy(self, input_path, temp_root):
        os.makedirs(temp_root, exist_ok=True)
        base = os.path.basename(input_path)
        name, ext = os.path.splitext(base)
        safe = re.sub(r"[^A-Za-z0-9_. -]+", "_", name).strip() or "word_export"
        temp_path = os.path.join(temp_root, safe + ext)
        i = 1
        while os.path.exists(temp_path):
            temp_path = os.path.join(temp_root, f"{safe}_{i}{ext}")
            i += 1
        shutil.copy2(input_path, temp_path)
        if not os.path.exists(temp_path) or os.path.getsize(temp_path) == 0:
            raise RuntimeError(f"Temporary Word copy could not be created for {base}")
        return temp_path

    def delete_comments_from_temp_doc(self, doc):
        try:
            doc.DeleteAllComments()
            return
        except Exception:
            pass
        try:
            for i in range(doc.Comments.Count, 0, -1):
                try:
                    doc.Comments(i).Delete()
                except Exception:
                    pass
        except Exception:
            pass

    def accept_revisions_in_range(self, rng):
        try:
            while rng.Revisions.Count > 0:
                rng.Revisions(1).Accept()
            return
        except Exception:
            pass
        try:
            rng.Revisions.AcceptAll()
        except Exception:
            pass

    def accept_header_footer_revisions_in_temp_doc(self, doc):
        try:
            for sec in doc.Sections:
                for typ in WD_HEADER_FOOTER_TYPES:
                    try:
                        self.accept_revisions_in_range(sec.Headers(typ).Range)
                    except Exception:
                        pass
                    try:
                        self.accept_revisions_in_range(sec.Footers(typ).Range)
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            story_ranges = doc.StoryRanges
            for i in range(1, story_ranges.Count + 1):
                try:
                    rng = story_ranges(i)
                    if rng.StoryType in WD_STORY_TYPES_HEADER_FOOTER:
                        self.accept_revisions_in_range(rng)
                except Exception:
                    pass
        except Exception:
            pass

    def apply_custom_markup_filters(self, doc):
        if not self.custom_show_comments_var.get():
            self.delete_comments_from_temp_doc(doc)
        hide_types = set()
        if not self.custom_show_ins_del_var.get():
            hide_types.update({WD_REVISION_INSERT, WD_REVISION_DELETE})
        if not self.custom_show_format_changes_var.get():
            hide_types.add(WD_REVISION_FORMAT_CHANGE)
        if not hide_types:
            return
        # Accepting a revision removes it from the collection and can shift other indices, so
        # re-scan from the start after every accept rather than assuming positions stay stable.
        while True:
            target_index = None
            try:
                count = doc.Revisions.Count
            except Exception:
                break
            for i in range(1, count + 1):
                try:
                    if doc.Revisions(i).Type in hide_types:
                        target_index = i
                        break
                except Exception:
                    continue
            if target_index is None:
                break
            try:
                doc.Revisions(target_index).Accept()
            except Exception:
                break

    def apply_word_export_options(self, word):
        self.capture_word_original_options(word)
        try:
            word.ActiveWindow.View.Type = WD_WINDOW_VIEW_TYPE_PRINT
        except Exception:
            pass
        try:
            word.ActiveWindow.View.RevisionsView = WD_REVISIONS_VIEW_FINAL
        except Exception:
            pass
        mode = self.word_export_mode_var.get()
        if mode == "revision":
            try:
                word.ActiveWindow.View.ShowRevisionsAndComments = True
            except Exception:
                pass
            try:
                word.ActiveWindow.View.RevisionsMode = WD_REVISIONS_MODE_INLINE
            except Exception:
                pass
            try:
                word.ActiveDocument.ShowRevisions = True
            except Exception:
                pass
            self._safe_set_word_option(word, "RevisedLinesMark", WD_REVISED_LINES_MARK_RIGHT_BORDER)
            for prop in ["InsertedTextMark", "InsertedTextColor", "DeletedTextMark", "DeletedTextColor", "RevisedPropertiesMark", "RevisedPropertiesColor", "CommentsColor"]:
                self._safe_set_word_option(word, prop, 0)
            try:
                word.ActiveWindow.View.ShowComments = False
            except Exception:
                pass
            try:
                word.ActiveWindow.View.ShowFormatChanges = False
            except Exception:
                pass
        elif mode == "custom":
            # Verified live: RevisionsFilter.Markup and the individual View.Show* flags have no
            # effect on ExportAsFixedFormat output on this Word version -- only ShowRevisionsAndComments
            # (master on/off) and RevisionsMode (balloon vs. inline) actually change the export.
            # Balloon mode is required for deleted-but-not-accepted text to render at all; the real
            # per-category granularity comes from apply_custom_markup_filters mutating a temp copy
            # before export, not from these view-level properties.
            try:
                word.ActiveWindow.View.ShowRevisionsAndComments = True
            except Exception:
                pass
            try:
                word.ActiveWindow.View.RevisionsMode = 0
            except Exception:
                pass
            try:
                word.ActiveDocument.ShowRevisions = True
            except Exception:
                pass
        else:
            try:
                word.ActiveWindow.View.ShowRevisionsAndComments = False
            except Exception:
                pass
            self._safe_set_word_option(word, "RevisedLinesMark", WD_REVISED_LINES_MARK_NONE)

    def export_doc_to_pdf(self, doc, output_path):
        mode = self.word_export_mode_var.get()
        if mode in ("revision", "custom"):
            doc.ExportAsFixedFormat(
                OutputFileName=output_path,
                ExportFormat=WD_EXPORT_FORMAT_PDF,
                OpenAfterExport=False,
                OptimizeFor=0,
                Range=0,
                From=1,
                To=1,
                Item=WD_EXPORT_ITEM_DOCUMENT_WITH_MARKUP,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=1,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )
        else:
            doc.ExportAsFixedFormat(output_path, WD_EXPORT_FORMAT_PDF)

    def convert_word_file_to_pdf(self, word, input_path, output_path, temp_root):
        doc = None
        temp_path = None
        try:
            mode = self.word_export_mode_var.get()
            if mode in ("revision", "custom"):
                temp_path = self.make_temp_word_copy(input_path, temp_root)
                doc = retry_com_call(
                    lambda: word.Documents.Open(temp_path, ReadOnly=False, AddToRecentFiles=False, ConfirmConversions=False),
                    f"open {os.path.basename(input_path)}",
                )
                doc.Activate()
                if mode == "revision":
                    self.delete_comments_from_temp_doc(doc)
                    if self.word_ignore_hf_revisions_var.get():
                        self.accept_header_footer_revisions_in_temp_doc(doc)
                else:
                    self.apply_custom_markup_filters(doc)
                try:
                    doc.Save()
                    time.sleep(0.2)
                except Exception:
                    pass
            else:
                doc = retry_com_call(
                    lambda: word.Documents.Open(input_path, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False),
                    f"open {os.path.basename(input_path)}",
                )
                doc.Activate()
            self.apply_word_export_options(word)
            retry_com_call(lambda: self.export_doc_to_pdf(doc, output_path), f"export {os.path.basename(input_path)}")
            self.verify_pdf_created(output_path, os.path.basename(input_path))
        finally:
            if doc:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if temp_path:
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

    def detect_revision_bar_pages(self, pdf_path):
        pages = []
        try:
            pdf = pymupdf.open(pdf_path)
            for idx in range(pdf.page_count):
                page = pdf[idx]
                rect = page.rect
                right_band_left = rect.x1 - max(72, rect.width * 0.14)
                min_h = max(8, rect.height * 0.012)
                found = False
                for drawing in page.get_drawings():
                    r = drawing.get("rect")
                    if not r:
                        continue
                    if r.x0 >= right_band_left and abs(r.y1 - r.y0) >= min_h and abs(r.x1 - r.x0) <= 5:
                        found = True
                        break
                if found:
                    pages.append(idx + 1)
            pdf.close()
        except Exception as e:
            print(f"Revision bar scan failed for {pdf_path}: {e}")
        return pages

    def append_change_report_pages(self, merged_doc, change_entries, total_docs, total_pages):
        filtered = [entry for entry in change_entries if entry["global_pages"]]
        docs_with_changes = len(filtered)
        changed_pages_total = sum(len(set(entry["global_pages"])) for entry in filtered)
        generated = datetime.now().strftime("%B %d, %Y %I:%M %p")
        rows = [["Document", "Changed Pages", "Count"]]
        for entry in filtered:
            count = len(set(entry["global_pages"]))
            rows.append([entry["file"], compress_page_ranges(entry["global_pages"]), str(count)])

        margin = 40
        line_h = 14
        max_rows_first = 34
        max_rows_next = 45
        row_index = 1
        page_number = 0
        while True:
            page = merged_doc.new_page()
            page_number += 1
            y = margin
            page.insert_text((margin, y), "CHANGE IDENTIFICATION REPORT", fontsize=16, fontname="helv", color=(0, 0, 0))
            y += 28
            if page_number == 1:
                summary = [
                    f"Package Generated: {generated}",
                    "",
                    "EXECUTIVE SUMMARY",
                    f"Documents Reviewed: {total_docs}",
                    f"Documents With Changes: {docs_with_changes}",
                    f"Combined PDF Pages Reviewed: {total_pages}",
                    f"Changed Pages Detected: {changed_pages_total}",
                    "",
                    "DETAILED CHANGE TABLE",
                ]
                for s in summary:
                    page.insert_text((margin, y), s, fontsize=10 if s not in ["EXECUTIVE SUMMARY", "DETAILED CHANGE TABLE"] else 11, fontname="helv")
                    y += line_h
                rows_this_page = max_rows_first
            else:
                page.insert_text((margin, y), "DETAILED CHANGE TABLE, CONTINUED", fontsize=11, fontname="helv")
                y += 20
                rows_this_page = max_rows_next
            x_doc, x_pages, x_count = margin, 330, 540
            page.insert_text((x_doc, y), rows[0][0], fontsize=9, fontname="helv")
            page.insert_text((x_pages, y), rows[0][1], fontsize=9, fontname="helv")
            page.insert_text((x_count, y), rows[0][2], fontsize=9, fontname="helv")
            y += line_h
            page.draw_line((margin, y - 8), (page.rect.width - margin, y - 8), color=(0, 0, 0), width=0.5)
            rows_written = 0
            if len(rows) == 1:
                page.insert_text((x_doc, y), "No changed pages detected.", fontsize=8.5, fontname="helv")
                y += line_h
            while row_index < len(rows) and rows_written < rows_this_page:
                doc_name, pages_text, count_text = rows[row_index]
                if len(doc_name) > 42:
                    doc_name = doc_name[:39] + "..."
                if len(pages_text) > 30:
                    pages_text = pages_text[:27] + "..."
                page.insert_text((x_doc, y), doc_name, fontsize=8.5, fontname="helv")
                page.insert_text((x_pages, y), pages_text, fontsize=8.5, fontname="helv")
                page.insert_text((x_count, y), count_text, fontsize=8.5, fontname="helv")
                y += line_h
                row_index += 1
                rows_written += 1
            page.insert_text((margin, page.rect.height - 30), f"Change Identification Report | Page {page_number}", fontsize=8, fontname="helv")
            if row_index >= len(rows):
                break

    def process_merging(self, ordered_paths, output_folder):
        try:
            merged = pymupdf.open()
            toc, labels, current = [], [], 0
            change_entries = []
            append_report = self.word_export_mode_var.get() == "revision" and self.word_append_change_index_var.get()
            for pdf_path in ordered_paths:
                filename = os.path.basename(pdf_path)
                title = os.path.splitext(filename)[0]
                local_pages = self.detect_revision_bar_pages(pdf_path) if append_report else []
                with pymupdf.open(pdf_path) as src:
                    num_pages = src.page_count
                    toc.append([1, title, current + 1])
                    labels.append({"startpage": current, "prefix": title + "-", "style": "D", "firstpagenum": 1})
                    global_pages = [current + p for p in local_pages if 1 <= p <= num_pages]
                    if append_report:
                        change_entries.append({"file": filename, "global_pages": global_pages})
                    merged.insert_pdf(src)
                    current += num_pages
            if append_report:
                report_start_page = merged.page_count + 1
                toc.append([1, "Change Identification Report", report_start_page])
                labels.append({"startpage": report_start_page - 1, "prefix": "Change Identification Report-", "style": "D", "firstpagenum": 1})
                self.append_change_report_pages(merged, change_entries, len(ordered_paths), current)
                toc[-1] = [1, "Change Identification Report", report_start_page]
            merged.set_toc(toc)
            if hasattr(merged, "set_page_labels"):
                merged.set_page_labels(labels)
            combined_path = os.path.join(output_folder, "_combined.pdf")
            merged.save(combined_path, garbage=4, deflate=True)
            merged.close()
            flash_window(self.root)
            if messagebox.askyesno("Success!", "Files combined successfully!\nSaved in 'Converted_PDFs' folder.\n\nWould you like to open the output folder?"):
                os.startfile(output_folder)
            return True
        except Exception as e:
            messagebox.showerror("Merge Error", f"Could not combine files: {e}")
            return False

    def should_offer_revision_baseline(self, word_paths, failed):
        return self.word_export_mode_var.get() == "revision" and bool(word_paths) and not failed

    def offer_revision_baseline_update(self, word_paths, failed, output_folder):
        if not self.should_offer_revision_baseline(word_paths, failed):
            return
        flash_window(self.root)
        prompt = (
            "Revision-Bar PDF Export Complete\n\n"
            "Would you like to establish these local Word files as the new revision baseline now?\n\n"
            "If you choose Yes, the tool will first create backup copies, then accept all tracked changes in the original local Word files and save them.\n\n"
            "Benefits:\n- Future revision-bar exports will only show new changes made after this issue.\n- Old revision bars will not continue appearing in future packages.\n\n"
            "Important:\n- This modifies the original local Word files in this folder.\n- A backup folder will be created before any changes are accepted.\n- Remember to re-upload the new saved Word files to SharePoint after this step."
        )
        if messagebox.askyesno("Establish New Revision Baseline?", prompt):
            self.accept_all_changes_in_word_files(word_paths, output_folder)

    def backup_word_files_before_baseline(self, word_paths, output_folder):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_folder = os.path.join(output_folder, f"Word_Backups_Before_Baseline_{stamp}")
        os.makedirs(backup_folder, exist_ok=True)
        for path in word_paths:
            shutil.copy2(path, os.path.join(backup_folder, os.path.basename(path)))
        return backup_folder

    def accept_all_changes_in_word_files(self, word_paths, output_folder):
        word = None
        failed = []
        backup_folder = None
        try:
            backup_folder = self.backup_word_files_before_baseline(word_paths, output_folder)
            word = self.create_word_application()
            for path in word_paths:
                try:
                    doc = retry_com_call(
                        lambda: word.Documents.Open(path, ReadOnly=False, AddToRecentFiles=False, ConfirmConversions=False),
                        f"open {os.path.basename(path)}",
                    )
                    try:
                        doc.AcceptAllRevisions()
                    except Exception:
                        doc.Revisions.AcceptAll()
                    doc.Save()
                    doc.Close(False)
                except Exception as e:
                    failed.append(f"{os.path.basename(path)}: {e}")
        finally:
            if word:
                word.Quit()
        if failed:
            messagebox.showwarning("Some Files Could Not Be Baselined", "Backups were created here:\n" + f"{backup_folder}\n\n" + "The following files could not have tracked changes accepted:\n\n" + "\n".join(failed))
        else:
            messagebox.showinfo("New Revision Baseline Established", "All tracked changes were accepted in the converted Word files.\n\n" + f"Backup copies were saved here:\n{backup_folder}\n\n" + "Remember to re-upload the new saved Word files to SharePoint so the project files reflect this new baseline.")

    def write_conversion_summary(self, output_folder, expected, actual, word_paths, original_pdf_count, failed, combined_created, elapsed_seconds, ledger_buckets=None, combined_pdf_path=None):
        path = os.path.join(output_folder, "Conversion_Summary.txt")
        mode = {"revision": "Revision Bars Only", "custom": "Custom Markup Detail"}.get(self.word_export_mode_var.get(), "Clean PDF")
        elapsed = time.strftime("%H:%M:%S", time.gmtime(elapsed_seconds))
        lines = []
        lines.append(f"BMcD Office to PDF Converter v{VERSION}")
        lines.append(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"Source Folder: {self.source_folder}")
        lines.append("")
        lines.append(f"Export Mode: {mode}")
        if self.word_export_mode_var.get() == "custom":
            lines.append(f"Custom Markup Shows Comments: {'Yes' if self.custom_show_comments_var.get() else 'No'}")
            lines.append(f"Custom Markup Shows Insertions/Deletions: {'Yes' if self.custom_show_ins_del_var.get() else 'No'}")
            lines.append(f"Custom Markup Shows Formatting Changes: {'Yes' if self.custom_show_format_changes_var.get() else 'No'}")
        lines.append(f"Header/Footer Revisions Ignored: {'Yes' if self.word_ignore_hf_revisions_var.get() else 'No'}")
        lines.append(f"Change Identification Report Appended: {'Yes' if self.word_append_change_index_var.get() and self.word_export_mode_var.get() == 'revision' else 'No'}")
        lines.append("")
        lines.append(f"Word Files Converted: {len(word_paths)}")
        lines.append(f"Excel Tabs Converted: {sum(len(v) for v in self.excel_tabs_to_convert.values())}")
        lines.append(f"Existing PDFs Included: {original_pdf_count}")
        lines.append(f"Expected Converted PDFs: {expected}")
        lines.append(f"Actual Converted PDFs: {actual}")
        lines.append(f"Combined PDF Created: {'Yes' if combined_created else 'No'}")
        lines.append(f"Elapsed Time: {elapsed}")
        lines.append("")
        lines.append("Failures:")
        if failed:
            lines.extend([f"- {item}" for item in failed])
        else:
            lines.append("None")
        if ledger_buckets is not None:
            lines.append("")
            lines.append("COMPLETENESS RECONCILIATION")
            lines.append(self.format_ledger_report(ledger_buckets))
            lines.append("")
            lines.append("FILE INTEGRITY (SHA-256)")
            for filename, entry in sorted(self.file_ledger.items()):
                if entry["status"] == STATUS_EXCLUDED_NOT_APPLICABLE:
                    continue
                lines.append(f"{filename}")
                lines.append(f"  source: {entry['sha256_source'] or 'unavailable'}")
                for out_name, out_hash in entry["sha256_outputs"].items():
                    lines.append(f"  -> {out_name}: {out_hash or 'unavailable'}")
            if combined_pdf_path:
                combined_hash = compute_sha256(combined_pdf_path)
                lines.append(f"_combined.pdf: {combined_hash or 'unavailable'}")
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return path

    def run_conversion_and_combine(self):
        self.start_btn.config(state=tk.DISABLED)
        self.precheck_btn.config(state=tk.DISABLED)
        self.select_btn.config(state=tk.DISABLED)
        self.word_options_btn.config(state=tk.DISABLED)
        self.start_time = time.time()
        self.created_pdf_count = 0
        self.current_file_index = 0
        self.status_label.config(text="Processing... The application may be unresponsive during conversion.")
        self.root.update_idletasks()
        files_to_convert = [f for f in self.all_files if not f.lower().endswith(".pdf")]
        self.total_files_to_convert = len(files_to_convert)
        output_folder = os.path.join(self.source_folder, "Converted_PDFs")
        os.makedirs(output_folder, exist_ok=True)
        self.expected_pdf_count = self.calculate_expected_converted_pdf_count()
        self.progress_bar["maximum"] = len(files_to_convert)
        self.update_status_panel("None")
        new_pdf_paths, failed_files = [], []
        word_paths = []
        baseline_prompt_was_offered = False
        temp_word_export_folder = tempfile.mkdtemp(prefix="word_export_temp_", dir=output_folder)
        word, excel = None, None
        combined_created = False
        try:
            for i, filename in enumerate(files_to_convert):
                self.current_file_index = i + 1
                self.status_label.config(text=f"Converting: {filename}")
                self.progress_bar["value"] = i + 1
                self.update_status_panel(filename)
                ext = os.path.splitext(filename)[1].lower()
                input_path = os.path.abspath(os.path.join(self.source_folder, filename))
                try:
                    if ext in [".docx", ".doc", ".docm"]:
                        output = os.path.abspath(os.path.join(output_folder, os.path.splitext(filename)[0] + ".pdf"))
                        if not word:
                            word = self.create_word_application()
                        self.convert_word_file_to_pdf(word, input_path, output, temp_word_export_folder)
                        new_pdf_paths.append(output)
                        word_paths.append(input_path)
                        self.created_pdf_count += 1
                        self.record_ledger_output(filename, output)
                        entry = self.file_ledger.get(filename)
                        if entry is not None:
                            entry["status"] = STATUS_CONVERTED
                    elif ext in [".xlsx", ".xls", ".xlsm"]:
                        if not excel:
                            excel = self.create_excel_application()
                        wb = retry_com_call(lambda: excel.Workbooks.Open(input_path, ReadOnly=True), f"open {filename}")
                        sheets = self.excel_tabs_to_convert.get(filename, [])
                        sheet_failures = []
                        if sheets:
                            if len(sheets) == 1 and len(wb.Sheets) == 1:
                                sheet_name = sheets[0]
                                try:
                                    output = os.path.abspath(os.path.join(output_folder, os.path.splitext(filename)[0] + ".pdf"))
                                    wb.Sheets(sheet_name).Select()
                                    retry_com_call(lambda: wb.ActiveSheet.ExportAsFixedFormat(XL_FIXED_FORMAT_PDF, output), f"export {filename} - {sheet_name}")
                                    self.verify_pdf_created(output, f"{filename} - {sheet_name}")
                                    new_pdf_paths.append(output)
                                    self.created_pdf_count += 1
                                    self.record_ledger_output(filename, output)
                                except Exception as sheet_err:
                                    sheet_failures.append(f"{filename} - {sheet_name}: {sheet_err}")
                            else:
                                base = os.path.splitext(filename)[0]
                                for sheet_name in sheets:
                                    try:
                                        output = os.path.abspath(os.path.join(output_folder, f"{base} - {sheet_name}.pdf"))
                                        wb.Sheets(sheet_name).Select()
                                        retry_com_call(lambda: wb.ActiveSheet.ExportAsFixedFormat(XL_FIXED_FORMAT_PDF, output), f"export {filename} - {sheet_name}")
                                        self.verify_pdf_created(output, f"{filename} - {sheet_name}")
                                        new_pdf_paths.append(output)
                                        self.created_pdf_count += 1
                                        self.record_ledger_output(filename, output)
                                    except Exception as sheet_err:
                                        sheet_failures.append(f"{filename} - {sheet_name}: {sheet_err}")
                        wb.Close(False)
                        entry = self.file_ledger.get(filename)
                        if entry is not None:
                            if entry["output_paths"]:
                                entry["status"] = STATUS_CONVERTED
                                if sheet_failures:
                                    entry["reason"] = "Partial export: " + "; ".join(sheet_failures)
                            elif sheet_failures:
                                entry["status"] = STATUS_FAILED
                                entry["reason"] = "; ".join(sheet_failures)
                        if sheet_failures:
                            failed_files.extend(sheet_failures)
                    self.update_status_panel(filename)
                except Exception as e:
                    failed_files.append(f"{filename}: {e}")
                    entry = self.file_ledger.get(filename)
                    if entry is not None and entry["status"] != STATUS_CONVERTED:
                        entry["status"] = STATUS_FAILED
                        entry["reason"] = str(e)
            if word:
                self.restore_word_original_options(word)
                word.Quit()
            if excel:
                excel.Quit()
            try:
                if os.path.isdir(temp_word_export_folder):
                    shutil.rmtree(temp_word_export_folder)
            except Exception:
                pass
            self.update_status_panel("Conversion complete")
            actual = len(new_pdf_paths)
            original_pdfs = [os.path.join(self.source_folder, f) for f in self.all_files if f.lower().endswith(".pdf")]
            for f in self.all_files:
                if f.lower().endswith(".pdf"):
                    entry = self.file_ledger.get(f)
                    if entry is not None:
                        entry["status"] = STATUS_INCLUDED_EXISTING_PDF
                        entry["output_paths"] = [entry["full_path"]]
                        entry["sha256_outputs"] = {f: entry["sha256_source"]}
            ledger_buckets = self.reconcile_file_ledger()
            unaccounted = ledger_buckets.get(STATUS_PENDING, [])
            if actual != self.expected_pdf_count or unaccounted:
                flash_window(self.root)
                self.show_reconciliation_dialog(ledger_buckets)
                msg = (
                    f"Expected converted PDFs: {self.expected_pdf_count}\nActual converted PDFs: {actual}\n\n"
                    "This usually means one or more Word/Excel exports failed, or a file could not be accounted "
                    "for (see the detail window). Review the conversion issues before issuing.\n\n"
                    "Continue to combine the PDFs that were created?"
                )
                if not messagebox.askyesno("Completeness Check", msg):
                    if failed_files:
                        messagebox.showwarning("Conversion Issues", "The following files failed to convert:\n\n" + "\n".join(failed_files))
                    self.write_conversion_summary(output_folder, self.expected_pdf_count, actual, word_paths, len(original_pdfs), failed_files, False, int(time.time() - self.start_time), ledger_buckets=ledger_buckets)
                    messagebox.showinfo("Process Stopped", "Conversion stopped because the completeness check failed.")
                    return
            all_pdf_paths = new_pdf_paths + original_pdfs
            all_pdf_paths.sort(key=natural_sort_key)
            proceed, dialog_ok = False, False
            if len(all_pdf_paths) > 1:
                flash_window(self.root)
                if messagebox.askyesno("Combine Files?", f"Conversion complete.\nFound {len(all_pdf_paths)} total PDFs.\n\nCombine them into a single PDF?"):
                    proceed = True
            elif all_pdf_paths:
                messagebox.showinfo("Process Complete", "Conversion finished. Only one PDF exists, so no combination was performed.")
            else:
                messagebox.showinfo("Process Complete", "Conversion finished. No PDFs were found to combine.")
            combined_pdf_path = None
            if proceed:
                dialog = ReorderDialog(self.root, all_pdf_paths)
                if dialog.result_list:
                    dialog_ok = True
                    combined_created = self.process_merging(dialog.result_list, output_folder)
                    if combined_created:
                        combined_pdf_path = os.path.join(output_folder, "_combined.pdf")
                        self.offer_revision_baseline_update(word_paths, failed_files, output_folder)
                        baseline_prompt_was_offered = True
            if failed_files:
                messagebox.showwarning("Conversion Issues", "The following files failed to convert:\n\n" + "\n".join(failed_files))
            if not baseline_prompt_was_offered:
                self.offer_revision_baseline_update(word_paths, failed_files, output_folder)
            ledger_buckets = self.reconcile_file_ledger()
            summary_path = self.write_conversion_summary(
                output_folder,
                self.expected_pdf_count,
                actual,
                word_paths,
                len(original_pdfs),
                failed_files,
                combined_created,
                int(time.time() - self.start_time),
                ledger_buckets=ledger_buckets,
                combined_pdf_path=combined_pdf_path,
            )
            flash_window(self.root)
            if not (proceed and dialog_ok):
                messagebox.showinfo("Process Complete", f"File conversion is finished.\n\nConversion summary saved to:\n{summary_path}")
        except Exception as e:
            if word:
                try:
                    self.restore_word_original_options(word)
                except Exception:
                    pass
                word.Quit()
            if excel:
                excel.Quit()
            try:
                if os.path.isdir(temp_word_export_folder):
                    shutil.rmtree(temp_word_export_folder)
            except Exception:
                pass
            messagebox.showerror("Critical Error", f"Engine failure: {e}")
        finally:
            sys.exit(0)


def show_initial_warning(root):
    dialog = tk.Toplevel(root)
    dialog.title("Warning")
    dialog.resizable(False, False)
    dialog.update_idletasks()
    width, height = 450, 120
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f"{width}x{height}+{x}+{y}")
    tk.Label(dialog, text="THIS TOOL REQUIRES WORD AND EXCEL TO BE CLOSED, OTHERWISE IT WILL KILL ANY OPEN PROCESSES.\nPLEASE SAVE YOUR WORK IN THOSE PROGRAMS BEFORE PROCEEDING.", wraplength=430, justify=tk.CENTER).pack(pady=10, padx=10)
    choice = tk.BooleanVar(value=False)

    def cont():
        choice.set(True)
        dialog.destroy()

    def exit_app():
        choice.set(False)
        dialog.destroy()

    frame = tk.Frame(dialog)
    frame.pack(pady=10)
    tk.Button(frame, text="Continue", width=10, command=cont).pack(side=tk.LEFT, padx=10)
    tk.Button(frame, text="Exit", width=10, command=exit_app).pack(side=tk.LEFT, padx=10)
    dialog.protocol("WM_DELETE_WINDOW", exit_app)
    dialog.grab_set()
    bring_window_to_front(dialog)
    dialog.wait_window()
    return choice.get()


if __name__ == "__main__":
    root = tk.Tk()
    root.title(APP_TITLE)
    root.withdraw()
    if show_initial_warning(root):
        root.deiconify()
        app = BatchConverterApp(root)
        root.protocol("WM_DELETE_WINDOW", lambda: sys.exit(0))
        root.mainloop()
    else:
        sys.exit(0)
