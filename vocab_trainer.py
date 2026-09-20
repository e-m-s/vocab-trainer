#!/usr/bin/env python3
"""Vocabulary practice app backed by an OpenDocument Spreadsheet (.ods).

Run with:  python vocab_trainer.py
Requires:  pandas, odfpy  (pip install pandas odfpy)
"""

import json
import os
import random
import shutil
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, ttk

import pandas as pd

COOLDOWN = timedelta(minutes=30)
CREATE_NEW = "<create new column…>"
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".vocab_trainer_config.json")
COLUMN_ROLES = ("native", "foreign", "fails", "last", "attempts")


# ---------------------------------------------------------------------------
# Data helpers
# ---------------------------------------------------------------------------

def parse_last(value):
    """Parse a 'last attempt' cell value into a datetime, or None if unset."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, float) and pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        try:
            parsed = pd.to_datetime(text)
        except (ValueError, TypeError):
            return None
        return None if pd.isna(parsed) else parsed.to_pydatetime()


def format_last(dt):
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def split_alternatives(text):
    """Split a foreign-phrase cell on '|' into acceptable alternative answers."""
    parts = (p.strip() for p in str(text).split("|"))
    return [p for p in parts if p]


def to_int(value, default=0):
    try:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return default
        return int(float(value))
    except (ValueError, TypeError):
        return default


class ColumnMap:
    def __init__(self, native, foreign, fails, last, attempts):
        self.native = native
        self.foreign = foreign
        self.fails = fails
        self.last = last
        self.attempts = attempts


class VocabStore:
    """Holds every sheet of one .ods file and knows how to save them back."""

    def __init__(self, path):
        self.path = path
        self.sheets = pd.read_excel(path, sheet_name=None, engine="odf")
        self.backed_up = False

    @property
    def sheet_names(self):
        return list(self.sheets.keys())

    def columns(self, sheet_name):
        return list(self.sheets[sheet_name].columns)

    def backup(self):
        if self.backed_up:
            return
        backup_path = self.path + ".bak"
        if not os.path.exists(backup_path):
            shutil.copy2(self.path, backup_path)
        self.backed_up = True

    def save(self, sheet_name, df, path=None):
        """Write all sheets back to disk. Only cell data is preserved --
        formatting/styles from the original file are not round-tripped."""
        target = path or self.path
        self.sheets[sheet_name] = df
        if target == self.path:
            self.backup()
        root, ext = os.path.splitext(target)
        tmp_path = f"{root}.tmp{ext}"
        with pd.ExcelWriter(tmp_path, engine="odf") as writer:
            for name, sheet_df in self.sheets.items():
                sheet_df.to_excel(writer, sheet_name=name, index=False)
        os.replace(tmp_path, target)
        self.path = target


# ---------------------------------------------------------------------------
# Selection / scoring logic
# ---------------------------------------------------------------------------

def get_fails(df, idx, cols):
    return to_int(df.at[idx, cols.fails])


def get_attempts(df, idx, cols):
    return to_int(df.at[idx, cols.attempts])


def get_last(df, idx, cols):
    return parse_last(df.at[idx, cols.last])


def practice_indices(df, cols):
    """Rows usable for practice: both native and foreign text present."""
    def has_text(v):
        return v is not None and str(v).strip() != "" and str(v).lower() != "nan"

    return [
        i for i in df.index
        if has_text(df.at[i, cols.native]) and has_text(df.at[i, cols.foreign])
    ]


def pick_next_index(df, cols, now=None):
    """Pick the next phrase to ask.

    Rule 1: phrases with fails > 0 take priority, but are skipped if they
    were last asked less than COOLDOWN ago.
    Rule 2: among the eligible pool, prefer phrases tested longest ago, and
    among ties prefer lower total attempts. A little randomness is kept by
    choosing among the top few candidates rather than always the single best.
    """
    now = now or datetime.now()
    indices = practice_indices(df, cols)
    if not indices:
        return None

    due_fail = []
    for i in indices:
        if get_fails(df, i, cols) > 0:
            last = get_last(df, i, cols)
            if last is None or now - last >= COOLDOWN:
                due_fail.append(i)

    pool = due_fail if due_fail else indices

    def sort_key(i):
        last = get_last(df, i, cols)
        last_ts = last.timestamp() if last else -1.0
        return (last_ts, get_attempts(df, i, cols))

    ranked = sorted(pool, key=sort_key)
    frontrunners = ranked[: min(5, len(ranked))]
    return random.choice(frontrunners)


def record_attempt(df, idx, cols, correct, now=None):
    now = now or datetime.now()
    fails = get_fails(df, idx, cols)
    attempts = get_attempts(df, idx, cols) + 1
    if correct:
        fails = max(0, fails - 1)
    else:
        fails += 1
    df.at[idx, cols.fails] = fails
    df.at[idx, cols.attempts] = attempts
    df.at[idx, cols.last] = format_last(now)


# ---------------------------------------------------------------------------
# GUI: setup screen
# ---------------------------------------------------------------------------

class SetupFrame(ttk.Frame):
    ROLES = [
        ("native", "Native-language phrase column"),
        ("foreign", "Foreign-language phrase column"),
        ("fails", "Fail-count column"),
        ("last", "Last-attempt-date column"),
        ("attempts", "Total-attempts column"),
    ]
    GUESSES = {
        "native": {"native", "english"},
        "foreign": {"foreign", "target"},
        "fails": {"fails", "errors", "error", "failcount"},
        "last": {"last", "lastattempt", "lastdate", "lasttest", "lasttested"},
        "attempts": {"attempts", "tests", "total", "totalattempts"},
    }

    def __init__(self, master, app):
        super().__init__(master, padding=20)
        self.app = app
        self.role_vars = {}
        self.role_entries = {}

        ttk.Label(self, text="Vocabulary Trainer — Setup", font=("", 16, "bold")).grid(
            row=0, column=0, columnspan=3, pady=(0, 15), sticky="w"
        )

        ttk.Button(self, text="Open .ods file…", command=self.open_file).grid(row=1, column=0, sticky="w")
        self.file_label = ttk.Label(self, text="No file selected")
        self.file_label.grid(row=1, column=1, columnspan=2, sticky="w")

        ttk.Label(self, text="Sheet:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(self, textvariable=self.sheet_var, state="readonly", width=30)
        self.sheet_combo.grid(row=2, column=1, sticky="w", pady=(10, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)

        self.mapping_frame = ttk.Frame(self)
        self.mapping_frame.grid(row=3, column=0, columnspan=3, sticky="w", pady=15)

        self.start_button = ttk.Button(self, text="Start Practice", command=self.start_practice, state="disabled")
        self.start_button.grid(row=4, column=0, pady=15, sticky="w")

    def on_show(self):
        pass

    def open_file(self):
        path = filedialog.askopenfilename(
            title="Select vocabulary .ods file",
            filetypes=[("OpenDocument Spreadsheet", "*.ods")],
        )
        if not path:
            return
        try:
            store = VocabStore(path)
        except Exception as exc:
            messagebox.showerror("Could not open file", str(exc))
            return
        self.app.store = store
        self.file_label.config(text=os.path.basename(path))
        self.sheet_combo["values"] = store.sheet_names
        if store.sheet_names:
            self.sheet_var.set(store.sheet_names[0])
            self.build_mapping(store.sheet_names[0])

    def on_sheet_selected(self, event=None):
        sheet = self.sheet_var.get()
        if sheet:
            self.build_mapping(sheet)

    def build_mapping(self, sheet):
        for child in self.mapping_frame.winfo_children():
            child.destroy()
        self.role_vars.clear()
        self.role_entries.clear()

        columns = self.app.store.columns(sheet)

        for row, (role, label) in enumerate(self.ROLES):
            ttk.Label(self.mapping_frame, text=label + ":").grid(row=row, column=0, sticky="w", pady=3)
            values = list(columns)
            if role in ("fails", "last", "attempts"):
                values = values + [CREATE_NEW]

            var = tk.StringVar()
            combo = ttk.Combobox(self.mapping_frame, textvariable=var, values=values, state="readonly", width=28)
            for col in columns:
                normalized = str(col).strip().lower().replace(" ", "").replace("_", "")
                if normalized in self.GUESSES[role]:
                    var.set(col)
                    break
            combo.grid(row=row, column=1, sticky="w", padx=5)

            entry = ttk.Entry(self.mapping_frame, width=20)
            entry.grid(row=row, column=2, sticky="w")
            entry.grid_remove()

            def on_change(event, entry=entry, var=var):
                if var.get() == CREATE_NEW:
                    entry.grid()
                else:
                    entry.grid_remove()
                self.validate()

            combo.bind("<<ComboboxSelected>>", on_change)
            entry.bind("<KeyRelease>", lambda e: self.validate())

            self.role_vars[role] = var
            self.role_entries[role] = entry

        self.validate()

    def validate(self):
        native = self.role_vars["native"].get()
        foreign = self.role_vars["foreign"].get()
        ok = bool(native) and bool(foreign) and native != foreign
        for role in ("fails", "last", "attempts"):
            val = self.role_vars[role].get()
            if not val:
                ok = False
            elif val == CREATE_NEW and not self.role_entries[role].get().strip():
                ok = False
        self.start_button.config(state="normal" if ok else "disabled")

    def start_practice(self):
        sheet = self.sheet_var.get()
        store = self.app.store
        df = store.sheets[sheet]

        resolved = {}
        for role in ("native", "foreign", "fails", "last", "attempts"):
            val = self.role_vars[role].get()
            if val == CREATE_NEW:
                name = self.role_entries[role].get().strip()
                if name in df.columns:
                    messagebox.showerror("Column exists", f"Column '{name}' already exists.")
                    return
                default = 0 if role in ("fails", "attempts") else ""
                df[name] = default
                resolved[role] = name
            else:
                resolved[role] = val

        if len(set(resolved.values())) != len(resolved):
            messagebox.showerror("Duplicate columns", "Each role must map to a distinct column.")
            return

        self.app.sheet_name = sheet
        self.app.df = df
        self.app.cols = ColumnMap(**resolved)
        self.app.on_data_ready()


# ---------------------------------------------------------------------------
# GUI: practice screen
# ---------------------------------------------------------------------------

class PracticeFrame(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=20)
        self.app = app
        self.state = "ask"
        self.current_idx = None

        self.progress_label = ttk.Label(self, text="", foreground="#555555")
        self.progress_label.grid(row=0, column=0, columnspan=2, sticky="w")

        self.native_label = ttk.Label(self, text="", font=("", 22, "bold"), wraplength=560)
        self.native_label.grid(row=1, column=0, columnspan=2, pady=(30, 20), sticky="w")

        self.native_var = tk.StringVar()
        self.native_entry = ttk.Entry(self, textvariable=self.native_var, font=("", 22), width=36)
        self.native_entry.grid(row=1, column=0, columnspan=2, pady=(30, 20), sticky="w")
        self.native_entry.grid_remove()
        self.native_entry.bind("<Return>", self.on_enter)
        self.native_entry.bind("<Escape>", lambda e: self.cancel_edit())

        self.answer_var = tk.StringVar()
        self.answer_entry = ttk.Entry(self, textvariable=self.answer_var, font=("", 16), width=36)
        self.answer_entry.grid(row=2, column=0, pady=5, sticky="w")
        self.answer_entry.bind("<Return>", self.on_enter)
        self.answer_entry.bind("<Escape>", lambda e: self.cancel_edit())

        self.action_button = ttk.Button(self, text="Submit", command=self.on_enter)
        self.action_button.grid(row=2, column=1, padx=10, sticky="w")

        self.feedback_label = ttk.Label(self, text="", font=("", 14, "bold"))
        self.feedback_label.grid(row=3, column=0, columnspan=2, pady=(20, 5), sticky="w")

        self.detail_label = ttk.Label(self, text="", font=("", 12), foreground="#333333", justify="left")
        self.detail_label.grid(row=4, column=0, columnspan=2, sticky="w")

        self.edit_button = ttk.Button(self, text="Edit Phrase…", command=self.start_edit)
        self.edit_button.grid(row=5, column=0, pady=(40, 0), sticky="w")

        self.save_edit_button = ttk.Button(self, text="Save", command=self.save_edit)
        self.save_edit_button.grid(row=5, column=0, pady=(40, 0), sticky="w")
        self.save_edit_button.grid_remove()

        self.cancel_edit_button = ttk.Button(self, text="Cancel", command=self.cancel_edit)
        self.cancel_edit_button.grid(row=5, column=1, pady=(40, 0), sticky="w")
        self.cancel_edit_button.grid_remove()

    def on_show(self):
        self.load_next()

    def load_next(self):
        self.show_phrase(pick_next_index(self.app.df, self.app.cols))

    def show_phrase(self, idx):
        if idx is None or idx not in self.app.df.index:
            self.current_idx = None
            self.native_label.config(text="No phrases available to practice.")
            self.answer_entry.config(state="disabled")
            self.action_button.config(state="disabled")
            self.edit_button.config(state="disabled")
            self.feedback_label.config(text="")
            self.detail_label.config(text="")
            return

        self.current_idx = idx
        self.state = "ask"
        cols = self.app.cols
        self.native_label.config(text=str(self.app.df.at[idx, cols.native]))
        self.answer_var.set("")
        self.feedback_label.config(text="")
        self.detail_label.config(text="")
        self.action_button.config(text="Submit", state="normal")
        self.answer_entry.config(state="normal")
        self.answer_entry.focus_set()
        self.edit_button.config(state="normal")
        self.update_progress()

    def update_progress(self):
        df, cols = self.app.df, self.app.cols
        idxs = practice_indices(df, cols)
        due = sum(1 for i in idxs if get_fails(df, i, cols) > 0)
        self.progress_label.config(text=f"{len(idxs)} phrases loaded — {due} currently marked with fails")

    def on_enter(self, event=None):
        if self.state == "editing":
            self.save_edit()
            return
        if self.current_idx is None:
            return
        if self.state == "ask":
            self.submit_answer()
        else:
            self.load_next()

    def start_edit(self):
        if self.current_idx is None:
            return
        idx = self.current_idx
        cols = self.app.cols
        self.state = "editing"

        self.native_var.set(str(self.app.df.at[idx, cols.native]))
        self.native_label.grid_remove()
        self.native_entry.grid()
        self.native_entry.focus_set()
        self.native_entry.select_range(0, "end")

        self.answer_entry.config(state="normal")
        self.answer_var.set(str(self.app.df.at[idx, cols.foreign]))

        self.feedback_label.config(text="Editing phrase…", foreground="#555555")
        self.detail_label.config(text="")

        self.action_button.grid_remove()
        self.edit_button.grid_remove()
        self.save_edit_button.grid()
        self.cancel_edit_button.grid()

    def save_edit(self):
        new_native = self.native_var.get().strip()
        new_foreign = self.answer_var.get().strip()
        if not new_native or not new_foreign:
            messagebox.showerror("Missing text", "Both the phrase and its translation must be non-empty.")
            return
        cols = self.app.cols
        idx = self.current_idx
        self.app.df.at[idx, cols.native] = new_native
        self.app.df.at[idx, cols.foreign] = new_foreign
        self.app.mark_dirty()
        self.end_edit()

    def cancel_edit(self):
        if self.state != "editing":
            return
        self.end_edit()

    def end_edit(self):
        self.native_entry.grid_remove()
        self.native_label.grid()
        self.save_edit_button.grid_remove()
        self.cancel_edit_button.grid_remove()
        self.action_button.grid()
        self.edit_button.grid()
        self.show_phrase(self.current_idx)

    def submit_answer(self):
        cols = self.app.cols
        idx = self.current_idx
        user_answer = self.answer_var.get().strip()
        alternatives = split_alternatives(self.app.df.at[idx, cols.foreign])
        is_correct = any(user_answer.lower() == alt.lower() for alt in alternatives)
        display_answer = " / ".join(alternatives)

        record_attempt(self.app.df, idx, cols, is_correct)
        self.app.mark_dirty()

        if is_correct:
            self.feedback_label.config(text="Correct!", foreground="#1a7f37")
        else:
            self.feedback_label.config(text="Wrong.", foreground="#c0392b")
        self.detail_label.config(
            text=f"Your answer:    {user_answer or '(empty)'}\nCorrect answer: {display_answer}"
        )
        self.answer_entry.config(state="readonly")
        self.action_button.config(text="Next")
        self.state = "feedback"


# ---------------------------------------------------------------------------
# GUI: edit screen
# ---------------------------------------------------------------------------

class EditFrame(ttk.Frame):
    COLUMNS = ("native", "foreign", "fails", "attempts", "last")

    def __init__(self, master, app):
        super().__init__(master, padding=15)
        self.app = app
        self._edit_widget = None

        ttk.Label(self, text="Edit Vocabulary", font=("", 16, "bold")).pack(anchor="w")

        self.tree = ttk.Treeview(self, columns=self.COLUMNS, show="headings", height=18)
        for c in self.COLUMNS:
            self.tree.heading(c, text=c.capitalize())
            self.tree.column(c, width=160 if c in ("native", "foreign") else 100, anchor="w")
        self.tree.pack(fill="both", expand=True, pady=10)
        self.tree.bind("<Double-1>", self.on_double_click)

        btns = ttk.Frame(self)
        btns.pack(fill="x")
        ttk.Button(btns, text="Add phrase", command=self.add_row).pack(side="left")
        ttk.Button(btns, text="Delete selected", command=self.delete_row).pack(side="left", padx=5)
        ttk.Button(btns, text="Back to Practice", command=self.back).pack(side="right")

    def on_show(self):
        self.refresh()

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        df, cols = self.app.df, self.app.cols
        for i in df.index:
            values = (
                df.at[i, cols.native],
                df.at[i, cols.foreign],
                to_int(df.at[i, cols.fails]),
                to_int(df.at[i, cols.attempts]),
                df.at[i, cols.last],
            )
            self.tree.insert("", "end", iid=str(i), values=values)

    def add_row(self):
        df, cols = self.app.df, self.app.cols
        new_idx = (int(df.index.max()) + 1) if len(df.index) else 0
        df.loc[new_idx, cols.native] = ""
        df.loc[new_idx, cols.foreign] = ""
        df.loc[new_idx, cols.fails] = 0
        df.loc[new_idx, cols.attempts] = 0
        df.loc[new_idx, cols.last] = ""
        self.app.mark_dirty()
        self.refresh()

    def delete_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        if not messagebox.askyesno("Delete", f"Delete {len(sel)} selected phrase(s)?"):
            return
        df = self.app.df
        for iid in sel:
            df.drop(index=int(iid), inplace=True)
        self.app.mark_dirty()
        self.refresh()

    def on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id:
            return
        col_index = int(col_id.replace("#", "")) - 1
        col_name = self.tree["columns"][col_index]

        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        value = self.tree.set(row_id, col_name)

        if self._edit_widget is not None:
            self._edit_widget.destroy()

        var = tk.StringVar(value=value)
        entry = ttk.Entry(self.tree, textvariable=var)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        entry.select_range(0, "end")
        self._edit_widget = entry

        def commit(event=None):
            if self._edit_widget is None:
                return
            new_value = var.get()
            self.tree.set(row_id, col_name, new_value)
            self.apply_edit(int(row_id), col_name, new_value)
            entry.destroy()
            self._edit_widget = None

        def cancel(event=None):
            entry.destroy()
            self._edit_widget = None

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", cancel)

    def apply_edit(self, idx, col_name, value):
        df, cols = self.app.df, self.app.cols
        target_col = getattr(cols, col_name)
        if col_name in ("fails", "attempts"):
            value = to_int(value)
        df.at[idx, target_col] = value
        self.app.mark_dirty()

    def back(self):
        self.app.show_frame("practice")


# ---------------------------------------------------------------------------
# Main application window
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Vocabulary Trainer")
        self.geometry("720x580")
        self.minsize(600, 480)

        self.store = None
        self.sheet_name = None
        self.df = None
        self.cols = None
        self.dirty = False

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        self.frames = {
            "setup": SetupFrame(container, self),
            "practice": PracticeFrame(container, self),
            "edit": EditFrame(container, self),
        }
        for frame in self.frames.values():
            frame.place(relwidth=1, relheight=1)

        self.build_menu()
        if not self.autoload_remembered_setup():
            self.show_frame("setup")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def build_menu(self):
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open…", command=lambda: self.show_frame("setup"))
        file_menu.add_command(label="Save", command=self.save, accelerator="Ctrl+S")
        file_menu.add_command(label="Save As…", command=self.save_as)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)
        menubar.add_cascade(label="File", menu=file_menu)

        vocab_menu = tk.Menu(menubar, tearoff=0)
        vocab_menu.add_command(label="Practice", command=lambda: self.show_frame("practice"))
        vocab_menu.add_command(label="Edit Phrases", command=lambda: self.show_frame("edit"))
        menubar.add_cascade(label="Vocabulary", menu=vocab_menu)

        self.config(menu=menubar)
        self.bind_all("<Control-s>", lambda e: self.save())

    def show_frame(self, name):
        frame = self.frames[name]
        frame.tkraise()
        if hasattr(frame, "on_show"):
            frame.on_show()

    def on_data_ready(self):
        self.update_title_dirty()
        self.show_frame("practice")
        self.remember_setup()

    def remember_setup(self):
        data = {
            "path": self.store.path,
            "sheet": self.sheet_name,
            "columns": {role: getattr(self.cols, role) for role in COLUMN_ROLES},
        }
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except OSError:
            pass

    def autoload_remembered_setup(self):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError):
            return False

        path = data.get("path")
        sheet = data.get("sheet")
        columns = data.get("columns", {})
        if not path or not sheet or not all(role in columns for role in COLUMN_ROLES):
            return False
        if not os.path.exists(path):
            return False

        try:
            store = VocabStore(path)
        except Exception:
            return False
        if sheet not in store.sheet_names:
            return False
        df = store.sheets[sheet]
        if not all(columns[role] in df.columns for role in COLUMN_ROLES):
            return False

        self.store = store
        self.sheet_name = sheet
        self.df = df
        self.cols = ColumnMap(**{role: columns[role] for role in COLUMN_ROLES})
        self.on_data_ready()
        return True

    def mark_dirty(self):
        self.dirty = True
        self.update_title_dirty()

    def update_title_dirty(self):
        base = "Vocabulary Trainer"
        if self.store:
            base += f" — {os.path.basename(self.store.path)} / {self.sheet_name}"
        self.title(base + (" *" if self.dirty else ""))

    def save(self):
        if not self.store or self.df is None:
            return
        proceed = messagebox.askyesno(
            "Save",
            "Saving will overwrite the original .ods file with the current data.\n\n"
            "Note: only the phrase data is written back — any custom cell "
            "formatting, colors, or column widths in the original file will "
            "not be preserved. A one-time backup is kept as '<file>.bak'.\n\n"
            "Continue?",
        )
        if not proceed:
            return
        try:
            self.store.save(self.sheet_name, self.df)
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))
            return
        self.dirty = False
        self.update_title_dirty()
        messagebox.showinfo("Saved", "Vocabulary saved.")

    def save_as(self):
        if not self.store or self.df is None:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".ods",
            filetypes=[("OpenDocument Spreadsheet", "*.ods")],
        )
        if not path:
            return
        try:
            self.store.save(self.sheet_name, self.df, path=path)
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))
            return
        self.dirty = False
        self.update_title_dirty()
        messagebox.showinfo("Saved", f"Vocabulary saved to {path}")

    def on_close(self):
        if self.dirty:
            resp = messagebox.askyesnocancel("Unsaved changes", "You have unsaved changes. Save before exiting?")
            if resp is None:
                return
            if resp:
                self.save()
                if self.dirty:
                    return
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
