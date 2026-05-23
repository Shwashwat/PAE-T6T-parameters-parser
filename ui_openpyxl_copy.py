import os
import sys
import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import ctypes

from tkinterdnd2 import TkinterDnD

from t6t_processor_openpyxl_copy import get_n_save_settings, monthly_params

# ---------------- DPI FIX ----------------
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

CONFIG_FILE = "config.json"


# ---------------- CONFIG ----------------
def load_config():
    if not os.path.exists(CONFIG_FILE):
        return {
            "auto_open": True,
            "open_folder": False,
            "last_folder": "",
            "recent_folders": [],
            "show_advanced": False
        }
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)


def save_config(cfg):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)


# ---------------- LOG REDIRECT ----------------
class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

        # Define color tags once
        self.widget.tag_config("WARN", foreground="orange")
        self.widget.tag_config("ERROR", foreground="red")
        self.widget.tag_config("DEFAULT", foreground="black")
        self.widget.tag_config("INFO_REPORT", foreground="green")

    def write(self, msg):
        self.widget.after(0, lambda: self._append(msg))

    def _append(self, msg):
        if not msg:
            return

        timestamp = datetime.now().strftime("[%H:%M:%S] ")
        formatted = timestamp + msg if msg.strip() else msg

        # --- Severity detection ---
        if "[ERROR]" in msg:
            tag = "ERROR"
        elif "[WARN]" in msg:
            tag = "WARN"
        elif "[INFO_REPORT]" in msg:
            tag = "INFO_REPORT"
        else:
            tag = "DEFAULT"

        # Insert with tag
        self.widget.insert(tk.END, formatted, tag)
        self.widget.see(tk.END)

    def flush(self):
        pass


# ---------------- APP ----------------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("T6T Parameters Processor")
        self.root.geometry("1000x640")
        self.root.iconbitmap("icondraft1.ico") ## icon on the taskbar

        self.config = load_config()

        self.folder_path = tk.StringVar(value=self.config.get("last_folder", ""))
        self.mode = tk.StringVar(value="monthly")
        self.custom_params = tk.StringVar()
        self.extra_params = tk.StringVar()
        self.filename = tk.StringVar()

        self.auto_open = tk.BooleanVar(value=self.config.get("auto_open", True))
        self.open_folder = tk.BooleanVar(value=self.config.get("open_folder", False))
        self.show_advanced = tk.BooleanVar(value=self.config.get("show_advanced", False))

        self.style = ttk.Style()

        self.style.configure(
            "Run.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=8
        )
        
        self.output_file = None

        self.build_ui()

        sys.stdout = TextRedirector(self.log)

        self.enable_drag_drop()
        self.update_mode_ui()
        self.toggle_advanced(init=True)

    # ---------- UI ----------
    def build_ui(self):
        # -------- TOP ROW 1 --------
        frame_top1 = ttk.Frame(self.root)
        frame_top1.pack(fill="x", padx=10, pady=5)

        self.folder_combo = ttk.Combobox(
            frame_top1,
            textvariable=self.folder_path,
            values=self.config.get("recent_folders", [])
        )
        self.folder_combo.pack(side="left", fill="x", expand=True)

        ttk.Button(frame_top1, text="Browse", command=self.browse)\
            .pack(side="left", padx=5)

        # -------- TOP ROW 2 --------
        frame_top2 = ttk.Frame(self.root)
        frame_top2.pack(fill="x", padx=10, pady=0)

        ttk.Checkbutton(
            frame_top2,
            text="Show Advanced",
            variable=self.show_advanced,
            command=self.toggle_advanced
        ).pack(side="left")

        ttk.Checkbutton(
            frame_top2,
            text="Auto-open Excel",
            variable=self.auto_open,
            command=self.update_config
        ).pack(side="left", padx=10)

        ttk.Checkbutton(
            frame_top2,
            text="Open Folder",
            variable=self.open_folder,
            command=self.update_config
        ).pack(side="left")

        ttk.Button(
            frame_top2,
            text="▶ RUN",
            command=self.run,
            style="Run.TButton"
        ).pack(side="right", padx=5)

        # -------- ADVANCED (DO NOT PACK HERE) --------
        self.frame_advanced = ttk.LabelFrame(self.root, text="Advanced Options")

        self.advanced_content = ttk.Frame(self.frame_advanced)
        self.advanced_content.pack(fill="x")

        # Mode
        frame_mode = ttk.LabelFrame(self.advanced_content, text="Mode")
        frame_mode.pack(fill="x", pady=5)

        for text, val in [("Monthly", "monthly"), ("All", "all"), ("Custom", "custom")]:
            ttk.Radiobutton(
                frame_mode,
                text=text,
                value=val,
                variable=self.mode,
                command=self.update_mode_ui,
            ).pack(anchor="w")

        # Params
        frame_params = ttk.LabelFrame(self.advanced_content, text="Parameters")
        frame_params.pack(fill="x", pady=5)

        ttk.Label(frame_params, text="Custom Params:").pack(anchor="w")
        self.entry_params = ttk.Entry(frame_params, textvariable=self.custom_params)
        self.entry_params.pack(fill="x")

        ttk.Label(frame_params, text="Append Params:").pack(anchor="w")
        ttk.Entry(frame_params, textvariable=self.extra_params).pack(fill="x")

        ttk.Label(frame_params, text="Output File Name:").pack(anchor="w")
        ttk.Entry(frame_params, textvariable=self.filename).pack(fill="x")

        # -------- TABLE --------
        frame_table = ttk.LabelFrame(self.root, text="Preview")
        frame_table.pack(fill="both", expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(frame_table)
        self.tree.pack(fill="both", expand=True)

        # -------- LOGS --------
        frame_log = ttk.LabelFrame(self.root, text="Logs")
        frame_log.pack(fill="both", padx=10, pady=5)

        self.log = tk.Text(frame_log, height=8)
        self.log.pack(fill="both", expand=True)

    # ---------- Logic ----------
    def toggle_advanced(self, init=False):
        if self.show_advanced.get():
            self.frame_advanced.pack(fill="x", padx=10, pady=5)
        else:
            self.frame_advanced.pack_forget()

        if not init:
            self.config["show_advanced"] = self.show_advanced.get()
            save_config(self.config)

    def update_mode_ui(self):
        state = "normal" if self.mode.get() == "custom" else "disabled"
        self.entry_params.config(state=state)

    def update_config(self):
        self.config["auto_open"] = self.auto_open.get()
        self.config["open_folder"] = self.open_folder.get()
        save_config(self.config)

    def update_history(self, folder):
        history = self.config.get("recent_folders", [])

        if folder in history:
            history.remove(folder)

        history.insert(0, folder)
        history = history[:5]

        self.config["recent_folders"] = history
        self.config["last_folder"] = folder

        save_config(self.config)
        self.folder_combo["values"] = history

    def enable_drag_drop(self):
        self.root.drop_target_register("DND_Files")
        self.root.dnd_bind("<<Drop>>", self.on_drop)

    def on_drop(self, event):
        folder = event.data.strip("{}")
        if os.path.isdir(folder):
            self.folder_path.set(folder)
            self.update_history(folder)
            print(f"Dropped folder: {folder}")

    def browse(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)
            self.update_history(folder)

    def build_args(self):
        mode = self.mode.get()

        params = None
        extra_params = None
        all_flag = False

        if mode == "monthly":
            params = monthly_params

        elif mode == "all":
            all_flag = True

        elif mode == "custom":
            params = [x.strip() for x in self.custom_params.get().split(",") if x.strip()]

        if self.extra_params.get():
            extra_params = [x.strip() for x in self.extra_params.get().split(",") if x.strip()]

        name = self.filename.get().strip() or None

        return params, extra_params, all_flag, name

    def run(self):
        folder = self.folder_path.get()

        if not os.path.isdir(folder):
            messagebox.showerror("Error", "Invalid folder")
            return

        self.update_history(folder)

        threading.Thread(target=self.process).start()

    def process(self):
        try:
            folder = self.folder_path.get()
            params, extra_params, all_flag, name = self.build_args()

            dictionary, file_path = get_n_save_settings(
                folder,
                params=params,
                extra_params=extra_params,
                all_flag=all_flag,
                save_file_name=name
            )

            self.output_file = file_path

            if dictionary is not None and dictionary:
                self.root.after(0, lambda: self.show_table(dictionary))

            if self.auto_open.get():
                os.startfile(file_path)

            if self.open_folder.get():
                os.startfile(os.path.dirname(file_path))

        except Exception as e:
            print(f"[ERROR] {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))

    def show_table(self, dictionary:dict):
        self.tree.delete(*self.tree.get_children())

        shown_columns = list(dictionary.keys())[:14]
        self.tree["columns"] = shown_columns
        self.tree["show"] = "headings"

        # configure columns
        for col in shown_columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)

        # convert column data -> row data
        rows = zip(*(dictionary[col] for col in shown_columns))    
        
        # show first 40 rows

        for i, row in enumerate(rows):

            if i>=40:
                break
            self.tree.insert("", "end", values=list(row))


# ---------------- MAIN ----------------
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = App(root)
    root.mainloop()