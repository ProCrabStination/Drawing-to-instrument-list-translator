"""Simple GUI for the Instrument List Transcriber.

Pick the drawings, pick a model, press Run. Everything else is worked out for you (endpoint, API host,
token limit, render DPI, reference images, output file names) and lives under "Advanced" if you want to
override it. Every setting is saved automatically as you change it and restored on the next launch.
"""
import argparse
import datetime
import json
import os
import queue
import re
import sys
import threading
import time
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Dict, List, Optional

import Claude_Code_Backend as claude_code
import Cable_Workbook_Step as cable_step
import Instrument_List_Transcriber as ilt

APP_DIR = Path(__file__).resolve().parent
SETTINGS_FILE = APP_DIR / "instrument_list_settings.json"
OLD_SETTINGS_FILE = APP_DIR / "instrument_list_gui_inputs.txt"  # read once to carry over earlier choices
REFERENCE_IMAGES = {
    "controls": APP_DIR / "AI Reference - Controls - smaller.png",
    "pids": APP_DIR / "AI Reference - P&IDs - smaller.png",
}
LOG_DIR = APP_DIR / "logs"  # one log file per run, written live
LOCAL_HOST = "http://localhost:11434"
CLOUD_HOSTS = {
    "Anthropic": "https://api.anthropic.com",
    "OpenAI": "https://api.openai.com",
    "Google": "https://generativelanguage.googleapis.com",
}
DEFAULT_MAX_TOKENS = 8192
DEFAULT_DPI = 300
HISTORY_LENGTH = 8

MODES = [
    ("all", "Full run: controls + P&ID, then combine them"),
    ("controls", "Controls drawings only"),
    ("pids", "P&ID drawings only"),
    ("combine", "Combine two saved results"),
]

# Every setting the window remembers: name -> default. Booleans and strings only; the rest is stored separately.
DEFAULTS: Dict[str, object] = {
    "mode": "all",
    "controls_pdf": "", "controls_pages": "",
    "pids_pdf": "", "pids_pages": "",
    "controls_json": "", "pids_json": "",
    "model": "", "api_key": "",
    "out_dir": "",
    "save_excel": True,
    "use_ocr": True,
    "advanced_open": False,
    # advanced
    "controls_model": "", "pids_model": "", "combine_model": "",
    "max_tokens": "", "dpi": "",
    "auto_reference": True, "controls_reference": "", "pids_reference": "",
    "ocr_scale": "3.0", "ocr_tables": True, "ocr_pictures": False,
    # final step: fill the Cable Calculations workbook from the combined list
    "fill_workbook": True, "workbook_template": str(cable_step.DEFAULT_TEMPLATE),
}


# --------------------------------------------------------------------------------------------- settings
def load_settings() -> dict:
    """Saved settings, seeded from the older settings file the first time this window is used."""
    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return data
    except (OSError, ValueError):
        pass

    seeded: dict = {}
    try:
        old = json.loads(OLD_SETTINGS_FILE.read_text(encoding="utf-8"))
        if isinstance(old, dict):
            for new_key, old_key in (("controls_pdf", "controls_pdf"), ("pids_pdf", "pids_pdf"),
                                     ("controls_pages", "controls_pages"),
                                     ("pids_pages", "pids_pages")):
                if old.get(old_key):
                    seeded[new_key] = old[old_key]
            seeded["model"] = old.get("controls_model") or old.get("model") or ""
            seeded["api_key"] = next((old[k] for k in ("api_key", "controls_api_key", "pids_api_key") if old.get(k)
                                      and old[k] != "anything"), "")
            if old.get("run_mode") in {m for m, _ in MODES}:
                seeded["mode"] = old["run_mode"]
    except (OSError, ValueError):
        pass
    return seeded


def write_settings(data: dict) -> None:
    tmp = SETTINGS_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
    tmp.replace(SETTINGS_FILE)  # atomic, so a crash mid-save can't leave a half-written file


# ------------------------------------------------------------------------------------------ model logic
class RunLog:
    """One log file per run: everything shown in the window (status lines and streamed model output) is also
    written here as it happens, and flushed on every write so the file is complete even if the program dies."""

    def __init__(self, mode: str) -> None:
        LOG_DIR.mkdir(exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.path = LOG_DIR / f"run_{stamp}_{mode}.log"
        self._file = self.path.open("w", encoding="utf-8", buffering=1)
        self._lock = threading.Lock()
        self.started = time.time()

    def line(self, message: str) -> None:
        stamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.chunk(f"[{stamp}] {message}\n")

    def chunk(self, text: str) -> None:
        with self._lock:
            if not self._file.closed:
                self._file.write(text)
                self._file.flush()

    def close(self, outcome: str) -> None:
        self.line(f"{outcome} (took {time.time() - self.started:.0f}s)")
        with self._lock:
            self._file.close()


def _guess_provider(model: str, api_key: str) -> str:
    lowered = model.lower()
    if ilt._is_claude_model(model):
        return "Anthropic"
    if lowered.startswith(("gpt", "o1", "o3", "o4", "chatgpt")):
        return "OpenAI"
    if lowered.startswith("gemini"):
        return "Google"
    if api_key.startswith("sk-ant"):
        return "Anthropic"
    if api_key.startswith("AIza"):
        return "Google"
    return "OpenAI"


def _providers_for_key(api_key: str) -> List[str]:
    if api_key.startswith("sk-ant"):
        return ["Anthropic"]
    if api_key.startswith("AIza"):
        return ["Google"]
    if api_key.startswith("sk-"):
        return ["OpenAI"]
    return []


def _auto_name(pdf: str, suffix: str) -> str:
    stem = re.sub(r" - (controls|pids)$", "", Path(pdf).stem)  # saved stage results feed the combine step
    return f"{stem} - {suffix}"


def run_gui(parser: argparse.ArgumentParser) -> None:
    saved = load_settings()
    root = tk.Tk()
    root.title("Instrument List Transcriber")
    root.geometry(saved.get("geometry") or "1280x820")
    root.minsize(1000, 640)

    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    style = ttk.Style()
    style.configure("Header.TLabel", font=("Segoe UI", 15, "bold"))
    style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
    style.configure("Run.TButton", font=("Segoe UI", 11, "bold"), padding=(18, 8))
    style.configure("Hint.TLabel", foreground="#666666")
    style.configure("Bad.TLabel", foreground="#b00020")
    style.configure("Good.TLabel", foreground="#1b7f3b")

    # ---- state -------------------------------------------------------------------------------------
    vars_: Dict[str, tk.Variable] = {}
    for key, default in DEFAULTS.items():
        value = saved.get(key, default)
        vars_[key] = tk.BooleanVar(value=bool(value)) if isinstance(default, bool) else tk.StringVar(value=str(value))
    history: Dict[str, List[str]] = {k: list(v) for k, v in (saved.get("history") or {}).items() if isinstance(v, list)}
    catalog: Dict[str, bool] = {m: bool(local) for m, local in (saved.get("known_models") or {}).items()}  # name -> is_local
    ui_queue: "queue.Queue[Callable[[], None]]" = queue.Queue()
    state = {"running": False, "last_outputs": [], "save_job": None}

    def post(fn: Callable[[], None]) -> None:
        ui_queue.put(fn)  # worker threads never touch tkinter directly

    def pump() -> None:
        try:
            while True:
                ui_queue.get_nowait()()
        except queue.Empty:
            pass
        root.after(80, pump)

    # ---- autosave ----------------------------------------------------------------------------------
    def save_now() -> None:
        if state["save_job"] is not None:
            root.after_cancel(state["save_job"])
            state["save_job"] = None
        payload = {k: (v.get() if isinstance(v, tk.BooleanVar) else v.get().strip()) for k, v in vars_.items()}
        payload["history"] = history
        payload["known_models"] = catalog
        try:
            payload["geometry"] = root.geometry()
        except tk.TclError:
            pass
        try:
            write_settings(payload)
        except OSError as exc:
            status("error", f"Could not save settings: {exc}")

    def schedule_save(*_args: object) -> None:
        if state["save_job"] is not None:
            root.after_cancel(state["save_job"])
        state["save_job"] = root.after(400, save_now)

    for var in vars_.values():
        var.trace_add("write", schedule_save)

    def remember(key: str, path: str) -> None:
        path = path.strip()
        if not path:
            return
        entries = [p for p in history.get(key, []) if p != path]
        history[key] = ([path] + entries)[:HISTORY_LENGTH]

    # ---- layout skeleton ---------------------------------------------------------------------------
    paned = ttk.Panedwindow(root, orient=tk.HORIZONTAL)
    paned.pack(fill="both", expand=True)
    left_outer = ttk.Frame(paned)
    right = ttk.Frame(paned)
    paned.add(left_outer, weight=2)
    paned.add(right, weight=3)

    canvas = tk.Canvas(left_outer, highlightthickness=0, borderwidth=0)
    scrollbar = ttk.Scrollbar(left_outer, orient="vertical", command=canvas.yview)
    form = ttk.Frame(canvas, padding=(16, 12))
    window_id = canvas.create_window((0, 0), window=form, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    form.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))

    def _wheel(event: tk.Event) -> None:
        if canvas.bbox("all") and canvas.bbox("all")[3] > canvas.winfo_height():
            canvas.yview_scroll(int(-event.delta / 120), "units")

    canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", _wheel))
    canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))
    form.columnconfigure(0, weight=1)

    ttk.Label(form, text="Instrument List Transcriber", style="Header.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(form, text="Choose your drawings and a model, then press Run. Settings are saved automatically.",
              style="Hint.TLabel").grid(row=1, column=0, sticky="w", pady=(0, 10))

    # ---- log panel (right) -------------------------------------------------------------------------
    ttk.Label(right, text="Progress", style="Header.TLabel").pack(anchor="w", padx=12, pady=(12, 4))
    status_var = tk.StringVar(value="Ready")
    status_label = ttk.Label(right, textvariable=status_var, wraplength=600, justify="left")
    status_label.pack(anchor="w", padx=12)
    progress = ttk.Progressbar(right, mode="indeterminate")
    progress.pack(fill="x", padx=12, pady=8)

    log_frame = ttk.Frame(right)
    log_frame.pack(fill="both", expand=True, padx=12)
    log_text = tk.Text(log_frame, wrap="word", state="disabled", font=("Consolas", 9))
    log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
    log_text.configure(yscrollcommand=log_scroll.set)
    log_text.pack(side="left", fill="both", expand=True)
    log_scroll.pack(side="left", fill="y")

    def append_log(message: str) -> None:
        stamp = datetime.datetime.now().strftime("%H:%M:%S")
        log_text.configure(state="normal")
        log_text.insert(tk.END, f"[{stamp}] {message}\n")
        log_text.configure(state="disabled")
        log_text.see(tk.END)
        match = re.search(r"Processing page (\d+)/(\d+)", message)
        if match and state["running"]:
            status("busy", f"Working on page {match.group(1)} of {match.group(2)}...")

    def append_chunk(text: str) -> None:
        log_text.configure(state="normal")
        log_text.insert(tk.END, text)
        log_text.configure(state="disabled")
        log_text.see(tk.END)

    def status(kind: str, message: str) -> None:
        status_var.set(message)
        status_label.configure(style={"error": "Bad.TLabel", "ok": "Good.TLabel"}.get(kind, "TLabel"))

    buttons = ttk.Frame(right)
    buttons.pack(anchor="w", padx=12, pady=10)
    open_folder_btn = ttk.Button(buttons, text="Open output folder", state="disabled",
                                 command=lambda: _open_output_folder())
    open_folder_btn.pack(side="left", padx=(0, 8))
    open_log_btn = ttk.Button(buttons, text="Open run log", state="disabled", command=lambda: _open_run_log())
    open_log_btn.pack(side="left", padx=(0, 8))

    def save_log() -> None:
        target = filedialog.asksaveasfilename(title="Save log", defaultextension=".txt",
                                              filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if target:
            Path(target).write_text(log_text.get("1.0", tk.END), encoding="utf-8")

    def clear_log() -> None:
        log_text.configure(state="normal")
        log_text.delete("1.0", tk.END)
        log_text.configure(state="disabled")

    ttk.Button(buttons, text="Save log...", command=save_log).pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="Clear log", command=clear_log).pack(side="left")

    # ---- form helpers ------------------------------------------------------------------------------
    form_row = {"n": 2}

    def next_row() -> int:
        form_row["n"] += 1
        return form_row["n"]

    def section(title: str) -> ttk.Labelframe:
        frame = ttk.Labelframe(form, text=title, style="Section.TLabelframe", padding=(12, 8))
        frame.grid(row=next_row(), column=0, sticky="ew", pady=(0, 10))
        frame.columnconfigure(1, weight=1)
        return frame

    mode_rows: Dict[str, List[tk.Widget]] = {}  # mode -> widgets shown for it (managed by update_mode_visibility)
    all_mode_widgets: List[tk.Widget] = []

    def file_row(parent: ttk.Frame, row: int, label: str, key: str, hist_key: str, filetypes: list,
                 modes: List[str], pages_key: Optional[str] = None, save: bool = False) -> None:
        widgets: List[tk.Widget] = []
        lab = ttk.Label(parent, text=label)
        lab.grid(row=row, column=0, sticky="e", padx=(0, 8), pady=4)
        combo = ttk.Combobox(parent, textvariable=vars_[key], values=history.get(hist_key, []))
        combo.grid(row=row, column=1, sticky="ew", pady=4)
        combo.bind("<Button-1>", lambda _e, c=combo, h=hist_key: c.configure(values=history.get(h, [])))

        def browse() -> None:
            chosen = (filedialog.asksaveasfilename if save else filedialog.askopenfilename)(
                title=label, filetypes=filetypes,
                initialdir=str(Path(vars_[key].get()).parent) if vars_[key].get().strip() else None)
            if chosen:
                vars_[key].set(chosen)

        btn = ttk.Button(parent, text="Browse...", command=browse)
        btn.grid(row=row, column=2, padx=(8, 0), pady=4)
        widgets += [lab, combo, btn]
        if pages_key:
            pages_label = ttk.Label(parent, text="Pages")
            pages_label.grid(row=row, column=3, padx=(12, 4))
            pages = ttk.Entry(parent, textvariable=vars_[pages_key], width=10)
            pages.grid(row=row, column=4, pady=4)
            widgets += [pages_label, pages]
        for mode in modes:
            mode_rows.setdefault(mode, []).extend(widgets)
        all_mode_widgets.extend(widgets)

    # ---- 1. what to do -----------------------------------------------------------------------------
    sec_mode = section("1. What do you want to do?")
    for index, (value, text) in enumerate(MODES):
        ttk.Radiobutton(sec_mode, text=text, value=value, variable=vars_["mode"]).grid(
            row=index, column=0, columnspan=3, sticky="w", pady=1)

    # ---- 2. files ----------------------------------------------------------------------------------
    sec_files = section("2. Drawings")
    pdf_types = [("PDF files", "*.pdf"), ("All files", "*.*")]
    json_types = [("JSON files", "*.json"), ("All files", "*.*")]
    file_row(sec_files, 0, "Controls PDF", "controls_pdf", "controls_pdf", pdf_types, ["all", "controls"], "controls_pages")
    file_row(sec_files, 1, "P&ID PDF", "pids_pdf", "pids_pdf", pdf_types, ["all", "pids"], "pids_pages")
    file_row(sec_files, 2, "Controls results", "controls_json", "controls_json", json_types, ["combine"])
    file_row(sec_files, 3, "P&ID results", "pids_json", "pids_json", json_types, ["combine"])
    pages_hint = ttk.Label(sec_files, text="Leave Pages empty for every page, or use e.g. 1,3-5.", style="Hint.TLabel")
    pages_hint.grid(row=5, column=0, columnspan=5, sticky="w", pady=(4, 0))
    for _mode in ("all", "controls", "pids"):
        mode_rows.setdefault(_mode, []).append(pages_hint)
    all_mode_widgets.append(pages_hint)

    # ---- 3. model ----------------------------------------------------------------------------------
    sec_model = section("3. AI model")
    ttk.Label(sec_model, text="Model").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=4)
    model_combo = ttk.Combobox(sec_model, textvariable=vars_["model"])  # values filled in at start-up
    model_combo.grid(row=0, column=1, sticky="ew", pady=4)
    refresh_btn = ttk.Button(sec_model, text="Refresh list", command=lambda: refresh_models())
    refresh_btn.grid(row=0, column=2, padx=(8, 0))
    model_note = ttk.Label(sec_model, text="", style="Hint.TLabel", wraplength=420, justify="left")
    model_note.grid(row=1, column=1, columnspan=2, sticky="w")

    key_label = ttk.Label(sec_model, text="API key")
    key_entry = ttk.Entry(sec_model, textvariable=vars_["api_key"], show="•")
    show_key = tk.BooleanVar(value=False)
    key_toggle = ttk.Checkbutton(sec_model, text="Show", variable=show_key,
                                 command=lambda: key_entry.configure(show="" if show_key.get() else "•"))
    key_widgets = (key_label, key_entry, key_toggle)

    def ordered_names() -> List[str]:
        """Local Ollama models first, then cloud models, then the Claude Code (CLI) choices."""
        def rank(name: str) -> tuple:
            if claude_code.is_claude_code_model(name):
                return (2, claude_code.MODEL_CHOICES.index(name[len(claude_code.PREFIX):])
                        if name[len(claude_code.PREFIX):] in claude_code.MODEL_CHOICES else 99, name)
            return (0 if catalog.get(name, True) else 1, 0, name.lower())
        return sorted(catalog, key=rank)

    def is_cloud(model: str) -> bool:
        if claude_code.is_claude_code_model(model):
            return False  # uses the CLI's own login
        if model in catalog:
            return not catalog[model]
        return model.lower().startswith(("claude", "gpt", "gemini", "o1", "o3", "o4", "chatgpt"))

    def update_model_widgets(*_args: object) -> None:
        model = vars_["model"].get().strip()
        if not model:
            model_note.configure(text="Pick the vision model to use. Local models come from Ollama.")
        elif claude_code.is_claude_code_model(model):
            if claude_code.cli_path():
                model_note.configure(text="Runs through your installed Claude Code CLI (uses your Claude login; no API key needed).")
            else:
                model_note.configure(text="The Claude Code CLI ('claude') isn't on PATH, so this model can't run.")
        elif is_cloud(model):
            model_note.configure(text=f"Cloud model ({_guess_provider(model, vars_['api_key'].get().strip())}); needs an API key.")
        else:
            model_note.configure(text="Runs locally on Ollama; no API key needed.")
        if model and is_cloud(model):
            key_label.grid(row=2, column=0, sticky="e", padx=(0, 8), pady=4)
            key_entry.grid(row=2, column=1, sticky="ew", pady=4)
            key_toggle.grid(row=2, column=2, padx=(8, 0))
        else:
            for widget in key_widgets:
                widget.grid_remove()

    vars_["model"].trace_add("write", update_model_widgets)

    # ---- 4. output ---------------------------------------------------------------------------------
    sec_out = section("4. Results")
    ttk.Label(sec_out, text="Save to").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=4)
    ttk.Entry(sec_out, textvariable=vars_["out_dir"]).grid(row=0, column=1, sticky="ew", pady=4)

    def browse_out() -> None:
        chosen = filedialog.askdirectory(title="Folder for results")
        if chosen:
            vars_["out_dir"].set(chosen)

    ttk.Button(sec_out, text="Browse...", command=browse_out).grid(row=0, column=2, padx=(8, 0))
    ttk.Label(sec_out, text="Empty = next to each PDF. File names are generated from the PDF names.",
              style="Hint.TLabel").grid(row=1, column=1, columnspan=2, sticky="w")
    ttk.Checkbutton(sec_out, text="Also save Excel (.xlsx) copies", variable=vars_["save_excel"]).grid(
        row=2, column=1, columnspan=2, sticky="w", pady=(4, 0))
    ttk.Checkbutton(sec_out, text="Read the drawing text with OCR first (recommended; cached, so re-runs are fast)",
                    variable=vars_["use_ocr"]).grid(row=3, column=1, columnspan=2, sticky="w")
    ttk.Checkbutton(sec_out, text="At the end, fill the Cable Calculations workbook with the combined list",
                    variable=vars_["fill_workbook"]).grid(row=4, column=1, columnspan=2, sticky="w")
    ttk.Label(sec_out, text="Applies to Full run and Combine. Your template is copied, never changed.",
              style="Hint.TLabel").grid(row=5, column=1, columnspan=2, sticky="w")

    # ---- run row -----------------------------------------------------------------------------------
    run_frame = ttk.Frame(form)
    run_frame.grid(row=next_row(), column=0, sticky="w", pady=(0, 10))
    run_btn = ttk.Button(run_frame, text="Run", style="Run.TButton", command=lambda: start_run())
    run_btn.pack(side="left")

    # ---- advanced ----------------------------------------------------------------------------------
    adv_toggle_var = tk.StringVar()
    adv_btn = ttk.Button(form, textvariable=adv_toggle_var, command=lambda: toggle_advanced())
    adv_btn.grid(row=next_row(), column=0, sticky="w")
    adv = ttk.Frame(form)
    adv.grid(row=next_row(), column=0, sticky="ew", pady=(6, 0))
    adv.columnconfigure(0, weight=1)

    adv_models = ttk.Labelframe(adv, text="Use a different model for one stage (empty = the model above)",
                                style="Section.TLabelframe", padding=(12, 8))
    adv_models.grid(row=0, column=0, sticky="ew", pady=(0, 10))
    adv_models.columnconfigure(1, weight=1)
    stage_combos: List[ttk.Combobox] = []
    for i, (label, key) in enumerate((("Controls", "controls_model"), ("P&ID", "pids_model"), ("Combine", "combine_model"))):
        ttk.Label(adv_models, text=label).grid(row=i, column=0, sticky="e", padx=(0, 8), pady=3)
        combo = ttk.Combobox(adv_models, textvariable=vars_[key], values=ordered_names())
        combo.grid(row=i, column=1, sticky="ew", pady=3)
        stage_combos.append(combo)

    adv_tuning = ttk.Labelframe(adv, text="Tuning (empty = automatic)", style="Section.TLabelframe", padding=(12, 8))
    adv_tuning.grid(row=1, column=0, sticky="ew", pady=(0, 10))
    adv_tuning.columnconfigure(1, weight=1)
    for i, (label, key, hint) in enumerate((
        ("Max output tokens", "max_tokens", f"automatic: the model's limit (else {DEFAULT_MAX_TOKENS})"),
        ("Render DPI", "dpi", f"automatic: {DEFAULT_DPI}"),
        ("OCR image scale", "ocr_scale", "3 to 4 is plenty; memory grows with its square"),
    )):
        ttk.Label(adv_tuning, text=label).grid(row=i, column=0, sticky="e", padx=(0, 8), pady=3)
        ttk.Entry(adv_tuning, textvariable=vars_[key], width=12).grid(row=i, column=1, sticky="w", pady=3)
        ttk.Label(adv_tuning, text=hint, style="Hint.TLabel").grid(row=i, column=2, sticky="w", padx=(8, 0))
    ttk.Checkbutton(adv_tuning, text="OCR: analyse table structure", variable=vars_["ocr_tables"]).grid(
        row=3, column=1, columnspan=2, sticky="w")
    ttk.Checkbutton(adv_tuning, text="OCR: describe pictures (slow)", variable=vars_["ocr_pictures"]).grid(
        row=4, column=1, columnspan=2, sticky="w")

    adv_ref = ttk.Labelframe(adv, text="Reference images (notation legends sent with each page)",
                             style="Section.TLabelframe", padding=(12, 8))
    adv_ref.grid(row=2, column=0, sticky="ew", pady=(0, 10))
    adv_ref.columnconfigure(1, weight=1)
    ttk.Checkbutton(adv_ref, text="Use the reference images that ship with this tool", variable=vars_["auto_reference"]).grid(
        row=0, column=0, columnspan=3, sticky="w")
    image_types = [("Images", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff"), ("All files", "*.*")]
    for i, (label, key) in enumerate((("Controls image", "controls_reference"), ("P&ID image", "pids_reference")), start=1):
        ttk.Label(adv_ref, text=label).grid(row=i, column=0, sticky="e", padx=(0, 8), pady=3)
        ttk.Entry(adv_ref, textvariable=vars_[key]).grid(row=i, column=1, sticky="ew", pady=3)
        ttk.Button(adv_ref, text="Browse...",
                   command=lambda k=key: vars_[k].set(filedialog.askopenfilename(filetypes=image_types) or vars_[k].get())
                   ).grid(row=i, column=2, padx=(8, 0))
    ttk.Label(adv_ref, text="A custom image here overrides the built-in one.", style="Hint.TLabel").grid(
        row=3, column=1, sticky="w")

    adv_wb = ttk.Labelframe(adv, text="Cable Calculations workbook step", style="Section.TLabelframe", padding=(12, 8))
    adv_wb.grid(row=3, column=0, sticky="ew", pady=(0, 10))
    adv_wb.columnconfigure(1, weight=1)
    ttk.Label(adv_wb, text="Template").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=3)
    ttk.Entry(adv_wb, textvariable=vars_["workbook_template"]).grid(row=0, column=1, sticky="ew", pady=3)
    ttk.Button(adv_wb, text="Browse...",
               command=lambda: vars_["workbook_template"].set(
                   filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm"), ("All files", "*.*")])
                   or vars_["workbook_template"].get())).grid(row=0, column=2, padx=(8, 0))
    def toggle_advanced() -> None:
        vars_["advanced_open"].set(not vars_["advanced_open"].get())
        apply_advanced()

    def apply_advanced() -> None:
        if vars_["advanced_open"].get():
            adv.grid()
            adv_toggle_var.set("▾ Advanced settings")
        else:
            adv.grid_remove()
            adv_toggle_var.set("▸ Advanced settings")

    def update_mode_visibility(*_args: object) -> None:
        mode = vars_["mode"].get()
        for widget in all_mode_widgets:
            widget.grid_remove()
        for widget in mode_rows.get(mode, []):
            widget.grid()
        sec_files.configure(text="2. Saved results to combine" if mode == "combine" else "2. Drawings")

    vars_["mode"].trace_add("write", update_mode_visibility)

    # ---- model discovery ---------------------------------------------------------------------------
    def apply_models(found: Dict[str, bool], local_count: int, error: Optional[str]) -> None:
        if found:
            catalog.clear()
            catalog.update(found)
        names = ordered_names()
        for combo in [model_combo] + stage_combos:
            combo.configure(values=names)
        if not vars_["model"].get().strip() and names:
            preferred = [n for n in names if catalog[n] and not claude_code.is_claude_code_model(n)
                         and ("vl" in n.lower() or "gemma" in n.lower())]
            vars_["model"].set((preferred or names)[0])
        update_model_widgets()
        if error:
            model_note.configure(text=error)
        refresh_btn.configure(state="normal")
        schedule_save()

    def refresh_models() -> None:
        refresh_btn.configure(state="disabled")
        api_key = vars_["api_key"].get().strip()

        def work() -> None:
            found: Dict[str, bool] = {}
            error = None
            local = ilt._get_local_ollama_models(LOCAL_HOST)
            for name in local:
                found[name] = True
            if not local:
                error = "Ollama isn't reachable on this computer (is it running?). Claude Code and cloud models still work."
            for name in claude_code.available_models():
                found[name] = True  # no API key involved
            for provider in _providers_for_key(api_key):
                for name in ilt._get_cloud_models(api_key, CLOUD_HOSTS[provider], provider):
                    found.setdefault(name, False)
            post(lambda: apply_models(found, len(local), error))

        threading.Thread(target=work, daemon=True).start()

    vars_["api_key"].trace_add("write", update_model_widgets)

    # ---- running -----------------------------------------------------------------------------------
    def resolve_endpoint(cfg: dict, model: str, api_key: str) -> dict:
        """Host, key and token limit for a model, worked out from the model name (runs on the worker thread)."""
        if claude_code.is_claude_code_model(model):
            return {"model": model, "host": "", "key": "", "max_tokens": DEFAULT_MAX_TOKENS}  # the CLI sets its own limits
        cloud = is_cloud(model)
        provider = _guess_provider(model, api_key) if cloud else "Local"
        host = CLOUD_HOSTS[provider] if cloud else LOCAL_HOST
        key = api_key if cloud else ""
        override = cfg["max_tokens"].strip()
        if override:
            tokens = int(override)
        else:
            try:
                tokens = int(ilt._get_model_capabilities(model, host, key, not cloud).get("max_tokens") or 0)
            except Exception:
                tokens = 0
            tokens = tokens or DEFAULT_MAX_TOKENS
        return {"model": model, "host": host, "key": key, "max_tokens": tokens}

    def out_folder(cfg: dict, pdf: str) -> Path:
        chosen = cfg["out_dir"].strip()
        return Path(chosen).expanduser() if chosen else Path(pdf).expanduser().resolve().parent

    def reference_for(cfg: dict, kind: str) -> Optional[str]:
        custom = cfg[f"{kind}_reference"].strip()
        if custom:
            return custom
        auto = REFERENCE_IMAGES[kind]
        return str(auto) if cfg["auto_reference"] and auto.is_file() else None

    def validate() -> List[str]:
        """Everything wrong with the current settings, in plain language (empty list = good to go)."""
        problems: List[str] = []
        mode = vars_["mode"].get()

        def need_file(key: str, label: str) -> Optional[Path]:
            text = vars_[key].get().strip()
            if not text:
                problems.append(f"Choose the {label}.")
                return None
            path = Path(text).expanduser()
            if not path.is_file():
                problems.append(f"{label} not found: {path}")
                return None
            return path

        controls = need_file("controls_pdf", "controls PDF") if mode in ("all", "controls") else None
        pids = need_file("pids_pdf", "P&ID PDF") if mode in ("all", "pids") else None
        if mode == "combine":
            need_file("controls_json", "controls results file")
            need_file("pids_json", "P&ID results file")
        for path, key, label in ((controls, "controls_pages", "Controls pages"), (pids, "pids_pages", "P&ID pages")):
            if path and vars_[key].get().strip():
                try:
                    ilt._resolve_pages_for_pdf(path, vars_[key].get().strip())
                except Exception as exc:
                    problems.append(f"{label}: {exc}")
        if vars_["fill_workbook"].get() and mode in ("all", "combine"):
            template = Path(vars_["workbook_template"].get().strip()).expanduser()
            if not template.is_file():
                problems.append(f"Cable Calculations template not found: {template} (or untick the workbook step).")
        if not vars_["model"].get().strip():
            problems.append("Pick a model.")
        elif is_cloud(vars_["model"].get().strip()) and not vars_["api_key"].get().strip():
            problems.append("This is a cloud model, so it needs an API key.")
        for key, label in (("max_tokens", "Max output tokens"), ("dpi", "Render DPI")):
            raw = vars_[key].get().strip()
            if raw and not raw.isdigit():
                problems.append(f"{label} must be a whole number (or empty for automatic).")
        try:
            if float(vars_["ocr_scale"].get().strip() or "3.0") <= 0:
                raise ValueError
        except ValueError:
            problems.append("OCR image scale must be a positive number.")
        return problems

    def build_args(cfg: dict) -> argparse.Namespace:
        """Translate the window into the namespace run_pipeline expects (endpoint lookups happen here)."""
        mode = cfg["mode"]
        api_key = cfg["api_key"].strip().strip()
        main_model = cfg["model"].strip()
        endpoints: Dict[str, dict] = {}

        def endpoint(model: str) -> dict:
            if model not in endpoints:
                endpoints[model] = resolve_endpoint(cfg, model, api_key)
            return endpoints[model]

        args = parser.parse_args([])
        args.run_mode = mode
        base = endpoint(main_model)
        args.model, args.api_host, args.api_key, args.max_tokens = base["model"], base["host"], base["key"], base["max_tokens"]
        args.dpi = int(cfg["dpi"].strip() or DEFAULT_DPI)

        for stage in ("controls", "pids", "combine"):
            ep = endpoint(cfg[f"{stage}_model"].strip() or main_model)
            setattr(args, f"{stage}_model", ep["model"])
            setattr(args, f"{stage}_api_host", ep["host"])
            setattr(args, f"{stage}_api_key", ep["key"])
            setattr(args, f"{stage}_max_tokens", ep["max_tokens"])

        controls_pdf = cfg["controls_pdf"].strip()
        pids_pdf = cfg["pids_pdf"].strip()
        if mode in ("all", "controls"):
            args.controls_pdf = controls_pdf
            args.controls_pages = cfg["controls_pages"].strip() or None
            args.controls_reference_image = reference_for(cfg, "controls")
            args.controls_output = str(out_folder(cfg, controls_pdf) / _auto_name(controls_pdf, "controls.json"))
        if mode in ("all", "pids"):
            args.pids_pdf = pids_pdf
            args.pids_pages = cfg["pids_pages"].strip() or None
            args.pids_reference_image = reference_for(cfg, "pids")
            args.pids_output = str(out_folder(cfg, pids_pdf) / _auto_name(pids_pdf, "pids.json"))
        if mode == "combine":
            args.controls_input_json = cfg["controls_json"].strip()
            args.pids_input_json = cfg["pids_json"].strip()
        if mode in ("all", "combine"):
            anchor = pids_pdf if mode == "all" else args.pids_input_json
            stem_source = anchor or controls_pdf
            args.combine_output = str(out_folder(cfg, stem_source) / _auto_name(stem_source, "combined.json"))

        args.docling_enable = cfg["use_ocr"]
        args.docling_do_ocr = True
        args.docling_do_table_structure = cfg["ocr_tables"]
        args.docling_do_picture_description = cfg["ocr_pictures"]
        args.docling_images_scale = float(cfg["ocr_scale"].strip() or "3.0")
        return args

    def export_excel(json_path: Path, log: Callable[[str], None]) -> Optional[Path]:
        """Write an .xlsx beside a results .json (per-page instrument lists, or the combined list)."""
        try:
            text = json_path.read_text(encoding="utf-8")
            data = json.loads(text)
            target = json_path.with_suffix(".xlsx")
            if "pages" in data:
                ilt.write_instruments_to_excel(text, target)
            elif isinstance(data.get("combined_instruments"), list):
                ilt.write_combined_to_excel(data, target)
            else:
                return None
            log(f"Saved Excel: {target}")
            return target
        except Exception as exc:
            log(f"Could not write Excel for {json_path.name}: {exc}")
            return None

    def set_running(running: bool) -> None:
        state["running"] = running
        run_btn.configure(state="disabled" if running else "normal", text="Running..." if running else "Run")
        if running:
            progress.start(12)
        else:
            progress.stop()

    def finish(error: Optional[str], outputs: List[Path], step_error: Optional[str] = None) -> None:
        set_running(False)
        state["last_outputs"] = [p for p in outputs if p.exists()]
        if step_error and not error:
            status("error", f"Results saved, but the workbook step failed: {step_error}")
            append_log(f"Workbook step failed: {step_error}")
            for path in state["last_outputs"]:
                append_log(f"Result: {path}")
        elif error:
            status("error", f"Failed: {error}")
            append_log(f"Run failed: {error}")
        else:
            status("ok", "Done. " + (f"{len(state['last_outputs'])} file(s) saved." if state["last_outputs"] else ""))
            if state.get("last_log"):
                append_log(f"Run log: {state['last_log']}")
            append_log("Run completed successfully.")
            for path in state["last_outputs"]:
                append_log(f"Result: {path}")
        open_folder_btn.configure(state="normal" if state["last_outputs"] else "disabled")

    def start_run() -> None:
        if state["running"]:
            return
        problems = validate()
        if problems:
            status("error", problems[0])
            messagebox.showwarning("Check your settings", "\n".join(f"• {p}" for p in problems))
            return
        for key, hist in (("controls_pdf", "controls_pdf"), ("pids_pdf", "pids_pdf"),
                          ("controls_json", "controls_json"), ("pids_json", "pids_json")):
            if vars_[key].get().strip():
                remember(hist, vars_[key].get())
        save_now()
        cfg = {k: (v.get() if isinstance(v, tk.BooleanVar) else v.get().strip()) for k, v in vars_.items()}
        set_running(True)
        open_folder_btn.configure(state="disabled")
        status("busy", "Starting...")

        try:
            run_log: Optional[RunLog] = RunLog(cfg["mode"])
        except OSError as exc:
            run_log = None
            append_log(f"Could not create a run log file: {exc}")
        if run_log:
            state["last_log"] = run_log.path
            open_log_btn.configure(state="normal")
            # The settings that shaped this run, minus the API key and the housekeeping entries.
            shown = {k: v for k, v in cfg.items() if k not in ("api_key", "advanced_open") and v not in ("", None)}
            run_log.line("Run settings: " + json.dumps(shown, ensure_ascii=False))
        append_log("Starting run..." + (f" (logging to {run_log.path})" if run_log else ""))

        def work() -> None:
            outputs: List[Path] = []
            error: Optional[str] = None

            def log(message: str) -> None:
                if run_log:
                    run_log.line(message)
                post(lambda: append_log(message))

            def stream(text: str) -> None:
                if run_log:
                    run_log.chunk(text)
                post(lambda: append_chunk(text))

            try:
                args = build_args(cfg)
                ilt.run_pipeline(args, log, stream)
                for attr in ("output", "controls_output", "pids_output", "combine_output"):
                    value = getattr(args, attr, None)
                    # Only count files this run could have produced for the chosen mode.
                    if value and Path(value).is_file() and Path(value).stat().st_mtime >= started:
                        outputs.append(Path(value))
                if cfg["save_excel"]:
                    for json_path in list(outputs):
                        excel = export_excel(json_path, log)
                        if excel:
                            outputs.append(excel)
            except Exception as exc:  # noqa: BLE001
                error = str(exc) or exc.__class__.__name__
                if run_log:
                    run_log.chunk(traceback.format_exc())

            step_error: Optional[str] = None
            combined = getattr(args, "combine_output", None) if not error else None
            if (not error and cfg["fill_workbook"] and cfg["mode"] in ("all", "combine")
                    and combined and Path(combined).is_file()):
                target = Path(combined).with_name(Path(combined).stem.replace(" - combined", "") + " - Cable Calculations.xlsx")

                def ask_retry(reason: str) -> bool:
                    """Ask on the UI thread whether to retry the workbook step; blocks this worker until answered."""
                    answer: Dict[str, bool] = {}
                    answered = threading.Event()

                    def show() -> None:
                        status("error", "The workbook step failed - waiting for your choice.")
                        answer["retry"] = messagebox.askretrycancel(
                            "Workbook step failed",
                            f"The final Cable Calculations workbook step failed:\n\n{reason}\n\n"
                            "Retry runs this step again (fix the problem first if you can, e.g. save the template).\n"
                            "Cancel ends the run; the results saved so far are kept.")
                        answered.set()

                    post(show)
                    answered.wait()
                    return answer["retry"]

                attempt = 0
                while True:
                    attempt += 1
                    post(lambda: status("busy", "Filling the Cable Calculations workbook..."))
                    try:
                        outputs.append(cable_step.fill_cable_workbook(
                            Path(combined), target,
                            template=Path(cfg["workbook_template"]).expanduser(), log=log))
                        step_error = None
                        break
                    except Exception as exc:  # noqa: BLE001 - keep the results already produced
                        step_error = str(exc) or exc.__class__.__name__
                        log(f"Workbook step attempt {attempt} failed: {step_error}")
                        if run_log:
                            run_log.chunk(traceback.format_exc())
                        if not ask_retry(step_error):
                            log("Run ended by the user after the workbook step failed.")
                            break
                        log("Retrying the workbook step...")
            if run_log:
                for path in outputs:
                    run_log.line(f"Result file: {path}")
                if error:
                    run_log.close(f"FAILED: {error}")
                elif step_error:
                    run_log.close(f"Pipeline completed, workbook step FAILED: {step_error}")
                else:
                    run_log.close("Run completed successfully")
            post(lambda: finish(error, outputs, step_error))

        started = datetime.datetime.now().timestamp() - 1
        threading.Thread(target=work, daemon=True).start()

    def _open_run_log() -> None:
        if state.get("last_log"):
            os.startfile(str(state["last_log"]))  # noqa: S606 - Windows-only tool

    def _open_output_folder() -> None:
        if state["last_outputs"]:
            os.startfile(str(state["last_outputs"][0].parent))  # noqa: S606 - Windows-only tool

    # ---- start up ----------------------------------------------------------------------------------
    def on_close() -> None:
        if state["running"] and not messagebox.askyesno("Run in progress", "A run is still in progress. Close anyway?"):
            return
        save_now()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    model_combo.configure(values=ordered_names())
    apply_advanced()
    update_mode_visibility()
    update_model_widgets()
    append_log("Ready. Choose your drawings and a model, then press Run.")
    pump()
    schedule_save()  # write the file on first launch too, so carried-over settings are kept
    refresh_models()
    root.mainloop()

    # Window closed: end the whole program. Stop any in-flight request and exit hard, because a model
    # worker thread (or a Docling/torch thread) could otherwise keep the console process alive.
    claude_code.kill_all()
    cable_step.kill_all()
    sys.stdout.flush()
    sys.stderr.flush()
    os._exit(0)
