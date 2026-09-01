from __future__ import annotations

import os
import subprocess
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import messagebox, ttk

from build_presentation import DEFAULT_CONFIG, ROOT, build_presentation, configured_output_path, load_model
from preview import generate_preview


def open_path(path: Path) -> None:
    path = path.resolve()
    if os.name == "nt":
        os.startfile(path)  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", str(path)])


class Manager(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("USD Cash Allocation — Python Presentation Builder")
        self.geometry("820x570")
        self.minsize(760, 520)
        self.configure(bg="#EEF2F5")
        self._build_ui()
        self.refresh_summary()

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg="#071B27", padx=24, pady=20)
        header.pack(fill="x")
        tk.Label(header, text="Presentation Builder", bg="#071B27", fg="white", font=("Arial", 22, "bold")).pack(anchor="w")
        tk.Label(header, text="Spreadsheet-driven · no VBA · no ActiveX · no PowerPoint required", bg="#071B27", fg="#B9C8D6", font=("Arial", 10)).pack(anchor="w", pady=(4, 0))

        body = tk.Frame(self, bg="#EEF2F5", padx=24, pady=20)
        body.pack(fill="both", expand=True)
        self.summary = tk.StringVar(value="Loading model…")
        tk.Label(body, textvariable=self.summary, bg="#EEF2F5", fg="#4A5C6E", font=("Arial", 11)).pack(anchor="w", pady=(0, 14))

        button_row = tk.Frame(body, bg="#EEF2F5")
        button_row.pack(fill="x")
        self.build_button = ttk.Button(button_row, text="1  Build presentation", command=self.build_async)
        self.build_button.grid(row=0, column=0, padx=(0, 10), pady=5, sticky="ew")
        ttk.Button(button_row, text="2  Refresh preview", command=self.preview_async).grid(row=0, column=1, padx=(0, 10), pady=5, sticky="ew")
        ttk.Button(button_row, text="Open preview", command=self.open_preview).grid(row=0, column=2, pady=5, sticky="ew")
        ttk.Button(button_row, text="Open spreadsheet", command=lambda: open_path(DEFAULT_CONFIG)).grid(row=1, column=0, padx=(0, 10), pady=5, sticky="ew")
        ttk.Button(button_row, text="Open output folder", command=lambda: open_path(ROOT / "output")).grid(row=1, column=1, padx=(0, 10), pady=5, sticky="ew")
        ttk.Button(button_row, text="Open presentation", command=self.open_presentation).grid(row=1, column=2, pady=5, sticky="ew")
        for col in range(3):
            button_row.columnconfigure(col, weight=1)

        tk.Label(body, text="Activity", bg="#EEF2F5", fg="#0B1826", font=("Arial", 11, "bold")).pack(anchor="w", pady=(20, 6))
        self.log_box = tk.Text(body, height=13, bg="white", fg="#0B1826", relief="flat", padx=12, pady=10, font=("Consolas", 9), state="disabled")
        self.log_box.pack(fill="both", expand=True)
        self.log("Ready. Edit the yellow cells in the spreadsheet, then build.")

    def log(self, message: str) -> None:
        def write() -> None:
            self.log_box.configure(state="normal")
            self.log_box.insert("end", message + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, write)

    def refresh_summary(self) -> None:
        try:
            model = load_model(DEFAULT_CONFIG)
            enabled = sum(bool(row.get("Enabled")) for row in model["objects"])
            exact = model["settings"].get("USE_EXACT_RUN_FORMATTING", True)
            self.summary.set(f"{len(model['slides'])} slides · {enabled} objects · {len(model['runs'])} text runs · exact formatting: {exact}")
        except Exception as exc:
            self.summary.set(f"Configuration error: {exc}")

    def _run(self, action) -> None:
        self.build_button.configure(state="disabled")
        try:
            action()
            self.refresh_summary()
        except Exception as exc:
            self.log(f"ERROR: {exc}")
            self.after(0, lambda: messagebox.showerror("Presentation Builder", str(exc)))
        finally:
            self.after(0, lambda: self.build_button.configure(state="normal"))

    def build_async(self) -> None:
        threading.Thread(target=lambda: self._run(lambda: build_presentation(DEFAULT_CONFIG, logger=self.log)), daemon=True).start()

    def preview_async(self) -> None:
        def action() -> None:
            self.log("Refreshing browser preview…")
            output = generate_preview(DEFAULT_CONFIG)
            self.log(f"Created {output.name}")
        threading.Thread(target=lambda: self._run(action), daemon=True).start()

    def open_preview(self) -> None:
        output = generate_preview(DEFAULT_CONFIG)
        webbrowser.open(output.as_uri())

    def open_presentation(self) -> None:
        output_path = configured_output_path(DEFAULT_CONFIG)
        if not output_path.exists():
            messagebox.showinfo("Presentation Builder", "Build the presentation first.")
            return
        open_path(output_path)


if __name__ == "__main__":
    Manager().mainloop()
