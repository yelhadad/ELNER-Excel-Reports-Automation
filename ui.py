"""UI module: desktop GUI for the Excel Working Paper Generator."""
from __future__ import annotations

import os
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
from pathlib import Path

from processor import load_workbook_data
from constructor import generate_working_paper


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Working Paper Generator")
        self.geometry("500x250")
        self.resizable(False, False)

        # File selection row
        tk.Label(self, text="File:").grid(row=0, column=0, padx=10, pady=20, sticky="w")
        self._path_var = tk.StringVar()
        self._path_entry = tk.Entry(self, textvariable=self._path_var, width=45, state="readonly")
        self._path_entry.grid(row=0, column=1, padx=5, pady=20)
        tk.Button(self, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=10)

        # Generate button
        self._generate_btn = tk.Button(self, text="Generate", command=self.generate, width=20)
        self._generate_btn.grid(row=1, column=0, columnspan=3, pady=10)

        # Status label
        self._status_label = tk.Label(self, text="Ready", fg="gray")
        self._status_label.grid(row=2, column=0, columnspan=3, pady=5)

        # Output path label (hidden initially)
        self._output_label = tk.Label(self, text="", fg="blue", cursor="hand2")
        self._output_label.grid(row=3, column=0, columnspan=3, pady=2)
        self._output_label.grid_remove()

        # Open File button (hidden initially)
        self._open_btn = tk.Button(self, text="Open File")
        self._open_btn.grid(row=4, column=0, columnspan=3, pady=5)
        self._open_btn.grid_remove()

        self._output_path: Path | None = None

    def browse_file(self) -> None:
        path = tkinter.filedialog.askopenfilename(
            title="Select Excel workbook",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self._path_var.set(path)
            self._output_label.grid_remove()
            self._open_btn.grid_remove()
            self._status_label.config(text="Ready", fg="gray")

    def open_file(self, path: str) -> None:
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif sys.platform == "win32":
            os.startfile(path)
        else:
            subprocess.Popen(["xdg-open", path])

    def generate(self) -> None:
        path_str = self._path_var.get().strip()
        if not path_str:
            tkinter.messagebox.showerror("No file selected", "Please select an Excel file first.")
            return
        if not Path(path_str).exists():
            tkinter.messagebox.showerror("File not found", f"File does not exist:\n{path_str}")
            return

        self._generate_btn.config(state="disabled")
        self._status_label.config(text="Processing...", fg="blue")
        self._output_label.grid_remove()
        self._open_btn.grid_remove()

        thread = threading.Thread(target=self._run_pipeline, args=(path_str,), daemon=True)
        thread.start()

    def _run_pipeline(self, path_str: str) -> None:
        try:
            path = Path(path_str)
            data = load_workbook_data(path)
            output_path = generate_working_paper(data)
            self.after(0, self._on_success, output_path)
        except Exception as e:
            self.after(0, self._on_error, str(e))

    def _on_success(self, output_path: Path) -> None:
        self._output_path = output_path
        self._status_label.config(text="Done!", fg="green")
        self._output_label.config(text=str(output_path))
        self._output_label.grid()
        self._open_btn.config(command=lambda: self.open_file(str(output_path)))
        self._open_btn.grid()
        self._generate_btn.config(state="normal")

    def _on_error(self, message: str) -> None:
        tkinter.messagebox.showerror("Error", message)
        self._status_label.config(text="Error", fg="red")
        self._generate_btn.config(state="normal")


def run_app() -> None:
    """Create and run the application."""
    app = App()
    app.mainloop()
