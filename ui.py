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
        self.geometry("600x340")
        self.resizable(False, False)
        self._full_path: str = ""
        self._dir_full: str = ""

        # Centre frame
        frame = tk.Frame(self)
        frame.place(relx=0.5, rely=0.5, anchor="center")

        # File row
        tk.Label(frame, text="File:").grid(row=0, column=0, padx=8, pady=12, sticky="e")
        self._filename_var = tk.StringVar(value="No file selected")
        tk.Label(frame, textvariable=self._filename_var, width=30, anchor="w",
                 relief="sunken", bg="white", fg="black").grid(row=0, column=1, padx=6, pady=12)
        tk.Button(frame, text="Browse…", command=self.browse_file).grid(row=0, column=2, padx=8)

        # Mode radios
        self._mode_var = tk.StringVar(value="inplace")
        tk.Radiobutton(frame, text="Edit file in place", variable=self._mode_var, value="inplace",
                       command=self._on_mode_change).grid(row=1, column=0, columnspan=3, sticky="w", padx=8, pady=2)
        tk.Radiobutton(frame, text="Create new file", variable=self._mode_var, value="newfile",
                       command=self._on_mode_change).grid(row=2, column=0, sticky="w", padx=8)

        # Dir picker (hidden until "new file" selected)
        self._dirname_var = tk.StringVar(value="No folder selected")
        self._dir_label = tk.Label(frame, textvariable=self._dirname_var, width=30, anchor="w",
                                   relief="sunken", bg="white", fg="black")
        self._dir_btn = tk.Button(frame, text="Choose Dir…", command=self.browse_dir)
        self._dir_label.grid(row=2, column=1, padx=6)
        self._dir_btn.grid(row=2, column=2, padx=8)
        self._dir_label.grid_remove()
        self._dir_btn.grid_remove()

        # Generate
        self._generate_btn = tk.Button(frame, text="Generate", command=self.generate, width=22)
        self._generate_btn.grid(row=3, column=0, columnspan=3, pady=16)

        # Status
        self._status_label = tk.Label(frame, text="Ready", fg="gray")
        self._status_label.grid(row=4, column=0, columnspan=3)

        # Output name (hidden)
        self._output_label = tk.Label(frame, text="", fg="blue", cursor="hand2")
        self._output_label.grid(row=5, column=0, columnspan=3, pady=2)
        self._output_label.grid_remove()

        # Open button (hidden)
        self._open_btn = tk.Button(frame, text="Open File")
        self._open_btn.grid(row=6, column=0, columnspan=3, pady=4)
        self._open_btn.grid_remove()

        self._output_path: Path | None = None

    def _on_mode_change(self) -> None:
        if self._mode_var.get() == "newfile":
            self._dir_label.grid()
            self._dir_btn.grid()
        else:
            self._dir_label.grid_remove()
            self._dir_btn.grid_remove()

    def browse_file(self) -> None:
        path = tkinter.filedialog.askopenfilename(
            title="Select Excel workbook",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self._full_path = path
            self._filename_var.set(Path(path).name)
            self._output_label.grid_remove()
            self._open_btn.grid_remove()
            self._status_label.config(text="Ready", fg="gray")

    def browse_dir(self) -> None:
        d = tkinter.filedialog.askdirectory(title="Select output directory")
        if d:
            self._dir_full = d
            self._dirname_var.set(Path(d).name or d)

    def open_file(self, path: str) -> None:
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif sys.platform == "win32":
            os.startfile(path)
        else:
            subprocess.Popen(["xdg-open", path])

    def generate(self) -> None:
        if not self._full_path:
            tkinter.messagebox.showerror("No file selected", "Please select an Excel file first.")
            return
        if not Path(self._full_path).exists():
            tkinter.messagebox.showerror("File not found", f"File does not exist:\n{self._full_path}")
            return

        if self._mode_var.get() == "inplace":
            output_path = Path(self._full_path)
        else:
            if not self._dir_full:
                tkinter.messagebox.showerror("No directory", "Please choose an output directory.")
                return
            output_path = Path(self._dir_full) / Path(self._full_path).name

        self._generate_btn.config(state="disabled")
        self._status_label.config(text="Processing…", fg="blue")
        self._output_label.grid_remove()
        self._open_btn.grid_remove()

        threading.Thread(target=self._run_pipeline, args=(self._full_path, output_path), daemon=True).start()

    def _run_pipeline(self, path_str: str, output_path: Path) -> None:
        try:
            data = load_workbook_data(Path(path_str))
            result = generate_working_paper(data, output_path=output_path)
            self.after(0, self._on_success, result)
        except Exception as e:
            self.after(0, self._on_error, str(e))

    def _on_success(self, output_path: Path) -> None:
        self._output_path = output_path
        self._status_label.config(text="Done!", fg="green")
        self._output_label.config(text=output_path.name)
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
