"""UI module: desktop GUI for the Excel Working Paper Generator."""
from __future__ import annotations

import os
import subprocess
import sys
import threading
from pathlib import Path

import customtkinter as ctk
import tkinter.filedialog
import tkinter.messagebox

from processor import load_workbook_data
from constructor import generate_working_paper

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Working Paper Generator")
        self.geometry("580x380")
        self.resizable(False, False)
        self._full_path: str = ""
        self._dir_full: str = ""
        self._output_path: Path | None = None

        # Main padding frame
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=32, pady=28)

        # Title
        ctk.CTkLabel(
            main,
            text="Working Paper Generator",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(anchor="w", pady=(0, 20))

        # File picker row
        file_frame = ctk.CTkFrame(main, fg_color="transparent")
        file_frame.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(file_frame, text="Excel File", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(0, 4))

        file_row = ctk.CTkFrame(file_frame, fg_color="transparent")
        file_row.pack(fill="x")

        self._filename_var = ctk.StringVar(value="No file selected")
        self._file_entry = ctk.CTkEntry(
            file_row,
            textvariable=self._filename_var,
            state="readonly",
            font=ctk.CTkFont(size=13),
        )
        self._file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkButton(
            file_row,
            text="Browse",
            width=90,
            command=self.browse_file,
        ).pack(side="right")

        # Output mode
        mode_frame = ctk.CTkFrame(main, fg_color="transparent")
        mode_frame.pack(fill="x", pady=(12, 0))

        ctk.CTkLabel(mode_frame, text="Output Mode", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(0, 6))

        self._mode_var = ctk.StringVar(value="inplace")

        ctk.CTkRadioButton(
            mode_frame,
            text="Edit file in place",
            variable=self._mode_var,
            value="inplace",
            command=self._on_mode_change,
        ).pack(anchor="w", pady=2)

        newfile_row = ctk.CTkFrame(mode_frame, fg_color="transparent")
        newfile_row.pack(fill="x", pady=2)

        ctk.CTkRadioButton(
            newfile_row,
            text="Create new file",
            variable=self._mode_var,
            value="newfile",
            command=self._on_mode_change,
        ).pack(side="left")

        self._dirname_var = ctk.StringVar(value="No folder selected")
        self._dir_entry = ctk.CTkEntry(
            newfile_row,
            textvariable=self._dirname_var,
            state="readonly",
            font=ctk.CTkFont(size=13),
            width=220,
        )
        self._dir_btn = ctk.CTkButton(
            newfile_row,
            text="Choose Folder",
            width=110,
            command=self.browse_dir,
        )
        self._dir_entry.pack(side="left", padx=(12, 8), fill="x", expand=True)
        self._dir_btn.pack(side="left")
        self._dir_entry.pack_forget()
        self._dir_btn.pack_forget()

        # Generate button
        self._generate_btn = ctk.CTkButton(
            main,
            text="Generate",
            height=42,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.generate,
        )
        self._generate_btn.pack(fill="x", pady=(24, 0))

        # Status + progress row
        status_frame = ctk.CTkFrame(main, fg_color="transparent")
        status_frame.pack(fill="x", pady=(10, 0))

        self._status_label = ctk.CTkLabel(
            status_frame,
            text="Ready",
            font=ctk.CTkFont(size=13),
            text_color="gray",
        )
        self._status_label.pack(side="left")

        self._progress = ctk.CTkProgressBar(status_frame, mode="indeterminate", width=140)
        self._progress.pack(side="right")
        self._progress.pack_forget()

        # Output row (hidden until success)
        self._output_frame = ctk.CTkFrame(main, fg_color="transparent")
        self._output_frame.pack(fill="x", pady=(6, 0))

        self._output_label = ctk.CTkLabel(
            self._output_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("#1a73e8", "#4da6ff"),
            cursor="hand2",
        )
        self._output_label.pack(side="left")

        self._open_btn = ctk.CTkButton(
            self._output_frame,
            text="Open File",
            width=90,
            height=30,
            fg_color="transparent",
            border_width=1,
            command=self._open_output,
        )
        self._open_btn.pack(side="right")

        self._output_frame.pack_forget()

    def _on_mode_change(self) -> None:
        if self._mode_var.get() == "newfile":
            self._dir_entry.pack(side="left", padx=(12, 8), fill="x", expand=True)
            self._dir_btn.pack(side="left")
        else:
            self._dir_entry.pack_forget()
            self._dir_btn.pack_forget()

    def browse_file(self) -> None:
        path = tkinter.filedialog.askopenfilename(
            title="Select Excel workbook",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self._full_path = path
            self._filename_var.set(Path(path).name)
            self._output_frame.pack_forget()
            self._status_label.configure(text="Ready", text_color="gray")

    def browse_dir(self) -> None:
        d = tkinter.filedialog.askdirectory(title="Select output directory")
        if d:
            self._dir_full = d
            self._dirname_var.set(Path(d).name or d)

    def _open_output(self) -> None:
        if self._output_path:
            if sys.platform == "darwin":
                subprocess.Popen(["open", str(self._output_path)])
            elif sys.platform == "win32":
                os.startfile(str(self._output_path))
            else:
                subprocess.Popen(["xdg-open", str(self._output_path)])

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

        self._generate_btn.configure(state="disabled")
        self._output_frame.pack_forget()
        self._status_label.configure(text="Processing…", text_color=("#1a73e8", "#4da6ff"))
        self._progress.pack(side="right")
        self._progress.start()

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
        self._progress.stop()
        self._progress.pack_forget()
        self._status_label.configure(text="Done!", text_color=("#1e8c3a", "#4caf50"))
        self._output_label.configure(text=output_path.name)
        self._output_frame.pack(fill="x", pady=(6, 0))
        self._generate_btn.configure(state="normal")

    def _on_error(self, message: str) -> None:
        self._progress.stop()
        self._progress.pack_forget()
        tkinter.messagebox.showerror("Error", message)
        self._status_label.configure(text="Error", text_color=("#c0392b", "#e74c3c"))
        self._generate_btn.configure(state="normal")


def run_app() -> None:
    """Create and run the application."""
    app = App()
    app.mainloop()
