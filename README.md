# Excel Working Paper Generator

Generates the missing `<year>` working paper sheet inside a `.xlsx` financial workbook, using the prior year's sheet as a template.

## Requirements

- Python 3.14 + `uv`
- macOS: `brew install python-tk@3.14` (needed for the desktop UI)
- The workbook must have a `מאזן <year>` sheet with no matching `<year>` working paper sheet.
- The prior year `<year-1>` working paper sheet must exist (used as the template).

---

## Run without UI (CLI)

```bash
uv run python run.py path/to/your-file.xlsx
```

Output is saved to `output/<same-filename>.xlsx` by default.

---

## Run the desktop UI

```bash
uv run python main.py
```

In the UI:
- **Edit file in place** (default) — overwrites the source file with the new sheet added.
- **Create new file** — choose an output directory; saves a copy there.

---

## Build a standalone executable (PyInstaller)

```bash
uv run pyinstaller --onefile --windowed --name excel-generator main.py
```

Output binary:
- macOS/Linux: `dist/excel-generator`
- Windows: `dist/excel-generator.exe`

---

## Folder layout

| Folder | Purpose |
|--------|---------|
| `example/` | Reference workbooks (all sheets present) |
| `input/` | Workbooks to process |
| `output/` | Default output location for CLI runs |
| `dist/` | PyInstaller build output |
