# Excel Working Paper Generator

Generates the missing `<year>` working paper sheet inside a `.xlsx` financial workbook, using the prior year's sheet as a template.

## Run on an input file (no UI)

```bash
uv run python run.py input/your-file.xlsx
```

Replace `input/your-file.xlsx` with the actual path to your workbook.

The output is saved to `output/<same-filename>.xlsx`.

## Requirements

- The workbook must have a `מאזן <year>` sheet with no matching `<year>` working paper sheet.
- The prior year `<year-1>` working paper sheet must exist (used as the template).

## Run the desktop UI

```bash
uv run python main.py
```

## Input / Output folders

| Folder | Purpose |
|--------|---------|
| `example/` | Reference workbooks (already have all sheets) |
| `input/` | Workbooks to process (missing the new year sheet) |
| `output/` | Generated workbooks (written here automatically) |
