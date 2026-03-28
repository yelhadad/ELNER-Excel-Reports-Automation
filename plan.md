# Implementation Plan — Excel Year Working Paper Generator

## Overview

Build a Python desktop tool that generates a new year's working paper sheet (`<year>`) inside an existing `.xlsx` workbook, using the `מאזן <year>` trial balance as input and the prior year's working paper as a pattern template.

---

## File Structure

```
excel-automation-4/
├── main.py           # Entry point: launches UI
├── processor.py      # Sheet detection + trial balance extraction
├── parser.py         # Claude CLI integration → structured new-year rows
├── constructor.py    # Writes new <year> sheet into the workbook
├── ui.py             # Desktop GUI (tkinter)
├── tests/
│   ├── test_deterministic.py
│   └── test_agent_validator.py
├── example/          # Reference .xlsx files (already present)
├── input/            # Test inputs (already present)
├── output/           # Generated outputs
├── pyproject.toml
├── prd.md
├── CLAUDE.md
└── plan.md
```

---

## Step-by-Step Implementation

### Step 1 — Project Setup

- Initialize `uv` project: `uv init`, configure `pyproject.toml`.
- Dependencies: `openpyxl`, `tkinter` (stdlib), `pytest`.
- Add PyInstaller as a dev dependency for `.exe` packaging.

**pyproject.toml key sections:**
```toml
[project]
name = "excel-automation"
requires-python = ">=3.14"
dependencies = ["openpyxl"]

[tool.pytest.ini_options]
testpaths = ["tests"]
```

---

### Step 2 — `processor.py`

Responsibilities:
1. Open the `.xlsx` workbook.
2. Identify all sheet pairs (`מאזן <year>` + `<year>`).
3. Find the latest `מאזן <year>` that has **no** matching `<year>` sheet.
4. Identify the **prior year** working paper sheet.
5. Extract trial balance rows from the target `מאזן <year>` sheet.
6. Extract opening balances (col G values keyed by account number) from the prior year sheet.

**Key types:**
```python
@dataclass
class TrialBalanceRow:
    row_type: Literal["section_header", "account", "group_total", "title", "header"]
    group: str | None          # col B
    account_number: str | None # col C
    account_name: str | None   # col D
    debit: float | None        # col E
    credit: float | None       # col F
    net: float | None          # col G (הפרש)
    label: str | None          # col A (section label or 'סה"כ לקבוצה:')

@dataclass
class WorkbookData:
    workbook_path: Path
    new_year: int
    prior_year: int
    trial_balance_rows: list[TrialBalanceRow]
    prior_year_balances: dict[str, float]  # account_number → prior year G (יתרה)
```

**Row classification logic (col A):**
- `None` and B/C populated → `account`
- Non-None and B/C are None → `section_header`
- `== 'סה"כ לקבוצה:'` → `group_total`
- Rows 1–3 → `title`
- Row 4 → `header`

---

### Step 3 — `parser.py`

Responsibilities:
- Call Claude CLI with the prior year's working paper structure + new year's trial balance rows.
- Receive structured JSON back describing the new year's rows (including section header placement and subtotals).
- Return a list of `WorkingPaperRow` objects.

**Approach:**
- Serialize prior year working paper and new trial balance to JSON/CSV strings.
- Invoke `claude -p "<prompt>"` via `subprocess`.
- Parse the JSON response.

**Prompt structure (sent to Claude):**
```
You are generating a working paper sheet for year <new_year>.

Prior year (<prior_year>) working paper (JSON):
<prior_year_sheet_json>

New year (<new_year>) trial balance (JSON):
<trial_balance_json>

Rules:
- Every account row must appear exactly once.
- Section headers precede their group (same placement as prior year).
- Group-total rows follow their group.
- D and E are mutually exclusive (debit > 0 → D, credit > 0 → E).
- Col F = prior year col G for the same account number.
- Col G formula = F + D − E.
- Preserve col H (notes) from prior year where account matches.

Return JSON array of rows:
[
  {"row_type": "section_header"|"account"|"group_total",
   "group": "<group_number>",
   "account_number": "<acct>",
   "details": "<label or name>",
   "debit": <number or null>,
   "credit": <number or null>,
   "opening_balance": <number or null>,
   "notes": "<text or null>"}
]
```

**Key type:**
```python
@dataclass
class WorkingPaperRow:
    row_type: Literal["section_header", "account", "group_total"]
    group: str | None
    account_number: str | None
    details: str | None
    debit: float | None         # col D (positive or None)
    credit: float | None        # col E (positive or None)
    opening_balance: float | None  # col F
    notes: str | None           # col H
```

---

### Step 4 — `constructor.py`

Responsibilities:
- Accept a `WorkbookData` + `list[WorkingPaperRow]`.
- Create a new sheet named `<year>` in the workbook.
- Write all rows with correct column mapping.
- Insert Excel formula for col G.
- Save the workbook (to output path or in-place).

**Column mapping:**

| Col | Content |
|-----|---------|
| A | group number (account/total rows) or empty (section headers) |
| B | account number |
| C | details / section label |
| D | debit (if debit row, else empty) |
| E | credit (if credit row, else empty) |
| F | opening balance |
| G | `=F{n}+D{n}-E{n}` formula |
| H | notes |
| I–K | empty |

**Row 1**: title (e.g., company name from prior year or blank).
**Row 2**: headers `['קבוצה', "מס' כרטיס", 'פרטים', 'חובה', 'זכות', 'פ.נ', 'יתרה', 'הערות']`.
**Row 3+**: data rows.

**Sheet insertion position**: insert after the last `מאזן <year>` sheet (before `דוחות כספיים` if present).

---

### Step 5 — `ui.py`

Simple `tkinter` GUI:

```
┌─────────────────────────────────────────┐
│  Excel Working Paper Generator          │
│                                         │
│  File: [path/to/file.xlsx    ] [Browse] │
│                                         │
│            [ Generate ]                 │
│                                         │
│  Status: Ready / Processing... / Done   │
│  Output: path/to/output.xlsx            │
└─────────────────────────────────────────┘
```

- File picker filters for `.xlsx`.
- "Generate" triggers the full pipeline in a background thread to keep UI responsive.
- Errors shown in a message box in plain Hebrew/English.
- On success: display output path, offer "Open File" button.

---

### Step 6 — `main.py`

```python
from ui import run_app

if __name__ == "__main__":
    run_app()
```

---

### Step 7 — Tests

#### `tests/test_deterministic.py`

Uses files from `input/` (or `example/`) — runs the full pipeline and checks the generated sheet mechanically:

| Test | Assertion |
|------|-----------|
| `test_headers` | Row 2 = expected header list |
| `test_d_e_mutual_exclusivity` | For every account row: exactly one of D, E is non-empty |
| `test_g_formula` | For every account row: G == F + D − E (±0.01 tolerance) |
| `test_all_accounts_present` | Every account number from trial balance is in output |
| `test_group_numbers_match` | Col A group numbers match trial balance |
| `test_d_positive` | All col D values > 0 |
| `test_e_positive` | All col E values > 0 |

#### `tests/test_agent_validator.py`

After generation, invokes Claude CLI as a validator agent. Passes:
1. Rules summary from `prd.md`.
2. Generated `<year>` sheet as JSON.
3. Source `מאזן <year>` as JSON.
4. Prior year `<year>` working paper as JSON.

Expects `{"passed": true, "violations": [], "summary": "..."}`.
Test fails if `passed` is `false` or `violations` is non-empty.

---

### Step 8 — PyInstaller Packaging

```bash
uv run pyinstaller --onefile --windowed --name excel-generator main.py
```

- `--windowed`: no console window on Windows.
- Output: `dist/excel-generator.exe`.

---

## Data Flow

```
.xlsx file
    │
    ▼
processor.py
    ├── trial_balance_rows (new year)
    └── prior_year_balances (account → G value)
    │
    ▼
parser.py  ←──── Claude CLI
    └── working_paper_rows
    │
    ▼
constructor.py
    └── writes <year> sheet → output .xlsx
    │
    ▼
ui.py
    └── displays output path to user
```

---

## Implementation Order

1. `pyproject.toml` + `uv` setup
2. `processor.py` — read/classify/extract (testable immediately with example files)
3. `constructor.py` — write sheet (testable with manually constructed rows)
4. `parser.py` — Claude CLI integration
5. `tests/test_deterministic.py`
6. `ui.py`
7. `main.py`
8. `tests/test_agent_validator.py`
9. PyInstaller packaging

---

## Open Questions / Risks

| Item | Notes |
|------|-------|
| Claude CLI availability | Must be installed and on PATH; `parser.py` should fail fast with a clear error if not found |
| Hebrew RTL in openpyxl | Column order is LTR in the file; display is RTL — this is handled by Excel, not the code |
| `.xls` vs `.xlsx` | Example folder has `.xls` files; tool targets `.xlsx` only (per PRD) |
| Group-total row reconstruction | Claude parser handles this; deterministic tests verify totals match |
| Sheet insertion order | Insert new `<year>` sheet immediately after its `מאזן <year>` |
