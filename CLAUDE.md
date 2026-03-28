# CLAUDE.md

## Project Purpose

Generate the missing `<year>` working paper sheet inside a `.xlsx` financial workbook, using the `מאזן <year>` trial balance as input and the prior year's working paper as the pattern template.

Full requirements: see `prd.md`.

## Stack

- Python 3.14, `uv` (no pip), full type annotations
- PyInstaller for `.exe` distribution
- Minimal file count — keep it simple

## Workbook Pattern

Every workbook has paired sheets: `מאזן <year>` (trial balance, 7 cols) + `<year>` (working paper, 11 cols).
The latest `מאזן <year>` will have no matching `<year>` sheet — that is what to generate.

## Source Sheet: `מאזן <year>` columns (A–G)

| Col | Header | Role |
|-----|--------|------|
| A | — | `None` on account rows; category label on section headers; `'סה"כ לקבוצה:'` on totals |
| B | מיון | Group number |
| C | חשבון | Account number |
| D | שם חשבון | Account name |
| E | חובה | Debit (or None) |
| F | זכות | Credit (or None) |
| G | הפרש | Debit − Credit |

## Target Sheet: `<year>` columns (A–K)

| Col | Header | Rule |
|-----|--------|------|
| A | קבוצה | Group number (from מיון) |
| B | מס' כרטיס | Account number |
| C | פרטים | Account name / section label |
| D | חובה | **Debit amount if account has debit, else empty** |
| E | זכות | **Credit amount if account has credit, else empty** |
| F | פ.נ | Opening balance = prior year's G (יתרה) for same account |
| G | יתרה | Formula: `=F+D−E` |
| H | הערות | Notes |
| I–K | — | Auxiliary (sometimes used) |

**D and E are mutually exclusive** — one has a value, the other is always empty.

## Processing Steps

1. Find latest `מאזן <year>` with no matching `<year>` sheet.
2. Load prior year's `<year>` working paper to get F (opening balances by account number).
3. Call Claude CLI with: prior year working paper structure + new year trial balance rows → returns structured new year rows.
4. Map trial balance → working paper: חובה→D, זכות→E, prior G→F, formula→G.
5. Write new sheet preserving section headers, group ordering, and subtotal rows.

## Components

| Name | Role |
|---|---|
| Processor | Read `.xlsx`, identify sheets, extract trial balance rows |
| Parser | Call Claude CLI to infer section structure and produce new year rows |
| Constructor | Write new `<year>` sheet with correct columns, formulas, and layout |
| UI | Desktop GUI: file picker → generate → output path |

## Testing

Two layers — run both in `pytest`:

**1. Deterministic (no AI)**
- D/E mutual exclusivity: exactly one of col D, col E is non-empty per account row.
- G = F + D − E (within float tolerance) for every account row.
- All account numbers from `מאזן <year>` are present in the output.
- Group numbers in col A match the trial balance groups.
- Col D and col E values are always positive.
- Headers in row 2 are correct.

**2. Claude CLI agent validator**
After generation, invoke a Claude CLI agent with:
- These rules (from `prd.md`).
- Generated `<year>` sheet content (JSON/CSV).
- Source `מאזן <year>` content (JSON/CSV).
- Prior year `<year>` working paper (JSON/CSV).

Agent checks what mechanical tests cannot: section header placement, col F matches prior year col G, subtotal rows correct, no phantom rows, overall structure mirrors prior year.

Agent returns `{ "passed": bool, "violations": [...], "summary": "..." }`.
Test fails if `passed` is `false` or `violations` is non-empty.

**Files**: `example/` = reference files, `input/` = test inputs, `output/` = generated output.

## Packaging

To build the standalone executable:

```bash
uv run pyinstaller --onefile --windowed --name excel-generator main.py
```

The binary is created at `dist/excel-generator` (macOS/Linux) or `dist/excel-generator.exe` (Windows).

Run it directly — no Python installation required.

Note: on macOS, Homebrew Python 3.14 does not ship with Tk support by default.
Install it first: `brew install python-tk@3.14`

## Restrictions
- Do not run commands with sudo or admin privileges.
- Do not write to system directories (e.g., `/usr`, `C:\Windows`).