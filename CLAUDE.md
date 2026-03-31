# CLAUDE.md

## Project Purpose

Generate the missing `<year>` working paper sheet inside a `.xlsx` financial workbook, using the `מאזן <year>` trial balance as input and the prior year's working paper as the pattern template.

## Stack

- Python 3.14, `uv` (no pip), full type annotations
- PyInstaller for `.exe` distribution
- Minimal file count — keep it simple
- **Always edit files with `cat > file << 'EOF'` heredoc, never with the built-in Edit/Write tools**

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
| F | פ.נ | **Always empty** — do NOT populate from prior year balances |
| G | יתרה | Formula: `=D+E` (always positive — deposit balances are credit so G = D+E, not D−E) |
| H | הערות | Notes / formulas from prior year template (preserved with row-ref adjustments) |
| I–K | — | Auxiliary (sometimes used) |

**D and E are mutually exclusive** — one has a value, the other is always empty. Both always positive.

**sum(D) must equal sum(E)** across all account rows in the generated sheet.

## Row Classification

| Category | Condition | Fill colour |
|----------|-----------|-------------|
| **Existing** (green) | In prior year template AND in new trial balance | `FF92D050` |
| **Missing** (yellow) | In prior year template but NOT in new trial balance | `FFFFF00` — D and E empty |
| **New** (blue) | In new trial balance but NOT in prior year template | `FF00B0F0` — inserted after last account of same group |

## Account Number Normalisation (`_normalize_account`)

Canonical form is always a plain integer string. Both int `253` and string `'0253'` normalise to `'253'`. This is critical for matching accounts across sheets where the same account may be stored as int in one sheet and zero-padded string in another.

## Account Cell Detection (`_is_account_cell`)

Accepts `int`, non-bool `float`, and purely-numeric strings (e.g. `'1005'`). Some workbooks store account numbers as strings in the prior year working paper — the check must handle both formats.

## Duplicate Account Rows

Some prior year templates contain the same account number twice. Handling:
- First occurrence: green with D/E values from trial balance
- Second occurrence (if account IS in new TB): green but D/E empty (avoids double-counting in sums)
- Second occurrence (if account NOT in new TB): yellow, D/E empty

## Processing Steps

1. Find latest `מאזן <year>` with no matching `<year>` sheet.
2. Load prior year's `<year>` working paper as the layout/style template.
3. Build output rows: all prior-year template rows in order, with new accounts injected after the last existing account of their group.
4. Accounts with no matching group in template are inserted just before grand-total rows (NOT at the very end) to avoid circular references in SUM formulas.
5. Write new sheet: copy styles from template, update D/E from trial balance, leave F empty, set G = `=D+E`.
6. Copy RTL (right-to-left) setting from prior year sheet: `new_ws.sheet_view.rightToLeft = prior_ws.sheet_view.rightToLeft`
7. Second pass: remap all formula row-references in ALL columns; extend `=SUM(…)` ranges to cover newly inserted/appended accounts.

## SUM Formula Extension (critical)

`_extend_sum_formula` handles ALL columns (D, E, G, H, and others), not just H:
- **Inserted accounts** (`inserted_after_orig_row` is not None): extend a SUM if the account was inserted after a template row within the SUM's original row range (section-scoped).
- **Appended accounts** (`inserted_after_orig_row` is None — currently not used, see below): extend only for grand-total SUMs whose original end row reaches the last template account row.

**Grand total rows** (e.g. `=SUM(D2:D296)` in col D) must be remapped AND extended to cover all new accounts. Previously only H and G were remapped — D and E were left stale, causing the formula totals to differ from the raw sums.

**Circular reference prevention**: accounts with no group match are inserted just BEFORE the grand total rows (not after), so the grand total formula range never includes its own row.

## Components

| Name | Role |
|---|---|
| `processor.py` | Read `.xlsx`, identify sheet pairs, extract trial balance rows and prior year data |
| `parser.py` | (Legacy) Claude CLI integration — not used in current deterministic flow |
| `constructor.py` | Write new `<year>` sheet with correct columns, formulas, colours, RTL, and layout |
| `ui.py` | Desktop GUI: file picker → generate → output path |
| `run.py` | CLI entry point: `uv run python run.py <path>` |
| `main.py` | PyInstaller entry point |

## Key Implementation Details (constructor.py)

- `_normalize_account`: strips leading zeros from numeric strings; float/int → plain int string.
- `_is_account_cell`: accepts int, float, and purely-numeric strings.
- `_build_output_rows`: builds ordered row list; sets `inserted_after_orig_row` on each new row.
- `_write_output_rows`: two-pass write. First pass writes rows and builds `orig_to_new` + `new_account_insertions`. Second pass remaps and extends ALL formula columns.
- `_extend_sum_formula`: general SUM extension for any column. Section-scoped for inserted accounts, grand-total-scoped for appended accounts.
- `_split_debit_credit`: if both debit and credit are non-zero, nets them into dominant column only.
- RTL set via `new_ws.sheet_view.rightToLeft = prior_ws.sheet_view.rightToLeft` after sheet creation.

## Known Workbook Variations

- **Account numbers as strings**: some workbooks (e.g. ס.ג.ק) store account numbers as strings `'1005'` in the working paper col B. `_is_account_cell` handles this.
- **Leading-zero accounts**: e.g. `'0253'` in trial balance vs int `253` in prior year. Normalisation strips leading zeros.
- **Duplicate accounts in template**: handled by `assigned_accounts` set — second occurrence gets green colour but no D/E values.
- **No group match**: new accounts whose group doesn't appear in the prior year template are inserted just before grand-total rows (not appended at file end).

## Testing

**Deterministic checks (run on every output file):**
- sum(D) == sum(E) (within 0.01 tolerance)
- F column always empty (F non-null count = 0)
- No TB account incorrectly yellow (all TB accounts are green or blue)
- No TB account missing from output
- No real circular references (formula in cell Xcol_n references Xcol_n itself)

**Files**: `example/` = reference files, `input/` = test inputs, `output/` = generated output.

Run all checks: `uv run python - << 'EOF'` with the verification script used in development.

## Packaging

```bash
uv run pyinstaller --onefile --windowed --name excel-generator main.py
```

Binary at `dist/excel-generator` (macOS/Linux) or `dist/excel-generator.exe` (Windows).

Note: on macOS, Homebrew Python 3.14 does not ship with Tk support by default.
Install first: `brew install python-tk@3.14`

## Restrictions
- Do not run commands with sudo or admin privileges.
- Do not write to system directories (e.g., `/usr`, `C:\Windows`).
- Always edit files with `cat > file << 'ENDOFFILE'` heredoc. Never use the built-in Edit or Write tools.
