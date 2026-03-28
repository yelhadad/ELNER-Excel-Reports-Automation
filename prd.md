# Product Requirements Document — Excel Year Working Paper Generator

## Overview

A desktop tool for accountants that automatically generates a new year's working paper sheet (`<year>`) inside an existing Excel workbook. The user provides a `.xlsx` file containing a trial balance sheet (`מאזן <year>`), and the tool produces an updated file with the full working paper sheet populated correctly.

---

## Problem Statement

Financial workbooks follow a paired-sheet structure: each year has a raw trial balance sheet (`מאזן <year>`) and a derived working paper sheet (`<year>`). The working paper is manually built from the trial balance each year — a tedious, error-prone process that requires understanding the prior year's layout and formula logic.

---

## Goals

- Automate creation of the new year's working paper sheet (`<year>`) from its `מאזן <year>` trial balance.
- Eliminate the need for accountants to understand code or use the command line.
- Produce output that matches the expected format exactly (columns, groupings, formulas, section headers).

---

## Users

- **Primary**: Accountants / financial staff with no technical background.
- **Distribution**: Desktop executable (`.exe`) distributed directly to users.

---

## Workbook Structure (observed from real files)

Every workbook follows this sheet pattern:

```
מאזן 2020 | 2020 | מאזן 2021 | 2021 | ... | מאזן 2024 | דוחות כספיים
```

- Pairs: `מאזן <year>` (trial balance, 7 columns) + `<year>` (working paper, 11 columns).
- The final `מאזן <year>` sheet exists but its matching `<year>` working paper is missing — this is what the tool must generate.

---

## Sheet Schemas

### Source: `מאזן <year>` — Trial Balance (7 columns)

| Col | Header | Content |
|-----|--------|---------|
| A | _(none)_ | Category label row OR `'סה"כ לקבוצה:'` for group-total rows, else `None` |
| B | מיון | Group number (e.g. 100, 101, 120, 300…) |
| C | חשבון | Account number |
| D | שם חשבון | Account name |
| E | חובה | Debit amount (or `None`) |
| F | זכות | Credit amount (or `None`) |
| G | הפרש | Net difference (debit − credit) |

Row types in the trial balance:
1. **Title rows** (rows 1–3): company name, date range, report title.
2. **Header row** (row 4): column labels.
3. **Section header rows**: A = category name, B–G = None.
4. **Account rows**: A = None, B = group number, C = account number, D = account name, E = debit or None, F = credit or None.
5. **Group-total rows**: A = `'סה"כ לקבוצה:'`, B = group number or wildcard (`'10*'`, `'1**'`, `'***'`), E = total debit, F = total credit, G = net.

### Target: `<year>` — Working Paper (11 columns)

| Col | Header | Content |
|-----|--------|---------|
| A | קבוצה | Group number (from מיון) |
| B | מס' כרטיס | Account number |
| C | פרטים | Account name / section label |
| D | חובה | Debit amount if account has debit, else **empty** |
| E | זכות | Credit amount if account has credit, else **empty** |
| F | פ.נ | Opening balance (= closing יתרה from the **prior year's** working paper) |
| G | יתרה | Net balance = F + D − E (formula) |
| H | הערות | Notes (manual, carried over or empty) |
| I–K | _(none)_ | Auxiliary columns (calculations / notes, not always populated) |

Key rules for the working paper:
- **Columns D and E are mutually exclusive per row** — only one has a value, the other is empty.
- **Column G = F + D − E** (formula). When D is filled: G = F + D. When E is filled: G = F − E.
- **Column F (opening balance)** is sourced from the matching account's G (יתרה) in the **prior year's working paper**.
- Rows are **grouped by group number** (col A). Each group may be preceded by a section header row (category name in col C, all other cols empty).
- Group-total / subtotal rows (section totals) are included between groups.

---

## Processing Logic

1. **Detect** the latest `מאזן <year>` sheet that has no matching `<year>` working paper.
2. **Load** the prior year's working paper sheet to extract column F values (opening balances).
3. **Use Claude CLI** to analyze the prior year's pattern (section headers, grouping, row ordering, subtotal rows) and apply it to the new year's data.
4. **Map** each account row from the trial balance to the working paper:
   - If trial balance חובה > 0 → working paper col D = חובה, col E = empty.
   - If trial balance זכות > 0 → working paper col E = זכות, col D = empty.
   - Col F = prior year's G for the same account number.
   - Col G = formula `=F+D-E`.
5. **Preserve** section header rows, group structure, and subtotal rows from the prior year pattern.
6. **Write** the new sheet into the workbook.

---

## Functional Requirements

### Input
- A `.xlsx` workbook containing at least one prior-year `<year>` working paper and a `מאזן <year>` trial balance for the new year.

### Output
- The same `.xlsx` file (or a copy) with the new year's `<year>` working paper fully populated.
- No manual post-processing required.

### UI
- Simple desktop GUI (no command line needed).
- User selects the input `.xlsx` file via a file picker.
- User clicks a single "Generate" button.
- Output file is saved and the path is shown to the user.
- Errors are displayed in plain language.

---

## Non-Functional Requirements

| Requirement | Detail |
|---|---|
| Language | Python 3.14 |
| Package manager | `uv` (no pip) |
| Distribution | PyInstaller `.exe` |
| Codebase | As few files as possible; straightforward, readable code |
| Type safety | Full Python type annotations |
| AI integration | Claude CLI for pattern inference |

---

## Architecture

| Component | Responsibility |
|---|---|
| **Processor** | Read the `.xlsx`, identify the latest `מאזן <year>` without a matching working paper, extract trial balance rows |
| **Parser** | Call Claude CLI with the prior year's working paper + new year's trial balance to infer section structure → produce new year rows |
| **Constructor** | Write the new `<year>` sheet into the workbook: groups, section headers, D/E/F/G columns, formulas |
| **UI** | Desktop GUI — file picker, generate button, output path display, plain-language errors |

---

## Testing Strategy

### Structure

- `pytest` with real example `.xlsx` files in `example/`.
- `input/` — source files used as test inputs; `output/` — generated files for comparison.

### Deterministic checks (pytest, no AI)

These rules can be verified mechanically against the generated sheet:

| Check | Rule |
|---|---|
| Headers present | Row 2 of `<year>` sheet = `['קבוצה', "מס' כרטיס", 'פרטים', 'חובה', 'זכות', 'פ.נ', 'יתרה', 'הערות', ...]` |
| D/E mutual exclusivity | For every account row: exactly one of col D, col E is non-empty |
| G formula | For every account row: G = F + D − E (within float tolerance) |
| All trial balance accounts present | Every account number from `מאזן <year>` appears in the generated sheet |
| Group ordering | Rows within each group share the same group number in col A |
| No missing groups | Every group number from `מאזן <year>` appears in col A |
| D sign | Col D values are always positive |
| E sign | Col E values are always positive |

### AI-powered validation (Claude CLI agent)

After generation, run a Claude CLI agent as a second-pass validator. The agent receives:
1. The rules from this document (verbatim or summarised).
2. The full content of the generated `<year>` sheet (as JSON or CSV).
3. The source `מאזן <year>` trial balance (as JSON or CSV).
4. The prior year's `<year>` working paper (as JSON or CSV).

The agent checks everything a mechanical test cannot:
- Section header rows are placed correctly (each section label precedes its group).
- Account names (col C) match the trial balance account names (col D of `מאזן`).
- Col F (opening balance) correctly reflects the prior year's col G for the same account.
- Subtotal / group-total rows are present and match the trial balance group totals.
- No phantom rows (rows in the output with no corresponding trial balance account).
- The overall sheet structure mirrors the prior year's working paper pattern.

The agent returns a structured verdict:
```json
{
  "passed": true | false,
  "violations": [
    { "row": <n>, "column": "<col>", "issue": "<description>" }
  ],
  "summary": "<plain-language summary>"
}
```

A test fails if `passed` is `false` or if any `violations` are returned.

### Regression test

For each file in `example/`, the prior year's working paper acts as the reference:
- Generate the working paper from the `מאזן <year>` trial balance.
- Run both deterministic checks and the Claude CLI agent against the output.
- Compare the generated sheet cell-by-cell against the reference working paper in the example file.

---

## Out of Scope

- Cloud or web deployment.
- Multi-user or concurrent processing.
- Support for file formats other than `.xlsx`.
- Manual editing of generated sheets within the tool.
- Generating the `דוחות כספיים` (financial statements) sheet.
