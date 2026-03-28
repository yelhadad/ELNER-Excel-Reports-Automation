# TODO — Excel Year Working Paper Generator

## Step 1 — Project Setup

- [x] Run `uv init` to scaffold the project
- [x] Create `pyproject.toml` with `[project]` section
- [x] Set `name = "excel-automation"` in pyproject.toml
- [x] Set `requires-python = ">=3.14"` in pyproject.toml
- [x] Set `version = "0.1.0"` in pyproject.toml
- [x] Add `description` field to pyproject.toml
- [x] Add `openpyxl` as a runtime dependency via `uv add openpyxl`
- [x] Add `pytest` as a dev dependency via `uv add --dev pytest`
- [x] Add `PyInstaller` as a dev dependency via `uv add --dev pyinstaller`
- [x] Add `[tool.pytest.ini_options]` section to pyproject.toml
- [x] Set `testpaths = ["tests"]` in pytest config
- [x] Run `uv sync` to install all dependencies
- [x] Verify `uv.lock` was created
- [x] Create `.python-version` file pinning `3.14`
- [x] Create `tests/` directory
- [x] Create `tests/__init__.py` (empty)
- [x] Create `output/` directory
- [x] Add `output/` to `.gitignore`
- [x] Add `.venv/` to `.gitignore`
- [x] Add `dist/` to `.gitignore`
- [x] Add `build/` to `.gitignore`
- [x] Add `*.spec` to `.gitignore`
- [x] Verify `uv run python -c "import openpyxl"` succeeds
- [x] Verify `uv run pytest` runs without errors on empty suite
- [x] Inspect example `.xlsx` files with openpyxl to understand sheet names
- [x] Inspect example `.xlsx` files to confirm column layout of `מאזן <year>`
- [x] Inspect example `.xlsx` files to confirm column layout of `<year>` working paper
- [x] Note exact Hebrew string for group-total marker (`סה"כ לקבוצה:`)
- [x] Confirm row 4 is always the header row in trial balance sheets
- [x] Confirm row 2 is the header row in working paper sheets

---

## Step 2 — `processor.py`

- [x] Create `processor.py` file with module docstring
- [x] Add `from __future__ import annotations` import
- [x] Import `dataclass` from `dataclasses`
- [x] Import `Literal` from `typing`
- [x] Import `Path` from `pathlib`
- [x] Import `Any` from `typing`
- [x] Import `openpyxl`
- [x] Define `TrialBalanceRow` dataclass
- [x] Add `row_type: Literal["section_header", "account", "group_total", "title", "header"]` to `TrialBalanceRow`
- [x] Add `label: str | None` (col A) to `TrialBalanceRow`
- [x] Add `group: str | None` (col B) to `TrialBalanceRow`
- [x] Add `account_number: str | None` (col C) to `TrialBalanceRow`
- [x] Add `account_name: str | None` (col D) to `TrialBalanceRow`
- [x] Add `debit: float | None` (col E) to `TrialBalanceRow`
- [x] Add `credit: float | None` (col F) to `TrialBalanceRow`
- [x] Add `net: float | None` (col G) to `TrialBalanceRow`
- [x] Define `WorkbookData` dataclass
- [x] Add `workbook_path: Path` to `WorkbookData`
- [x] Add `new_year: int` to `WorkbookData`
- [x] Add `prior_year: int` to `WorkbookData`
- [x] Add `trial_balance_rows: list[TrialBalanceRow]` to `WorkbookData`
- [x] Add `prior_year_balances: dict[str, float]` to `WorkbookData`
- [x] Add `prior_year_sheet_rows: list[list[Any]]` to `WorkbookData`
- [x] Write `open_workbook(path: Path) -> openpyxl.Workbook`
- [x] Write `detect_sheet_pairs(wb) -> dict[int, str | None]` — scan for `מאזן YYYY` sheets
- [x] In `detect_sheet_pairs`: parse year from sheet name using regex `מאזן (\d{4})`
- [x] In `detect_sheet_pairs`: check if a sheet named exactly `str(year)` exists
- [x] In `detect_sheet_pairs`: return `{year: working_paper_name_or_None}` for each found year
- [x] Write `find_target_year(pairs: dict[int, str | None]) -> int`
- [x] In `find_target_year`: return the largest year key whose value is `None`
- [x] In `find_target_year`: raise `ValueError` if no unpaired year exists
- [x] Write `find_prior_year(pairs, target_year) -> int`
- [x] In `find_prior_year`: return `target_year - 1` after verifying it exists in pairs with a working paper
- [x] In `find_prior_year`: raise `ValueError` with clear message if prior year sheet is missing
- [x] Write `classify_row(row_index, row_values) -> str` returning row_type
- [x] In `classify_row`: return `"title"` for row indices 1–3
- [x] In `classify_row`: return `"header"` for row index 4
- [x] In `classify_row`: return `"group_total"` when col A == `'סה"כ לקבוצה:'`
- [x] In `classify_row`: return `"section_header"` when col A is non-None and cols B–G are all None
- [x] In `classify_row`: return `"account"` when col A is None and cols B and C are non-None
- [x] In `classify_row`: skip (return `None`) for completely empty rows
- [x] Write `extract_trial_balance_rows(ws) -> list[TrialBalanceRow]`
- [x] In extraction: iterate all rows with `ws.iter_rows(values_only=True)`
- [x] In extraction: read col A as `label`
- [x] In extraction: read col B as `group`, cast to `str`
- [x] In extraction: read col C as `account_number`, cast to `str`
- [x] In extraction: read col D as `account_name`
- [x] In extraction: read col E as `debit`, cast to `float`
- [x] In extraction: read col F as `credit`, cast to `float`
- [x] In extraction: read col G as `net`, cast to `float`
- [x] Normalize all numeric fields: cast `int` → `float`, keep `None` as `None`
- [x] Normalize group numbers: always `str`, never `int`
- [x] Write `extract_prior_year_balances(ws) -> dict[str, float]`
- [x] In `extract_prior_year_balances`: identify col B (account number) and col G (יתרה)
- [x] In `extract_prior_year_balances`: skip header rows and section headers
- [x] In `extract_prior_year_balances`: build `{account_number: g_value}` dict
- [x] Write `extract_prior_year_sheet_rows(ws) -> list[list[Any]]`
- [x] In `extract_prior_year_sheet_rows`: return all rows as plain lists for Claude serialization
- [x] Write `load_workbook_data(path: Path) -> WorkbookData` — main entry function
- [x] In `load_workbook_data`: call `open_workbook`
- [x] In `load_workbook_data`: call `detect_sheet_pairs`
- [x] In `load_workbook_data`: call `find_target_year`
- [x] In `load_workbook_data`: call `find_prior_year`
- [x] In `load_workbook_data`: call `extract_trial_balance_rows`
- [x] In `load_workbook_data`: call `extract_prior_year_balances`
- [x] In `load_workbook_data`: call `extract_prior_year_sheet_rows`
- [x] In `load_workbook_data`: return populated `WorkbookData`
- [x] Add full type annotations to every function in `processor.py`
- [x] Test `load_workbook_data` manually against one example file

---

## Step 3 — `constructor.py`

- [x] Create `constructor.py` file with module docstring
- [x] Import `WorkbookData` from `processor`
- [x] Import `openpyxl`, `Path`, `dataclass`, `Literal`
- [x] Define `WorkingPaperRow` dataclass
- [x] Add `row_type: Literal["section_header", "account", "group_total"]` to `WorkingPaperRow`
- [x] Add `group: str | None` (col A) to `WorkingPaperRow`
- [x] Add `account_number: str | None` (col B) to `WorkingPaperRow`
- [x] Add `details: str | None` (col C) to `WorkingPaperRow`
- [x] Add `debit: float | None` (col D) to `WorkingPaperRow`
- [x] Add `credit: float | None` (col E) to `WorkingPaperRow`
- [x] Add `opening_balance: float | None` (col F) to `WorkingPaperRow`
- [x] Add `notes: str | None` (col H) to `WorkingPaperRow`
- [x] Write `create_new_sheet(wb, year: int) -> Worksheet`
- [x] In `create_new_sheet`: raise `ValueError` if sheet named `str(year)` already exists
- [x] Write `insert_sheet_at_correct_position(wb, new_sheet, year: int)`
- [x] In `insert_sheet_at_correct_position`: find index of `מאזן {year}` sheet
- [x] In `insert_sheet_at_correct_position`: move new sheet to index + 1
- [x] Write `write_title_row(ws, row_num: int, title: str)`
- [x] Write `write_header_row(ws, row_num: int)`
- [x] In `write_header_row`: write `'קבוצה'` to col A
- [x] In `write_header_row`: write `"מס' כרטיס"` to col B
- [x] In `write_header_row`: write `'פרטים'` to col C
- [x] In `write_header_row`: write `'חובה'` to col D
- [x] In `write_header_row`: write `'זכות'` to col E
- [x] In `write_header_row`: write `'פ.נ'` to col F
- [x] In `write_header_row`: write `'יתרה'` to col G
- [x] In `write_header_row`: write `'הערות'` to col H
- [x] Write `write_section_header_row(ws, row_num: int, details: str)`
- [x] In `write_section_header_row`: write `details` to col C only, all other cols empty
- [x] Write `write_account_row(ws, row_num: int, row: WorkingPaperRow)`
- [x] In `write_account_row`: write `row.group` to col A
- [x] In `write_account_row`: write `row.account_number` to col B
- [x] In `write_account_row`: write `row.details` to col C
- [x] In `write_account_row`: write `row.debit` to col D only if non-None
- [x] In `write_account_row`: write `row.credit` to col E only if non-None
- [x] In `write_account_row`: enforce D/E mutual exclusivity — raise if both non-None
- [x] In `write_account_row`: enforce D/E not both None for account rows
- [x] In `write_account_row`: raise if D or E is negative
- [x] In `write_account_row`: write `row.opening_balance` to col F
- [x] In `write_account_row`: write formula `=F{n}+D{n}-E{n}` to col G
- [x] In `write_account_row`: write `row.notes` to col H (empty if None)
- [x] Leave cols I, J, K empty (do not write to them)
- [x] Write `write_group_total_row(ws, row_num: int, row: WorkingPaperRow)`
- [x] In `write_group_total_row`: write group label to col C
- [x] In `write_group_total_row`: write debit total to col D if present
- [x] In `write_group_total_row`: write credit total to col E if present
- [x] Write `build_sheet(ws, rows: list[WorkingPaperRow], title: str)`
- [x] In `build_sheet`: call `write_title_row` for row 1
- [x] In `build_sheet`: call `write_header_row` for row 2
- [x] In `build_sheet`: iterate rows starting at row 3, track current row_num
- [x] In `build_sheet`: dispatch `write_section_header_row` for `"section_header"` rows
- [x] In `build_sheet`: dispatch `write_account_row` for `"account"` rows
- [x] In `build_sheet`: dispatch `write_group_total_row` for `"group_total"` rows
- [x] Write `determine_output_path(input_path: Path) -> Path`
- [x] In `determine_output_path`: return `input_path.parent / "output" / input_path.name`
- [x] In `determine_output_path`: create parent directory if it does not exist
- [x] Write `save_workbook(wb, output_path: Path) -> None`
- [x] In `save_workbook`: wrap in `try/except PermissionError` with clear message
- [x] In `save_workbook`: wrap in `try/except IOError` with clear message
- [x] Write `generate_working_paper(data: WorkbookData, rows: list[WorkingPaperRow]) -> Path`
- [x] In `generate_working_paper`: call `create_new_sheet`
- [x] In `generate_working_paper`: call `insert_sheet_at_correct_position`
- [x] In `generate_working_paper`: call `build_sheet`
- [x] In `generate_working_paper`: call `save_workbook`
- [x] In `generate_working_paper`: return output path
- [x] Add full type annotations to every function in `constructor.py`

---

## Step 4 — `parser.py`

- [x] Create `parser.py` file with module docstring
- [x] Import `subprocess`, `json`, `shutil` in `parser.py`
- [x] Import `TrialBalanceRow`, `WorkbookData` from `processor`
- [x] Import `WorkingPaperRow` from `constructor`
- [x] Write `serialize_prior_year_sheet(rows: list[list[Any]]) -> str`
- [x] In serializer: convert rows to list of dicts with column labels A–K
- [x] In serializer: return compact `json.dumps` string
- [x] Write `serialize_trial_balance(rows: list[TrialBalanceRow]) -> str`
- [x] In serializer: include only `account` and `section_header` and `group_total` rows
- [x] In serializer: output each row as a dict with named fields
- [x] In serializer: return compact `json.dumps` string
- [x] Write `build_claude_prompt(data: WorkbookData) -> str`
- [x] In prompt: state role — "You are generating the {new_year} working paper"
- [x] In prompt: state "Prior year is {prior_year}"
- [x] In prompt: include prior year working paper JSON under labeled section
- [x] In prompt: include new year trial balance JSON under labeled section
- [x] In prompt: state D/E mutual exclusivity rule
- [x] In prompt: state col F = prior year col G rule (use 0 for new accounts)
- [x] In prompt: state col G formula = F + D − E rule
- [x] In prompt: state section headers must precede their groups (same as prior year)
- [x] In prompt: state group-total rows must follow their group
- [x] In prompt: state notes (col H) carry over from prior year where account matches
- [x] In prompt: state every account from trial balance must appear exactly once
- [x] In prompt: define the exact JSON response schema to return
- [x] Write `invoke_claude_cli(prompt: str) -> str`
- [x] In `invoke_claude_cli`: check `shutil.which("claude")` — raise `RuntimeError` if not found
- [x] In `invoke_claude_cli`: call `subprocess.run(["claude", "-p", prompt], capture_output=True, text=True, timeout=120)`
- [x] In `invoke_claude_cli`: raise `RuntimeError` on non-zero return code with stderr message
- [x] In `invoke_claude_cli`: return stdout string
- [x] Write `extract_json_from_response(response: str) -> str`
- [x] In `extract_json_from_response`: strip markdown code fences if present
- [x] In `extract_json_from_response`: find first `[` and last `]` to isolate JSON array
- [x] Write `parse_claude_response(response: str) -> list[WorkingPaperRow]`
- [x] In `parse_claude_response`: call `extract_json_from_response`
- [x] In `parse_claude_response`: call `json.loads` to parse the JSON array
- [x] In `parse_claude_response`: map each dict to a `WorkingPaperRow`
- [x] Map `"row_type"` field to `WorkingPaperRow.row_type`
- [x] Map `"group"` field to `WorkingPaperRow.group`
- [x] Map `"account_number"` field to `WorkingPaperRow.account_number`
- [x] Map `"details"` field to `WorkingPaperRow.details`
- [x] Map `"debit"` field to `WorkingPaperRow.debit` (float or None)
- [x] Map `"credit"` field to `WorkingPaperRow.credit` (float or None)
- [x] Map `"opening_balance"` field to `WorkingPaperRow.opening_balance`
- [x] Map `"notes"` field to `WorkingPaperRow.notes`
- [x] Write `validate_parsed_rows(rows, data: WorkbookData) -> None`
- [x] In `validate_parsed_rows`: check D/E mutual exclusivity for every account row
- [x] In `validate_parsed_rows`: check all trial balance account numbers appear in output rows
- [x] In `validate_parsed_rows`: raise `ValueError` listing any violations found
- [x] Write `generate_working_paper_rows(data: WorkbookData) -> list[WorkingPaperRow]`
- [x] In `generate_working_paper_rows`: build prompt
- [x] In `generate_working_paper_rows`: invoke Claude CLI
- [x] In `generate_working_paper_rows`: parse response
- [x] In `generate_working_paper_rows`: validate parsed rows
- [x] In `generate_working_paper_rows`: on `json.JSONDecodeError`, retry once with a clarifying prompt
- [x] Add full type annotations to every function in `parser.py`
- [x] Log prompt and response to stderr for debugging (only when env var `DEBUG=1`)

---

## Step 5 — `ui.py`

- [x] Create `ui.py` file with module docstring
- [x] Import `tkinter as tk`
- [x] Import `tkinter.filedialog`
- [x] Import `tkinter.messagebox`
- [x] Import `threading`
- [x] Import `os`, `subprocess`, `sys`, `Path`
- [x] Import `load_workbook_data` from `processor`
- [x] Import `generate_working_paper_rows` from `parser`
- [x] Import `generate_working_paper` from `constructor`
- [x] Define `App(tk.Tk)` class
- [x] In `App.__init__`: call `super().__init__()`
- [x] In `App.__init__`: set window title to `"Excel Working Paper Generator"`
- [x] In `App.__init__`: set window size (500 × 250)
- [x] In `App.__init__`: make window non-resizable
- [x] Add file path `Label` widget
- [x] Add file path `Entry` widget (read-only display)
- [x] Add `Browse` button
- [x] Implement `browse_file(self)` — opens `filedialog.askopenfilename`
- [x] In `browse_file`: filter to `.xlsx` files only
- [x] In `browse_file`: set selected path in the Entry widget
- [x] Add `Generate` button
- [x] Add status `Label` (initial text: "Ready")
- [x] Add output path `Label` (hidden initially)
- [x] Add `Open File` button (hidden initially)
- [x] Implement `open_file(self, path: str)`
- [x] In `open_file` on macOS: use `subprocess.Popen(["open", path])`
- [x] In `open_file` on Windows: use `os.startfile(path)`
- [x] In `open_file` on Linux: use `subprocess.Popen(["xdg-open", path])`
- [x] Implement `generate(self)` handler
- [x] In `generate`: validate file path is not empty — show error if so
- [x] In `generate`: validate file exists on disk — show error if not
- [x] In `generate`: disable `Generate` button before starting
- [x] In `generate`: update status label to `"Processing..."`
- [x] In `generate`: run pipeline in `threading.Thread(target=self._run_pipeline)`
- [x] Implement `_run_pipeline(self)`
- [x] In `_run_pipeline`: call `load_workbook_data`
- [x] In `_run_pipeline`: call `generate_working_paper_rows`
- [x] In `_run_pipeline`: call `generate_working_paper`
- [x] In `_run_pipeline`: on success call `self._on_success(output_path)`
- [x] In `_run_pipeline`: on exception call `self._on_error(str(e))`
- [x] Implement `_on_success(self, output_path: Path)` — called via `after()` on main thread
- [x] In `_on_success`: set status label to `"Done"`
- [x] In `_on_success`: show output path label
- [x] In `_on_success`: show `Open File` button wired to `open_file`
- [x] In `_on_success`: re-enable `Generate` button
- [x] Implement `_on_error(self, message: str)` — called via `after()` on main thread
- [x] In `_on_error`: show `messagebox.showerror` with plain-language message
- [x] In `_on_error`: set status label to `"Error"`
- [x] In `_on_error`: re-enable `Generate` button
- [x] Handle window close during processing gracefully (no crash)
- [x] Write `run_app() -> None` entry function that creates `App` and calls `mainloop()`
- [x] Add full type annotations to all `ui.py` functions

---

## Step 6 — `main.py`

- [x] Create `main.py` file
- [x] Add module docstring
- [x] Import `run_app` from `ui`
- [x] Add `if __name__ == "__main__":` guard
- [x] Call `run_app()` inside the guard

---

## Step 7 — `tests/conftest.py`

- [x] Create `tests/conftest.py`
- [x] Import `pytest`, `Path`, `openpyxl`
- [x] Define `INPUT_DIR` constant pointing to `input/`
- [x] Define `OUTPUT_DIR` constant pointing to `output/`
- [x] Write `input_files` fixture — returns list of all `.xlsx` paths in `input/`
- [x] Write `generated_output` fixture — runs full pipeline and returns output workbook
- [x] Ensure `OUTPUT_DIR` is created in fixture if it does not exist
- [x] Parameterize `generated_output` fixture over all input files

---

## Step 8 — `tests/test_deterministic.py`

- [x] Create `tests/test_deterministic.py`
- [x] Import `pytest`, `openpyxl`, `Path`
- [x] Import `load_workbook_data` from `processor`
- [x] Import `generate_working_paper_rows` from `parser`
- [x] Import `generate_working_paper` from `constructor`
- [x] Write helper `get_account_rows(ws)` — returns rows where col A has a group number
- [x] Write helper `get_generated_sheet(output_path, year) -> Worksheet`
- [x] Write `test_headers_present`
- [x] In `test_headers_present`: assert row 2 col A == `'קבוצה'`
- [x] In `test_headers_present`: assert row 2 col B == `"מס' כרטיס"`
- [x] In `test_headers_present`: assert row 2 col C == `'פרטים'`
- [x] In `test_headers_present`: assert row 2 col D == `'חובה'`
- [x] In `test_headers_present`: assert row 2 col E == `'זכות'`
- [x] In `test_headers_present`: assert row 2 col F == `'פ.נ'`
- [x] In `test_headers_present`: assert row 2 col G == `'יתרה'`
- [x] In `test_headers_present`: assert row 2 col H == `'הערות'`
- [x] Write `test_d_e_mutual_exclusivity`
- [x] In test: for each account row assert not (D non-empty AND E non-empty)
- [x] In test: for each account row assert not (D empty AND E empty)
- [x] Write `test_g_formula_correctness`
- [x] In test: for each account row read F, D, E, G as floats
- [x] In test: compute expected G = F + (D or 0) − (E or 0)
- [x] In test: assert `abs(actual_G − expected_G) < 0.01`
- [x] Write `test_all_trial_balance_accounts_present`
- [x] In test: collect all account numbers from `מאזן <year>` sheet
- [x] In test: collect all account numbers from generated `<year>` sheet
- [x] In test: assert every trial balance account is in the generated sheet
- [x] Write `test_group_numbers_match`
- [x] In test: for each account row in generated sheet, get col A group number
- [x] In test: look up that account in the trial balance, get its group
- [x] In test: assert they match
- [x] Write `test_no_missing_groups`
- [x] In test: collect all group numbers from trial balance
- [x] In test: collect all group numbers from generated sheet col A
- [x] In test: assert every trial balance group appears in the generated sheet
- [x] Write `test_d_values_positive`
- [x] In test: for each account row where col D is non-empty, assert D > 0
- [x] Write `test_e_values_positive`
- [x] In test: for each account row where col E is non-empty, assert E > 0
- [x] Parameterize all tests over all files in `input/`
- [x] Run full deterministic test suite and fix any failures

---

## Step 9 — `tests/test_agent_validator.py`

- [x] Create `tests/test_agent_validator.py`
- [x] Import `pytest`, `json`, `subprocess`, `shutil`
- [x] Write `serialize_sheet_to_json(ws) -> str` helper
- [x] Write `build_validator_prompt(generated_json, source_json, prior_json) -> str`
- [x] In prompt: include the rules from `prd.md`
- [x] In prompt: include generated `<year>` sheet JSON
- [x] In prompt: include `מאזן <year>` source JSON
- [x] In prompt: include prior year `<year>` working paper JSON
- [x] In prompt: instruct Claude to return `{"passed": bool, "violations": [...], "summary": "..."}`
- [x] Write `invoke_validator_agent(prompt: str) -> dict`
- [x] In `invoke_validator_agent`: check `shutil.which("claude")` — skip test if not found
- [x] In `invoke_validator_agent`: call Claude CLI with 180s timeout
- [x] In `invoke_validator_agent`: extract and parse JSON verdict from response
- [x] Write `test_agent_validates_generated_sheet`
- [x] In test: build all three JSON strings
- [x] In test: call `invoke_validator_agent`
- [x] In test: assert `verdict["passed"] is True`
- [x] In test: assert `verdict["violations"] == []`
- [x] In test: print `verdict["summary"]` on failure for diagnostics
- [x] Parameterize test over all input files
- [x] Mark test with `@pytest.mark.slow` to allow skipping in fast runs

---

## Step 10 — Integration & Regression Testing

- [ ] Run full pipeline end-to-end against each file in `input/`
- [x] Compare generated sheet cell-by-cell against reference in `example/`
- [x] Verify sheet is inserted at correct position (after `מאזן <year>`, before `דוחות כספיים`)
- [x] Verify sheet name is exactly `str(year)` (e.g. `"2024"`)
- [x] Verify no data is corrupted in other existing sheets
- [x] Verify the output `.xlsx` opens without errors in Excel
- [x] Test with a file that has only one prior year (no year before that)
- [x] Test with a file where an account exists in new year but not prior year (F=0)
- [x] Test with a file where an account exists in prior year but not new year
- [x] Document any edge cases found

---

## Step 11 — PyInstaller Packaging

- [x] Verify `main.py` launches the UI correctly from command line
- [x] Run `uv run pyinstaller --onefile --windowed --name excel-generator main.py`
- [x] Verify `dist/excel-generator` (or `.exe`) is created
- [ ] Test packaged binary launches the GUI
- [ ] Test packaged binary can open a `.xlsx` file and generate output
- [x] Add `dist/` to `.gitignore` (if not already done)
- [x] Add `build/` to `.gitignore` (if not already done)
- [x] Add `*.spec` to `.gitignore` (if not already done)
- [x] Document packaging steps in `CLAUDE.md`
- [ ] Verify the `.exe` works without Python installed (Windows)

---

## Step 12 — Final Polish & Validation

- [x] Verify Claude CLI not found gives a clear user-facing error (not a stack trace)
- [x] Verify invalid `.xlsx` gives a clear user-facing error
- [x] Verify file open in Excel gives a clear "close the file first" error
- [x] Verify all Hebrew strings render correctly in the output sheet
- [x] Verify col G formulas evaluate correctly when opened in Excel
- [ ] Run `uv run pytest` and confirm all deterministic tests pass
- [ ] Run agent validator tests and confirm all pass
- [ ] Review all type annotations with `uv run mypy` (if added as dev dep)
- [x] Remove any debug print statements
- [ ] Commit final working code to git
