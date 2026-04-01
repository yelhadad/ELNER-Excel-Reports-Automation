"""Constructor module: writes the new working paper sheet from the prior year template."""
from __future__ import annotations

import re
from copy import copy as copy_obj
from dataclasses import dataclass
from typing import Literal, Any
from pathlib import Path

import openpyxl
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import PatternFill

from processor import WorkbookData, TrialBalanceRow


# ── Legacy dataclass kept for deterministic tests ────────────────────────────

@dataclass
class WorkingPaperRow:
    row_type: Literal["section_header", "account", "group_total"]
    group: str | None
    account_number: str | None
    details: str | None
    debit: float | None
    credit: float | None
    opening_balance: float | None
    notes: str | None


# ── Style constants ───────────────────────────────────────────────────────────

EXISTING_FILL = PatternFill(fill_type="solid", fgColor="FF92D050")  # green
MISSING_FILL  = PatternFill(fill_type="solid", fgColor="FFFFFF00")  # yellow
NEW_FILL      = PatternFill(fill_type="solid", fgColor="FF00B0F0")  # blue

GREEN_FILL  = EXISTING_FILL
YELLOW_FILL = MISSING_FILL


# ── Helpers ───────────────────────────────────────────────────────────────────

def _normalize_account(val: Any) -> str | None:
    """Normalise account/group numbers to a canonical string.

    int/float -> str without decimal: 1001.0 -> '1001'
    Numeric strings with leading zeros stripped: '0253' -> '253'
    """
    if val is None:
        return None
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    if isinstance(val, int):
        return str(val)
    s = str(val).strip()
    if s and s.isdigit():
        return str(int(s))
    return s or None


def _is_account_cell(val: Any) -> bool:
    """True if val is a numeric account/group number (int, float, or numeric string)."""
    if isinstance(val, bool):
        return False
    if isinstance(val, (int, float)):
        return True
    if isinstance(val, str):
        s = val.strip()
        return bool(s) and s.isdigit()
    return False


def _split_debit_credit(
    debit: float | None, credit: float | None
) -> tuple[float | None, float | None]:
    d = debit if (debit is not None and debit > 0) else None
    c = credit if (credit is not None and credit > 0) else None
    if d is not None and c is not None:
        net = d - c
        return (net, None) if net >= 0 else (None, -net)
    if d is not None:
        return d, None
    if c is not None:
        return None, c
    return 0.0, None


def _remap_row_refs(formula: str, orig_to_new: dict[int, int]) -> str:
    """Replace all cell row-number references in a formula using orig_to_new."""
    def replacer(m: re.Match) -> str:
        col = m.group(1)
        row = int(m.group(2))
        return f"{col}{orig_to_new.get(row, row)}"
    return re.sub(r'([A-Za-z]+)(\d+)', replacer, formula)


def _parse_col_sum_range(formula: str) -> tuple[str, int, int] | None:
    """Parse =SUM(Xn:Xm) for a single same-column range.

    Returns (col_letter, start_row, end_row) or None.
    Works for any column: G, D, E, H, etc.
    """
    m = re.match(r"=SUM\(([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)\)", formula or "", re.IGNORECASE)
    if m and m.group(1).upper() == m.group(3).upper():
        return (m.group(1).upper(), int(m.group(2)), int(m.group(4)))
    return None


def _extend_sum_formula(
    formula: str,
    orig_to_new: dict[int, int],
    new_account_insertions: list[tuple[int, int | None]],
    last_template_acct_orig: int,
) -> str:
    """Remap row refs in a formula and extend SUM range to cover new accounts."""
    orig_range = _parse_col_sum_range(formula)
    remapped = _remap_row_refs(formula, orig_to_new)

    if orig_range is None:
        return remapped

    col_letter, orig_start, orig_end = orig_range
    parsed_new = _parse_col_sum_range(remapped)
    if parsed_new is None:
        return remapped

    _, new_start, new_end = parsed_new

    for acct_row, after_orig in new_account_insertions:
        if after_orig is not None:
            should_extend = orig_start <= after_orig <= orig_end
        else:
            should_extend = orig_end >= last_template_acct_orig
        if should_extend and acct_row > new_end:
            new_end = acct_row

    return f"=SUM({col_letter}{new_start}:{col_letter}{new_end})"


def _copy_cell(src: Any, dst: Any) -> None:
    dst.value = src.value
    if src.has_style:
        dst.font = copy_obj(src.font)
        dst.fill = copy_obj(src.fill)
        dst.border = copy_obj(src.border)
        dst.alignment = copy_obj(src.alignment)
        dst.number_format = src.number_format


def _remap_cross_sheet_refs(
    wb: Workbook,
    new_sheet_name: str,
    orig_to_new: dict[int, int],
) -> None:
    """Update row references to the newly generated sheet in all other sheets.

    When new accounts are inserted into the generated working paper, rows shift.
    Any sheet (e.g. 'דוחות כספיים') that references specific rows of the new
    working paper by row number needs those references remapped using orig_to_new
    (which maps prior-year template row -> new working paper row).

    Only remaps references that contain no #REF! to avoid mangling already-broken
    formulas.
    """
    if not orig_to_new:
        return

    escaped_name = re.escape(new_sheet_name)
    pattern = re.compile(
        r"('" + escaped_name + r"'!|" + escaped_name + r"!)([A-Z]+)(\d+)",
        re.IGNORECASE,
    )

    def _remap_match(m: re.Match) -> str:
        prefix = m.group(1)
        col = m.group(2)
        row = int(m.group(3))
        new_row = orig_to_new.get(row, row)
        return f"{prefix}{col}{new_row}"

    for ws_name in wb.sheetnames:
        if ws_name == new_sheet_name:
            continue
        ws = wb[ws_name]
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                if not cell.value.startswith("="):
                    continue
                if "#REF!" in cell.value:
                    continue
                if new_sheet_name not in cell.value:
                    continue
                new_formula = pattern.sub(_remap_match, cell.value)
                if new_formula != cell.value:
                    cell.value = new_formula


# ── Core template-based generation ───────────────────────────────────────────

def _build_output_rows(
    prior_ws: Worksheet,
    new_accounts: list[TrialBalanceRow],
) -> list[dict]:
    """Build the ordered list of rows for the new sheet."""
    new_by_group: dict[str, list[TrialBalanceRow]] = {}
    for tb in new_accounts:
        g = _normalize_account(tb.group) or ""
        new_by_group.setdefault(g, []).append(tb)

    template: list[dict] = []
    for row in prior_ws.iter_rows():
        cells = list(row)
        col_b = cells[1].value if len(cells) > 1 else None
        group_val = cells[0].value if cells else None
        is_account = _is_account_cell(col_b)
        template.append({
            "orig_row": cells[0].row,
            "cells": cells,
            "account_num": _normalize_account(col_b),
            "group": _normalize_account(group_val) if is_account else None,
            "is_account": is_account,
        })

    group_last_idx: dict[str, int] = {}
    for i, t in enumerate(template):
        if t["is_account"] and t["group"]:
            group_last_idx[t["group"]] = i

    output: list[dict] = []
    inserted_groups: set[str] = set()

    for i, t in enumerate(template):
        output.append(t)
        if t["is_account"] and t["group"] and group_last_idx.get(t["group"]) == i:
            group = t["group"]
            if group in new_by_group and group not in inserted_groups:
                for tb in sorted(
                    new_by_group[group],
                    key=lambda x: _normalize_account(x.account_number) or "",
                ):
                    output.append({"is_new": True, "tb": tb, "inserted_after_orig_row": t["orig_row"]})
                inserted_groups.add(group)

    remaining = sorted(
        [tb for tb in new_accounts if (_normalize_account(tb.group) or "") not in inserted_groups],
        key=lambda tb: (
            _normalize_account(tb.group) or "",
            _normalize_account(tb.account_number) or "",
        ),
    )
    if remaining:
        last_tmpl_acct_idx = max(
            (i for i, row in enumerate(output)
             if not row.get("is_new") and row["is_account"]),
            default=len(output) - 1,
        )
        last_tmpl_acct_orig = output[last_tmpl_acct_idx]["orig_row"]
        insert_at = last_tmpl_acct_idx + 1
        for tb in remaining:
            output.insert(insert_at, {
                "is_new": True,
                "tb": tb,
                "inserted_after_orig_row": last_tmpl_acct_orig,
            })
            insert_at += 1

    return output


def _write_output_rows(
    ws: Worksheet,
    output_rows: list[dict],
    tb_by_account: dict[str, TrialBalanceRow],
    prior_year_balances: dict[str, float],
    prior_ws: Worksheet,
) -> dict[int, int]:
    """Write all rows to the new worksheet, updating values and formulas.

    Returns orig_to_new mapping (prior-year template row -> new sheet row) so
    the caller can update cross-sheet references in other sheets.
    """

    for col_letter, col_dim in prior_ws.column_dimensions.items():
        ws.column_dimensions[col_letter].width = col_dim.width
        ws.column_dimensions[col_letter].hidden = col_dim.hidden

    orig_to_new: dict[int, int] = {}
    assigned_accounts: set[str] = set()
    new_account_insertions: list[tuple[int, int | None]] = []
    template_rows_written: list[tuple[int, dict]] = []

    # ── First pass ────────────────────────────────────────────────────────────
    for new_row_num, row in enumerate(output_rows, start=1):

        if row.get("is_new"):
            tb: TrialBalanceRow = row["tb"]
            group = _normalize_account(tb.group) or ""
            debit, credit = _split_debit_credit(tb.debit, tb.credit)
            acct_str = _normalize_account(tb.account_number)

            ws.cell(new_row_num, 1).value = int(group) if group.isdigit() else group
            ws.cell(new_row_num, 2).value = (
                int(acct_str) if (acct_str and acct_str.isdigit()) else acct_str
            )
            ws.cell(new_row_num, 3).value = tb.account_name
            ws.cell(new_row_num, 4).value = debit
            ws.cell(new_row_num, 5).value = credit
            ws.cell(new_row_num, 7).value = f"=D{new_row_num}+E{new_row_num}"
            for col in range(1, 6):
                ws.cell(new_row_num, col).fill = copy_obj(NEW_FILL)

            new_account_insertions.append((new_row_num, row.get("inserted_after_orig_row")))
            continue

        orig_row: int = row["orig_row"]
        orig_to_new[orig_row] = new_row_num
        cells: list = row["cells"]

        for cell in cells:
            _copy_cell(cell, ws.cell(new_row_num, cell.column))

        src_height = prior_ws.row_dimensions[orig_row].height
        if src_height:
            ws.row_dimensions[new_row_num].height = src_height

        if row["is_account"]:
            account_num = row["account_num"] or ""
            ws.cell(new_row_num, 6).value = None  # F always empty

            tb_row = tb_by_account.get(account_num)
            if tb_row and account_num not in assigned_accounts:
                debit, credit = _split_debit_credit(tb_row.debit, tb_row.credit)
                ws.cell(new_row_num, 4).value = debit
                ws.cell(new_row_num, 5).value = credit
                assigned_accounts.add(account_num)
                for col in range(1, 6):
                    ws.cell(new_row_num, col).fill = copy_obj(EXISTING_FILL)
            elif tb_row:
                ws.cell(new_row_num, 4).value = None
                ws.cell(new_row_num, 5).value = None
                for col in range(1, 6):
                    ws.cell(new_row_num, col).fill = copy_obj(EXISTING_FILL)
            else:
                ws.cell(new_row_num, 4).value = None
                ws.cell(new_row_num, 5).value = None
                for col in range(1, 6):
                    ws.cell(new_row_num, col).fill = copy_obj(MISSING_FILL)

            ws.cell(new_row_num, 7).value = f"=D{new_row_num}+E{new_row_num}"

        template_rows_written.append((new_row_num, row))

    # ── Second pass: remap and extend all formula cells ───────────────────────
    last_template_acct_orig = max(
        (row["orig_row"] for _, row in template_rows_written if row["is_account"]),
        default=0,
    )

    for new_row_num, row in template_rows_written:
        h_cell = ws.cell(new_row_num, 8)
        if isinstance(h_cell.value, str) and h_cell.value.startswith("="):
            h_cell.value = _extend_sum_formula(
                h_cell.value, orig_to_new, new_account_insertions, last_template_acct_orig
            )

        if not row["is_account"]:
            for col_idx in range(1, 12):
                if col_idx == 8:
                    continue
                cell = ws.cell(new_row_num, col_idx)
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = _extend_sum_formula(
                        cell.value, orig_to_new, new_account_insertions, last_template_acct_orig
                    )

    return orig_to_new


# ── Public entry point ────────────────────────────────────────────────────────

def generate_working_paper(data: WorkbookData) -> Path:
    """Generate the new working paper sheet using the prior year sheet as a template."""

    wb = openpyxl.load_workbook(data.workbook_path)
    prior_ws: Worksheet = wb[str(data.prior_year)]

    tb_by_account: dict[str, TrialBalanceRow] = {}
    for tb in data.trial_balance_rows:
        if tb.row_type == "account" and tb.account_number:
            key = _normalize_account(tb.account_number) or ""
            if key:
                tb_by_account[key] = tb

    prior_accounts: set[str] = set()
    for row in prior_ws.iter_rows():
        cells = list(row)
        col_b = cells[1].value if len(cells) > 1 else None
        if _is_account_cell(col_b):
            acct = _normalize_account(col_b)
            if acct:
                prior_accounts.add(acct)

    new_accounts = [
        tb
        for tb in data.trial_balance_rows
        if tb.row_type == "account"
        and tb.account_number
        and (_normalize_account(tb.account_number) or "") not in prior_accounts
    ]

    output_rows = _build_output_rows(prior_ws, new_accounts)

    sheet_name = str(data.new_year)
    if sheet_name in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' already exists")
    new_ws: Worksheet = wb.create_sheet(sheet_name)

    new_ws.sheet_view.rightToLeft = prior_ws.sheet_view.rightToLeft

    target_name = f"מאזן {data.new_year}"
    if target_name in wb.sheetnames:
        target_idx = wb.sheetnames.index(target_name)
        current_idx = wb.sheetnames.index(sheet_name)
        offset = (target_idx + 1) - current_idx
        if offset:
            wb.move_sheet(sheet_name, offset)

    orig_to_new = _write_output_rows(
        new_ws, output_rows, tb_by_account, data.prior_year_balances, prior_ws
    )

    # Update cross-sheet references in all other sheets (e.g. דוחות כספיים)
    _remap_cross_sheet_refs(wb, sheet_name, orig_to_new)

    for merged_range in prior_ws.merged_cells.ranges:
        try:
            new_ws.merge_cells(str(merged_range))
        except Exception:
            pass

    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / data.workbook_path.name

    try:
        wb.save(output_path)
    except PermissionError:
        raise RuntimeError(
            f"Cannot save: '{output_path}' is open in another application. Close it and try again."
        )
    except IOError as e:
        raise RuntimeError(f"Failed to save workbook: {e}")

    return output_path


# ── Legacy helpers (kept for deterministic tests) ─────────────────────────────

def determine_output_path(input_path: Path) -> Path:
    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir / input_path.name


def build_sheet(ws: Worksheet, rows: list[WorkingPaperRow], title: str) -> None:
    ws.cell(row=1, column=1, value=title)
    headers = [
        (1, "קבוצה"), (2, "מס' כרטיס"), (3, "פרטים"),
        (4, "חובה"), (5, "זכות"), (6, "פ.נ"), (7, "יתרה"), (8, "הערות"),
    ]
    for col, header in headers:
        ws.cell(row=2, column=col, value=header)

    current_row = 3
    for row in rows:
        if row.row_type == "section_header":
            ws.cell(row=current_row, column=3, value=row.details or "")
        elif row.row_type == "account":
            ws.cell(row=current_row, column=1, value=row.group)
            ws.cell(row=current_row, column=2, value=row.account_number)
            ws.cell(row=current_row, column=3, value=row.details)
            if row.debit is not None:
                ws.cell(row=current_row, column=4, value=row.debit)
            if row.credit is not None:
                ws.cell(row=current_row, column=5, value=row.credit)
            ws.cell(row=current_row, column=6, value=row.opening_balance)
            ws.cell(row=current_row, column=7, value=f"=F{current_row}+D{current_row}-E{current_row}")
        elif row.row_type == "group_total":
            ws.cell(row=current_row, column=3, value=row.details)
            if row.debit is not None:
                ws.cell(row=current_row, column=4, value=row.debit)
            if row.credit is not None:
                ws.cell(row=current_row, column=5, value=row.credit)
        current_row += 1
