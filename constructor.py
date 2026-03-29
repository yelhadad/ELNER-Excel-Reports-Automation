"""Constructor module: writes the new working paper sheet from the prior year template."""
from __future__ import annotations

import re
from copy import copy as copy_obj
from dataclasses import dataclass
from typing import Literal, Any
from pathlib import Path

import openpyxl
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

GREEN_FILL = PatternFill(fill_type="solid", fgColor="FF92D050")
YELLOW_FILL = PatternFill(fill_type="solid", fgColor="FFFFFF00")


# ── Helpers ───────────────────────────────────────────────────────────────────

def _normalize_account(val: Any) -> str | None:
    """Normalise account/group numbers: int 1001 or float 1001.0 → '1001'."""
    if val is None:
        return None
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    if isinstance(val, int):
        return str(val)
    return str(val).strip() or None


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


def _parse_sum_range(formula: str) -> tuple[int, int] | None:
    """'=SUM(G3:G9)' → (3, 9), otherwise None."""
    m = re.match(r"=SUM\(G(\d+):G(\d+)\)", formula or "")
    return (int(m.group(1)), int(m.group(2))) if m else None


def _copy_cell(src: Any, dst: Any) -> None:
    dst.value = src.value
    if src.has_style:
        dst.font = copy_obj(src.font)
        dst.fill = copy_obj(src.fill)
        dst.border = copy_obj(src.border)
        dst.alignment = copy_obj(src.alignment)
        dst.number_format = src.number_format


# ── Core template-based generation ───────────────────────────────────────────

def _build_output_rows(
    prior_ws: Worksheet,
    new_accounts: list[TrialBalanceRow],
) -> list[dict]:
    """
    Build the ordered list of rows for the new sheet:
    - All prior-year rows in their original order.
    - New accounts inserted immediately after the last existing row
      whose group matches their group.
    """
    template: list[dict] = []
    for row in prior_ws.iter_rows():
        cells = list(row)
        col_b = cells[1].value if len(cells) > 1 else None
        col_h = cells[7].value if len(cells) > 7 else None
        group_val = cells[0].value if cells else None

        # Account rows have a numeric col B (int or float), not a string header
        is_account = isinstance(col_b, (int, float)) and not isinstance(col_b, bool)
        template.append({
            "orig_row": cells[0].row,
            "cells": cells,
            "account_num": _normalize_account(col_b),
            "group": _normalize_account(group_val),
            "is_account": is_account,
            "sum_range": (
                _parse_sum_range(col_h)
                if isinstance(col_h, str)
                else None
            ),
        })

    # Last template-list index for each group
    group_last_idx: dict[str, int] = {}
    for i, t in enumerate(template):
        if t["is_account"] and t["group"]:
            group_last_idx[t["group"]] = i

    # New accounts by group
    new_by_group: dict[str, list[TrialBalanceRow]] = {}
    for tb in new_accounts:
        g = _normalize_account(tb.group) or ""
        new_by_group.setdefault(g, []).append(tb)

    output: list[dict] = []
    inserted_groups: set[str] = set()

    for i, t in enumerate(template):
        output.append(t)
        if t["is_account"] and t["group"]:
            g = t["group"]
            if (
                i == group_last_idx.get(g)
                and g not in inserted_groups
                and g in new_by_group
            ):
                for tb in new_by_group[g]:
                    output.append({"is_new": True, "tb": tb})
                inserted_groups.add(g)

    # Groups with no existing template rows → append at end
    for g, tbs in new_by_group.items():
        if g not in inserted_groups:
            for tb in tbs:
                output.append({"is_new": True, "tb": tb})

    return output


def _write_output_rows(
    ws: Worksheet,
    output_rows: list[dict],
    tb_by_account: dict[str, TrialBalanceRow],
    prior_year_balances: dict[str, float],
    prior_ws: Worksheet,
) -> None:
    """Write all rows to the new worksheet, updating values and formulas."""

    # Column widths / hidden flags
    for col_letter, col_dim in prior_ws.column_dimensions.items():
        ws.column_dimensions[col_letter].width = col_dim.width
        ws.column_dimensions[col_letter].hidden = col_dim.hidden

    # orig_row → new_row_num (for SUM formula translation)
    orig_to_new: dict[int, int] = {}
    # group → new row numbers of newly inserted accounts
    new_acct_rows_by_group: dict[str, list[int]] = {}
    # (new_anchor_row, orig_start, orig_end) collected for post-processing
    sum_anchors: list[tuple[int, int, int]] = []

    # Track which accounts have already received D/E values (to avoid duplicates)
    assigned_accounts: set[str] = set()

    for new_row_num, row in enumerate(output_rows, start=1):

        # ── Newly inserted account (not in prior year) ────────────────────
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
            # F column: not written (feature pending)
            ws.cell(new_row_num, 7).value = f"=F{new_row_num}+D{new_row_num}-E{new_row_num}"
            # Yellow = new account, did not exist in prior year
            for col in range(1, 6):
                ws.cell(new_row_num, col).fill = copy_obj(YELLOW_FILL)
            new_acct_rows_by_group.setdefault(group, []).append(new_row_num)
            continue

        # ── Template row ──────────────────────────────────────────────────
        orig_row = row["orig_row"]
        orig_to_new[orig_row] = new_row_num
        cells: list = row["cells"]

        # Copy every cell (value + style)
        for cell in cells:
            _copy_cell(cell, ws.cell(new_row_num, cell.column))

        # Row height
        src_height = prior_ws.row_dimensions[orig_row].height
        if src_height:
            ws.row_dimensions[new_row_num].height = src_height

        if row["is_account"]:
            account_num = row["account_num"] or ""

            # F column: not written (feature pending)
            ws.cell(new_row_num, 6).value = None

            # D / E from new year trial balance — only for the first occurrence
            # of each account number to prevent duplicate rows showing the same value
            tb_row = tb_by_account.get(account_num)
            if tb_row and account_num not in assigned_accounts:
                debit, credit = _split_debit_credit(tb_row.debit, tb_row.credit)
                ws.cell(new_row_num, 4).value = debit
                ws.cell(new_row_num, 5).value = credit
                assigned_accounts.add(account_num)
            else:
                # Duplicate occurrence or not in 2024 — clear D/E
                ws.cell(new_row_num, 4).value = None
                ws.cell(new_row_num, 5).value = None

            # Green = existed in prior year (all template rows)
            for col in range(1, 6):
                ws.cell(new_row_num, col).fill = copy_obj(GREEN_FILL)

            # G formula
            ws.cell(new_row_num, 7).value = f"=F{new_row_num}+D{new_row_num}-E{new_row_num}"

            if row["sum_range"]:
                sum_anchors.append((new_row_num, row["sum_range"][0], row["sum_range"][1]))

        else:
            # Non-account row: rewrite G formula if present
            g_cell = ws.cell(new_row_num, 7)
            if isinstance(g_cell.value, str) and re.search(r"=[DF]\d+", g_cell.value):
                g_cell.value = f"=F{new_row_num}+D{new_row_num}-E{new_row_num}"

    # ── Update H SUM formulas ─────────────────────────────────────────────
    for new_anchor_row, orig_start, orig_end in sum_anchors:
        new_start = orig_to_new.get(orig_start, orig_start)
        new_end = orig_to_new.get(orig_end, orig_end)

        # Collect groups whose accounts were in the original SUM range
        groups_in_range: set[str] = set()
        for row in output_rows:
            if not row.get("is_new") and row["is_account"]:
                if orig_start <= row["orig_row"] <= orig_end and row["group"]:
                    groups_in_range.add(row["group"])

        # Extend end to include any new accounts inserted for those groups
        for g in groups_in_range:
            inserted = new_acct_rows_by_group.get(g, [])
            if inserted:
                new_end = max(new_end, max(inserted))

        ws.cell(new_anchor_row, 8).value = f"=SUM(G{new_start}:G{new_end})"


# ── Public entry point ────────────────────────────────────────────────────────

def generate_working_paper(data: WorkbookData) -> Path:
    """Generate the new working paper sheet using the prior year sheet as a template."""

    # Load with formulas preserved (not data_only) so we can copy styles + formula strings
    wb = openpyxl.load_workbook(data.workbook_path)
    prior_ws: Worksheet = wb[str(data.prior_year)]

    # Trial balance lookup by normalised account number
    tb_by_account: dict[str, TrialBalanceRow] = {}
    for tb in data.trial_balance_rows:
        if tb.row_type == "account" and tb.account_number:
            key = _normalize_account(tb.account_number) or ""
            if key:
                tb_by_account[key] = tb

    # Accounts already in the prior year template
    prior_accounts: set[str] = set()
    for row in prior_ws.iter_rows():
        cells = list(row)
        acct = _normalize_account(cells[1].value if len(cells) > 1 else None)
        if acct:
            prior_accounts.add(acct)

    # New accounts: in 2024 trial balance but absent from 2023 template
    new_accounts = [
        tb
        for tb in data.trial_balance_rows
        if tb.row_type == "account"
        and tb.account_number
        and (_normalize_account(tb.account_number) or "") not in prior_accounts
    ]

    # Build ordered output row list
    output_rows = _build_output_rows(prior_ws, new_accounts)

    # Create the new sheet
    sheet_name = str(data.new_year)
    if sheet_name in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' already exists")
    new_ws: Worksheet = wb.create_sheet(sheet_name)

    # Position immediately after מאזן <year>
    target_name = f"מאזן {data.new_year}"
    if target_name in wb.sheetnames:
        target_idx = wb.sheetnames.index(target_name)
        current_idx = wb.sheetnames.index(sheet_name)
        offset = (target_idx + 1) - current_idx
        if offset:
            wb.move_sheet(sheet_name, offset)

    # Write rows with updated values
    _write_output_rows(new_ws, output_rows, tb_by_account, data.prior_year_balances, prior_ws)

    # Copy merged cells
    for merged_range in prior_ws.merged_cells.ranges:
        try:
            new_ws.merge_cells(str(merged_range))
        except Exception:
            pass

    # Save to output/
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
