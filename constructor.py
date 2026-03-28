"""Constructor module: writes the new working paper sheet into the workbook."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, Any
from pathlib import Path

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from processor import WorkbookData


@dataclass
class WorkingPaperRow:
    row_type: Literal["section_header", "account", "group_total"]
    group: str | None           # col A (קבוצה)
    account_number: str | None  # col B (מס' כרטיס)
    details: str | None         # col C (פרטים)
    debit: float | None         # col D (חובה) - mutually exclusive with credit
    credit: float | None        # col E (זכות) - mutually exclusive with debit
    opening_balance: float | None  # col F (פ.נ)
    notes: str | None           # col H (הערות)


def create_new_sheet(wb: openpyxl.Workbook, year: int) -> Worksheet:
    """Create a new sheet named str(year). Raises ValueError if it already exists."""
    sheet_name = str(year)
    if sheet_name in wb.sheetnames:
        raise ValueError(f"Sheet '{year}' already exists")
    return wb.create_sheet(sheet_name)


def insert_sheet_at_correct_position(wb: openpyxl.Workbook, new_sheet: Worksheet, year: int) -> None:
    """Move new_sheet to immediately after the מאזן <year> sheet."""
    target_name = f"מאזן {year}"
    if target_name not in wb.sheetnames:
        return
    target_index = wb.sheetnames.index(target_name)
    current_index = wb.sheetnames.index(new_sheet.title)
    offset = (target_index + 1) - current_index
    if offset != 0:
        wb.move_sheet(new_sheet.title, offset)


def write_title_row(ws: Worksheet, row_num: int, title: str) -> None:
    """Write title string to col A of the given row."""
    ws.cell(row=row_num, column=1, value=title)


def write_header_row(ws: Worksheet, row_num: int) -> None:
    """Write column headers to the given row."""
    headers: list[tuple[int, str]] = [
        (1, "קבוצה"),
        (2, "מס' כרטיס"),
        (3, "פרטים"),
        (4, "חובה"),
        (5, "זכות"),
        (6, "פ.נ"),
        (7, "יתרה"),
        (8, "הערות"),
    ]
    for col, header in headers:
        ws.cell(row=row_num, column=col, value=header)


def write_section_header_row(ws: Worksheet, row_num: int, details: str) -> None:
    """Write section header text to col C only."""
    ws.cell(row=row_num, column=3, value=details)


def write_account_row(ws: Worksheet, row_num: int, row: WorkingPaperRow) -> None:
    """Write an account row, enforcing D/E mutual exclusivity."""
    ws.cell(row=row_num, column=1, value=row.group)
    ws.cell(row=row_num, column=2, value=row.account_number)
    ws.cell(row=row_num, column=3, value=row.details)

    if row.debit is not None and row.credit is not None:
        raise ValueError(f"Row {row_num}: D and E are both set (mutual exclusivity violation)")
    if row.debit is None and row.credit is None:
        raise ValueError(f"Row {row_num}: account row has neither D nor E set")

    if row.debit is not None:
        if row.debit < 0:
            raise ValueError(f"Row {row_num}: debit value must be positive, got {row.debit}")
        ws.cell(row=row_num, column=4, value=row.debit)
    if row.credit is not None:
        if row.credit < 0:
            raise ValueError(f"Row {row_num}: credit value must be positive, got {row.credit}")
        ws.cell(row=row_num, column=5, value=row.credit)

    ws.cell(row=row_num, column=6, value=row.opening_balance)
    ws.cell(row=row_num, column=7, value=f"=F{row_num}+D{row_num}-E{row_num}")

    if row.notes is not None:
        ws.cell(row=row_num, column=8, value=row.notes)


def write_group_total_row(ws: Worksheet, row_num: int, row: WorkingPaperRow) -> None:
    """Write a group total row."""
    ws.cell(row=row_num, column=3, value=row.details)
    if row.debit is not None:
        ws.cell(row=row_num, column=4, value=row.debit)
    if row.credit is not None:
        ws.cell(row=row_num, column=5, value=row.credit)


def build_sheet(ws: Worksheet, rows: list[WorkingPaperRow], title: str) -> None:
    """Write title row, header row, then all data rows starting from row 3."""
    write_title_row(ws, 1, title)
    write_header_row(ws, 2)

    current_row = 3
    for row in rows:
        if row.row_type == "section_header":
            write_section_header_row(ws, current_row, row.details or "")
        elif row.row_type == "account":
            write_account_row(ws, current_row, row)
        elif row.row_type == "group_total":
            write_group_total_row(ws, current_row, row)
        current_row += 1


def determine_output_path(input_path: Path) -> Path:
    """Return the output path under input_path.parent/output/, creating it if needed."""
    output_dir = input_path.parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir / input_path.name


def save_workbook(wb: openpyxl.Workbook, output_path: Path) -> None:
    """Save the workbook, raising RuntimeError on PermissionError or IOError."""
    try:
        wb.save(output_path)
    except PermissionError:
        raise RuntimeError(
            f"Cannot save: '{output_path}' is open in another application. Please close it and try again."
        )
    except IOError as e:
        raise RuntimeError(f"Failed to save workbook: {e}")


def generate_working_paper(data: WorkbookData, rows: list[WorkingPaperRow]) -> Path:
    """Generate the new working paper sheet and save it to the output path."""
    wb = openpyxl.load_workbook(data.workbook_path)
    new_sheet = create_new_sheet(wb, data.new_year)
    insert_sheet_at_correct_position(wb, new_sheet, data.new_year)
    build_sheet(new_sheet, rows, title=str(data.new_year))
    output_path = determine_output_path(data.workbook_path)
    save_workbook(wb, output_path)
    return output_path
