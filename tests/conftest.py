"""pytest configuration and shared fixtures."""
from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from processor import load_workbook_data, WorkbookData
from parser import generate_working_paper_rows
from constructor import generate_working_paper, WorkingPaperRow

INPUT_DIR = Path(__file__).parent.parent / "input"
OUTPUT_DIR = Path(__file__).parent.parent / "output"


@pytest.fixture
def input_files() -> list[Path]:
    """Return all .xlsx paths in input/."""
    return sorted(INPUT_DIR.glob("*.xlsx"))


@pytest.fixture(params=sorted(Path(__file__).parent.parent.joinpath("input").glob("*.xlsx")))
def generated_output(request: pytest.FixtureRequest) -> tuple[WorkbookData, list[WorkingPaperRow], Path]:
    """Run the full pipeline for each input file and return (data, rows, output_path)."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    input_path: Path = request.param
    data = load_workbook_data(input_path)
    rows = generate_working_paper_rows(data)
    output_path = generate_working_paper(data, rows)
    return data, rows, output_path
