"""Run the generator from the command line: uv run python run.py <path-to-xlsx>"""
import sys
from pathlib import Path

from processor import load_workbook_data
from constructor import generate_working_paper


def main() -> None:
    if len(sys.argv) != 2:
        print("Usage: uv run python run.py <path-to-xlsx>")
        sys.exit(1)

    path = Path(sys.argv[1])
    if not path.exists():
        print(f"Error: file not found: {path}")
        sys.exit(1)

    data = load_workbook_data(path)
    print(f"Generating {data.new_year} sheet from {data.prior_year} template...")
    output = generate_working_paper(data)
    print(f"Saved to: {output}")


if __name__ == "__main__":
    main()
