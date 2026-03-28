"""Parser module: calls Claude CLI to generate working paper rows from trial balance data."""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from typing import Any

from processor import TrialBalanceRow, WorkbookData
from constructor import WorkingPaperRow


def serialize_prior_year_sheet(rows: list[list[Any]]) -> str:
    """Convert prior year sheet rows to JSON with column labels A-K."""
    col_labels = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
    result = []
    for row in rows:
        row_dict = {}
        for i, val in enumerate(row):
            if i < len(col_labels) and val is not None:
                row_dict[col_labels[i]] = val
        if row_dict:
            result.append(row_dict)
    return json.dumps(result, ensure_ascii=False)


def serialize_trial_balance(rows: list[TrialBalanceRow]) -> str:
    """Serialize trial balance rows to JSON (account, section_header, group_total only)."""
    included_types = {"account", "section_header", "group_total"}
    result = []
    for row in rows:
        if row.row_type not in included_types:
            continue
        result.append({
            "row_type": row.row_type,
            "label": row.label,
            "group": row.group,
            "account_number": row.account_number,
            "account_name": row.account_name,
            "debit": row.debit,
            "credit": row.credit,
            "net": row.net,
        })
    return json.dumps(result, ensure_ascii=False)


def build_claude_prompt(data: WorkbookData) -> str:
    """Build the prompt for Claude CLI to generate working paper rows."""
    prior_year_json = serialize_prior_year_sheet(data.prior_year_sheet_rows)
    trial_balance_json = serialize_trial_balance(data.trial_balance_rows)

    return f"""You are generating the {data.new_year} working paper sheet for a Hebrew financial workbook.
Prior year is {data.prior_year}.

Rules:
1. Columns D (חובה/debit) and E (זכות/credit) are MUTUALLY EXCLUSIVE per account row — only one has a value, the other must be null.
2. Column F (opening balance / פ.נ) = prior year column G (יתרה) for the same account number. Use 0 for new accounts not in the prior year.
3. Column G formula = F + D - E (do not compute it; just set it to null in the JSON - the constructor adds the formula).
4. Section header rows must precede their groups, in the same order as the prior year.
5. Group-total rows (סה"כ לקבוצה:) must follow their group.
6. Notes (column H) carry over from the prior year where account numbers match.
7. Every account from the trial balance must appear exactly once in the output.
8. Use the trial balance col E (חובה) as the working paper col D value when it is non-null and > 0.
9. Use the trial balance col F (זכות) as the working paper col E value when it is non-null and > 0.
10. All D and E values must be positive numbers.

## Prior year {data.prior_year} working paper (columns A-K):
{prior_year_json}

## New year {data.new_year} trial balance:
{trial_balance_json}

Return a JSON array of row objects. Each object must have these fields:
- "row_type": "section_header" | "account" | "group_total"
- "group": string or null (col A)
- "account_number": string or null (col B)
- "details": string or null (col C - account name or section label)
- "debit": number or null (col D - positive value if account has debit, else null)
- "credit": number or null (col E - positive value if account has credit, else null)
- "opening_balance": number or null (col F - prior year G value)
- "notes": string or null (col H)

Return ONLY the JSON array, no explanation, no markdown fences."""


def invoke_claude_cli(prompt: str) -> str:
    """Invoke Claude CLI and return stdout."""
    if shutil.which("claude") is None:
        raise RuntimeError(
            "Claude CLI not found. Please install it and ensure 'claude' is on your PATH."
        )
    if os.environ.get("DEBUG") == "1":
        print("[DEBUG] Claude prompt:", prompt[:500], "...", file=sys.stderr)

    result = subprocess.run(
        ["claude", "-p", prompt],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"Claude CLI failed (exit {result.returncode}): {result.stderr.strip()}"
        )
    if os.environ.get("DEBUG") == "1":
        print("[DEBUG] Claude response:", result.stdout[:500], "...", file=sys.stderr)
    return result.stdout


def extract_json_from_response(response: str) -> str:
    """Strip markdown fences and isolate the JSON array from Claude's response."""
    text = response.strip()
    # Remove markdown code fences
    if text.startswith("```"):
        lines = text.splitlines()
        # Remove first and last fence lines
        inner = []
        in_block = False
        for line in lines:
            if line.startswith("```") and not in_block:
                in_block = True
                continue
            if line.startswith("```") and in_block:
                break
            if in_block:
                inner.append(line)
        text = "\n".join(inner).strip()

    # Find first [ and last ]
    start = text.find("[")
    end = text.rfind("]")
    if start == -1 or end == -1:
        raise ValueError(f"No JSON array found in Claude response: {text[:200]}")
    return text[start : end + 1]


def parse_claude_response(response: str) -> list[WorkingPaperRow]:
    """Parse Claude CLI JSON response into WorkingPaperRow list."""
    json_str = extract_json_from_response(response)
    items: list[dict[str, Any]] = json.loads(json_str)
    rows: list[WorkingPaperRow] = []
    for item in items:
        row = WorkingPaperRow(
            row_type=item["row_type"],
            group=item.get("group"),
            account_number=item.get("account_number"),
            details=item.get("details"),
            debit=float(item["debit"]) if item.get("debit") is not None else None,
            credit=float(item["credit"]) if item.get("credit") is not None else None,
            opening_balance=float(item["opening_balance"]) if item.get("opening_balance") is not None else None,
            notes=item.get("notes"),
        )
        rows.append(row)
    return rows


def validate_parsed_rows(rows: list[WorkingPaperRow], data: WorkbookData) -> None:
    """Validate parsed rows against the trial balance data."""
    violations: list[str] = []

    # Check D/E mutual exclusivity for account rows
    for i, row in enumerate(rows):
        if row.row_type == "account":
            if row.debit is not None and row.credit is not None:
                violations.append(
                    f"Row {i}: account '{row.account_number}' has both debit and credit set"
                )
            if row.debit is None and row.credit is None:
                violations.append(
                    f"Row {i}: account '{row.account_number}' has neither debit nor credit"
                )

    # Check all trial balance account numbers appear in output
    tb_accounts = {
        row.account_number
        for row in data.trial_balance_rows
        if row.row_type == "account" and row.account_number is not None
    }
    output_accounts = {
        row.account_number
        for row in rows
        if row.row_type == "account" and row.account_number is not None
    }
    missing = tb_accounts - output_accounts
    if missing:
        violations.append(f"Missing accounts from trial balance: {sorted(missing)}")

    if violations:
        raise ValueError("Validation failed:\n" + "\n".join(violations))


def generate_working_paper_rows(data: WorkbookData) -> list[WorkingPaperRow]:
    """Call Claude CLI and return validated WorkingPaperRow list."""
    prompt = build_claude_prompt(data)
    try:
        response = invoke_claude_cli(prompt)
        rows = parse_claude_response(response)
    except json.JSONDecodeError:
        # Retry once with a clarifying prompt
        retry_prompt = (
            prompt
            + "\n\nIMPORTANT: Your previous response could not be parsed as JSON. "
            "Return ONLY a valid JSON array, starting with [ and ending with ]. "
            "No text before or after."
        )
        response = invoke_claude_cli(retry_prompt)
        rows = parse_claude_response(response)

    validate_parsed_rows(rows, data)
    return rows
