"""Parser module: converts trial balance rows directly into working paper rows.

The trial balance already encodes the full section structure (section_header rows,
account rows, group_total rows in the correct order), so no AI call is needed for
the core mapping. Opening balances come from the prior-year working paper.
"""
from __future__ import annotations

from processor import WorkbookData
from constructor import WorkingPaperRow

GROUP_TOTAL_LABEL = 'סה"כ לקבוצה:'


def _split_debit_credit(debit: float | None, credit: float | None) -> tuple[float | None, float | None]:
    """Return (D, E) that are mutually exclusive and positive.

    Rules:
    - If only debit > 0  → D = debit,  E = None
    - If only credit > 0 → D = None,   E = credit
    - If both > 0        → use net; positive net → D, negative net → E
    - If both zero/None  → D = 0.0, E = None  (rare; keeps constructor happy)
    """
    d = debit if (debit is not None and debit > 0) else None
    c = credit if (credit is not None and credit > 0) else None

    if d is not None and c is not None:
        net = d - c
        if net >= 0:
            return net, None
        else:
            return None, -net

    if d is not None:
        return d, None
    if c is not None:
        return None, c

    # Both zero — write debit=0 so constructor doesn't raise "neither D nor E"
    return 0.0, None


def generate_working_paper_rows(data: WorkbookData) -> list[WorkingPaperRow]:
    """Convert trial balance rows to working paper rows deterministically.

    The trial balance already contains section headers, account rows, and
    group-total rows in the correct order — we just map the columns.
    """
    rows: list[WorkingPaperRow] = []

    for tb in data.trial_balance_rows:
        if tb.row_type in ("title", "header"):
            continue

        if tb.row_type == "section_header":
            rows.append(WorkingPaperRow(
                row_type="section_header",
                group=None,
                account_number=None,
                details=tb.label,
                debit=None,
                credit=None,
                opening_balance=None,
                notes=None,
            ))

        elif tb.row_type == "account":
            debit, credit = _split_debit_credit(tb.debit, tb.credit)
            # None for new accounts (no prior-year entry); 0.0 only if explicitly zero
            opening_val = data.prior_year_balances.get(tb.account_number or "")
            opening = float(opening_val) if opening_val is not None else None
            rows.append(WorkingPaperRow(
                row_type="account",
                group=tb.group,
                account_number=tb.account_number,
                details=tb.account_name,
                debit=debit,
                credit=credit,
                opening_balance=opening,
                notes=None,
            ))

        elif tb.row_type == "group_total":
            d, c = _split_debit_credit(tb.debit, tb.credit)
            rows.append(WorkingPaperRow(
                row_type="group_total",
                group=None,
                account_number=None,
                details=GROUP_TOTAL_LABEL,
                debit=d,
                credit=c,
                opening_balance=None,
                notes=None,
            ))

    return rows
