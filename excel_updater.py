"""
excel_updater.py (CLEAN VERSION)

- Fixes string date issues
- Supports both start/end and start_date/end_date
- Safer matching + iteration
"""

from datetime import datetime, date, timedelta
from typing import List, Dict, Tuple
import openpyxl
import re

# Optional fuzzy matching
try:
    from rapidfuzz import process as fz_process, fuzz
    HAS_FUZZ = True
except ImportError:
    HAS_FUZZ = False


MONTH_SHEETS = {
    1: "Jan", 2: "Feb", 3: "Mar",  4: "Apr",
    5: "May", 6: "Jun", 7: "Jul",  8: "Aug",
    9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}

EMPLOYEE_START_ROW = 8
DAY_1_COL = 4
REMARKS_OFFSET = 10


# -------------------------
# Helpers
# -------------------------

def _normalise(name: str) -> str:
    return re.sub(r"\s+", " ", name.lower().strip())


def _to_date(value):
    """Convert string/datetime/date → date safely"""
    if isinstance(value, date):
        return value
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, str):
        try:
            return datetime.fromisoformat(value).date()
        except:
            return None
    return None


def _day_column(day: int) -> int:
    return DAY_1_COL + day - 1


def _days_in_month(year: int, month: int) -> int:
    from calendar import monthrange
    return monthrange(year, month)[1]


def _remarks_column(days_in_month: int) -> int:
    return DAY_1_COL + days_in_month + REMARKS_OFFSET - 1


def _get_employee_map(sheet) -> Dict[str, int]:
    emp_map = {}
    for row_idx in range(EMPLOYEE_START_ROW, sheet.max_row + 1):
        val = sheet.cell(row=row_idx, column=2).value
        if not val:
            continue
        name = str(val).strip()
        if name.startswith("="):
            continue
        emp_map[_normalise(name)] = row_idx
    return emp_map


def _fuzzy_match(name: str, candidates: Dict[str, int], threshold=75):
    norm = _normalise(name)

    # Exact
    if norm in candidates:
        return candidates[norm]

    # Contains
    for key, row in candidates.items():
        if norm in key or key in norm:
            return row

    if not HAS_FUZZ:
        return None

    match = fz_process.extractOne(norm, list(candidates.keys()), scorer=fuzz.token_set_ratio)
    if match and match[1] >= threshold:
        return candidates[match[0]]

    return None


def resolve_status(record: Dict, default="UA"):
    text = (record.get("type") or "").lower()

    if "holiday" in text or "leave" in text:
        return "H"
    if "sick" in text:
        return "SA"
    if "emergency" in text:
        return "EA"
    if "authorised" in text:
        return "AA"

    return record.get("excel_status", default)


# -------------------------
# MAIN FUNCTION
# -------------------------

def update_excel(
    excel_path: str,
    output_path: str,
    cancellations: List[Dict],
    default_status="UA",
    fuzzy_threshold=75,
):

    wb = openpyxl.load_workbook(excel_path)
    stats = {"updated": 0, "unmatched": 0, "unmatched_names": []}

    for record in cancellations:

        name = (record.get("name") or "").strip()

        # 🔧 SUPPORT BOTH KEY TYPES
        start_date = record.get("start_date") or record.get("start")
        end_date   = record.get("end_date") or record.get("end")

        start_date = _to_date(start_date)
        end_date   = _to_date(end_date)

        if not name or not start_date or not end_date:
            continue

        partial_week = record.get("partial_week", False)
        status_code = resolve_status(record, default_status)

        current = start_date
        matched_any = False

        while current <= end_date:

            sheet_name = MONTH_SHEETS.get(current.month)
            if sheet_name not in wb.sheetnames:
                current += timedelta(days=1)
                continue

            sheet = wb[sheet_name]
            emp_map = _get_employee_map(sheet)

            row = _fuzzy_match(name, emp_map, fuzzy_threshold)

            if not row:
                current += timedelta(days=1)
                continue

            matched_any = True

            col = _day_column(current.day)
            cell = sheet.cell(row=row, column=col)

            cell.value = status_code
            cell.number_format = "@"

            stats["updated"] += 1

            # Handle part-time note
            if partial_week:
                dim = _days_in_month(current.year, current.month)
                rem_col = _remarks_column(dim)

                existing = sheet.cell(row=row, column=rem_col).value or ""
                note = "Working 2 days/week — update P days manually"

                if note not in str(existing):
                    sheet.cell(row=row, column=rem_col).value = (
                        f"{existing}; {note}".lstrip("; ") if existing else note
                    )

            current += timedelta(days=1)

        if not matched_any:
            stats["unmatched"] += 1
            if name not in stats["unmatched_names"]:
                stats["unmatched_names"].append(name)

    wb.save(output_path)
    return stats