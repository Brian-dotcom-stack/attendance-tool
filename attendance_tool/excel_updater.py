"""
excel_updater.py

Reads the Yearly Staff Attendance Master Excel file,
matches cancellation records to employee rows using fuzzy name matching,
and writes the correct absence codes into the right day columns.

Sheet structure (Jan–Dec):
  Row 6  : Header (Employee ID | Employee Name | Reported To | 1 … 31)
  Row 7  : WD labels
  Row 8+ : One row per employee
  Col A  : Employee ID
  Col B  : Employee Name (formula referencing Employee Master)
  Col C  : Reported To
  Col D  : Day 1  →  col D+N-1 : Day N
"""

import shutil
from datetime import date, timedelta
from typing import List, Dict, Tuple

import openpyxl

try:
    from rapidfuzz import process as fz_process, fuzz
    HAS_RAPIDFUZZ = True
except ImportError:
    try:
        from fuzzywuzzy import process as fz_process, fuzz
        HAS_RAPIDFUZZ = True
    except ImportError:
        HAS_RAPIDFUZZ = False

MONTH_SHEETS = {
    1: "Jan", 2: "Feb", 3: "Mar",  4: "Apr",
    5: "May", 6: "Jun", 7: "Jul",  8: "Aug",
    9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}

# Row where employees start (1-based)
EMPLOYEE_START_ROW = 8
# Columns: A=1, B=2, C=3, D=4 (day 1)
DAY_1_COL = 4

REMARKS_OFFSET = 10  # after day cols: Present,AA,UA,SA,EA,OR,WO,H,AbsReason,Remarks


def _day_column(day: int) -> int:
    """Return openpyxl column index for a given day number (1-based)."""
    return DAY_1_COL + day - 1


def _remarks_column(days_in_month: int) -> int:
    """Return column index for the Remarks cell."""
    return DAY_1_COL + days_in_month + REMARKS_OFFSET - 1


def _days_in_month(year: int, month: int) -> int:
    from calendar import monthrange
    return monthrange(year, month)[1]


def _get_employee_map_from_master(wb) -> Dict[str, int]:
    """
    Read employee names and row numbers from the Employee Master sheet.
    Returns {normalised_name: row_number} — row numbers match monthly sheets.
    """
    emp_map = {}
    if "Employee Master" not in wb.sheetnames:
        return emp_map
    master = wb["Employee Master"]
    for row_idx in range(EMPLOYEE_START_ROW, master.max_row + 1):
        cell_b = master.cell(row=row_idx, column=2).value
        if not cell_b:
            continue
        name = str(cell_b).strip()
        if not name or name == "Employee Name":
            continue
        emp_map[_normalise(name)] = row_idx
    return emp_map


def _get_employee_map(sheet) -> Dict[str, int]:
    """
    Return {normalised_name: row_number} for all employees in the sheet.
    Names come from col B (could be formula text or actual value).
    """
    emp_map = {}
    for row_idx in range(EMPLOYEE_START_ROW, sheet.max_row + 1):
        cell_b = sheet.cell(row=row_idx, column=2).value
        if not cell_b:
            continue
        name = str(cell_b).strip()
        if name.startswith("="):
            continue  # formula cell — value not loaded
        emp_map[_normalise(name)] = row_idx
    return emp_map


def _normalise(name: str) -> str:
    """Lowercase, collapse whitespace, strip punctuation for matching."""
    import re
    return re.sub(r"\s+", " ", name.lower().strip())


def _fuzzy_match(name: str, candidates: Dict[str, int], threshold: int) -> Tuple[str | None, int | None]:
    """
    Return (matched_key, row) or (None, None).
    Handles:
      - Full names: "Zaheer Abbas" vs "Zaheer Abbas"
      - Nicknames:  "Zaheer" vs "Zaheer Abbas"
      - Middle names: "Favour" vs "Nkechiyere Favour Chukwuma"
      - Typos:      "Gurijinder" vs "Gurjinder Kaur"
    """
    norm = _normalise(name)

    # 1. Exact match first (fastest)
    if norm in candidates:
        return norm, candidates[norm]

    # 2. Substring containment — handles "zaheer" inside "zaheer abbas"
    for key, row in candidates.items():
        if norm in key or key in norm:
            return key, row

    if not HAS_RAPIDFUZZ:
        return None, None

    keys = list(candidates.keys())

    # 3. token_set_ratio: best for partial names (ignores extra tokens)
    r1 = fz_process.extractOne(norm, keys, scorer=fuzz.token_set_ratio)
    # 4. token_sort_ratio: good for same tokens in different order
    r2 = fz_process.extractOne(norm, keys, scorer=fuzz.token_sort_ratio)
    # 5. partial_ratio: matches substrings
    r3 = fz_process.extractOne(norm, keys, scorer=fuzz.partial_ratio)

    best = max(
        [r for r in [r1, r2, r3] if r],
        key=lambda r: r[1],
        default=None
    )
    if best and best[1] >= threshold:
        return best[0], candidates[best[0]]
    return None, None

def resolve_status(record: Dict, default_status="UA") -> str:
    text = (record.get("type") or "").lower()
    
    if "holiday" in text or "leave" in text:
        return "H"
    if "sick" in text:
        return "SA"
    if "emergency" in text:
        return "EA"
    if "authorised" in text:
        return "AA"
    if "cancel" in text and record.get("reinstated"):
        return "UA"
    
    return record.get("excel_status", default_status)

def update_excel(
    excel_path: str,
    output_path: str,
    cancellations: List[Dict],
    default_status: str = "UA",
    fuzzy_threshold: int = 75,
) -> Dict:
    """
    Write cancellation records into the Excel file.

    Returns stats dict: {updated, unmatched, unmatched_names}
    """
    wb = openpyxl.load_workbook(excel_path)
    stats = {"updated": 0, "unmatched": 0, "unmatched_names": []}

    # Load employee name→row map from Employee Master (names are formulas in monthly sheets)
    master_emp_map = _get_employee_map_from_master(wb)

    # Fallback: build per-sheet map if Employee Master unavailable
    sheet_emp_maps: Dict[str, Dict[str, int]] = {}

    def get_emp_map(sheet_name: str):
        if master_emp_map:
            return master_emp_map
        if sheet_name not in sheet_emp_maps:
            if sheet_name not in wb.sheetnames:
                return {}
            sheet_emp_maps[sheet_name] = _get_employee_map(wb[sheet_name])
        return sheet_emp_maps[sheet_name]

    for record in cancellations:
        name         = record.get("name", "").strip()
        start_date   = record["start_date"]
        end_date     = record["end_date"]
        partial_week = record.get("partial_week", False)
        status_code  = resolve_status(record, default_status)

        if not name:
            continue

        # Iterate day by day across the cancellation range
        current = start_date
        matched_any = False

        while current <= end_date:
            month      = current.month
            sheet_name = MONTH_SHEETS.get(month)
            if not sheet_name or sheet_name not in wb.sheetnames:
                current += timedelta(days=1)
                continue

            emp_map = get_emp_map(sheet_name)
            matched_key, row = _fuzzy_match(name, emp_map, fuzzy_threshold)

            if row is None:
                # Will be reported once after the loop
                current += timedelta(days=1)
                continue

            matched_any = True
            sheet = wb[sheet_name]
            col = _day_column(current.day)
            cell = sheet.cell(row=row, column=col)

            cell.value = status_code
            cell.number_format = "@"  # keeps text format consistent


            stats["updated"] += 1

            # Add a note in Remarks column for partial-week workers (once per month)
            if partial_week:
                dim = _days_in_month(current.year, month)
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
