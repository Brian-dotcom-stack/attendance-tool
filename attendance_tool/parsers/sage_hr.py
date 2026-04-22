"""
parsers/sage_hr.py

Fetches absence / leave records from Sage HR (formerly CakeHR) via their REST API.

Sage HR API docs:  https://sagehr.docs.apiary.io/
Auth:              X-Auth-Token header
Base URL:          https://{subdomain}.sage.hr/api/v1

Key endpoints used:
  GET /employees           — list all employees (id + name)
  GET /leave-management    — list all leave requests (with date ranges)

To get your API key:
  Sage HR → Settings → Integrations → API → Enable API Access
"""

from datetime import date, datetime
from typing import List, Dict

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False


# Status type mapping from Sage HR leave types → Excel status codes
SAGE_STATUS_MAP = {
    "holiday":               "H",
    "annual leave":          "H",
    "sick":                  "SA",
    "sickness":              "SA",
    "sick leave":            "SA",
    "emergency":             "EA",
    "emergency leave":       "EA",
    "authorised":            "AA",
    "authorised absence":    "AA",
    "unauthorised":          "UA",
    "unauthorised absence":  "UA",
    "other":                 "OR",
    "other reason":          "OR",
}


def _map_leave_type(leave_type: str) -> str:
    """Map a Sage HR leave type label to an Excel status code."""
    if not leave_type:
        return "AA"
    return SAGE_STATUS_MAP.get(leave_type.lower().strip(), "AA")


class SageHRClient:
    def __init__(self, api_key: str, subdomain: str):
        if not HAS_REQUESTS:
            raise ImportError("requests is not installed. Run: pip install requests")
        self.base_url = BASE_URL = "https://medicaresupportandhousingltd.sage.hr"
        self.headers = {
            "X-Auth-Token": api_key,
            "Accept": "application/json",
            "Content-Type": "application/json",
        }

    def _get(self, endpoint: str, params: dict = None) -> dict | list:
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        resp = requests.get(url, headers=self.headers, params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def get_employees(self) -> List[Dict]:
        """Return list of {id, name, email} for all active employees."""
        data = self._get("/api/employees")
        employees = data.get("data", data) if isinstance(data, dict) else data
        result = []
        for emp in employees:
            result.append({
                "id":    str(emp.get("id", "")),
                "name":  emp.get("full_name") or emp.get("name", ""),
                "email": emp.get("email", ""),
            })
        return result

    def get_leave_requests(
        self,
        start_date: date,
        end_date: date,
        status: str = "approved"
    ) -> List[Dict]:
        """
        Return leave requests within the date range.
        status: "approved" | "pending" | "declined" | "all"
        """
        params = {
            "date_from": start_date.isoformat(),
            "date_to":   end_date.isoformat(),
        }
        if status != "all":
            params["status"] = status

        data = self._get("/api/leave-requests", params=params)
        leaves = data.get("data", data) if isinstance(data, dict) else data
        return leaves if isinstance(leaves, list) else []


def fetch_sage_hr_absences(cfg: dict, month: int, year: int) -> List[Dict]:
    """
    Fetch absences from Sage HR for the given month/year and return
    them in the same format as the WhatsApp parser:
      [{name, start_date, end_date, partial_week, note, excel_status}]
    """
    api_key   = cfg.get("sage_hr_api_key", "")
    subdomain = cfg.get("sage_hr_subdomain", "")

    if not api_key or not subdomain:
        print(
            "ERROR: Sage HR credentials missing.\n"
            "  Set 'sage_hr_api_key' and 'sage_hr_subdomain' in config.json\n"
            "  OR export SAGE_HR_API_KEY and SAGE_HR_SUBDOMAIN as environment variables."
        )
        return []

    if not HAS_REQUESTS:
        print("ERROR: 'requests' package not installed. Run: pip install requests")
        return []

    # Cover the whole month (± 1 day buffer to catch overlapping absences)
    from calendar import monthrange
    _, last_day = monthrange(year, month)
    range_start = date(year, month, 1)
    range_end   = date(year, month, last_day)

    client = SageHRClient(api_key, subdomain)

    # Build id→name map
    print("   Fetching employee list…")
    employees = client.get_employees()
    id_to_name = {emp["id"]: emp["name"] for emp in employees}
    print(f"   Found {len(employees)} employee(s).")

    # Fetch leaves
    print("   Fetching leave records…")
    leaves = client.get_leave_requests(range_start, range_end, status="all")
    print(f"   Found {len(leaves)} leave record(s) in date range.")

    results = []
    for leave in leaves:
        emp_id   = str(leave.get("employee_id", ""))
        name     = id_to_name.get(emp_id) or leave.get("employee_name", emp_id)
        start_str = leave.get("date_from") or leave.get("start_date") or ""
        end_str   = leave.get("date_to")   or leave.get("end_date")   or ""
        leave_type = leave.get("leave_type_name") or leave.get("type", "")
        status_val = leave.get("status", "")

        if not start_str or not end_str:
            continue

        try:
            start = datetime.fromisoformat(start_str[:10]).date()
            end   = datetime.fromisoformat(end_str[:10]).date()
        except ValueError:
            continue

        # Clip to the requested month
        start = max(start, range_start)
        end   = min(end, range_end)
        if start > end:
            continue

        excel_code = _map_leave_type(leave_type)

        results.append({
            "name":         name,
            "start_date":   start.isoformat(),
            "end_date":     end.isoformat(),
            "partial_week": False,
            "note":         f"{leave_type} [{status_val}]",
            "excel_status": excel_code,     # overrides --status flag when present
        })

    return results
