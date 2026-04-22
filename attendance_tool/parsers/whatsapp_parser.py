"""
parsers/whatsapp_parser.py

Uses Claude (Anthropic API) to extract structured cancellation data
from freeform WhatsApp messages, handling all the messy variations:
  • "cancelled from X to Y"
  • "cancelled from X to current"
  • "working 2 days a week"
  • "has not returned to work"
  • multiple overlapping date ranges per person
  • WhatsApp timestamps, phone numbers, etc.
"""

import json
import re
from datetime import date, datetime
from typing import List, Dict

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

try:
    import dateparser
    HAS_DATEPARSER = True
except ImportError:
    HAS_DATEPARSER = False


# ── System prompt for Claude ───────────────────────────────────────────────
SYSTEM_PROMPT = """You are a data extraction assistant for a UK care staffing team.
Extract shift cancellation records from the WhatsApp message text provided.
Return ONLY a valid JSON array (no markdown, no explanation).

Each element must have:
  "name"         : string  — full name as written (e.g. "Winifred Nyingi")
  "start_date"   : string  — ISO format YYYY-MM-DD
  "end_date"     : string  — ISO format YYYY-MM-DD, or "current" if still ongoing
  "partial_week" : boolean — true if "working 2 days a week" or similar
  "note"         : string  — any extra context (e.g. "has not returned to work")

Rules:
- If a person has multiple date ranges, create one entry per range.
- "current" or "to date" means today's date: {today}.
- Ignore WhatsApp metadata lines ([time], phone numbers, "Messages and calls are end-to-end encrypted", etc).
- Normalise all dates to the year {year} unless another year is clear.
- "cancelled" = set absent for ALL days in range (end_date inclusive).
- If "working 2 days a week" appears, set partial_week=true; still set the full range.
- Return [] if no cancellations are found.
"""

# ── Fallback regex parser (used when no API key is configured) ─────────────
MONTH_MAP = {
    "january": 1, "jan": 1,
    "february": 2, "feb": 2,
    "march": 3, "mar": 3,
    "april": 4, "apr": 4,
    "may": 5,
    "june": 6, "jun": 6,
    "july": 7, "jul": 7,
    "august": 8, "aug": 8,
    "september": 9, "sep": 9, "sept": 9,
    "october": 10, "oct": 10,
    "november": 11, "nov": 11,
    "december": 12, "dec": 12,
}


def _parse_date_str(s: str, default_year: int) -> date | None:
    s = s.strip().lower()
    if s in ("current", "date", "today", "now", "present"):
        return date.today()

    # Try dateparser first if available
    if HAS_DATEPARSER:
        parsed = dateparser.parse(
            s,
            settings={"PREFER_DAY_OF_MONTH": "first", "RETURN_AS_TIMEZONE_AWARE": False}
        )
        if parsed:
            return parsed.date()

    # Manual regex fallback: "24 march", "march 24", "24 march 2026"
    pattern = re.compile(
        r"(\d{1,2})\s+([a-z]+)(?:\s+(\d{4}))?|([a-z]+)\s+(\d{1,2})(?:\s+(\d{4}))?",
        re.IGNORECASE
    )
    m = pattern.search(s)
    if m:
        if m.group(1):
            day, month_str, year_str = m.group(1), m.group(2), m.group(3)
        else:
            month_str, day, year_str = m.group(4), m.group(5), m.group(6)
        month = MONTH_MAP.get(month_str.lower())
        if month:
            year = int(year_str) if year_str else default_year
            try:
                return date(year, month, int(day))
            except ValueError:
                pass
    return None


def _regex_parse(text: str, default_year: int) -> List[Dict]:
    """Simple regex fallback for when no API key is available."""
    results = []

    # Strip WhatsApp metadata lines
    lines = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        # Skip WhatsApp system lines
        if re.match(r"^\[?\d{1,2}[:/]\d{2}", line):
            continue
        if "end-to-end encrypted" in line.lower():
            continue
        lines.append(line)

    clean = "\n".join(lines)

    # Pattern: "Name cancelled from DATE to DATE"
    cancellation_pattern = re.compile(
        r"-?\s*([A-Z][a-zA-Z\s]+?)\s+cancelled?\s+from\s+"
        r"(.+?)\s+to\s+(.+?)(?=\s*(?:,|and|\n|-|$))",
        re.IGNORECASE
    )

    for m in cancellation_pattern.finditer(clean):
        name = m.group(1).strip().rstrip(",")
        start_str = m.group(2).strip()
        end_str = m.group(3).strip()

        start = _parse_date_str(start_str, default_year)
        end = _parse_date_str(end_str, default_year)

        if not start or not end:
            continue

        # Check for "2 days a week" nearby
        context = clean[max(0, m.start() - 20): m.end() + 120]
        partial = bool(re.search(r"2\s+days?\s+(a|per|every)\s+week", context, re.IGNORECASE))

        results.append({
            "name": name,
            "start_date": start.isoformat(),
            "end_date": end.isoformat(),
            "partial_week": partial,
            "note": "working 2 days/week" if partial else ""
        })

    return results


# ── Main entry point ───────────────────────────────────────────────────────

def parse_whatsapp_text(text: str, cfg: dict) -> List[Dict]:
    """
    Parse cancellation text and return a list of cancellation dicts.
    Uses Claude API if an anthropic_api_key is configured, otherwise falls
    back to regex parsing.
    """
    today = date.today()
    year = today.year

    api_key = cfg.get("anthropic_api_key", "")

    if api_key and HAS_ANTHROPIC:
        return _claude_parse(text, api_key, today, year)
    else:
        if not api_key:
            print("   ℹ️  No Anthropic API key — using regex parser (less accurate).")
            print("      Add 'anthropic_api_key' to config.json for best results.")
        elif not HAS_ANTHROPIC:
            print("   ℹ️  anthropic package not installed — using regex parser.")
            print("      Run: pip install anthropic")
        return _regex_parse(text, year)


def _claude_parse(text: str, api_key: str, today: date, year: int) -> List[Dict]:
    client = anthropic.Anthropic(api_key=api_key)

    system = SYSTEM_PROMPT.format(today=today.isoformat(), year=year)

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            system=system,
            messages=[{"role": "user", "content": text}]
        )
        raw = response.content[0].text.strip()

        # Strip any accidental markdown fences
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)

        records = json.loads(raw)

        # Normalise: convert "current" end dates to today
        for r in records:
            if r.get("end_date") == "current":
                r["end_date"] = today.isoformat()

        return records

    except json.JSONDecodeError as e:
        print(f"   ⚠️  Claude returned invalid JSON: {e}. Falling back to regex.")
        return _regex_parse(text, year)
    except Exception as e:
        print(f"   ⚠️  Claude API error: {e}. Falling back to regex.")
        return _regex_parse(text, year)
