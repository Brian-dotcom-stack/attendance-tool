"""
whatsapp_parser.py

Parses WhatsApp shift cancellation messages.

Handles:
- Names on separate lines before their message
- Informal cancellation language ("not be available", "can't make the shift", etc.)
- Explicit dates (25/04) and relative dates (today, tomorrow, this week)
- Falls back to Claude AI parser when an API key is configured
"""

import re
from datetime import date, datetime, timedelta
from typing import List, Dict, Tuple, Optional

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def _is_name(line: str) -> bool:
    """
    Detect whether a line is likely a staff member's name.
    Criteria: 1-3 words, title-cased, no digits.
    """
    line = line.strip()
    return (
        1 <= len(line.split()) <= 3
        and line.istitle()
        and not any(ch.isdigit() for ch in line)
    )


# Cancellation phrases — order matters: more specific patterns first
CANCELLATION_PATTERNS = [
    r"not (?:be )?available",          # "not available" OR "not be available"
    r"cancel(?:led|ling)?",            # cancelled, cancelling, cancel
    r"can(?:not|'t) (?:make|work|do)", # can't make, can't work, cannot work
    r"won't (?:be able|work|make)",    # won't be able, won't work, won't make
    r"unable to (?:work|make|come)",   # unable to work/make/come
    r"not working",
    r"off sick",
    r"calling in sick",
    r"sick today",
    r"won't come",
    r"not (?:going to|gonna) (?:make|work|come)",
]
_CANCEL_RE = re.compile("|".join(CANCELLATION_PATTERNS), re.IGNORECASE)


def _is_cancellation(text: str) -> bool:
    return bool(_CANCEL_RE.search(text))


def _is_availability(text: str) -> bool:
    """Detect availability messages (not cancellations)."""
    text = text.lower()
    return "available" in text and not re.search(r"not (?:be )?available", text, re.IGNORECASE)


def _parse_date_range(text: str, default_year: int) -> Tuple[date, date]:
    """
    Extract a date range from free text.
    Returns (start_date, end_date) — both the same for single-day events.
    """
    today = date.today()
    text_lower = text.lower()

    # "this week" → today through end of that week (Sunday)
    if "this week" in text_lower or "rest of the week" in text_lower:
        # Find the coming Sunday (weekday 6)
        days_to_sunday = 6 - today.weekday()
        end = today + timedelta(days=days_to_sunday)
        return today, end

    # "next week" → Monday to Sunday of next week
    if "next week" in text_lower:
        days_to_monday = (7 - today.weekday()) % 7 or 7
        start = today + timedelta(days=days_to_monday)
        end = start + timedelta(days=6)
        return start, end

    # "tomorrow"
    if "tomorrow" in text_lower:
        tomorrow = today + timedelta(days=1)
        return tomorrow, tomorrow

    # "today"
    if "today" in text_lower:
        return today, today

    # Explicit dates like 25/04 or 5/4
    date_matches = re.findall(r"\b(\d{1,2})/(\d{1,2})\b", text)
    parsed_dates = []
    for day_str, month_str in date_matches:
        try:
            parsed_dates.append(date(default_year, int(month_str), int(day_str)))
        except ValueError:
            continue

    if parsed_dates:
        parsed_dates.sort()
        return parsed_dates[0], parsed_dates[-1]

    # No date found — fall back to today
    return today, today


# ─────────────────────────────────────────────────────────────
# SMART REGEX PARSER
# ─────────────────────────────────────────────────────────────

def _smart_parse(text: str, default_year: int) -> List[Dict]:
    """
    Parse WhatsApp-style messages line by line.
    Expects: name on its own line, then the message on the next line(s).
    """
    results = []
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    current_name: Optional[str] = None

    for line in lines:
        # Skip WhatsApp timestamp lines like "[12:30, 01/04/2026]"
        if re.match(r"^\[?\d{1,2}[:/]\d{2}", line):
            continue

        if _is_name(line):
            current_name = line
            continue

        if current_name is None:
            continue

        if _is_cancellation(line):
            start, end = _parse_date_range(line, default_year)
            results.append({
                "name":       current_name,
                "start_date": start.isoformat(),
                "end_date":   end.isoformat(),
                "type":       "cancelled",
                "note":       line,
            })
            # Don't reset current_name — the same person may send another message

        elif _is_availability(line):
            # Positive availability — skip
            pass

    return results


# ─────────────────────────────────────────────────────────────
# CLAUDE AI PARSER (optional)
# ─────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You are a staff attendance assistant. Extract shift cancellations from WhatsApp messages.

Return a JSON array ONLY — no preamble, no markdown fences, no extra text.

Each object must have:
  name        (string)  — the staff member's name
  start_date  (string)  — YYYY-MM-DD
  end_date    (string)  — YYYY-MM-DD
  type        (string)  — always "cancelled"
  note        (string)  — the original message text

Rules:
- Only include genuine cancellations or absences
- Ignore availability notices ("I can work on…", "I'm available…")
- Ignore complaints or rants that don't contain an actual cancellation
- For "this week", use today through the coming Sunday
- For "tomorrow", use tomorrow's date
- If no date is given, use today's date
- Do NOT invent dates
"""


def _claude_parse(text: str, api_key: str) -> List[Dict]:
    client = anthropic.Anthropic(api_key=api_key)
    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1500,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": text}],
        )
        import json
        raw = response.content[0].text.strip()
        # Strip any accidental markdown fences
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        return json.loads(raw)
    except Exception as e:
        print(f"   ⚠ Claude parser error: {e}")
        return []


# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────

def parse_whatsapp_text(text: str, cfg: dict) -> List[Dict]:
    year = date.today().year
    api_key = cfg.get("anthropic_api_key", "")

    if api_key and HAS_ANTHROPIC:
        print("   Using Claude AI parser…")
        results = _claude_parse(text, api_key)
        if results:
            return results
        print("   Falling back to smart regex parser…")

    print("   Using smart regex parser…")
    return _smart_parse(text, year)