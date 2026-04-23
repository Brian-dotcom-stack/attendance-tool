"""
config.py — loads settings from config.json or environment variables.

Copy config.json.example → config.json and fill in your details.
Alternatively set environment variables (useful for servers / CI).
"""

import os
import json
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_FILE = BASE_DIR / "config.json"

DEFAULTS = {
    "anthropic_api_key": "",
    "sage_hr_api_key": "",
    "sage_hr_subdomain": "",   # e.g. "mycompany" → mycompany.sage.hr
    "fuzzy_match_threshold": 80,
    "default_status_code": "UA",
    "two_days_week_note": "Working 2 days/week — update P days manually"
}


def load_config() -> dict:
    cfg = dict(DEFAULTS)

    # Load from JSON file if it exists
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r") as f:
            file_cfg = json.load(f)
        cfg.update(file_cfg)

    # Environment variables always override JSON
    env_map = {
        "ANTHROPIC_API_KEY":    "anthropic_api_key",
        "SAGE_HR_API_KEY":      "sage_hr_api_key",
        "SAGE_HR_SUBDOMAIN":    "sage_hr_subdomain",
    }
    for env_key, cfg_key in env_map.items():
        val = os.getenv(env_key)
        if val:
            cfg[cfg_key] = val

    return cfg
