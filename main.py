#!/usr/bin/env python3
"""
Attendance Tracker Automation Tool
Parses shift cancellations from WhatsApp messages or Sage HR
and writes them into the Excel attendance tracker.

Usage:
  python main.py --source whatsapp --input cancellations.txt --excel tracker.xlsx
  python main.py --source sageHR  --excel tracker.xlsx --month 4 --year 2026
  python main.py --source whatsapp --text "Winifred cancelled from 24 March to 30 March" --excel tracker.xlsx
"""

import argparse
import sys
import os
sys.path.append(os.path.join(os.getcwd(), 'attendance_tool'))
from datetime import date

from attendance_tool.config import load_config
from parsers.whatsapp_parser import parse_whatsapp_text
from parsers.sage_hr import fetch_sage_hr_absences
from excel_updater import update_excel


def main():
    parser = argparse.ArgumentParser(
        description="Sync shift cancellations into the Excel Attendance Tracker"
    )
    parser.add_argument(
        "--source",
        required=True,
        help="Where to read cancellations from"
    )
    parser.add_argument(
        "--excel",
        required=True,
        help="Path to the attendance Excel file"
    )
    parser.add_argument(
        "--input",
        help="[whatsapp] Path to a .txt file with WhatsApp messages"
    )
    parser.add_argument(
        "--text",
        help="[whatsapp] Cancellation text pasted directly on the command line"
    )
    parser.add_argument(
        "--month",
        type=int,
        default=date.today().month,
        help="Month to sync (default: current month)"
    )
    parser.add_argument(
        "--year",
        type=int,
        default=date.today().year,
        help="Year to sync (default: current year)"
    )
    parser.add_argument(
        "--status",
        default="UA",
        choices=["UA", "AA", "SA", "EA", "OR"],
        help="Absence status code to write (default: UA)"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print changes without writing to Excel"
    )
    parser.add_argument(
        "--output",
        help="Save updated Excel to a different file (default: overwrites --excel)"
    )

    args = parser.parse_args()

    # ── Load config (API keys, subdomain, etc.) ────────────────────────────
    cfg = load_config()

    # ── Gather cancellations ───────────────────────────────────────────────
    if args.source == "whatsapp":
        if args.input:
            with open(args.input, "r", encoding="utf-8") as f:
                raw_text = f.read()
        elif args.text:
            raw_text = args.text
        else:
            print("ERROR: --source whatsapp requires --input <file> or --text <text>")
            sys.exit(1)

        print("\n📲  Parsing WhatsApp cancellation messages…")
        cancellations = parse_whatsapp_text(raw_text, cfg)

    elif args.source == "sage_pdf":
        from parsers.sage_pdf import parse_sage_pdf
        cancellations = parse_sage_pdf(args.input)

    else:  # sageHR
        print("\n🔗  Fetching absences from Sage HR…")
        cancellations = fetch_sage_hr_absences(
            cfg,
            month=args.month,
            year=args.year
        )

    if not cancellations:
        print("⚠️  No cancellations found. Nothing to write.")
        sys.exit(0)

    print(f"\n✅  Found {len(cancellations)} cancellation record(s):")
    for c in cancellations:
        flag = "  ⚠ 2 days/week — update P days manually" if c.get("partial_week") else ""
        print(f"   • {c['name']:40s} {c['start_date']}  →  {c['end_date']}{flag}")

    # ── Write to Excel ─────────────────────────────────────────────────────
    output_path = args.output or args.excel
    if not os.path.exists(args.excel):
        print(f"\nERROR: Excel file not found: {args.excel}")
        sys.exit(1)

    if args.dry_run:
        print("\n🔎  Dry run — no changes written.")
    else:
        print(f"\n📝  Writing to Excel: {output_path}")
        stats = update_excel(
            excel_path=args.excel,
            output_path=output_path,
            cancellations=cancellations,
            default_status=args.status
        )
        print(f"✅  Done!  Cells updated: {stats['updated']}  |  Unmatched names: {stats['unmatched']}")
        if stats["unmatched_names"]:
            print("   Unmatched names (check spelling):")
            for n in stats["unmatched_names"]:
                print(f"     - {n}")


if __name__ == "__main__":
    main()
