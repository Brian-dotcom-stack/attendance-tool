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
import shutil
from datetime import date

# ── Flat imports (all files live in the same directory) ───────────────────────
from attendance_tool.config import load_config
from attendance_tool.parsers.whatsapp_parser import parse_whatsapp_text
from attendance_tool.parsers.sage_hr import fetch_sage_hr_absences
from excel_updater import update_excel


def main():
    parser = argparse.ArgumentParser(
        description="Sync shift cancellations into the Excel Attendance Tracker"
    )
    parser.add_argument(
        "--source",
        required=True,
        choices=["whatsapp", "sageHR", "sage_pdf"],
        help="Where to read cancellations from: whatsapp | sageHR | sage_pdf"
    )
    parser.add_argument(
        "--excel",
        required=True,
        help="Path to the attendance Excel template file"
    )
    parser.add_argument(
        "--input",
        help="[whatsapp / sage_pdf] Path to input file (.txt or folder of PDFs)"
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
        help="Save updated Excel to a different file (default: tracker_output.xlsx)"
    )

    args = parser.parse_args()

    # ── Load config (API keys, subdomain, etc.) ────────────────────────────────
    cfg = load_config()

    # ── Gather cancellations ───────────────────────────────────────────────────
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
        if not args.input:
            print("ERROR: --source sage_pdf requires --input <folder containing PDFs>")
            sys.exit(1)
        from sage_pdf import parse_sage_pdf
        print("\n📄  Parsing Sage PDF reports…")
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
        print(f"   • {c['name']:40s}  {c['start_date']}  →  {c['end_date']}{flag}")

    # ── Validate Excel path ────────────────────────────────────────────────────
    template_path = args.excel
    if not os.path.exists(template_path):
        print(f"\nERROR: Excel file not found: {template_path}")
        sys.exit(1)

    output_path = args.output or "tracker_output.xlsx"

    # Always copy template → output (never overwrite the template)
    if template_path != output_path:
        shutil.copy(template_path, output_path)
        print(f"\n📄  Template copied → {output_path}")
    else:
        print("\n⚠️  Writing directly to the template file (use --output to avoid this)")

    # ── Write to Excel ─────────────────────────────────────────────────────────
    if args.dry_run:
        print("\n🔎  Dry run — no changes written.")
    else:
        print(f"\n📝  Writing to Excel: {output_path}")
        stats = update_excel(
            excel_path=output_path,
            output_path=output_path,
            cancellations=cancellations,
            default_status=args.status
        )
        print(f"\n✅  Done!")
        print(f"   Cells updated  : {stats['updated']}")
        print(f"   Unmatched names: {stats['unmatched']}")
        if stats["unmatched_names"]:
            print("   ⚠ Names not found in Excel (check spelling):")
            for n in stats["unmatched_names"]:
                print(f"     - {n}")


if __name__ == "__main__":
    main()