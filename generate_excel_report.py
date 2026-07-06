#!/usr/bin/env python3
"""CLI: turn an AdImpact "Spending Chart" export into the flat, spreadsheet-
style competitive report (the same layout GPS Impact already uses for linear
TV), applied to digital spend.

Usage:
    python generate_excel_report.py --spending path/to/Spending.xlsx --output report.xlsx
"""
import argparse
import sys
from datetime import datetime

from report import parse
from report.excel_export import write_excel_report


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--spending", required=True, help="Path to the AdImpact Spending Chart .xlsx export")
    p.add_argument("--output", default="competitive_report.xlsx", help="Where to write the rendered workbook")
    p.add_argument("--title", help="Override the report title (default: derived from the export's Race)")
    p.add_argument(
        "--current-week",
        help='Pin the "This Week" tab to a specific media week (YYYY-MM-DD, the Tuesday it starts on) '
             "instead of deriving it from today's date in America/New_York. Use this to regenerate a "
             "past week's report or to test without waiting for a real Tuesday.",
    )
    p.add_argument(
        "--creative",
        help="Path to an AdImpact Topline Creatives .xlsx export. Adds a Creative Timeline tab showing "
             "which creatives ran where and when. Optional.",
    )
    args = p.parse_args()

    week_cols, week_starts, leaf_rows, meta = parse.load_spending_export(args.spending)
    week_labels_short, week_iso, index_map = parse.build_continuous_week_axis(week_starts)
    week_labels = [datetime.strptime(iso, "%Y-%m-%d").strftime("%B %d, %Y") for iso in week_iso]

    title = args.title
    if not title:
        race = meta.get("race")
        title = f"{race.upper()} DIGITAL COMPETITIVE REPORT" if race else "DIGITAL COMPETITIVE REPORT"

    this_week_iso = args.current_week or parse.current_media_week_iso()

    creative_rows = parse.load_creative_export(args.creative) if args.creative else None

    write_excel_report(
        leaf_rows, index_map, week_labels, args.output, title=title,
        week_iso=week_iso, this_week_iso=this_week_iso, creative_rows=creative_rows,
    )
    print(f"Wrote {args.output}")
    print(f"  This Week tab: {this_week_iso}" + (" (not in export)" if this_week_iso not in week_iso else ""))
    if creative_rows is not None:
        print(f"  Creative Timeline: {len(creative_rows)} creatives")


if __name__ == "__main__":
    sys.exit(main())
