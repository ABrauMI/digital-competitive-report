#!/usr/bin/env python3
"""CLI: turn an AdImpact "Spending Chart" export into the flat, spreadsheet-
style competitive report (the same layout GPS Impact already uses for linear
TV), applied to digital spend.

Usage:
    python generate_excel_report.py --spending path/to/Spending.xlsx --output report.xlsx
"""
import argparse
import sys

from report.pipeline import build_digital_competitive_report


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

    summary = build_digital_competitive_report(
        args.spending, args.output, creative_path=args.creative, title=args.title,
        current_week_iso=args.current_week,
    )
    print(f"Wrote {args.output}")
    note = "" if summary["this_week_in_export"] else " (not in export)"
    print(f"  This Week tab: {summary['this_week_iso']}{note}")
    if summary["creative_count"] is not None:
        print(f"  Creative Timeline: {summary['creative_count']} creatives")


if __name__ == "__main__":
    sys.exit(main())
