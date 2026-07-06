#!/usr/bin/env python3
"""CLI: turn an AdImpact "Spending Chart" export into a branded HTML report.

Usage:
    python generate_report.py --spending path/to/Spending.xlsx --output report.html
"""
import argparse
import sys

from report.build import build_report


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--spending", required=True, help="Path to the AdImpact Spending Chart .xlsx export")
    p.add_argument("--output", default="report.html", help="Where to write the rendered HTML report")
    p.add_argument("--top-n", type=int, default=6, help="How many advertisers to break out individually in the advertiser chart (default 6)")
    p.add_argument("--title", help="Override the report title (default: derived from the export's Race + Media Types)")
    p.add_argument("--race", help="Override the race label (default: read from the export header)")
    p.add_argument("--media-types", help="Override the media types label (default: read from the export header)")
    args = p.parse_args()

    payload = build_report(
        args.spending,
        args.output,
        top_n=args.top_n,
        title=args.title,
        race=args.race,
        media_types=args.media_types,
    )
    print(f"Wrote {args.output}")
    print(f"  {len(payload['advertiser_rows'])} advertisers, {len(payload['platform_series'])} platforms, "
          f"{payload['n_weeks']} weeks, ${payload['grand_total']:,.0f} total tracked spend")


if __name__ == "__main__":
    sys.exit(main())
