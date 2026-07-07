"""Shared build logic — the CLI (`generate_excel_report.py`) and the Slack
app both call this so the two surfaces can't drift apart.
"""
from . import parse
from .excel_export import write_excel_report


def build_digital_competitive_report(spending_path, output_path, creative_path=None, title=None,
                                      current_week_iso=None):
    """Build the workbook and return a summary dict for the caller to display.

    Raises whatever `report.parse` raises (ValueError) on a malformed
    export — callers are expected to catch that and show a friendly message.
    """
    week_cols, week_starts, leaf_rows, meta = parse.load_spending_export(spending_path)
    week_labels_short, week_iso, index_map = parse.build_continuous_week_axis(week_starts)
    week_labels = [parse.datetime.strptime(iso, "%Y-%m-%d").strftime("%m/%d/%Y") for iso in week_iso]

    if not title:
        race = meta.get("race")
        title = f"{race.upper()} DIGITAL COMPETITIVE REPORT" if race else "DIGITAL COMPETITIVE REPORT"

    this_week_iso = current_week_iso or parse.current_media_week_iso()

    creative_rows = parse.load_creative_export(creative_path) if creative_path else None

    write_excel_report(
        leaf_rows, index_map, week_labels, output_path, title=title,
        week_iso=week_iso, this_week_iso=this_week_iso, creative_rows=creative_rows,
    )

    grand_total = sum(sum(r["weeks"]) for r in leaf_rows)
    advertisers = {r["advertiser"] for r in leaf_rows}

    return {
        "title": title,
        "race": meta.get("race"),
        "grand_total": grand_total,
        "advertiser_count": len(advertisers),
        "n_weeks": len(week_labels),
        "this_week_iso": this_week_iso,
        "this_week_in_export": this_week_iso in week_iso,
        "creative_count": len(creative_rows) if creative_rows is not None else None,
    }
