"""Assemble the branded HTML report from a parsed AdImpact export."""
import base64
import json
from pathlib import Path

from . import colors, parse

ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
TEMPLATE_PATH = Path(__file__).resolve().parent / "template.html"

FONT_FILES = {
    "figtree_400": ("Figtree", 400, "fonts/figtree-400.woff2"),
    "figtree_600": ("Figtree", 600, "fonts/figtree-600.woff2"),
    "figtree_700": ("Figtree", 700, "fonts/figtree-700.woff2"),
    "figtree_800": ("Figtree", 800, "fonts/figtree-800.woff2"),
    "playfair_800": ("Superior Display", 800, "fonts/playfair-800.woff2"),
    "playfair_900": ("Superior Display", 900, "fonts/playfair-900.woff2"),
}
LOGO_WHITE_FILE = "logos/GPSImpact_White_Horizontal_2026.png"

# A handful of AdImpact station names are long/technical; prettify the ones
# we've seen without hiding anything for names we haven't.
STATION_LABEL_OVERRIDES = {
    "Ampersand Addressable-WI": "Addressable (Ampersand)",
}


def _b64_asset(rel_path):
    with open(ASSETS_DIR / rel_path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


def _font_faces():
    rules = []
    for family, weight, rel_path in FONT_FILES.values():
        rules.append(
            "@font-face {{\n"
            "  font-family: '{family}';\n"
            "  font-style: normal;\n"
            "  font-weight: {weight};\n"
            "  font-display: swap;\n"
            "  src: url(data:font/woff2;base64,{b64}) format('woff2');\n"
            "}}".format(family=family, weight=weight, b64=_b64_asset(rel_path))
        )
    return "\n".join(rules)


def build_report(spending_xlsx, output_html, top_n=6, title=None, race=None, media_types=None,
                  eyebrow="Digital Competitive Report", source="AdImpact"):
    """Parse `spending_xlsx` and write the rendered report to `output_html`."""
    week_cols, week_starts, leaf_rows, meta = parse.load_spending_export(spending_xlsx)
    week_labels, week_iso, index_map = parse.build_continuous_week_axis(week_starts)
    n_weeks = len(week_labels)

    agg = parse.aggregate(leaf_rows, index_map, n_weeks, top_n=top_n)

    for s in agg["platform_series"]:
        s["label"] = STATION_LABEL_OVERRIDES.get(s["key"], s["key"])
    colors.assign_platform_colors(agg["platform_series"])
    colors.assign_top_series_colors(agg["top_series"])

    parties_present = sorted({r["party"] for r in agg["advertiser_rows"]})
    party_colors = {p: colors.party_color(p) for p in parties_present}
    party_colors["__default__"] = colors.PARTY_NEUTRAL
    platform_colors = {s["key"]: s["color"] for s in agg["platform_series"]}

    # weeks AdImpact dropped entirely (no advertiser spent) — surfaced in the footer
    original_positions = set(index_map)
    filled_gap_weeks = [week_labels[i] for i in range(n_weeks) if i not in original_positions]

    race_label = race or meta.get("race")
    media_label = media_types or meta.get("media_types") or "CTV, Digital"
    if title:
        report_title = title
    elif race_label:
        report_title = f"{race_label} — {media_label} Ad Spend"
    else:
        report_title = "Digital Competitive Report"

    payload = {
        "week_labels": week_labels,
        "week_iso": week_iso,
        "n_weeks": n_weeks,
        "grand_total": agg["grand_total"],
        "platform_series": agg["platform_series"],
        "advertiser_rows": agg["advertiser_rows"],
        "top_series": agg["top_series"],
        "other_count": agg["other_count"],
        "party_colors": party_colors,
        "platform_colors": platform_colors,
        "meta": {
            "eyebrow": eyebrow,
            "title": report_title,
            "race": race_label,
            "media_types": media_label,
            "source": source,
            "export_date": meta.get("export_date"),
            "filled_gap_weeks": filled_gap_weeks,
        },
    }

    html = TEMPLATE_PATH.read_text()
    html = html.replace("/*__FONT_FACES__*/", _font_faces())
    html = html.replace("__LOGO_WHITE__", f"data:image/png;base64,{_b64_asset(LOGO_WHITE_FILE)}")
    html = html.replace("__PAGE_TITLE__", report_title)
    html = html.replace("/*__DATA_JSON__*/", json.dumps(payload, separators=(",", ":")))

    Path(output_html).write_text(html)
    return payload
