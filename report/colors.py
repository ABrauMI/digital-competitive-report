"""Color assignment for the report's charts and table.

All hues here are GPS Impact brand-derived and were run through the dataviz
skill's categorical validator (lightness band, chroma floor, CVD separation,
contrast) for both light and dark surfaces before being hardcoded — see
README.md#color-choices. If you add more platform slots or party labels,
re-validate rather than eyeballing a new hex.
"""

# Ordered by validated adjacency (see README). Extra stations beyond this
# list fold into the neutral "Other" slot rather than getting a fresh,
# unvalidated hue.
PLATFORM_PALETTE = [
    {"light": "#7a5aa8", "dark": "#a17fc9"},  # violet
    {"light": "#2f9678", "dark": "#3aab86"},  # green
    {"light": "#c9922f", "dark": "#b58a35"},  # amber (needs direct labels — sub-3:1 contrast)
    {"light": "#b0527a", "dark": "#c8749a"},  # magenta
    {"light": "#a15f26", "dark": "#b98549"},  # brown
]
PLATFORM_NEUTRAL = {"light": "#8a8fa0", "dark": "#838a9e"}

PARTY_COLORS = {
    "Republican": {"light": "#de5e4e", "dark": "#ea8272"},
    "Democrat": {"light": "#2f6ea3", "dark": "#6ea8d8"},
    "Independent": {"light": "#7a5aa8", "dark": "#a17fc9"},
    "Nonpartisan": {"light": "#2f9678", "dark": "#3aab86"},
}
PARTY_NEUTRAL = {"light": "#9aa0ae", "dark": "#838a9e"}

# Light-tint anchor each party family lightens toward for lower-ranked
# advertisers in the "top N" chart (mixed with white in _shade_ramp).
_PARTY_TINT_TARGET = {
    "Republican": {"light": "#f6ddd6", "dark": "#3a2b2c"},
    "Democrat": {"light": "#d9e8f3", "dark": "#22303f"},
    "Independent": {"light": "#e7ddf0", "dark": "#2e2740"},
    "Nonpartisan": {"light": "#d9ece6", "dark": "#1f332e"},
}
_PARTY_NEUTRAL_TINT = {"light": "#e5e7eb", "dark": "#2a2e38"}


def party_color(party):
    return PARTY_COLORS.get(party, PARTY_NEUTRAL)


def assign_platform_colors(platform_series):
    """Assign a validated hue to each platform, ranked by total spend."""
    for i, s in enumerate(platform_series):
        s["color"] = PLATFORM_PALETTE[i] if i < len(PLATFORM_PALETTE) else PLATFORM_NEUTRAL
    return platform_series


def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))


def _rgb_to_hex(rgb):
    return "#" + "".join(f"{max(0, min(255, round(c))):02x}" for c in rgb)


def _lerp_hex(a, b, t):
    ra, rb = _hex_to_rgb(a), _hex_to_rgb(b)
    return _rgb_to_hex(tuple(ra[i] + (rb[i] - ra[i]) * t for i in range(3)))


def assign_top_series_colors(top_series):
    """Shade each party family dark->light by rank within the top-N chart.

    The dominant advertiser in each party gets the full-strength brand hue;
    subsequent same-party advertisers step toward a lighter tint so the
    family reads as related without colliding on an exact hue. "All Other
    Advertisers" always gets the fixed neutral.
    """
    counts = {}
    for s in top_series:
        if s["name"] == "All Other Advertisers":
            s["color"] = PARTY_NEUTRAL
            continue
        party = s["party"]
        rank = counts.get(party, 0)
        counts[party] = rank + 1
        base = PARTY_COLORS.get(party, PARTY_NEUTRAL)
        tint = _PARTY_TINT_TARGET.get(party, _PARTY_NEUTRAL_TINT)
        # up to 4 visible steps before repeating the lightest tint
        t = min(rank, 3) / 3.0
        s["color"] = {
            "light": _lerp_hex(base["light"], tint["light"], t),
            "dark": _lerp_hex(base["dark"], tint["dark"], t),
        }
    return top_series
