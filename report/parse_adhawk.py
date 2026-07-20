"""Parse an AdHawk digital-spend export into the same tidy long-format
leaf rows `report.excel_export` builds a tree from.

AdHawk's export is a flat table, not AdImpact's collapsed pivot — one row
per Advertiser x Spend Platform x Spender Type, a running Total column,
then one 4-column block per week (DEM/GOP/oth/total), no Market/DMA column
at all (AdHawk only carries digital spend, and digital doesn't target by
broadcast DMA the way linear TV does).

Two things AdHawk does differently that this module works around:

- Party isn't a column value — it's whichever of the DEM/GOP/oth totals is
  actually nonzero for that row. GPS Impact's AdHawk usage never splits an
  advertiser's spend across more than one of those columns, so a row with
  more than one nonzero is treated as bad data and raises, rather than
  guessing which party it belongs to.
- The week-block "total" sub-column is unreliable — a handful of weeks
  (always on CTV rows, in every sample seen so far) report exactly double
  the correct amount. We never read it; the weekly total is always
  recomputed as DEM + GOP + oth for that week.
"""
import csv
import re
from datetime import datetime, timedelta

from . import parse

PARTY_LABELS = {"DEM": "Democrat", "GOP": "Republican", "OTH": "Other"}

_WEEK_RE = re.compile(r"^(\d{2}/\d{2})-(\d{2}/\d{2})$")
_FOOTER_MARKERS = ("data produced by adhawk", "ad-hawk.org")


def classify_export(path):
    """Peek at the header row to tell an AdHawk export from anything else.

    Mirrors `report.parse.classify_export`'s role for the Slack app: lets a
    caller check a file matches the source type the user said it would be,
    without fully parsing it.
    """
    try:
        with open(path, newline="", encoding="utf-8-sig") as f:
            header = next(csv.reader(f), None)
    except (OSError, UnicodeDecodeError, StopIteration):
        return None
    if not header:
        return None
    wanted = {"Advertiser", "Election Name", "Spend Platform", "Spender Type"}
    return "adhawk" if wanted.issubset(set(header)) else None


def _num(v):
    v = (v or "").strip()
    return float(v) if v else 0.0


def _find_week_groups(header):
    """Locate each week's 4-column block by its date-range label.

    Returns [(label, dem_idx, gop_idx, oth_idx)] in header order (left to
    right = oldest to newest, per every AdHawk export seen so far). Ignores
    the block's own "total" column — see the module docstring.
    """
    by_label = {}
    order = []
    for i, name in enumerate(header):
        if not isinstance(name, str):
            continue
        parts = name.rsplit(" ", 1)
        if len(parts) != 2:
            continue
        label, suffix = parts
        if not _WEEK_RE.match(label):
            continue
        suffix = suffix.strip().upper()
        if suffix not in ("DEM", "GOP", "OTH", "TOTAL"):
            continue
        if label not in by_label:
            by_label[label] = {}
            order.append(label)
        by_label[label][suffix] = i

    groups = []
    for label in order:
        cols = by_label[label]
        missing = {"DEM", "GOP", "OTH"} - cols.keys()
        if missing:
            raise ValueError(f"Week block {label!r} is missing column(s): {sorted(missing)}")
        groups.append((label, cols["DEM"], cols["GOP"], cols["OTH"]))
    return groups


def _reconstruct_week_dates(labels, as_of):
    """Attach the right year to each "MM/DD-MM/DD" label (AdHawk omits it).

    Anchors the last (most recent) label to a year that puts it on or
    before `as_of`, then walks backward assigning each earlier label a date
    that is some whole number of 7-day steps before the one after it,
    verifying the label's own month/day matches what that implies. Raises
    if a label's date doesn't fit that chain — that means either the
    export isn't in chronological order or the week grid has a gap this
    function doesn't know how to size.
    """
    starts = [datetime.strptime(l.split("-")[0], "%m/%d") for l in labels]

    last = starts[-1]
    candidates = [last.replace(year=as_of.year + delta) for delta in (-1, 0, 1)]
    on_or_before = [d for d in candidates if d <= as_of]
    anchor = max(on_or_before) if on_or_before else min(candidates)

    dates = [None] * len(labels)
    dates[-1] = anchor
    for i in range(len(labels) - 2, -1, -1):
        later = dates[i + 1]
        month, day = starts[i].month, starts[i].day
        found = None
        for n in range(1, 27):  # 27 weeks ~ generous headroom for a gappy export
            expected = later - timedelta(days=7 * n)
            if (expected.month, expected.day) == (month, day):
                found = expected
                break
        if found is None:
            raise ValueError(
                f"Week column {labels[i]!r} doesn't line up as a whole number of weeks "
                f"before {labels[i + 1]!r} — check the export's week columns are in order."
            )
        dates[i] = found
    return dates


def load_adhawk_export(path, as_of=None):
    """Read an AdHawk export and return (week_iso, leaf_rows, meta).

    leaf_rows is a list of dicts with keys: party, advertiser_type,
    advertiser, mediatype ("CTV" or "Digital"), station (the platform name
    — "CTV" itself for CTV rows, else e.g. "Facebook"/"Google"), weekly
    (list aligned to week_iso). meta carries `race` (AdHawk's "Election
    Name") when every row agrees on one, else None.
    """
    as_of = as_of or datetime.now(parse.MEDIA_WEEK_TZ).replace(tzinfo=None)

    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    if not rows:
        raise ValueError("Empty AdHawk export.")

    header = rows[0]
    required = ("Advertiser", "Election Name", "Spend Platform", "Spender Type")
    missing = [c for c in required if c not in header]
    if missing:
        raise ValueError(
            f"Doesn't look like an AdHawk export — missing column(s) {missing}. "
            "Expected Advertiser/Election Name/Spend Platform/Spender Type."
        )
    col = {name: i for i, name in enumerate(header)}
    week_groups = _find_week_groups(header)
    if not week_groups:
        raise ValueError("Couldn't find any weekly DEM/GOP/oth columns in this AdHawk export.")

    week_dates = _reconstruct_week_dates([g[0] for g in week_groups], as_of)
    week_iso = [d.strftime("%Y-%m-%d") for d in week_dates]

    leaf_rows = []
    races = set()
    for r in rows[1:]:
        if not r or not r[0].strip():
            continue
        if any(m in r[0].strip().lower() for m in _FOOTER_MARKERS):
            continue

        advertiser = r[col["Advertiser"]].strip()
        race = r[col["Election Name"]].strip()
        platform = r[col["Spend Platform"]].strip()
        spender_type = r[col["Spender Type"]].strip()

        dem_total = _num(r[col["DEM"]])
        gop_total = _num(r[col["GOP"]])
        oth_total = _num(r[col["Oth"]])
        nonzero = [k for k, v in (("DEM", dem_total), ("GOP", gop_total), ("OTH", oth_total)) if v]
        if len(nonzero) > 1:
            raise ValueError(
                f"{advertiser!r} ({platform}) has spend in more than one party column "
                f"(DEM={dem_total}, GOP={gop_total}, Oth={oth_total}) — expected exactly one."
            )
        if not nonzero:
            continue  # no spend at all on this row; nothing to attribute or plot
        party = PARTY_LABELS[nonzero[0]]

        weekly = [_num(r[dem_i]) + _num(r[gop_i]) + _num(r[oth_i]) for _, dem_i, gop_i, oth_i in week_groups]

        if platform.strip().upper() == "CTV":
            mediatype, station = "CTV", "CTV"
        else:
            mediatype, station = "Digital", platform

        races.add(race)
        leaf_rows.append(
            {
                "party": party,
                "advertiser_type": spender_type,
                "advertiser": advertiser,
                "mediatype": mediatype,
                "station": station,
                "weekly": weekly,
            }
        )

    meta = {"race": races.pop() if len(races) == 1 else None}
    return week_iso, leaf_rows, meta
