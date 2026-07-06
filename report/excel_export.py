"""Write the flat, spreadsheet-style competitive report GPS Impact already
produces for linear TV — same hierarchy, same shade ramp by rollup level,
same merged-cell nesting — reapplied to digital spend, in GPS Impact's own
2026 brand colors and with the brand logo embedded in the header band.

Layout (per the linear "Competitive TV Report" template this mirrors):
  CANDIDATE / COMMITTEE -> MARKET -> TYPE -> STATION/PLATFORM, with a Total
  Spend column and one column per week. Subtotal rows roll up Market+Type,
  then the advertiser, then the party, then a grand total.

Digital has no GRPs/CPP equivalent without a separate impressions export, so
those columns are simply omitted for now — spend only, matching how far the
underlying AdImpact export goes.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# GPS Impact Brand Guidelines 2026
BRAND_NAVY = "323b51"
BRAND_BLUE = "3d6a91"
BRAND_RED = "de5e4e"
BRAND_PALE_BLUE = "bed7d5"

# Rollup-level shade ramp: light-to-dark tints of BRAND_BLUE, ending in
# BRAND_BLUE itself (party total) and BRAND_NAVY (header/grand total) — see
# README.md#color-choices for how these tints were derived.
NAVY = BRAND_NAVY
HEADER_ACCENT = BRAND_RED
LEAF_FILL = "ecf0f4"
SUBTOTAL_FILL = "d4dee7"
ADV_TOTAL_FILL = "b5c6d5"
PARTY_TOTAL_FILL = BRAND_BLUE
GRAND_FILL = BRAND_NAVY
TEXT_DARK = BRAND_NAVY
BORDER_LIGHT = "E5E7EB"
BORDER_MED = "CCCCCC"
FOOTER_GRAY = "6B7280"

LEAF_CURRENCY = '$#,##0;-$#,##0;""'
TOTAL_CURRENCY = "$#,##0"

FONT_NAME = "Calibri"

FIRST_DATA_ROW = 4  # 1=title, 2=blank, 3=header

ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
LOGO_FILE = ASSETS_DIR / "logos" / "GPSImpact_White_Horizontal_2026.png"


def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def _thin_border(color):
    side = Side(style="thin", color=color)
    return Border(top=side, bottom=side)


def _medium_border(color):
    side = Side(style="medium", color=color)
    return Border(top=side, bottom=side)


def _sum_lists(lists, n_weeks):
    out = [0.0] * n_weeks
    for lst in lists:
        for i in range(n_weeks):
            out[i] += lst[i]
    return out


def build_hierarchy(leaf_rows, index_map, n_weeks):
    """party -> advertiser -> market -> mediatype -> station -> weekly[]."""

    def remap(values):
        out = [0.0] * n_weeks
        for i, v in enumerate(values):
            out[index_map[i]] += v
        return out

    tree = {}
    for r in leaf_rows:
        party_node = tree.setdefault(r["party"], {})
        adv_node = party_node.setdefault(r["advertiser"], {})
        market_node = adv_node.setdefault(r["market"], {})
        type_node = market_node.setdefault(r["mediatype"], {})
        arr = type_node.setdefault(r["station"], [0.0] * n_weeks)
        vals = remap(r["weeks"])
        for i in range(n_weeks):
            arr[i] += vals[i]
    return tree


class _Writer:
    def __init__(self, ws, n_weeks, week_labels, title):
        self.ws = ws
        self.n_weeks = n_weeks
        self.week_labels = week_labels
        self.row = 1
        self._write_scaffold(title)

    def _write_scaffold(self, title):
        ws = self.ws
        ws.sheet_view.showGridLines = False
        last_col = 5 + self.n_weeks
        frozen_cols = 5  # A:E — must match the column freeze_panes below fixes in place

        for col in range(1, last_col + 1):
            ws.cell(1, col).fill = _fill(NAVY)

        # Two merges, not one: a merged cell can't straddle a freeze-pane
        # split without rendering oddly once you scroll, so the title stays
        # inside the frozen columns and the rest of the navy band is a
        # second, separate merge entirely in the scrollable region.
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=frozen_cols)
        c = ws.cell(1, 1, title)
        c.font = Font(name=FONT_NAME, bold=True, size=14, color="FFFFFF")
        c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True, indent=1)
        if last_col > frozen_cols:
            ws.merge_cells(start_row=1, start_column=frozen_cols + 1, end_row=1, end_column=last_col)
        ws.row_dimensions[1].height = 49.5

        if LOGO_FILE.exists():
            logo = XLImage(str(LOGO_FILE))
            aspect = logo.width / logo.height
            logo.height = 34
            logo.width = 34 * aspect
            ws.add_image(logo, "A1")

        headers = ["CANDIDATE / COMMITTEE", "MARKET", "TYPE", "STATION / PLATFORM", "TOTAL SPEND"] + self.week_labels
        for col, text in enumerate(headers, start=1):
            c = ws.cell(3, col, text)
            c.font = Font(name=FONT_NAME, bold=True, size=9, color="FFFFFF")
            c.fill = _fill(NAVY)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = _medium_border(HEADER_ACCENT)
        ws.row_dimensions[3].height = 31.5

        widths = {"A": 28, "B": 24, "C": 10, "D": 30, "E": 14}
        for col_letter, w in widths.items():
            ws.column_dimensions[col_letter].width = w
        for i in range(self.n_weeks):
            ws.column_dimensions[get_column_letter(6 + i)].width = 13

        ws.freeze_panes = f"{get_column_letter(frozen_cols + 1)}{FIRST_DATA_ROW}"
        self.row = FIRST_DATA_ROW

    def _write_row(self, total_value, weekly_values, fill, bold, text_color, border, font_size=9,
                    currency=LEAF_CURRENCY, label=None, label_col=4):
        ws = self.ws
        r = self.row
        font = Font(name=FONT_NAME, bold=bold, size=font_size, color=text_color)
        last_col = 5 + self.n_weeks
        for col in range(1, last_col + 1):
            c = ws.cell(r, col)
            c.fill = _fill(fill)
            c.border = border
            c.font = font
        if label is not None:
            ws.cell(r, label_col, label)
        e = ws.cell(r, 5, total_value)
        e.number_format = currency
        for i, v in enumerate(weekly_values):
            cc = ws.cell(r, 6 + i, v)
            cc.number_format = currency
        self.row += 1
        return r

    def write_leaf(self, station, total, weekly):
        r = self.row
        ws = self.ws
        ws.cell(r, 4, station).font = Font(name=FONT_NAME, size=9, color=TEXT_DARK)
        for col in (1, 2, 3, 4, 5):
            ws.cell(r, col).fill = _fill(LEAF_FILL)
            ws.cell(r, col).border = _thin_border(BORDER_LIGHT)
        ws.cell(r, 5, total).number_format = LEAF_CURRENCY
        ws.cell(r, 5).font = Font(name=FONT_NAME, size=9, color=TEXT_DARK)
        for i, v in enumerate(weekly):
            cc = ws.cell(r, 6 + i, v if v else None)
            cc.number_format = LEAF_CURRENCY
            cc.font = Font(name=FONT_NAME, size=9, color=TEXT_DARK)
            cc.fill = _fill(LEAF_FILL)
            cc.border = _thin_border(BORDER_LIGHT)
        self.row += 1
        return r

    def write_subtotal(self, label, total, weekly):
        return self._write_row(total, weekly, SUBTOTAL_FILL, True, TEXT_DARK, _thin_border(BORDER_MED),
                                currency=TOTAL_CURRENCY, label=label)

    def write_advertiser_total(self, label, total, weekly):
        return self._write_row(total, weekly, ADV_TOTAL_FILL, True, TEXT_DARK, _medium_border(BORDER_MED),
                                currency=TOTAL_CURRENCY, label=label)

    def write_party_total(self, label, total, weekly):
        return self._write_row(
            total, weekly, PARTY_TOTAL_FILL, True, "FFFFFF", _medium_border("FFFFFF"), font_size=10,
            currency=TOTAL_CURRENCY, label=label, label_col=1,
        )

    def write_grand_total(self, label, total, weekly):
        return self._write_row(
            total, weekly, GRAND_FILL, True, "FFFFFF", _medium_border("FFFFFF"), font_size=10,
            currency=TOTAL_CURRENCY, label=label, label_col=1,
        )

    def blank_row(self):
        self.row += 1

    def write_footer(self):
        last_col = 5 + self.n_weeks
        self.blank_row()
        r = self.row
        self.ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        c = self.ws.cell(r, 1, "Report prepared by GPS Impact  |  Confidential")
        c.font = Font(name=FONT_NAME, italic=True, size=8, color=FOOTER_GRAY)
        self.ws.row_dimensions[r].height = 18
        self.row += 1

    def write_group_label(self, col, row, text):
        c = self.ws.cell(row, col, text)
        c.font = Font(name=FONT_NAME, bold=True, size=9, color=TEXT_DARK)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def merge_col(self, col, start_row, end_row):
        if end_row > start_row:
            self.ws.merge_cells(start_row=start_row, start_column=col, end_row=end_row, end_column=col)
            top = self.ws.cell(start_row, col)
            top.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def merge_label_row(self, row, start_col=1, end_col=4):
        self.ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)


def _write_main_sheet(wb, tree, week_labels, n_weeks, title):
    ws = wb.active
    ws.title = "Competitive Digital Report"
    w = _Writer(ws, n_weeks, week_labels, title)

    parties = sorted(tree.items(), key=lambda kv: -sum(_sum_lists(
        [s for adv in kv[1].values() for mk in adv.values() for tp in mk.values() for s in tp.values()], n_weeks
    )))

    for party, advertisers in parties:
        adv_totals = {
            adv: _sum_lists([s for mk in markets.values() for tp in mk.values() for s in tp.values()], n_weeks)
            for adv, markets in advertisers.items()
        }
        ranked_advs = sorted(advertisers.items(), key=lambda kv: -sum(adv_totals[kv[0]]))

        party_weekly = _sum_lists(adv_totals.values(), n_weeks)

        for adv, markets in ranked_advs:
            adv_start = w.row
            w.write_group_label(1, adv_start, adv)
            for market, types in markets.items():
                market_start = w.row
                w.write_group_label(2, market_start, market)
                for mtype, stations in types.items():
                    type_start = w.row
                    w.write_group_label(3, type_start, mtype)
                    type_weekly = _sum_lists(stations.values(), n_weeks)

                    if mtype.strip().upper() == "CTV":
                        # All CTV platforms (In-App, Device, Streaming, ...)
                        # roll into one combined line per market rather than
                        # breaking out by individual platform.
                        w.write_subtotal(f"{market} - {mtype} Total", sum(type_weekly), type_weekly)
                    else:
                        ranked_stations = sorted(stations.items(), key=lambda kv: -sum(kv[1]))
                        for station, weekly in ranked_stations:
                            w.write_leaf(station, sum(weekly), weekly)
                        type_end = w.row - 1
                        w.merge_col(3, type_start, type_end)  # TYPE column
                        w.write_subtotal(f"{market} - {mtype} Total", sum(type_weekly), type_weekly)
                market_end = w.row - 1
                w.merge_col(2, market_start, market_end)  # MARKET column

            w.write_advertiser_total(f"{adv} Total", sum(adv_totals[adv]), adv_totals[adv])
            adv_end = w.row - 1
            w.merge_col(1, adv_start, adv_end)  # CANDIDATE / COMMITTEE column
            w.blank_row()

        w.write_party_total(f"{party.upper()} PARTY TOTAL", sum(party_weekly), party_weekly)
        w.merge_label_row(w.row - 1)

    grand_weekly = _sum_lists(
        [_sum_lists(
            [s for adv in advertisers.values() for mk in adv.values() for tp in mk.values() for s in tp.values()],
            n_weeks,
        ) for _, advertisers in parties],
        n_weeks,
    )
    w.write_grand_total("GRAND TOTAL", sum(grand_weekly), grand_weekly)
    w.merge_label_row(w.row - 1)
    w.write_footer()
    return ws


def _write_market_summary_sheet(wb, tree, n_weeks):
    ws = wb.create_sheet("Market Summary")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 16

    c = ws.cell(1, 1, "Market Summary")
    c.font = Font(name=FONT_NAME, bold=True, size=14, color="FFFFFF")
    c.fill = _fill(NAVY)
    ws.merge_cells("A1:B1")

    parties = sorted(tree.items(), key=lambda kv: -sum(_sum_lists(
        [s for adv in kv[1].values() for mk in adv.values() for tp in mk.values() for s in tp.values()], n_weeks
    )))

    row = 3
    for _, advertisers in parties:
        adv_totals = {
            adv: sum(_sum_lists([s for mk in markets.values() for tp in mk.values() for s in tp.values()], n_weeks))
            for adv, markets in advertisers.items()
        }
        ranked_advs = sorted(advertisers.items(), key=lambda kv: -adv_totals[kv[0]])

        for adv, markets in ranked_advs:
            row += 1
            c = ws.cell(row, 1, adv)
            c.font = Font(name=FONT_NAME, bold=True, size=11, color=TEXT_DARK)
            row += 1
            for col, text in enumerate(["MARKET", "SPEND"], start=1):
                c = ws.cell(row, col, text)
                c.font = Font(name=FONT_NAME, bold=True, size=9, color="FFFFFF")
                c.fill = _fill(NAVY)
            row += 1

            # Group by media type first (all CTV markets together, all
            # Digital markets together), each with its own type total,
            # biggest type first.
            by_type = {}
            for market, types in markets.items():
                for mtype, stations in types.items():
                    total = sum(_sum_lists(stations.values(), n_weeks))
                    by_type.setdefault(mtype, []).append((market, total))
            ranked_types = sorted(by_type.items(), key=lambda kv: -sum(t for _, t in kv[1]))

            for mtype, market_totals in ranked_types:
                c = ws.cell(row, 1, mtype.upper())
                c.font = Font(name=FONT_NAME, bold=True, size=9, color=TEXT_DARK)
                row += 1
                for market, total in sorted(market_totals, key=lambda x: -x[1]):
                    ws.cell(row, 1, market).font = Font(name=FONT_NAME, size=9, color=TEXT_DARK)
                    e = ws.cell(row, 2, total)
                    e.number_format = TOTAL_CURRENCY
                    e.font = Font(name=FONT_NAME, size=9, color=TEXT_DARK)
                    row += 1
                c = ws.cell(row, 1, f"{mtype} Total")
                c.font = Font(name=FONT_NAME, bold=True, size=9, color=TEXT_DARK)
                c.fill = _fill(SUBTOTAL_FILL)
                e = ws.cell(row, 2, sum(t for _, t in market_totals))
                e.number_format = TOTAL_CURRENCY
                e.font = Font(name=FONT_NAME, bold=True, size=9, color=TEXT_DARK)
                e.fill = _fill(SUBTOTAL_FILL)
                row += 1

            c = ws.cell(row, 1, f"{adv} — Total")
            c.font = Font(name=FONT_NAME, bold=True, size=9, color=TEXT_DARK)
            c.fill = _fill(ADV_TOTAL_FILL)
            e = ws.cell(row, 2, adv_totals[adv])
            e.number_format = TOTAL_CURRENCY
            e.font = Font(name=FONT_NAME, bold=True, size=9, color=TEXT_DARK)
            e.fill = _fill(ADV_TOTAL_FILL)
            row += 2
    return ws


def write_excel_report(leaf_rows, index_map, week_labels, output_path, title="DIGITAL COMPETITIVE REPORT"):
    n_weeks = len(week_labels)
    tree = build_hierarchy(leaf_rows, index_map, n_weeks)
    wb = Workbook()
    _write_main_sheet(wb, tree, week_labels, n_weeks, title)
    _write_market_summary_sheet(wb, tree, n_weeks)
    wb.save(output_path)
