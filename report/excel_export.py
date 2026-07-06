"""Write the flat, spreadsheet-style competitive report GPS Impact already
produces for linear TV — same hierarchy, same navy/light-blue shade ramp by
rollup level, same merged-cell nesting — reapplied to digital spend.

Layout (per the linear "Competitive TV Report" template this mirrors):
  CANDIDATE / COMMITTEE -> MARKET -> TYPE -> STATION/PLATFORM, with a Total
  Spend column and one column per week. Subtotal rows roll up Market+Type,
  then the advertiser, then the party, then a grand total.

Digital has no GRPs/CPP equivalent without a separate impressions export, so
those columns are simply omitted for now — spend only, matching how far the
underlying AdImpact export goes.
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NAVY = "1F3355"
HEADER_ACCENT = "ED492E"
LEAF_FILL = "EBF3FA"
SUBTOTAL_FILL = "D8EDF8"
ADV_TOTAL_FILL = "B8D9F0"
PARTY_TOTAL_FILL = "2A669B"
GRAND_FILL = "1F3355"
TEXT_DARK = "1A1A2E"
BORDER_LIGHT = "E5E7EB"
BORDER_MED = "CCCCCC"

LEAF_CURRENCY = '$#,##0;-$#,##0;""'
TOTAL_CURRENCY = "$#,##0"

FONT_NAME = "Calibri"

FIRST_DATA_ROW = 4  # 1=title, 2=blank, 3=header


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
        ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=6)
        c = ws.cell(1, 4, title)
        c.font = Font(name=FONT_NAME, bold=True, size=16, color="FFFFFF")
        c.fill = _fill(NAVY)
        ws.row_dimensions[1].height = 49.5

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

        ws.freeze_panes = "F4"
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
                    ranked_stations = sorted(stations.items(), key=lambda kv: -sum(kv[1]))
                    for station, weekly in ranked_stations:
                        w.write_leaf(station, sum(weekly), weekly)
                    type_end = w.row - 1
                    w.merge_col(3, type_start, type_end)  # TYPE column

                    type_weekly = _sum_lists(stations.values(), n_weeks)
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
            for col, text in enumerate(["MARKET / TYPE", "SPEND"], start=1):
                c = ws.cell(row, col, text)
                c.font = Font(name=FONT_NAME, bold=True, size=9, color="FFFFFF")
                c.fill = _fill(NAVY)
            row += 1
            for market, types in markets.items():
                for mtype, stations in types.items():
                    total = sum(_sum_lists(stations.values(), n_weeks))
                    ws.cell(row, 1, f"{market} — {mtype}").font = Font(name=FONT_NAME, size=9, color=TEXT_DARK)
                    e = ws.cell(row, 2, total)
                    e.number_format = TOTAL_CURRENCY
                    e.font = Font(name=FONT_NAME, size=9, color=TEXT_DARK)
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
