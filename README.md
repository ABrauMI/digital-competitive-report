# Digital Competitive Report

Turns an AdImpact "Spending Chart" export into a competitive report on who's
spending what, where, on which platform, week by week. Two output modes:

- **Excel (start here)** — the same flat, spreadsheet-style competitive
  report GPS Impact already produces for linear TV (Candidate/Committee →
  Market → Type → Station, weekly spend columns, subtotals, party/grand
  totals), reapplied to digital spend.
- **HTML dashboard** — a richer, branded interactive report (charts,
  sparklines, a written summary) for when the flat table isn't enough.

## Quick start

```bash
pip install -r requirements.txt

# Flat Excel report — mirrors the linear TV competitive report format
python generate_excel_report.py --spending path/to/Spending.xlsx --output report.xlsx

# Branded HTML dashboard
python generate_report.py --spending path/to/Spending.xlsx --output report.html
```

## Getting the input file

In AdMo+ / AdImpact, pull a **Spending Chart** export:

- Rows: `Party, AdvertiserType, Advertiser, Market, MediaType, Station`
- Value Type: `Spend`
- Date Grouping: `Weeks`
- Filter to a single race (Race = `<your race>`) and whatever media types
  you want tracked (typically `CTV, Digital`)

Both generators auto-detect the race and media types from the export's
header block.

## `generate_excel_report.py`

Produces a workbook with two sheets:

- **Competitive Digital Report** — Candidate/Committee → Market → Type →
  Station/Platform, with a Total Spend column and one column per week.
  Subtotals roll up Market+Type, then the advertiser, then the party, then a
  grand total — same shade-ramp-by-rollup-level styling as the linear
  template (navy header, escalating light-blue fills, merged hierarchy
  cells, frozen panes so the label columns stay visible while scrolling
  through weeks).
- **Market Summary** — per-advertiser market/type totals, no weekly detail.

Digital has no GRPs/CPP equivalent without a separate impressions export, so
those columns are simply omitted for now — spend only. If you pull a weekly
*Impressions* export from AdImpact later, that's the natural next column to
add (impressions is the honest digital analog to GRPs — not something to
fabricate from spend).

| Flag | Default | What it does |
|---|---|---|
| `--spending` | *(required)* | Path to the AdImpact Spending Chart `.xlsx` |
| `--output` | `competitive_report.xlsx` | Output path |
| `--title` | derived | Overrides the report title |

## `generate_report.py` (HTML dashboard)

| Flag | Default | What it does |
|---|---|---|
| `--spending` | *(required)* | Path to the AdImpact Spending Chart `.xlsx` |
| `--output` | `report.html` | Output path |
| `--top-n` | `6` | How many advertisers get their own line in the "weekly spend by advertiser" chart; everyone else rolls into "All Other Advertisers" |
| `--title` | derived | Overrides the masthead title |
| `--race` | from file | Overrides the race label |
| `--media-types` | from file | Overrides the media types label |

Self-contained — fonts, logo, and data are all inlined, so `report.html` can
be emailed or dropped into a shared drive as a single file.

## How the parsing works

AdImpact's export is a collapsed pivot table — blank cells mean "same value
as the row above," and every group gets a redundant `<LEVEL> TOTAL` subtotal
row. `report/parse.py` forward-fills the hierarchy, drops the subtotal rows,
and re-aggregates from the leaf (Station-level) rows. It also re-derives a
continuous weekly axis: AdImpact drops any week that's $0 across every
advertiser instead of keeping a fixed weekly grid, which would otherwise
make the x-axis silently non-uniform.

## Repo layout

```
generate_excel_report.py   CLI entrypoint — flat Excel report
generate_report.py         CLI entrypoint — HTML dashboard
report/
  parse.py           reads the xlsx, forward-fills the pivot, aggregates
  excel_export.py    writes the flat, spreadsheet-style report (mirrors the linear TV template)
  colors.py           brand palette + color assignment (platforms, parties, advertisers) — HTML dashboard only
  build.py            wires parse -> colors -> template -> output html
  template.html      the HTML dashboard itself (HTML/CSS/vanilla-JS, no build step)
assets/
  fonts/              Figtree + Playfair Display, woff2, embedded at build time (HTML dashboard only)
  logos/               GPS Impact logo, embedded at build time (HTML dashboard only)
```

There's no bundler — `template.html` is plain HTML/CSS/JS with a few string
placeholders (`__PAGE_TITLE__`, `__LOGO_WHITE__`, `/*__FONT_FACES__*/`,
`/*__DATA_JSON__*/`) that `report/build.py` fills in.

## Color choices (HTML dashboard)

The palette is derived from `GPSImpact_BrandGuidelines_2026`: Navy `#323b51`,
Blue `#3d6a91`, Red `#de5e4e`, Pale Blue `#bed7d5`. The brand's own blue and
red are correct for the masthead/logo but read as low-chroma ("grayish") once
used as chart-series colors, so the chart palette in `report/colors.py` uses
boosted-saturation variants of the same hues plus a small set of
brand-adjacent extra hues for platforms that aren't Republican/Democrat.

Every categorical palette here (party colors, platform colors) was run
through the [dataviz skill](https://github.com/anthropics/skills)'s
`validate_palette.js` for both light and dark chart surfaces — lightness
band, chroma floor, CVD (colorblind) separation, and contrast — rather than
picked by eye. If you add a platform beyond the five validated slots in
`PLATFORM_PALETTE`, or a party beyond the four in `PARTY_COLORS`, re-run the
validator on the new set instead of eyeballing a hex value; both fall back to
a neutral gray for anything beyond what's been validated.

## Color choices (Excel report)

`report/excel_export.py` mirrors the linear TV report's *layout* (hierarchy,
merges, rollup shading) but its *colors* come straight from the brand
guidelines rather than the specific hex values sampled off that one linear
file (which was closer but not exact — e.g. its navy was `#1F3355` and its
header-accent red was `#ED492E`, not the brand's `#323b51` / `#de5e4e`):

- Header band, grand total, and the GPS Impact logo (embedded in the header,
  `assets/logos/GPSImpact_White_Horizontal_2026.png`): brand Navy `#323b51`
- Header underline accent: brand Red `#de5e4e`
- Body text: brand Navy `#323b51`
- Footer line ("Report prepared by GPS Impact | Confidential"): neutral gray,
  matching the linear template's own footer convention

Each advertiser's whole block (leaf rows → market/type subtotal → CTV/Digital
type total → advertiser total) is shaded in a light-to-dark tint ramp of
**its own party's color** rather than one blue ramp for everyone —
`PARTY_RAMPS` in `report/excel_export.py`:

| Party | Ramp | Party-total row |
|---|---|---|
| Republican | tints of brand Red `#de5e4e` | solid `#de5e4e` |
| Democrat | tints of brand Blue `#3d6a91` | solid `#3d6a91` |
| anything else (Independent, Nonpartisan, ...) | tints of neutral gray | solid gray |

Grand Total always stays neutral Navy — it isn't party-specific. To add a
party the ramp doesn't cover yet, add an entry to `PARTY_RAMPS` with the same
four tint steps (10%/22%/30%/38% toward white, landing on the full color) so
the hierarchy still reads the same way.

## Data & confidentiality

AdImpact exports carry a confidentiality notice and contain competitive
intelligence about live campaigns. `.gitignore` excludes `*.xlsx` and
generated report files so raw exports and rendered reports don't end up in
git history — keep them in a shared drive instead.
