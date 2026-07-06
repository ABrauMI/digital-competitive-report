# Digital Competitive Report

Turns an AdImpact "Spending Chart" export into a branded, interactive HTML
report: total spend, weekly spend by platform, weekly spend by advertiser,
and a full advertiser table with per-platform breakdown and a share-of-voice
narrative — all in GPS Impact's brand colors and type.

## Quick start

```bash
pip install -r requirements.txt
python generate_report.py --spending path/to/WIGov_Spending.xlsx --output report.html
```

Then open `report.html` in a browser (it's fully self-contained — fonts,
logo, and data are all inlined, so it can be emailed or dropped into a
shared drive as a single file).

## Getting the input file

In AdMo+ / AdImpact, pull a **Spending Chart** export:

- Rows: `Party, AdvertiserType, Advertiser, Market, MediaType, Station`
- Value Type: `Spend`
- Date Grouping: `Weeks`
- Filter to a single race (Race = `<your race>`) and whatever media types
  you want tracked (typically `CTV, Digital`)

The script auto-detects the race and media types from the export's header
block, so `--race` / `--media-types` / `--title` are only needed if you want
to override what's in the file.

## CLI options

| Flag | Default | What it does |
|---|---|---|
| `--spending` | *(required)* | Path to the AdImpact Spending Chart `.xlsx` |
| `--output` | `report.html` | Output path |
| `--top-n` | `6` | How many advertisers get their own line in the "weekly spend by advertiser" chart; everyone else rolls into "All Other Advertisers" |
| `--title` | derived | Overrides the masthead title |
| `--race` | from file | Overrides the race label |
| `--media-types` | from file | Overrides the media types label |

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
generate_report.py   CLI entrypoint
report/
  parse.py           reads the xlsx, forward-fills the pivot, aggregates
  colors.py           brand palette + color assignment (platforms, parties, advertisers)
  build.py            wires parse -> colors -> template -> output html
  template.html      the report itself (HTML/CSS/vanilla-JS, no build step)
assets/
  fonts/              Figtree + Playfair Display, woff2, embedded at build time
  logos/               GPS Impact logo, embedded at build time
```

There's no bundler — `template.html` is plain HTML/CSS/JS with a few string
placeholders (`__PAGE_TITLE__`, `__LOGO_WHITE__`, `/*__FONT_FACES__*/`,
`/*__DATA_JSON__*/`) that `report/build.py` fills in.

## Color choices

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

## Data & confidentiality

AdImpact exports carry a confidentiality notice and contain competitive
intelligence about live campaigns. `.gitignore` excludes `*.xlsx` and
generated `report.html` / `DigitalCompetitiveReport_*.html` files so raw
exports and rendered reports don't end up in git history — keep them in a
shared drive instead.
