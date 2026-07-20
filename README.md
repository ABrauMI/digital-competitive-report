# Digital Competitive Report

Turns a digital-spend export into a competitive report on who's spending
what, where, on which platform, week by week — the same flat,
spreadsheet-style competitive report GPS Impact already produces for linear
TV (Candidate/Committee → Market → Type → Station, weekly spend columns,
subtotals, party/grand totals), reapplied to digital spend.

Two source formats are supported, picked with `--source` (CLI) or the
modal (Slack):

- **AdImpact** (default) — full report: Competitive Digital Report, Market
  Summary, This Week, and (with `--creative`) Creative Timeline.
- **AdHawk** — Competitive Digital Report and This Week only. AdHawk is
  digital-only spend data and carries no market/DMA column at all (digital
  doesn't target by broadcast market the way linear TV does), so there's no
  Market Summary tab, and no Creative Timeline yet.

## Quick start

```bash
pip install -r requirements.txt
python generate_excel_report.py --spending path/to/Spending.xlsx --output report.xlsx

# AdHawk:
python generate_excel_report.py --source adhawk --spending path/to/export.csv --output report.xlsx
```

Or skip the CLI entirely and run it from Slack — see **Slack app** below.

## Getting the input file

### AdImpact

In AdMo+ / AdImpact, pull a **Spending Chart** export:

- Rows: `Party, AdvertiserType, Advertiser, Market, MediaType, Station`
- Value Type: `Spend`
- Date Grouping: `Weeks`
- Filter to a single race (Race = `<your race>`) and whatever media types
  you want tracked (typically `CTV, Digital`)

The generator auto-detects the race and media types from the export's
header block.

Optionally, also pull a **Topline Creatives** export (Rows: `Creative
(FPUUID)`, same race/media-type filters) to get a Creative Timeline tab —
see `--creative` below.

### AdHawk

Pull AdHawk's export as a `.csv` — one row per Advertiser × Spend Platform
× Spender Type, with `DEM`/`GOP`/`Oth` spend columns (overall and per week)
rather than a single Party column. The generator infers each row's party
from whichever of those columns actually has spend in it; a row with
nonzero spend in more than one is treated as bad data and raises, since
that's not expected to happen. It also recomputes every weekly total as
`DEM + GOP + Oth` itself rather than trusting AdHawk's own per-week "total"
sub-column — that column has been observed reporting exactly double the
correct amount on a handful of weeks (always on CTV rows).

AdHawk's week-range column headers (`MM/DD-MM/DD`) don't include a year;
the generator reconstructs one by walking backward from the most recent
column, so it needs `--current-week` or today's real date to anchor
correctly — don't feed it a stale system clock.

## `generate_excel_report.py`

With `--source adimpact` (the default), produces a workbook with three
sheets, in this tab order (Competitive Report opens as the active tab):

- **Competitive Digital Report** — Candidate/Committee → Market → Type →
  Station/Platform, with a Total Spend column and one column per week.
  Subtotals roll up Market+Type, then CTV/Digital, then the advertiser, then
  the party, then a grand total — same shade-ramp-by-rollup-level styling as
  the linear template (navy header, merged hierarchy cells, frozen panes so
  the label columns stay visible while scrolling through weeks).
- **Market Summary** — per-advertiser market/type totals, no weekly detail.
  CTV stays collapsed to one line per market, but Digital breaks out each
  market's platforms (Facebook, Google, ...) the same way the main sheet
  does, rolling up to a market total, then a Digital total.
- **This Week** — same hierarchy as the main sheet, but scoped to just the
  current media week, compared against the week before it (spend, $ change,
  % change), with any advertiser/market/platform that spent nothing this
  week dropped entirely — this is "what's running right now," not a
  historical view.

  "This week" means the **Tuesday on or before today, in America/New
  York** (GPS Impact's media-week convention — Tuesday through Monday),
  not just whichever column happens to be rightmost in the export. Those
  are only the same thing when the export is fully caught up to today; if
  it's a day or two stale, or you're regenerating an old report, they
  diverge. If the resolved week genuinely isn't in the export yet, the tab
  shows a plain "no spending data available for this week yet" placeholder
  instead of silently substituting some other week.

  Override with `--current-week YYYY-MM-DD` (the Tuesday it starts on) to
  pin a specific week — for regenerating a past week's report, or for
  testing without waiting for a real Tuesday.
- **Creative Timeline** (only with `--creative`, appended last) — one row
  per creative, grouped by party then advertiser (same red/blue shading as
  the rest of the workbook), sorted chronologically within each advertiser.
  Candidate/Committee, Creative title, Platform (CTV or Digital — a
  creative in this export is always exactly one, never both), Tone, Total
  Spend, Start/End Date, then a shaded bar across every media week the
  flight overlapped. A shaded week means "live at some point that week,"
  not necessarily the full week — flights rarely start or end on a
  Tuesday. Its own week axis, independent of the spending file's, spanning
  from the earliest creative's start to the latest one's end.

  The Topline Creatives export has no Party column of its own — party
  coloring is looked up by advertiser name from the Spending Chart export,
  so `--spending` and `--creative` should be pulled for the same race.
  Broadcast/Cable spend columns in that export are read and discarded; if
  you pulled it filtered to Media Type = CTV, Digital (as documented
  above) they're $0 anyway, not just irrelevant.

On both the main sheet and This Week, CTV platforms (In-App, Device,
Streaming, ...) roll into a single combined line per market instead of
breaking out individually — Digital platforms (Google, Facebook, ...) are
unaffected and still show per-platform detail.

Digital has no GRPs/CPP equivalent without a separate impressions export, so
those columns are simply omitted for now — spend only. If you pull a weekly
*Impressions* export from AdImpact later, that's the natural next column to
add (impressions is the honest digital analog to GRPs — not something to
fabricate from spend).

With `--source adhawk`, the workbook has just **Competitive Digital
Report** and **This Week** — both the same idea as above minus the Market
column (AdHawk has no market data to group by) and minus the CTV/Digital
Type column too (AdHawk's own "Digital" grouping just means Facebook or
Google, which reads as a redundant label once Market isn't there to nest
it under). Platforms sit flat under each advertiser instead: Candidate/
Committee → Platform, one row each for CTV, Facebook, Google, etc. There's
no Market Summary tab (nothing to summarize by) and no Creative Timeline
yet.

| Flag | Default | What it does |
|---|---|---|
| `--source` | `adimpact` | `adimpact` or `adhawk` — which export format `--spending` is |
| `--spending` | *(required)* | Path to the AdImpact `.xlsx` or AdHawk `.csv` export |
| `--output` | `competitive_report.xlsx` | Output path |
| `--title` | derived | Overrides the report title |
| `--current-week` | today's media week | Pins the This Week tab to a specific week (`YYYY-MM-DD`, the Tuesday it starts on) |
| `--creative` | *(none)* | Path to an AdImpact Topline Creatives `.xlsx` export — adds the Creative Timeline tab. `--source adimpact` only |

## Slack app

A `/digital-comp` slash command that runs the same pipeline as
`generate_excel_report.py` without anyone needing Python installed: type
the command, pick a data source (AdImpact or AdHawk) and optionally set a
title override in the modal, reply in the thread with your export(s) —
for AdImpact, a Spending Chart export and optionally a Topline Creatives
export (order doesn't matter, the bot tells them apart by their header
row); for AdHawk, a single `.csv` — click **Build Report**, get the
workbook back in the thread. Built the same way as GPS Impact's existing
Sample Buy bot (`broadcast-buy-optimizer`) — Slack Bolt in Socket Mode, so
it needs no public URL or webhook, just an outbound connection to Slack.

- **`SLACK_SETUP.md`** — one-time Slack dashboard setup (app creation,
  Socket Mode, bot scopes, the slash command, event subscriptions);
  this part can only be done by a human with access to your Slack workspace.
- **`RAILWAY_DEPLOY.md`** — running it as an always-on Railway worker so it
  doesn't depend on someone's laptop being open.

The Slack app and the CLI share `report/pipeline.py`, so anything that
changes report behavior (new columns, styling, tabs) applies to both
automatically — there's no separate code path to keep in sync.

## How the parsing works

AdImpact's export is a collapsed pivot table — blank cells mean "same value
as the row above," and every group gets a redundant `<LEVEL> TOTAL` subtotal
row. `report/parse.py` forward-fills the hierarchy, drops the subtotal rows,
and re-aggregates from the leaf (Station-level) rows. It also re-derives a
continuous weekly axis: AdImpact drops any week that's $0 across every
advertiser instead of keeping a fixed weekly grid, which would otherwise
make the x-axis silently non-uniform.

AdHawk's export (`report/parse_adhawk.py`) is already a flat, tidy table —
one row per Advertiser × Spend Platform × Spender Type, no forward-filling
or subtotal rows to strip. The two things it does need to work around: party
is inferred from whichever of the DEM/GOP/Oth columns is nonzero rather than
read from a Party column, and the week-range headers (`MM/DD-MM/DD`) don't
carry a year, so it's reconstructed by walking backward from the most recent
column in fixed 7-day steps.

## Repo layout

```
generate_excel_report.py   CLI entrypoint
report/
  parse.py           reads AdImpact exports, forward-fills the pivot, aggregates
  parse_adhawk.py    reads AdHawk's flat .csv export, infers party, reconstructs week dates
  pipeline.py         build_digital_competitive_report() / build_adhawk_competitive_report() — shared by the CLI and the Slack app
  excel_export.py    writes the flat, spreadsheet-style report (mirrors the linear TV template)
slack_app/
  app.py               /digital-comp slash command, modal, file intake, Build Report button
  session.py           in-memory session store keyed by thread ts
assets/
  logos/               GPS Impact logo, embedded in the report header at build time
```

## Color choices

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
| anything else (Independent, Nonpartisan, AdHawk's "Other", ...) | tints of neutral gray | solid gray |

Grand Total always stays neutral Navy — it isn't party-specific. To add a
party the ramp doesn't cover yet, add an entry to `PARTY_RAMPS` with the same
four tint steps (10%/22%/30%/38% toward white, landing on the full color) so
the hierarchy still reads the same way.

## Data & confidentiality

AdImpact and AdHawk exports both carry competitive intelligence about live
campaigns. `.gitignore` excludes `*.xlsx`, `*.xls`, and `*.csv` so raw
exports and rendered reports don't end up in git history — keep them in a
shared drive instead.
