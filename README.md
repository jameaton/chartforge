# ChartForge

### Open-source charting for PowerPoint

**Professional charts. No licence fees. No limits.**

Waterfalls, Mekkos, Gantts and stacked columns with growth arrows, drawn as native, editable PowerPoint shapes with one click and editable at any time.

ChartForge is a PowerPoint add-in (`.ppam`) written in VBA. It needs no server, no account and no IT deployment: you load one file through PowerPoint's own Add-ins menu.

---

## Features

| Chart | What it's for |
|---|---|
| **Stacked** | Column totals and composition, with totals above each column |
| **Clustered** | Side-by-side comparison of series |
| **100%** | Share of each column, labelled in percentages |
| **Waterfall** | Bridges from a start value to an end value; type `e` for a computed running total |
| **Mekko** | Variable-width 100% columns, e.g. market size × market share |
| **Gantt** | Timelines with bars and milestones; month, quarter or year scale chosen automatically |

- **Growth arrows** on column charts: CAGR, % change or absolute difference between any two columns.
- **Datasheet editing.** Data lives in an ordinary PowerPoint table, so you can type straight into it or paste from Excel. **Edit data** brings the table back for any chart.
- **Native shapes.** Every chart is a group of rectangles, lines and text boxes, so it's fully editable and looks identical for anyone who opens the deck, even without ChartForge.
- **Resize and redraw.** Resize a chart, click **Update**, and it re-lays itself out at the new size.
- **Labels and formatting.** Toggle value labels, totals and legend; change decimal places; add a prefix or suffix (`$`, `bn`, `%`).
- **House style.** Default palette: blue `#0047BA`, `#99B5E3`, grey `#575756`, yellow `#FFED00`. Default font: IBM Plex Sans, switchable to Calibri.

## Install

**Download [`ChartForge.ppam`](../../releases/latest) from the latest release**, then follow the [install guide](docs/INSTALL.md). In short:

- **Mac:** PowerPoint → **Tools → PowerPoint Add-ins… → +**, choose `ChartForge.ppam`, then **Enable Macros**.
- **Windows:** PowerPoint → **File → Options → Add-ins**, set **Manage** to *PowerPoint Add-ins*, click **Go… → Add New…**, choose `ChartForge.ppam`, then **Enable Macros**.

A **ChartForge** tab appears on the ribbon.

## Quick start

1. Select a slide and click **Stacked** on the ChartForge tab. A sample datasheet table appears.
2. Replace the sample data: type into the table, or paste a block of cells from Excel.
3. Keep the table selected and click **Stacked** again. The chart is drawn where the table was.
4. With the chart selected, try **CAGR**, **Values** or **More decimals**.
5. To change the data later, select the chart, click **Edit data**, edit the table, then click **Update**.

Clicking a different column type (**Clustered**, **100%**) while a chart is selected switches it, and the data is kept.

## Datasheet layouts

**Stacked, Clustered, 100%.** Categories across the top, one row per series:

|  | 2022 | 2023 | 2024 |
|---|---|---|---|
| Vaccines | 12 | 15 | 19 |
| Therapeutics | 8 | 9 | 12 |

**Waterfall.** The first value is the starting bar; the rest are changes. Type `e` to show a running total:

|  | 2024 | New launches | Pricing | FX | 2025 |
|---|---|---|---|---|---|
| Revenue | 100 | 25 | 8 | -4 | e |

**Mekko.** The first row sets column widths; the rows below split each column:

|  | US | EU | China |
|---|---|---|---|
| Market size | 120 | 80 | 50 |
| Company A | 30 | 25 | 10 |
| Company B | 25 | 15 | 5 |

**Gantt.** One task per row. Dates as `YYYY-MM-DD` or `DD/MM/YYYY`; leave **End** blank for a milestone:

| Task | Start | End |
|---|---|---|
| Phase I | 2026-01-01 | 2026-06-30 |
| Filing | 2028-09-01 | |

Numbers may include thousands separators, currency symbols, `%`, or brackets for negatives: `(12)` means −12.

## Ribbon reference

| Group | Button | Does |
|---|---|---|
| Charts | Stacked … Gantt | Inserts a sample datasheet, or draws / switches the selected datasheet or chart |
| Data | Edit data | Opens the selected chart's data as a table |
| Data | Update | Applies the selected datasheet, or redraws the selected chart after a resize |
| Growth arrow | CAGR / % change / Difference | Adds an arrow between two columns (first to last by default) |
| Growth arrow | Columns… | Chooses which columns the arrow connects, e.g. `2,5` |
| Growth arrow | Remove | Removes the arrow |
| Labels | Values / Totals / Legend | Toggles each on or off |
| Labels | More / Fewer decimals | Changes decimal places (automatic by default) |
| Labels | Prefix/suffix… | Adds text before or after every number |
| Labels | Switch font | Toggles IBM Plex Sans ↔ Calibri |

## Requirements and limitations

- PowerPoint 2016 or later (Microsoft 365), with macros allowed. *Disable all macros with notification* is enough: you approve ChartForge once.
- Developed and tested on **PowerPoint for Mac**. Windows should work but hasn't been tested yet.
- IBM Plex Sans must be installed for the default font; otherwise PowerPoint substitutes one. Use **Switch font** for Calibri.
- Charts have no value axis, by design (labels carry the values, think-cell style).
- Label fitting uses an approximate text width, so very long labels may need a wider chart.
- Each person who *edits* charts needs the add-in. Anyone can *view* them without it.

## Building from source

The add-in is built from three files in this repository:

```
src/ChartForge.bas        VBA module: chart layouts, datasheets, ribbon callbacks
ribbon/customUI.xml       The ChartForge ribbon tab
tools/build_addin.py      Packs a saved .pptm into a .ppam and adds the ribbon
```

See [docs/BUILD.md](docs/BUILD.md) for step-by-step instructions.

## How it works

Each chart is a PowerPoint group containing an invisible frame (which fixes the chart's bounds) plus the shapes that make up the chart. The datasheet and options are stored as tags on the group, which is why **Edit data** and **Update** can rebuild any chart without a separate file. Layout logic lives in `LayoutColumns`, `LayoutWaterfall`, `LayoutMekko` and `LayoutGantt` in `src/ChartForge.bas`.

## Licence

See [LICENSE](LICENSE).
