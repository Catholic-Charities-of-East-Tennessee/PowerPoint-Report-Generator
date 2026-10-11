# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

`Outcome Tracker PowerPoint Maker.html` is a single, self-contained web page for Catholic Charities of East Tennessee. Staff drop Outcome Tracker (VistaShare) report exports (.csv) on it, choose and adjust slides, and download a PowerPoint. The COO wanted one deck per program (CES and Samaritan Place first) plus agency decks combining programs. He also wanted to add explanations to the charts and to refresh the decks easily. The decks are kept on SharePoint, and there's no server.

The users are non-technical, so these constraints are deliberate:
- No install and no build step: one HTML file that opens from SharePoint or OneDrive.
- No upload: report data never leaves the browser.
- No stored Outcome Tracker credentials. An earlier Python/Selenium version that auto-downloaded reports was dropped, because the login uses reCAPTCHA and 2FA and kept breaking.

## Structure of the HTML file

- `<script id="report-logic">` defines `ReportLogic`: pure functions with no DOM access.
  - `parseCsv` and `interpret` turn an export into sections and tables. The export format is documented in the comment at the top of this script.
  - `planChart` picks the categories and series to chart. It leaves out Total rows and columns, transposes single-row tables, charts the Total column when there are more than 6 columns, and detects fractions as percentages.
  - `availableKinds` and `autoKind` decide which chart types a slide offers and which it starts with.
  - `chartData` produces the data for one chosen kind (pies drop zero slices; long bar lists become top 14 + "All others").
  - `tableCells` and `paginate` lay out the raw-number tables.
- `<script id="app">` handles everything else: loading files (UTF-8 with a windows-1252 fallback), slide cards with SVG previews, remembering edits, and building the deck with PptxGenJS (`LAYOUT_WIDE`, native editable charts).
  - Edits are saved to `localStorage` (key `ccetn-outcomes-pptx-maker:v1`).
  - A slide's saved edits are keyed by csv file name, section index and title, and table index and label. Renaming a report's export file, or Outcome Tracker reordering its tables, loses those edits.
- `<script id="brand">` holds the CCETN mountain logo from ccetn.org as an embedded PNG (`CCETN_LOGO`). The logo has a transparent background and black lettering, so it only goes on light backgrounds; on the navy divider slides it sits on a white circle. The brand colors (navy `#1D4A77`, `#455D9E`, the logo blues `#7B9DD1`/`#A3C2E3`, CC yellow `#FFCB05` and blue `#0088CB`) come from the ccetn.org theme. They are set in the CSS `:root` tokens and in `C`/`PALETTE` in the app script.
- PptxGenJS 3.12.0 is the only external dependency, loaded from jsDelivr with a pinned version.

## Checking changes

There is no automated test suite (the owner chose not to keep one). To check a change, open the page and drop in the sample exports in `app/exported reports from outcometracker/`. That folder is gitignored, like all `.csv`, `.pdf` and `.pptx` files, because the exports hold client data. Headless Chrome via Selenium (installed in `.venv`) can drive the page and call `ReportLogic` through `execute_script`.

To check slide layout visually, render a downloaded deck through PowerPoint COM (`PowerPoint.Application` → `Presentations.Open` → `Slide.Export(path, "PNG", 1333, 750)`). PowerPoint refuses paths over 255 characters, so copy the file to a short path first.

## Leftovers

`app/BOT_DATA.py` (gitignored) holds Outcome Tracker credentials for the retired downloader. Nothing uses it. Never print or commit it.
