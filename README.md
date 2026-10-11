# Outcome Tracker PowerPoint Maker
Turns Outcome Tracker reports into PowerPoints: one for a single program, or one agency PowerPoint that combines several programs. Every table in a report becomes a slide with a chart and the numbers behind it.

It's a single web page, **Outcome Tracker PowerPoint Maker.html**. There's nothing to install, and report data never leaves your computer.

## Making a PowerPoint
1. In Outcome Tracker, open the report, run it for the dates you want, and export it. You get a .pdf and a .csv; you need the **.csv**.
2. Open **Outcome Tracker PowerPoint Maker.html** in Chrome or Edge. From SharePoint, download it first or open it from your synced OneDrive folder.
3. Drag the .csv onto the page. For an agency PowerPoint, drag in several reports and use the arrows to put them in order.
4. Fill in the title and reporting period.
5. Go through the slides:
   - untick any you don't want,
   - change a title,
   - switch between column chart, bar chart, pie chart, big number or table only,
   - type an **Explanation** to show on the slide.
6. Click **Download PowerPoint**. It goes to your Downloads folder, ready to edit further in PowerPoint (charts stay editable).

**Next time:** export the same report again and drop it in. Your titles, chart choices, explanations and unticked slides come back automatically, so the new PowerPoint matches the last one with up-to-date numbers. These choices are kept in your browser on your computer. Someone using another computer starts with the defaults.

## How charts are chosen
- Totals stay in the tables but are left out of the charts.
- Parts of a whole with 5 or fewer pieces get a pie chart. Everything else gets a bar or column chart.
- Charts with more than 15 bars show the largest 14 and group the rest into "All others". Tables with more than 40 rows show the largest 25. The slide always says when it does this.
- Tables that break a number down across many programs chart the total. The full breakdown follows on table slides.
- Tables with no data this period start unticked.

## For whoever maintains this
Everything is in the one HTML file:
- The `report-logic` script reads Outcome Tracker exports and chooses charts.
- The `app` script runs the page and builds the PowerPoint with [PptxGenJS](https://gitbrent.github.io/PptxGenJS/), which loads from the internet.

The CCETN logo and colors are in the `brand` script and at the top of the `app` script (`C` and `PALETTE`).

The comments at the top of `report-logic` describe the export format. If Outcome Tracker changes its exports, start there.
