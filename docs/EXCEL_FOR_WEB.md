# Using this with Excel for the Web (no desktop Excel)

Some users have Microsoft 365 with **Excel for the Web only** — no desktop Excel installed. Here's what works and what doesn't.

## What works (≈80% of the value)

All **file-mode** tools work without desktop Excel because the server edits the `.xlsx` file directly using a Node library. That covers:

- Read sheets, ranges, cells, formulas
- Write cells, ranges, rows, formulas
- Sheet management (create, delete, rename, duplicate)
- Cell formatting (fonts, colors, borders, alignment, column widths)
- Chart creation (basic — placeholders for now)
- Pivot tables (computed values)
- Tables, named ranges (file-mode, when AppleScript port lands)
- **Find and replace** across whole workbook
- **CSV import/export**
- **Image insertion**
- Pagination for huge sheets
- All the validation, search, filter, and analysis tools

## What doesn't work

Anything labeled **live mode** in the docs — tools that talk to a running Excel application:

- Live cell editing while you watch
- Live native charts (the kind Excel renders, not the file-mode placeholder)
- VBA macros (Windows desktop Excel only anyway)
- Real-time formatting (`excel_batch_format`)
- Screenshot capture
- Power Query refresh

## How to use it productively

1. Pick the file-mode prompts from `EXAMPLES.md` (1, 2, 4 — reconcile, multi-file rollup, CSV cleanup).
2. Avoid prompts that need a chart to appear in real time (#3) or VBA editing (#5).
3. Open the resulting `.xlsx` in Excel for the Web after the server finishes editing it.

## If you need full live mode

You need a desktop Excel install (Excel 2021+ on Mac, Excel 2019+ on Windows). Talk to your IT department about adding Excel to your install — it's part of any Microsoft 365 plan that includes desktop apps.
