# What can I actually do with this?

Five prompts that exercise capabilities the native "Claude for Excel" add-in does not have. (For one-workbook tasks like "explain this formula" or "add a pivot to this sheet I'm staring at," the native add-in is the better tool. Use both.)

## 1. Reconcile two workbooks

> Reconcile `Q1-forecast.xlsx` against `Q1-actuals.xlsx` in my Downloads folder. New tab in actuals showing variance per line, red-fill anything off by more than 10%.

What you get: a populated variance tab in the actuals file, conditionally formatted, saved to disk. No download or re-upload.

Why the native add-in can't: it operates inside one open workbook only. It cannot open a second file from your filesystem.

## 2. Multi-file rollup

> Loop every `.xlsx` in `~/customers/2026/` and pull the 'MRR' cell from the Summary tab into one consolidated sheet with filename plus value. Save to `mrr-rollup.xlsx`.

What you get: one workbook, one row per customer file, done in seconds.

Why the native add-in can't: no filesystem access. It cannot walk a folder.

## 3. CSV consolidation with formatting

> Take all the CSVs in `~/Downloads/may-exports/` and combine them into one .xlsx with a tab per file. Format each tab as a proper table: freeze headers, currency-format columns C-E, Segoe UI 10pt body. Save next to the originals.

What you get: a clean multi-tab .xlsx assembled from the source CSVs, written to your folder.

Why the native add-in can't: same reason — no filesystem reach beyond the open workbook.

## 4. VBA macro authoring (Windows + VBA Trust enabled)

> Read the macro in `pricing.xlsm`, explain what it does in plain English, then write a new macro that does the same thing but logs each run to a 'RunLog' sheet with a timestamp.

What you get: a plain-English explanation of the existing macro plus a new VBA module installed in the file with the logging behavior.

Why the native add-in can't: the native Claude for Excel add-in supports `.xlsm` files but cannot modify or author VBA code. It analyzes the spreadsheet structure and formulas only.

## 5. Project-context-driven dashboard

> In my Soracom-Reports project, build a Q1 sales dashboard from `q1-data.xlsx` using my project's design rules and the layout in `report-template.xlsx`. Save as `q1-dashboard.xlsx`.

What you get: a dashboard styled per your project's CLAUDE.md instructions (colors, fonts, naming conventions), built off your project's template file, saved to disk.

Why the native add-in can't: the native add-in has its own separate chat history per Excel session. It does not inherit your Claude Desktop project's CLAUDE.md instructions, knowledge-base files, or persistent memory. This server runs inside the project, so every Excel call sees the full project context.

## Other things this server does well

- Find and replace across an entire workbook with regex and dry-run preview.
- Pagination on huge sheets so you can page through hundreds of thousands of rows without blowing past context limits.
- Sandboxed file access — the server refuses to read or write outside the folders you allow, so you can hand it to coworkers without worrying about accidental access to system files.
- Capability probe (`excel_check_environment`) tells you what works on your machine before you waste a turn on a tool that needs something not installed (Excel running, VBA Trust enabled, Mac Automation permission, etc.).

## Try this prompt template when you're stuck

> Run `excel_check_environment` first. Then, given what's available, [your goal]. If anything is missing or unsupported, tell me what would need to be installed or enabled.

Pinning the capability check at the start of a session means Claude won't waste a turn discovering mid-task that VBA trust isn't enabled.
