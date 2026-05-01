# What can I actually do with this?

Five prompts that beat both **Claude.ai-native Excel handling** (file-attached chat) and **Microsoft Copilot in Excel**.

## 1. Reconcile two workbooks

> Reconcile `Q1-forecast.xlsx` against `Q1-actuals.xlsx` in my Downloads folder. New tab in actuals showing variance % per line, red-fill anything off by more than 10%.

**You get:** A populated variance tab in the actuals file, conditionally formatted, saved to disk. No download/re-upload.

**Why Claude.ai-native can't:** It can't reach two files on your filesystem. You'd upload both, get text back, manually paste into Excel.

**Why Microsoft Copilot can't:** Copilot operates inside one open workbook. It can't open a sibling file from your disk and write into it.

## 2. Multi-file rollup

> Loop every `.xlsx` in `~/customers/2026/` and pull the 'MRR' cell from the Summary tab into one consolidated sheet with filename + value, then save to `mrr-rollup.xlsx`.

**You get:** One workbook, one row per customer, done in seconds.

**Why competitors can't:** Both are single-file by design. No filesystem walks, no batch reads.

## 3. Live edit while you watch

> I'm staring at this sheet in Excel right now. Add a column chart of column D vs F on the active selection, title it "Revenue vs Spend", and put it next to the data.

**You get:** The chart appears in the open workbook while you watch. (Windows: native COM chart; Mac: file-mode write — close + reopen to see, or use the styling on a chart you already inserted.)

**Why competitors can't:** Claude.ai-native has no live Excel link — output is a download. Copilot can chart, but only via its own UI flow, not from natural language with surrounding context (other tabs, other files).

## 4. CSV cleanup with formatting

> Take this CSV export at `~/Downloads/customers-raw.csv`, format it as a proper table (freeze headers, currency-format columns C-E, Segoe UI 10pt body), and save as `customers-clean.xlsx` next to the original.

**You get:** Cleaned `.xlsx` on disk, formatting applied, in the same folder as the source.

**Why Claude.ai-native can't:** Returns a downloadable file but no in-place save and loses your folder structure.

**Why Copilot can't:** Requires manual import into Excel first; can't write to disk on its own.

## 5. VBA macro authoring (Windows + VBA Trust enabled)

> Read the macro in `pricing.xlsm`, explain what it does in plain English, then write a new macro that does the same thing but logs each run to a 'RunLog' sheet with timestamp.

**You get:** Plain-English explanation of the existing macro + a new VBA module installed in the file with the logging behavior.

**Why Claude.ai-native can't:** Can read pasted VBA text, but can't install new code into a `.xlsm` file.

**Why Copilot can't:** Microsoft Copilot doesn't author or modify VBA — explicit policy.

## Other things this server does well that nothing else does

- **Find and replace across an entire workbook** with regex + dry-run preview.
- **Pagination on huge sheets** so you can ask Claude to "page through the next 500 rows" without hitting context limits.
- **Sandbox enforcement** — server refuses to read or write outside the folders you allow, so you can hand it to coworkers without worrying about accidental access to system files.
- **Friendly capability probe** — `excel_check_environment` tells you what works on your machine before you waste time on a tool that needs something not installed.

## Try this prompt template when you're stuck

> Run `excel_check_environment` first. Then, given what's available, [your goal]. If anything is missing or unsupported, tell me what would need to be installed/enabled.

Pinning the capability check at the start of a session means Claude won't waste a turn discovering "oh, VBA trust isn't enabled" mid-task.
