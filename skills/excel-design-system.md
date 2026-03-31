# Excel Design System Skill

## When This Fires
Apply this design system whenever the user asks to:
- Format, style, or design anything in Excel
- "Make it look professional" / "clean it up" / "polish this"
- Create or redesign a dashboard, report, summary, or table
- Create charts, pivot tables, or any visual element

**ALWAYS read this skill before formatting. Never improvise colors, fonts, or layouts.**

---

## Tool Selection Decision Tree

| Scenario | Tool | Why |
|----------|------|-----|
| **Step 0: Any dashboard** | `excel_set_display_options` | Hide gridlines + headers FIRST. This is mandatory. |
| Dashboard card/section containers | `excel_add_shape` | Rounded rectangles with fill, shadow, text = card layout |
| Format 3+ cells/ranges at once | `excel_batch_format` | One call, all cell formatting |
| Format 1-2 individual cells | `excel_format_cell` | Simple, fast |
| Set column widths only | `excel_set_column_width` | Dedicated tool |
| Set row heights only | `excel_set_row_height` | Dedicated tool |
| Merge cells only | `excel_merge_cells` | Dedicated tool |
| Conditional formatting (color scales, data bars) | `excel_apply_conditional_format` | Requires rule logic |
| Complex automation with loops/logic | `excel_set_vba_code` + `excel_run_vba_macro` | Last resort only |

**Dashboard workflow order:**
1. `excel_set_display_options` — hide gridlines, hide headers, set zoom
2. `excel_add_shape` — create card containers, title bars, section backgrounds
3. `excel_batch_format` — format the data cells, headers, and values
4. `excel_create_chart` — add charts positioned inside card areas

**Default choice: `excel_batch_format`** for cell formatting, `excel_add_shape` for visual containers.

**NEVER use VBA macros for formatting.** VBA is fragile, requires trust settings, and can freeze Excel. The batch format tool + shape tool do everything VBA formatting can do, more reliably.

---

## Color Palette — 26 Named Tokens

### Primary Brand Colors
| Token | Hex | RGB | Usage |
|-------|-----|-----|-------|
| `navy-dark` | `#1E3247` | 30, 50, 71 | Title bars, primary headers, dashboard header background |
| `navy-mid` | `#283D52` | 40, 61, 82 | Section headers, sidebar backgrounds |
| `navy-light` | `#3A5068` | 58, 80, 104 | Sub-headers, secondary backgrounds |
| `slate-dark` | `#4A6274` | 74, 98, 116 | Tertiary headers, muted text on dark |
| `slate-light` | `#6B8A9E` | 107, 138, 158 | Disabled/secondary text |

### Accent Colors
| Token | Hex | RGB | Usage |
|-------|-----|-----|-------|
| `cyan-bright` | `#00BCD4` | 0, 188, 212 | Accent bars, divider lines, highlights |
| `cyan-light` | `#B2EBF2` | 178, 235, 242 | Light accent backgrounds, hover states |
| `teal` | `#009688` | 0, 150, 136 | Secondary accent, chart color 2 |
| `blue-link` | `#2196F3` | 33, 150, 243 | Hyperlinks, interactive elements |

### Semantic Colors — Data States
| Token | Hex | RGB | Usage |
|-------|-----|-----|-------|
| `input-yellow` | `#FFF9C4` | 255, 249, 196 | User input cells — editable fields |
| `input-border` | `#F9A825` | 249, 168, 37 | Border for input cells |
| `calc-blue` | `#E3F2FD` | 227, 242, 253 | Calculated/formula cells |
| `calc-border` | `#90CAF9` | 144, 202, 249 | Border for calculated cells |
| `total-green` | `#E8F5E9` | 232, 245, 233 | Total/summary rows |
| `total-dark` | `#2E7D32` | 46, 125, 50 | Total row text, green KPI |
| `warning-orange` | `#FFF3E0` | 255, 243, 224 | Warning/attention cells |
| `warning-text` | `#E65100` | 230, 81, 0 | Warning text color |
| `error-red` | `#FFEBEE` | 255, 235, 238 | Error/negative cells |
| `error-text` | `#C62828` | 198, 40, 40 | Error text, negative values |
| `positive-green` | `#1B5E20` | 27, 94, 32 | Positive values, good KPIs |

### Neutral Colors
| Token | Hex | RGB | Usage |
|-------|-----|-----|-------|
| `white` | `#FFFFFF` | 255, 255, 255 | Text on dark backgrounds, clean backgrounds |
| `off-white` | `#FAFAFA` | 250, 250, 250 | Alternating row background (even rows) |
| `light-gray` | `#F5F5F5` | 245, 245, 245 | Section backgrounds, card fills |
| `mid-gray` | `#E0E0E0` | 224, 224, 224 | Borders, divider lines |
| `dark-gray` | `#424242` | 66, 66, 66 | Primary body text |
| `black` | `#212121` | 33, 33, 33 | Emphasis text, header text on light backgrounds |

---

## Typography Scale

| Level | Size | Weight | Color | Usage |
|-------|------|--------|-------|-------|
| Dashboard Title | 14pt | Bold | `#FFFFFF` on `navy-dark` | Main page title, one per sheet |
| Section Header | 11pt | Bold | `#FFFFFF` or `cyan-bright` on `navy-mid` | Section dividers (e.g., "QUICK STATS", "TOP 5") |
| Column Header | 10pt | Bold | `#FFFFFF` on `navy-light` | Table column headers |
| Body Text | 10pt | Normal | `#424242` | Standard data cells |
| Body Emphasis | 10pt | Bold | `#212121` | Important values, row labels |
| Small Label | 9pt | Normal | `#6B8A9E` | Captions, footnotes, metadata |
| Tiny Note | 8pt | Normal | `#6B8A9E` | Timestamps, version info |

**Font family: Always `Segoe UI`** — it's clean, modern, and available on all Windows machines. Fallback: `Calibri`.

---

## Number Formats

| Type | Format String | Example |
|------|---------------|---------|
| Currency | `$#,##0` | $12,500 |
| Currency (cents) | `$#,##0.00` | $12,500.00 |
| Integer | `#,##0` | 12,500 |
| Decimal | `#,##0.00` | 12,500.00 |
| Percentage | `0%` | 85% |
| Percentage (decimal) | `0.0%` | 85.3% |
| Date | `MM/DD/YYYY` | 03/12/2026 |
| Date (short) | `M/D/YY` | 3/12/26 |
| Time | `h:mm AM/PM` | 2:30 PM |
| Ratio | `0.00` | 8.15 |
| Score | `0.0` | 7.3 |
| Phone | `(###) ###-####` | (612) 555-1234 |
| Count | `#,##0` | 1,250 |

---

## Layout Patterns

### Pattern 1: Title Bar (Row 1)
```
Range: A1:{last_column}1
- Merge entire row across used columns
- Fill: navy-dark (#1E3247)
- Font: 14pt Segoe UI Bold, white (#FFFFFF)
- Alignment: center horizontal, center vertical
- Row height: 40
```

### Pattern 2: Accent Divider (Row 2)
```
Range: A2:{last_column}2
- Fill: cyan-bright (#00BCD4)
- Row height: 4 (thin accent line)
- No text
```

### Pattern 3: Section Header
```
Range: {start}:{end} (e.g., A3:C3)
- Merge range
- Fill: navy-mid (#283D52)
- Font: 11pt Segoe UI Bold, white (#FFFFFF)
- Text: Use bullet prefix → Chr(9679) & "  SECTION NAME" or just "  SECTION NAME"
- Alignment: left horizontal, center vertical
- Row height: 28
```

### Pattern 4: Column Headers (Table Head)
```
Range: full header row (e.g., A4:H4)
- Fill: navy-light (#3A5068)
- Font: 10pt Segoe UI Bold, white (#FFFFFF)
- Alignment: center horizontal, center vertical
- Row height: 24
- Bottom border: thin, cyan-bright (#00BCD4)
```

### Pattern 5: Data Grid (Table Body)
```
- Font: 10pt Segoe UI Normal, dark-gray (#424242)
- Alternating rows: white (#FFFFFF) / off-white (#FAFAFA)
- Row height: 20
- Borders: thin, mid-gray (#E0E0E0) on all edges
- Numbers: right-aligned with appropriate number format
- Text: left-aligned
- Alignment: center vertical
```

### Pattern 6: Total Row
```
- Fill: total-green (#E8F5E9)
- Font: 10pt Segoe UI Bold, total-dark (#2E7D32)
- Top border: medium, navy-dark (#1E3247)
- Bottom border: medium, navy-dark (#1E3247)
```

### Pattern 7: Dashboard Card
```
- Card area: light-gray (#F5F5F5) fill
- Card header: navy-mid (#283D52) fill, white text, 10pt bold
- Card body: white (#FFFFFF) fill
- Border: thin, mid-gray (#E0E0E0)
- 1-column gap between cards (leave empty)
```

### Pattern 8: KPI Metric Display
```
- Large number: 14pt Segoe UI Bold
  - Positive: positive-green (#1B5E20)
  - Negative: error-text (#C62828)
  - Neutral: navy-dark (#1E3247)
- Label below: 9pt Segoe UI Normal, slate-light (#6B8A9E)
```

### Pattern 9: Year/Month Banner
```
- Fill: navy-dark (#1E3247)
- Font: 11pt Segoe UI Bold, cyan-bright (#00BCD4)
- Merge across relevant columns
- Alignment: center
- Row height: 26
```

---

## Column Width Guidelines

| Content Type | Width | Example |
|-------------|-------|---------|
| Narrow label (rank, #) | 5-6 | "#", "Rank" |
| Short text (state, code) | 8-10 | "MN", "Yes/No" |
| Name/title | 20-25 | "School Name", "Category" |
| Description | 30-40 | Long text fields |
| Currency | 12-14 | "$12,500.00" |
| Percentage | 8-10 | "85.3%" |
| Date | 12 | "03/12/2026" |
| Score/rating | 8-10 | "8.15" |
| Count/integer | 8-10 | "1,250" |
| Spacer column | 2-3 | Visual gap between sections |

**Always set column widths explicitly.** Never leave default widths — they look unfinished.

---

## Chart Design Rules

### General
- **No 3D effects.** Ever. Flat only.
- **No gradients** in chart fills.
- **No chart border** — remove the outline.
- **Background: white** (#FFFFFF) or transparent.
- **Title: 11pt Segoe UI Bold**, dark-gray (#424242), left-aligned or centered.
- **Legend: 9pt Segoe UI**, positioned at bottom or right. Remove if only one series.
- **Gridlines: light-gray** (#E0E0E0), thin. Remove vertical gridlines for bar charts.
- **Axis labels: 9pt Segoe UI**, dark-gray (#424242).

### Bar/Column Charts
- Primary color: navy-dark (#1E3247)
- Secondary: cyan-bright (#00BCD4)
- Tertiary: teal (#009688)
- Gap width: 80-100%
- Sort bars by value (largest to smallest) unless categorical order matters

### Pie/Donut Charts
- Use max 6-8 slices. Group remaining as "Other".
- Color sequence: navy-dark, cyan-bright, teal, slate-dark, warning-text, positive-green, blue-link, slate-light
- Show percentage labels, not values
- No exploded slices

### Line/Sparkline Charts
- Line weight: 2pt
- Primary: navy-dark (#1E3247)
- Secondary: cyan-bright (#00BCD4)
- Markers: circle, 4pt, same color as line
- Area fill (if used): 20% opacity of line color

---

## Batch Format Usage Examples

### Example: Format a title bar + accent line + section header
```json
{
  "filePath": "C:\\path\\to\\file.xlsx",
  "sheetName": "Dashboard",
  "operations": [
    {
      "range": "A1:I1",
      "unmerge": true,
      "merge": true,
      "value": "DASHBOARD TITLE",
      "fillColor": "#1E3247",
      "fontName": "Segoe UI",
      "fontSize": 14,
      "fontBold": true,
      "fontColor": "#FFFFFF",
      "horizontalAlignment": "center",
      "verticalAlignment": "center",
      "rowHeight": 40
    },
    {
      "range": "A2:I2",
      "fillColor": "#00BCD4",
      "rowHeight": 4
    },
    {
      "range": "A3:C3",
      "unmerge": true,
      "merge": true,
      "value": "  SECTION NAME",
      "fillColor": "#283D52",
      "fontName": "Segoe UI",
      "fontSize": 11,
      "fontBold": true,
      "fontColor": "#FFFFFF",
      "horizontalAlignment": "left",
      "verticalAlignment": "center",
      "rowHeight": 28
    }
  ]
}
```

### Example: Format alternating data rows
```json
{
  "operations": [
    { "range": "A5:H5", "fillColor": "#FFFFFF", "fontName": "Segoe UI", "fontSize": 10, "fontColor": "#424242", "borderStyle": "thin", "borderColor": "#E0E0E0", "rowHeight": 20 },
    { "range": "A6:H6", "fillColor": "#FAFAFA", "fontName": "Segoe UI", "fontSize": 10, "fontColor": "#424242", "borderStyle": "thin", "borderColor": "#E0E0E0", "rowHeight": 20 },
    { "range": "A7:H7", "fillColor": "#FFFFFF", "fontName": "Segoe UI", "fontSize": 10, "fontColor": "#424242", "borderStyle": "thin", "borderColor": "#E0E0E0", "rowHeight": 20 }
  ]
}
```

---

## Pre-Flight Checklist

Before formatting any sheet, always:
1. **Read the sheet first** (`excel_read_sheet`) to understand the data layout
2. **Identify the used range** — don't format beyond the data
3. **Plan the layout** — title row, accent divider, section headers, data areas
4. **Set column widths first** — this defines the visual structure
5. **Apply formatting top-to-bottom** — title → headers → body → totals
6. **Use `unmerge: true`** on any range that might have been previously merged
7. **Save automatically** — `excel_batch_format` saves after all operations

---

## Dashboard Canvas Setup (ALWAYS DO FIRST)

Before any dashboard formatting, set up the canvas:
```json
{
  "filePath": "...",
  "showGridlines": false,
  "showRowColumnHeaders": false,
  "zoomLevel": 90
}
```
This transforms the spreadsheet into a clean white canvas — the foundation for everything else.

---

## Shape-Based Card Layouts (Cottrell Method)

The technique that makes Excel dashboards look like designed applications: use rounded rectangle shapes as visual containers over the cell grid.

### Card Container Pattern
```json
{
  "shapeType": "roundedRectangle",
  "left": 15, "top": 60, "width": 300, "height": 200,
  "name": "CardQuickStats",
  "fill": { "color": "#FFFFFF" },
  "line": { "visible": false },
  "shadow": {
    "visible": true,
    "color": "#000000",
    "offsetX": 2, "offsetY": 2,
    "blur": 8, "transparency": 0.8
  }
}
```

### Dark Header Card
```json
{
  "shapeType": "roundedRectangle",
  "left": 15, "top": 60, "width": 300, "height": 35,
  "name": "HeaderQuickStats",
  "fill": { "color": "#1E3247" },
  "line": { "visible": false },
  "shadow": { "visible": false },
  "text": {
    "value": "QUICK STATS",
    "fontName": "Segoe UI",
    "fontSize": 11,
    "fontBold": true,
    "fontColor": "#FFFFFF",
    "horizontalAlignment": "left",
    "verticalAlignment": "middle"
  }
}
```

### Gradient Card (Modern Look)
```json
{
  "fill": {
    "gradient": {
      "color1": "#1E3247",
      "color2": "#3A5068",
      "direction": "horizontal"
    }
  }
}
```

### KPI Metric Card
```json
{
  "shapeType": "roundedRectangle",
  "left": 15, "top": 100, "width": 140, "height": 80,
  "fill": { "color": "#F5F5F5" },
  "line": { "color": "#E0E0E0", "weight": 1 },
  "shadow": { "visible": true, "blur": 6, "transparency": 0.85 },
  "text": {
    "value": "16\nTotal Schools",
    "fontName": "Segoe UI",
    "fontSize": 20,
    "fontBold": true,
    "fontColor": "#1E3247",
    "horizontalAlignment": "center",
    "verticalAlignment": "middle"
  }
}
```

### Card Layout Grid
Position cards on a grid for alignment (AAE: Always Align Everything):
- Left margin: 15 points
- Card spacing: 10 points between cards
- Standard card widths: 140pt (small KPI), 300pt (medium), 620pt (full width)
- Standard heights: 35pt (header only), 80pt (KPI metric), 200pt (data section), 300pt (chart area)

### Dashboard Composition Order
1. **Canvas**: `excel_set_display_options` — hide gridlines, headers
2. **Background**: Full-width rectangle shape for page background (optional, use light gray #F5F5F5)
3. **Title bar**: Full-width dark shape with white title text
4. **Accent line**: Thin full-width shape in cyan (#00BCD4), height 4pt
5. **Section cards**: White rounded rectangles with subtle shadows
6. **Card headers**: Dark navy shapes overlapping top of each card
7. **Data**: `excel_batch_format` for cell data inside card areas
8. **Charts**: Position charts inside card boundaries
9. **KPI cards**: Small colored shapes with big numbers

---

## Anti-Patterns — NEVER Do These

- **Never leave gridlines visible on a dashboard** — use `excel_set_display_options` to hide them FIRST
- Never use default Excel blue (#4472C4) — use the palette above
- Never use Comic Sans, Times New Roman, or Arial — use Segoe UI
- Never leave default column widths
- Never use ALL CAPS in body text (headers only)
- Never use more than 3 font sizes on one sheet
- Never put data in row 1 — row 1 is always the title bar
- Never skip the accent divider (row 2)
- Never use borders without consistent color (#E0E0E0 for data, #1E3247 for totals)
- Never use VBA macros for formatting — use `excel_batch_format` + `excel_add_shape`
- Never improvise colors — every color must come from the palette above
- **Never use 3D effects, gradients on charts, or exploded pie slices**
- **Never create a dashboard without shape-based card containers** — cells alone look like a spreadsheet
- Never place charts without aligning them to a grid — use consistent positioning
