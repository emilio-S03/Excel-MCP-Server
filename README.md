# Excel MCP Server (v3.x — Soracom internal fork)

[![TypeScript](https://img.shields.io/badge/TypeScript-5.0+-3178C6?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![MCP](https://img.shields.io/badge/MCP-Server-green)](https://modelcontextprotocol.io/)
[![ExcelJS](https://img.shields.io/badge/ExcelJS-4.4+-217346)](https://github.com/exceljs/exceljs)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)
[![Node](https://img.shields.io/badge/Node.js-18+-339933?logo=node.js&logoColor=white)](https://nodejs.org/)

An MCP server that gives Claude Desktop (and Claude Code) the ability to operate Excel files **across folders, across workbooks, and inside your Claude project's full context** — for the work the in-Excel "Claude for Excel" add-in cannot do.

> **Attribution & fork history**
>
> This is a fork and substantial extension of the original Excel MCP Server by **[sbraind](https://github.com/sbraind/excel-mcp-server)** (MIT licensed via package.json + README badge + "Fork the repository" invitation in their Contributing section). The original server provided the foundation: ExcelJS file mode, the live-editing dispatcher, the PowerShell COM bridge, and many of the chart/pivot/VBA tools.
>
> This v3.x fork (maintained by [@emilio-S03](https://github.com/emilio-S03) for Soracom internal use) adds sandboxed file access, env-var-based config that actually works in both Claude Desktop and Claude Code, a unified test/CI harness, friendly Mac platform-error messages, ~50 additional tools, and full coworker onboarding docs. See [LICENSE](LICENSE) for the joint copyright.

---

## When to use this vs. the native "Claude for Excel" add-in

Anthropic ships a "Claude for Excel" add-in that lives inside the Excel sidebar. It's genuinely capable: it has Skills, Connectors, conditional formatting, pivot tables, charts, scenario testing, formula auditing, cross-tab navigation with citations, and shared context across Excel and PowerPoint. **For most one-workbook tasks, the native add-in is the right tool.**

This MCP server is for the four things the native add-in cannot do:

| Task | Native "Claude for Excel" add-in | This MCP server |
|---|---|---|
| Editing or analyzing the **single workbook open in Excel** in front of you | Best tool. Use it. | Works, but no advantage over native. |
| **Multi-file or multi-workbook** operations (loop a folder, reconcile two files, batch edits across many .xlsx) | Cannot reach beyond the open workbook | Direct filesystem access |
| **Authoring or modifying VBA macros** in a `.xlsm` file (Windows + VBA Trust required) | Cannot modify VBA — analyzes only | `excel_run_vba_macro`, `excel_set_vba_code`, `excel_get_vba_code` |
| Operating with your **Claude project's full context** (CLAUDE.md instructions, knowledge-base files, persistent memory, other MCP servers you've installed) | Has its own separate chat history; doesn't inherit Claude Desktop project context | Runs inside your Claude project — every Excel action sees your project's instructions, KBs, and memory |

If your Excel work is "open this file, help me debug formulas" — use the native add-in.
If your Excel work is "loop these 14 customer files and consolidate the MRR column into a report formatted to my project's design rules" — use this server.

Most Soracom power users will install both and learn when to reach for which.

---

## What's in v3.2.0

- **96 tools** covering reading, writing, formatting, sheet management, charts, pivots, tables, VBA, sparklines, find/replace, CSV import/export, image insertion, sheet/row/column visibility, hyperlinks, conditional formatting, data validation, named ranges, calculation modes, page setup, PDF export, formula auditing, and workbook stats. Run `excel_check_environment` after install to see what works on your machine.
- **Sandboxed file access** — defaults to your `Documents`, `Desktop`, `Downloads` folders. Configurable via the extension settings UI in Claude Desktop or the `EXCEL_ALLOWED_DIRS` env var.
- **Capability probe** (`excel_check_environment`) — structured report on Excel install state, VBA Trust state (Windows), Automation permission state (Mac), and which tool categories are usable on this machine right now.
- **Cross-platform** — most tools work on Windows, macOS, and Linux. The few Windows-only tools (live VBA, native COM chart creation) throw friendly platform errors on Mac with alternatives.
- **Live editing** when Excel is open (Windows COM, macOS AppleScript) — Claude's changes appear in real time.

**Built for:**
- Multi-workbook data analysis and report consolidation
- Cross-file ETL — loop folders of customer/product/report .xlsx files
- Template-driven report generation with project-scoped style rules
- Workbook auditing — formula errors, circular refs, complexity hotspots
- VBA macro authoring and installation (Windows)
- Anything else that requires reaching outside a single open workbook

## Installation

### 🚀 Quick Installation (Recommended) - One Click!

The easiest way to install this server is using the pre-built MCPB bundle:

1. **Download** the latest `excel-mcp-server-3.2.0.mcpb` file from the [releases page](https://github.com/emilio-S03/Excel-MCP-Server/releases)
2. **Double-click** the `.mcpb` file, or:
   - Open Claude Desktop
   - Go to **Settings** → **Extensions** → **Advanced Settings**
   - Click **"Install Extension..."**
   - Select the downloaded `.mcpb` file
3. **Restart** Claude Desktop
4. **Done!** No Node.js installation, no config files to edit

> **Note:** One-click installation works on Claude Desktop for macOS and Windows. All dependencies are bundled - no additional setup required!

For more details, see [BUNDLE.md](BUNDLE.md).

---

### 🛠️ Manual Installation (Advanced)

If you prefer to build from source:

#### Step 1: Clone and build the project

```bash
git clone https://github.com/emilio-S03/Excel-MCP-Server.git
cd Excel-MCP-Server
npm install
npm run build
```

#### Step 2: Configure Claude Desktop

Add this configuration to your Claude Desktop config file:

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": ["${__dirname}/dist/index.js"]
    }
  }
}
```

**Note**: When using the MCPB bundle or manual installation, use `${__dirname}` which automatically resolves to the server's directory. For manual installations without MCPB, you can also use absolute paths like `/path/to/excel-mcp-server/dist/index.js`.

#### Step 3: Restart Claude Desktop

Close and reopen Claude Desktop completely.

#### Step 4: Verify

The server should now be available in Claude. Try:
```
Create a new Excel file at ~/Documents/test.xlsx with a sheet called "Sales" containing sample data
```

For detailed installation instructions and troubleshooting, see [INSTALLATION.md](INSTALLATION.md).

---

## Configuration Options

The server supports several configuration options that can be set through Claude Desktop's MCP configuration:

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": ["${__dirname}/dist/index.js"],
      "config": {
        "createBackupByDefault": false,
        "defaultResponseFormat": "json",
        "allowedDirectories": []
      }
    }
  }
}
```

### Available Options:

- **`createBackupByDefault`** (boolean, default: `false`)
  Automatically create backup files (`.backup` extension) before modifying Excel files. When enabled, every destructive operation will create a backup unless explicitly disabled in the tool call.

- **`defaultResponseFormat`** (string: `"json"` or `"markdown"`, default: `"json"`)
  Default format for tool responses. Can be overridden per tool call with the `responseFormat` parameter.

- **`allowedDirectories`** (array of strings, default: `[]`)
  List of directories where the server is allowed to read/write Excel files. When empty, all directories are accessible. Use this to restrict file access for security:
  ```json
  "allowedDirectories": [
    "~/Documents/Excel",
    "~/Projects/data"
  ]
  ```
  The server will reject any file operations outside these directories.

### Input Validation

All tool inputs are validated using Zod schemas. Invalid parameters will return clear error messages indicating what's wrong:

- Cell addresses must match format `A1`, `B2`, etc.
- Ranges must match format `A1:D10`
- File paths are checked against `allowedDirectories` if configured
- Missing required parameters are reported immediately

---

## ✨ Real-Time Live Editing (macOS)

The Excel MCP Server features **automatic real-time editing** for Excel files that are already open in Microsoft Excel on macOS. When a file is open, changes are applied instantly and become visible immediately—no need to close and reopen the file!

### How It Works

The server automatically detects when:
1. **Microsoft Excel is running** on your Mac
2. **The target file is open** in Excel

When both conditions are met, the server uses **AppleScript** to modify the open Excel file directly. Otherwise, it falls back to file-based editing using ExcelJS.

### Supported Operations (16 Live-Editing Tools)

The following tools support real-time editing when files are open in Excel:

**Writing:**
- `excel_update_cell` - Update cell values instantly
- `excel_add_row` - Add rows and see them appear immediately
- `excel_write_range` - Write data ranges in real-time
- `excel_set_formula` - Set formulas that calculate instantly

**Formatting:**
- `excel_format_cell` - Apply formatting (fonts, colors, borders) live
- `excel_set_column_width` - Adjust column widths instantly
- `excel_set_row_height` - Adjust row heights instantly
- `excel_merge_cells` - Merge cells in real-time

**Sheet Management:**
- `excel_create_sheet` - Create new sheets that appear immediately
- `excel_delete_sheet` - Delete sheets with instant feedback
- `excel_rename_sheet` - Rename sheets in real-time

**Row/Column Operations:**
- `excel_delete_rows` - Delete rows and see them disappear instantly
- `excel_delete_columns` - Delete columns in real-time
- `excel_insert_rows` - Insert rows that appear immediately
- `excel_insert_columns` - Insert columns instantly

**Advanced:**
- `excel_unmerge_cells` - Unmerge cells in real-time

### Response Indicators

Tool responses include a `method` field indicating which approach was used:

```json
{
  "success": true,
  "message": "Cell A1 updated (via Excel)",
  "method": "applescript",
  "note": "Changes are visible immediately in Excel"
}
```

vs.

```json
{
  "success": true,
  "message": "Cell A1 updated",
  "method": "exceljs",
  "note": "File updated. Open in Excel to see changes."
}
```

### Requirements

- **Platform**: macOS only (AppleScript is a macOS technology)
- **Application**: Microsoft Excel for Mac must be installed
- **File State**: Target Excel file must be open in Excel

### Benefits

- **Instant Feedback**: See changes as they happen—perfect for interactive workflows
- **No File Conflicts**: Works directly with the open file without save/reload cycles
- **Seamless Experience**: Automatically falls back to file-based editing when Excel isn't available

### Note

Read-only operations (like `excel_read_sheet`, `excel_read_range`, etc.) don't require real-time editing as they don't modify files. Complex operations like pivot tables, charts, and conditional formatting use file-based editing for reliability.

---

## Quick Start

Once installed, you can start using the server immediately in Claude Desktop. Here are some example prompts:

```
Create a new Excel file with sales data for Q1 2024
```

```
Read the data from Sheet1 in ~/Documents/report.xlsx
```

```
Apply bold formatting and blue background to the header row in my sales spreadsheet
```

```
Create a pivot table showing total sales by product and month
```

```
Generate a column chart from the data in range A1:B10
```

For more examples and detailed use cases, see [FEATURE_SUMMARY.md](FEATURE_SUMMARY.md).

---

## Available Tools

### 📖 Reading (5 tools)

#### 1. `excel_read_workbook`
List all sheets and metadata of an Excel workbook.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "responseFormat": "json"
}
```

#### 2. `excel_read_sheet`
Read complete data from a sheet with optional range.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "range": "A1:D10",
  "responseFormat": "markdown"
}
```

#### 3. `excel_read_range`
Read a specific range of cells.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "range": "B2:E20",
  "responseFormat": "json"
}
```

#### 4. `excel_get_cell`
Read value from a specific cell.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "cellAddress": "A1",
  "responseFormat": "json"
}
```

#### 5. `excel_get_formula`
Read the formula from a specific cell.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "cellAddress": "D5",
  "responseFormat": "json"
}
```

### ✏️ Writing (5 tools)

#### 6. `excel_write_workbook`
Create a new Excel file with data.

**Example:**
```json
{
  "filePath": "./output.xlsx",
  "sheetName": "MyData",
  "data": [
    ["Name", "Age", "City"],
    ["Alice", 30, "New York"],
    ["Bob", 25, "Los Angeles"]
  ],
  "createBackup": false
}
```

#### 7. `excel_update_cell`
Update value of a specific cell.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "cellAddress": "B2",
  "value": 1500,
  "createBackup": true
}
```

#### 8. `excel_write_range`
Write multiple cells simultaneously.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "range": "A1:C2",
  "data": [
    ["Header1", "Header2", "Header3"],
    [100, 200, 300]
  ],
  "createBackup": false
}
```

#### 9. `excel_add_row`
Add a row at the end of the sheet.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "data": ["Product X", 150, "2024-01-15"],
  "createBackup": false
}
```

#### 10. `excel_set_formula`
Set or modify a formula in a cell.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "cellAddress": "D2",
  "formula": "SUM(B2:C2)",
  "createBackup": false
}
```

### 🎨 Formatting (4 tools)

#### 11. `excel_format_cell`
Change cell formatting (color, font, borders, alignment).

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "cellAddress": "A1",
  "format": {
    "font": {
      "bold": true,
      "size": 14,
      "color": "FF0000"
    },
    "fill": {
      "type": "pattern",
      "pattern": "solid",
      "fgColor": "FFFF00"
    },
    "alignment": {
      "horizontal": "center",
      "vertical": "middle"
    }
  },
  "createBackup": false
}
```

#### 12. `excel_set_column_width`
Adjust width of a column.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "column": "A",
  "width": 20,
  "createBackup": false
}
```

#### 13. `excel_set_row_height`
Adjust height of a row.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "row": 1,
  "height": 30,
  "createBackup": false
}
```

#### 14. `excel_merge_cells`
Merge cells in a range.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "range": "A1:D1",
  "createBackup": false
}
```

### 📑 Sheet Management (4 tools)

#### 15. `excel_create_sheet`
Create a new sheet in the workbook.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "NewSheet",
  "createBackup": false
}
```

#### 16. `excel_delete_sheet`
Delete a sheet from the workbook.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "OldSheet",
  "createBackup": true
}
```

#### 17. `excel_rename_sheet`
Rename a sheet.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "oldName": "Sheet1",
  "newName": "Sales2024",
  "createBackup": false
}
```

#### 18. `excel_duplicate_sheet`
Duplicate a complete sheet.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sourceSheetName": "Template",
  "newSheetName": "January",
  "createBackup": false
}
```

### 🔧 Operations (3 tools)

#### 19. `excel_delete_rows`
Delete specific rows.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "startRow": 5,
  "count": 3,
  "createBackup": true
}
```

#### 20. `excel_delete_columns`
Delete specific columns.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "startColumn": "C",
  "count": 2,
  "createBackup": true
}
```

#### 21. `excel_copy_range`
Copy range to another location.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sourceSheetName": "Sales",
  "sourceRange": "A1:D10",
  "targetSheetName": "Backup",
  "targetCell": "A1",
  "createBackup": false
}
```

### 📊 Analysis (2 tools)

#### 22. `excel_search_value`
Search for a value in sheet/range.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "searchValue": "Apple",
  "range": "A1:Z100",
  "caseSensitive": false,
  "responseFormat": "markdown"
}
```

#### 23. `excel_filter_rows`
Filter rows by condition.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "column": "B",
  "condition": "greater_than",
  "value": 1000,
  "responseFormat": "json"
}
```

### 📈 Charts (1 tool)

#### 24. `excel_create_chart`
Create charts from data ranges.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "chartType": "column",
  "dataRange": "A1:B10",
  "position": "D2",
  "title": "Monthly Sales",
  "showLegend": true,
  "createBackup": false
}
```

**Note**: ExcelJS has limited native chart support. This creates a chart placeholder with metadata.

### 🔄 Pivot Tables (1 tool)

#### 25. `excel_create_pivot_table`
Create pivot tables for data analysis.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sourceSheetName": "Sales",
  "sourceRange": "A1:D100",
  "targetSheetName": "Pivot",
  "targetCell": "A1",
  "rows": ["Product"],
  "columns": ["Month"],
  "values": [
    { "field": "Amount", "aggregation": "sum" }
  ],
  "createBackup": false
}
```

### 📋 Excel Tables (1 tool)

#### 26. `excel_create_table`
Convert ranges to formatted Excel tables.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Data",
  "range": "A1:D50",
  "tableName": "SalesTable",
  "tableStyle": "TableStyleMedium2",
  "showRowStripes": true,
  "createBackup": false
}
```

### ✅ Validation (3 tools)

#### 27. `excel_validate_formula_syntax`
Validate formula syntax without applying it.

**Example:**
```json
{
  "formula": "SUM(A1:A10) / COUNT(B1:B10)"
}
```

#### 28. `excel_validate_range`
Validate if a range string is valid.

**Example:**
```json
{
  "range": "A1:Z100"
}
```

#### 29. `excel_get_data_validation_info`
Get data validation rules for a cell.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Input",
  "cellAddress": "A1",
  "responseFormat": "json"
}
```

### 🔧 Advanced Operations (4 tools)

#### 30. `excel_insert_rows`
Insert rows at a specific position.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "startRow": 5,
  "count": 3,
  "createBackup": false
}
```

#### 31. `excel_insert_columns`
Insert columns at a specific position.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "startColumn": "C",
  "count": 2,
  "createBackup": false
}
```

#### 32. `excel_unmerge_cells`
Unmerge previously merged cells.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Report",
  "range": "A1:D1",
  "createBackup": false
}
```

#### 33. `excel_get_merged_cells`
List all merged cell ranges in a sheet.

**Example:**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Report",
  "responseFormat": "markdown"
}
```

### 🎨 Conditional Formatting (1 tool)

#### 34. `excel_apply_conditional_format`
Apply conditional formatting to ranges.

**Example (Cell Value):**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "range": "B2:B100",
  "ruleType": "cellValue",
  "condition": {
    "operator": "greaterThan",
    "value": 1000
  },
  "style": {
    "fill": {
      "type": "pattern",
      "pattern": "solid",
      "fgColor": "FF00FF00"
    },
    "font": {
      "bold": true
    }
  },
  "createBackup": false
}
```

**Example (Color Scale):**
```json
{
  "filePath": "./data.xlsx",
  "sheetName": "Sales",
  "range": "C2:C100",
  "ruleType": "colorScale",
  "colorScale": {
    "minColor": "FFFF0000",
    "maxColor": "FF00FF00"
  },
  "createBackup": false
}
```

## Development

### Build
```bash
npm run build
```

### Watch mode
```bash
npm run watch
```

### Run
```bash
npm start
```

## Error Handling

All tools include robust error handling and will return descriptive error messages for:
- File not found
- Sheet not found
- Invalid cell addresses or ranges
- Invalid formatting options
- Write errors

## Features

### Backup Support
Most write operations support an optional `createBackup` parameter. When set to `true`, a backup of the original file will be created with a `.backup` extension before modifications.

### Response Formats
Read operations support both `json` and `markdown` response formats:
- **JSON**: Structured data, ideal for programmatic processing
- **Markdown**: Human-readable tables and formatted output

### Data Preview
When reading large datasets, the markdown format automatically shows a preview of the first 100 rows.

## Dependencies

- `@modelcontextprotocol/sdk` - Official MCP SDK
- `exceljs` - Excel file manipulation
- `zod` - Schema validation
- `typescript` - Type safety

## Links & Resources

- **This fork (v3.x)**: [github.com/emilio-S03/Excel-MCP-Server](https://github.com/emilio-S03/Excel-MCP-Server)
- **Issues & bug reports** (this fork): [GitHub Issues](https://github.com/emilio-S03/Excel-MCP-Server/issues)
- **Original upstream**: [github.com/sbraind/excel-mcp-server](https://github.com/sbraind/excel-mcp-server) (the source this fork extends — credit to sbraind)
- **Model Context Protocol**: [modelcontextprotocol.io](https://modelcontextprotocol.io/)
- **Claude Desktop**: [claude.ai/download](https://claude.ai/download)
- **ExcelJS Documentation**: [github.com/exceljs/exceljs](https://github.com/exceljs/exceljs)

## Coworker docs (read these first)

- **[docs/INSTALL.md](docs/INSTALL.md)** — full install funnel, Mac + Windows
- **[docs/EXAMPLES.md](docs/EXAMPLES.md)** — 5 prompts that exercise capabilities the native Claude for Excel add-in does not have
- **[docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md)** — common errors with actionable fixes
- **[docs/PLATFORM_PARITY.md](docs/PLATFORM_PARITY.md)** — full Windows / macOS / Linux tool matrix
- **[docs/EXCEL_FOR_WEB.md](docs/EXCEL_FOR_WEB.md)** — what works for browser-only Excel users
- **[docs/MAC_VERIFICATION_NEEDED.md](docs/MAC_VERIFICATION_NEEDED.md)** — Mac coworkers please read

## License

MIT — see [LICENSE](LICENSE). Joint copyright: original work © sbraind, v3.x fork © Emilio Soria.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### How to Contribute

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Support

If you encounter any issues or have questions:

1. Check the [Installation Guide](docs/INSTALL.md) and [Troubleshooting](docs/TROUBLESHOOTING.md) for common setup issues
2. Review [existing issues](https://github.com/emilio-S03/Excel-MCP-Server/issues) to see if your problem has been addressed
3. Open a [new issue](https://github.com/emilio-S03/Excel-MCP-Server/issues/new) with detailed information about your problem

---

**Built with ❤️ using [Claude Code](https://claude.com/claude-code)**
