# Excel Parser MCP Server

A universal Excel-to-text MCP (Model Context Protocol) server for AI agents. Feed any `.xlsx` or `.xls` file into your LLM workflow with zero configuration.

## What It Does

| Problem | Solution |
|---------|----------|
| LLMs can't read Excel files directly | Converts Excel to structured, LLM-friendly text |
| Different Excel formats across projects | Format-agnostic — no adapters or configs needed |
| Merged cells break data parsing | Auto-resolves merged cells by inheriting top-left values |
| Too many irrelevant sheets | Smart keyword-based filtering reduces token usage |
| Both `.xlsx` and legacy `.xls` exist | Unified reader handles both transparently |

## Quick Start

### Install

```bash
git clone https://github.com/<your-org>/excel-parser-mcp.git
cd excel-parser-mcp
pip install -e .
```

### Configure in Claude Code

Add to `.claude/settings.json`:

```json
{
  "mcpServers": {
    "excel-parser": {
      "command": "python",
      "args": ["/path/to/excel-parser-mcp/mcp_server.py"]
    }
  }
}
```

Restart Claude Code. You now have three new tools available.

## Tools

### 1. `list_sheets`

List all sheets with metadata (rows, columns, merged cells).

```
> Use the list_sheets tool to show me what's in "report.xlsx"
```

Returns:
```json
[
  {"name": "Sales Data", "rows": 150, "cols": 12, "merged_cells": 5},
  {"name": "Summary", "rows": 30, "cols": 8, "merged_cells": 2},
  {"name": "Change History", "rows": 20, "cols": 4, "merged_cells": 0}
]
```

### 2. `read_excel`

Read the entire workbook as LLM-friendly text. Automatically filters out irrelevant sheets.

```
> Read "report.xlsx" and summarize the sales data
```

Features:
- **Smart filtering** — skips sheets like "Change History", "Cover", "Instructions"
- **Merged cells resolved** — child cells inherit the top-left value
- **Auto-truncation** — caps at 200 rows per sheet to control token usage
- **Sentinel detection** — stops at `#endofdata` markers

Customize filtering with your own keywords:

```json
{
  "file_path": "report.xlsx",
  "relevant_keywords": ["sales", "revenue", "quarterly"],
  "exclude_keywords": ["draft", "internal"],
  "max_rows": 100
}
```

### 3. `read_single_sheet`

Read one specific sheet when you don't need the whole file.

```
> Read the "Sales Data" sheet from "report.xlsx"
```

## Architecture

```
excel-parser-mcp/
├── mcp_server.py              # MCP Server entry point (3 tools)
├── excel_parser/
│   ├── reader.py              # Unified .xlsx/.xls reader (~100 LOC)
│   ├── text_converter.py      # Sheet-to-text + smart filtering (~80 LOC)
│   └── value_normalizer.py    # Bool/hex/time normalization (~40 LOC)
└── pyproject.toml
```

### How It Works

```
┌──────────────┐     ┌─────────────────┐     ┌──────────────┐
│  ExcelReader  │────▶│ SheetTextConverter│────▶│  LLM Agent   │
│  (.xlsx/.xls)│     │ (filter + format) │     │ (Claude/GPT) │
└──────────────┘     └─────────────────┘     └──────────────┘
       │                      │
       ▼                      ▼
  SheetData              Plain text
  (rows + merges)     "Row 1: A | B | C"
```

**Key design decisions:**

1. **Format-agnostic** — No format adapters or YAML configs. Any new Excel format works immediately because format understanding is delegated to the LLM.

2. **Merged cell resolution** — Pre-computes a `{(row,col): top_left}` mapping. When a Service ID spans 3 rows in a merged cell, all 3 rows see the same value.

3. **Sentinel detection** — Recognizes `#endofdata` / `#EndOfData` markers (case-insensitive) to stop reading beyond the actual data region.

4. **Smart filtering** — Two keyword lists: `relevant` (keep) and `exclude` (skip). Sheets matching neither list are kept by default ("prefer inclusion over omission").

## Use as a Python Library

The parser can also be used directly without MCP:

```python
from excel_parser import ExcelReader, SheetTextConverter

# Read any Excel file
reader = ExcelReader("any_file.xlsx")
sheets = reader.read_all_sheets()

# Convert to text (with filtering)
converter = SheetTextConverter(
    relevant_keywords=["your", "keywords"],
    exclude_keywords=["skip", "these"],
)
text = converter.convert_workbook(sheets)

# Feed to any LLM
print(text)
```

## Requirements

- Python 3.12+
- openpyxl >= 3.1.0 (for .xlsx)
- xlrd >= 2.0.0 (for .xls)
- mcp >= 1.0.0 (for MCP server mode)

## License

MIT
