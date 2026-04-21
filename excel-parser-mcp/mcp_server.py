"""Excel Parser MCP Server — Universal Excel-to-text tool for AI agents.

Provides 3 tools:
  1. list_sheets      — List all sheets with metadata
  2. read_excel       — Read entire workbook as LLM-friendly text
  3. read_single_sheet — Read a specific sheet

Usage in Claude Code (.claude/settings.json):
  {
    "mcpServers": {
      "excel-parser": {
        "command": "python",
        "args": ["path/to/mcp_server.py"]
      }
    }
  }
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from excel_parser import ExcelReader, SheetTextConverter
from mcp.server.fastmcp import FastMCP

mcp = FastMCP(
    "excel-parser",
    instructions=(
        "Universal Excel file parser. Supports .xlsx and .xls formats. "
        "Handles merged cells, multi-row headers, and auto-detects data boundaries. "
        "Use list_sheets to explore, then read_excel or read_single_sheet to get text."
    ),
)


@mcp.tool()
def list_sheets(file_path: str) -> str:
    """List all sheets in an Excel file with row/col counts and merge info.

    Args:
        file_path: Path to the Excel file (.xlsx or .xls)
    """
    try:
        reader = ExcelReader(file_path)
        sheets = reader.read_all_sheets()
        return json.dumps([
            {"name": name, "rows": sd.max_row, "cols": sd.max_col, "merged_cells": len(sd.merged_ranges)}
            for name, sd in sheets.items()
        ], ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def read_excel(
    file_path: str,
    filter_sheets: bool = True,
    relevant_keywords: list[str] | None = None,
    exclude_keywords: list[str] | None = None,
    max_rows: int = 200,
) -> str:
    """Read an Excel file and convert to LLM-friendly text.

    Args:
        file_path: Path to the Excel file (.xlsx or .xls)
        filter_sheets: Whether to filter out irrelevant sheets (default True)
        relevant_keywords: Custom keywords to keep sheets (default: built-in generic keywords)
        exclude_keywords: Custom keywords to exclude sheets (default: built-in list)
        max_rows: Max rows per sheet (default 200)
    """
    try:
        reader = ExcelReader(file_path, max_rows_per_sheet=max_rows)
        sheets = reader.read_all_sheets()
        converter = SheetTextConverter(
            relevant_keywords=relevant_keywords,
            exclude_keywords=exclude_keywords,
            max_rows=max_rows,
        )
        text = converter.convert_workbook(sheets, filter_sheets=filter_sheets)
        filtered = converter.filter_relevant_sheets(list(sheets.keys())) if filter_sheets else list(sheets.keys())
        return json.dumps({
            "text": text,
            "stats": {
                "total_sheets": len(sheets),
                "filtered_sheets": len(filtered),
                "text_length": len(text),
                "sheet_names": filtered,
            },
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def read_single_sheet(file_path: str, sheet_name: str, max_rows: int = 200) -> str:
    """Read a single sheet from an Excel file.

    Args:
        file_path: Path to the Excel file (.xlsx or .xls)
        sheet_name: Name of the sheet to read
        max_rows: Max rows to read (default 200)
    """
    try:
        reader = ExcelReader(file_path, max_rows_per_sheet=max_rows)
        sd = reader.read_sheet(sheet_name)
        converter = SheetTextConverter(max_rows=max_rows)
        return json.dumps({
            "sheet": sheet_name, "rows": sd.max_row, "cols": sd.max_col,
            "text": converter.convert_sheet(sd),
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
