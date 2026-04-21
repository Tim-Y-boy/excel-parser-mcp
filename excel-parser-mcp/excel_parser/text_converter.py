"""SheetData to LLM-readable text converter with smart sheet filtering."""

from __future__ import annotations

import re
from typing import Any

from .reader import SheetData

_SENTINEL_RE = re.compile(r"#endofdata", re.IGNORECASE)


class SheetTextConverter:
    """Convert SheetData to LLM-friendly text with smart sheet filtering."""

    DEFAULT_RELEVANT_KEYWORDS = [
        "basic", "diagnostic", "service", "did", "dtc", "routine",
        "session", "nrc", "timing", "general", "data", "table", "list",
        "parameter", "config", "result", "report", "summary",
    ]

    DEFAULT_EXCLUDE_KEYWORDS = [
        "history", "change", "revision", "instruction", "formula",
        "predefined", "template", "legend", "document", "cover",
        "readme", "toc", "index",
    ]

    def __init__(
        self,
        relevant_keywords: list[str] | None = None,
        exclude_keywords: list[str] | None = None,
        max_rows: int = 200,
    ):
        self._relevant = relevant_keywords or self.DEFAULT_RELEVANT_KEYWORDS
        self._exclude = exclude_keywords or self.DEFAULT_EXCLUDE_KEYWORDS
        self._max_rows = max_rows

    def convert_workbook(self, sheets: dict[str, SheetData], filter_sheets: bool = True) -> str:
        if filter_sheets:
            names = self.filter_relevant_sheets(list(sheets.keys()))
        else:
            names = list(sheets.keys())
        parts = []
        for name in names:
            if name in sheets:
                text = self.convert_sheet(sheets[name])
                if text.strip():
                    parts.append(text)
        return "\n\n".join(parts)

    def convert_sheet(self, sheet_data: SheetData) -> str:
        lines = [f"=== Sheet: {sheet_data.sheet_name} ({sheet_data.max_row} rows x {sheet_data.max_col} cols) ==="]
        for i, row in enumerate(sheet_data.rows):
            if i >= self._max_rows:
                lines.append(f"... (truncated at {self._max_rows} rows)")
                break
            first_val = str(row[0]).strip() if row and row[0] is not None else ""
            if _SENTINEL_RE.search(first_val):
                break
            vals = [_format_cell(v) for v in row]
            if all(v == "" for v in vals):
                continue
            lines.append(f"Row {i + 1}: {' | '.join(vals)}")
        return "\n".join(lines)

    def filter_relevant_sheets(self, sheet_names: list[str]) -> list[str]:
        result = []
        for name in sheet_names:
            lower = name.lower()
            if any(kw.lower() in lower for kw in self._exclude):
                continue
            result.append(name)
        return result


def _format_cell(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() in ("none", "nan", "null"):
        return ""
    return s
