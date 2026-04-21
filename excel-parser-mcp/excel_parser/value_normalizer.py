"""Generic value normalization utilities for post-LLM-processing."""

from __future__ import annotations

import re
from typing import Any


class ValueNormalizer:
    """Normalize values extracted by LLM into canonical formats."""

    _TRUE_VALUES = {"x", "X", "y", "Y", "yes", "Yes", "YES", "M", "m", "true", "True"}
    _FALSE_VALUES = {"n", "N", "no", "No", "NO", "-", "/", "\\", "false", "False", ""}

    @staticmethod
    def to_bool(value: Any) -> bool:
        s = str(value).strip()
        if s in ValueNormalizer._TRUE_VALUES:
            return True
        if s in ValueNormalizer._FALSE_VALUES:
            return False
        return bool(s)

    @staticmethod
    def to_hex(value: Any, strip_prefix: bool = True) -> str:
        s = str(value).strip()
        m = re.search(r"0x([0-9a-fA-F]+)", s, re.IGNORECASE)
        if m:
            return m.group(1) if strip_prefix else f"0x{m.group(1)}"
        m = re.search(r"([0-9a-fA-F]{2,})", s)
        if m:
            return m.group(1) if strip_prefix else f"0x{m.group(1)}"
        return s

    @staticmethod
    def to_int_ms(value: Any) -> int:
        s = str(value).strip()
        m = re.search(r"[\d.]+", s)
        if not m:
            return 0
        num = float(m.group())
        if "s" in s.lower() and "ms" not in s.lower():
            num *= 1000
        return int(num)
