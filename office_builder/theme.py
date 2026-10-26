"""Single source of truth: design_tokens.json.

Every builder and the auditor read the same file, so "consensus" is
measured against numbers, not opinions. See STANDARDS.md.
"""
from __future__ import annotations

import json
from pathlib import Path


def load_tokens(path: str | Path | None = None) -> dict:
    if path is None:
        path = Path(__file__).resolve().parent.parent / "design_tokens.json"
    return json.loads(Path(path).read_text(encoding="utf-8"))


def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    h = hex_color.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _linear(c: float) -> float:
    return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4


def contrast_ratio(fg_hex: str, bg_hex: str) -> float:
    """WCAG relative-luminance contrast ratio."""

    def lum(hex_color: str) -> float:
        r, g, b = (v / 255 for v in hex_to_rgb(hex_color))
        return 0.2126 * _linear(r) + 0.7152 * _linear(g) + 0.0722 * _linear(b)

    l1, l2 = lum(fg_hex), lum(bg_hex)
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def check_contrast(tokens: dict, fg: str, bg: str) -> bool:
    return contrast_ratio(fg, bg) >= tokens["contrast"]["normal_text_min"]
