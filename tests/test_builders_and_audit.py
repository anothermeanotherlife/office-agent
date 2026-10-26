import json
from pathlib import Path

import pytest

from office_builder.audit import audit_file
from office_builder.docx_builder import build_docx
from office_builder.pptx_builder import build_pptx
from office_builder.pptx_motion import add_fade_click_beats
from office_builder.theme import contrast_ratio, load_tokens
from office_builder.xlsx_builder import build_xlsx

ROOT = Path(__file__).resolve().parent.parent
DATA = json.loads((ROOT / "examples" / "data.json").read_text(encoding="utf-8"))


def test_contrast_primary_on_white():
    tokens = load_tokens()
    assert contrast_ratio(tokens["colors"]["primary"], "FFFFFF") >= 4.5


def test_builders_and_audit(tmp_path):
    tokens = load_tokens()
    docx_path = build_docx(DATA, tmp_path / "report.docx", tokens)
    xlsx_path = build_xlsx(DATA, tmp_path / "report.xlsx", tokens)
    raw = build_pptx(DATA, tmp_path / "raw.pptx", tokens)
    pptx_path = add_fade_click_beats(raw, tmp_path / "report.pptx", tokens)
    for path in (docx_path, xlsx_path, pptx_path):
        assert path.exists() and path.stat().st_size > 0
        report = audit_file(path, tokens)
        assert report["blockers"] == [], report
        assert report["score"] >= 85, report


def test_xlsx_summary_is_formula(tmp_path):
    from openpyxl import load_workbook

    path = build_xlsx(DATA, tmp_path / "report.xlsx", load_tokens())
    wb = load_workbook(str(path), data_only=False)
    assert wb["Summary"]["B2"].value == "=SUM(Data!B2:B4)"
    assert wb["Summary"]["B4"].value == "=B2-B3"
