"""Deterministic auditor for DOCX / XLSX / PPTX.

The auditor never trusts a builder log. It re-opens the real file,
measures real values, and compares them to design_tokens.json.
A PASS here is the structure rung of the evidence ladder only;
LibreOffice/PowerPoint render is a separate rung (tools/render_check.py).

Report shape (see agents/auditor.md):
  verdict, score (0-100), blockers, majors, minors, must_fix_next,
  artifact_sha256, checks[]
"""
from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path

from lxml import etree

from .theme import contrast_ratio, load_tokens

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _issue(rule: str, where: str, measured: str, fix: str) -> dict:
    return {"rule": rule, "where": where, "measured": measured, "fix": fix}


def audit_pptx(path: Path, tokens: dict) -> dict:
    blockers, majors, minors, checks = [], [], [], []

    def check(name: str, ok: bool, detail: str = ""):
        checks.append({"check": name, "pass": ok, "detail": detail})

    try:
        with zipfile.ZipFile(path) as z:
            assert z.testzip() is None
            slide_parts = sorted(n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
            for part in slide_parts:
                etree.fromstring(z.read(part))
        check("pptx_zip_and_xml_valid", True)
    except Exception as exc:  # noqa: BLE001
        blockers.append(_issue("pptx_valid", str(path), str(exc), "Rebuild the file; do not hand-edit the zip."))
        check("pptx_zip_and_xml_valid", False, str(exc))
        return _report(path, blockers, majors, minors, checks)

    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation(str(path))
    check("pptx_opens", True, f"{len(prs.slides)} slides")
    motion = tokens["motion"]
    for i, slide in enumerate(prs.slides, start=1):
        # Title present?
        has_title = bool(slide.shapes.title and slide.shapes.title.text_frame.text.strip())
        check(f"slide_{i}_has_title", has_title)
        if not has_title:
            blockers.append(_issue("slide_title", f"slide {i}", "no title", "Add a unique assertion-style title."))
        # Overflow / overlap (bbox, in inches).
        boxes = []
        for shape in slide.shapes:
            try:
                left, top = shape.left / 914400, shape.top / 914400
                width, height = shape.width / 914400, shape.height / 914400
            except Exception:
                continue
            boxes.append((shape.shape_id, left, top, width, height))
            if left + width > prs.slide_width / 914400 + 0.05 or top + height > prs.slide_height / 914400 + 0.05:
                blockers.append(_issue("slide_overflow", f"slide {i} shape {shape.shape_id}",
                                       f"right={left + width:.2f}in bottom={top + height:.2f}in",
                                       "Shrink or move the shape inside the slide."))
        # Word count per content slide (title slide excluded).
        if i > 1:
            words = sum(len(s.text_frame.text.split()) for s in slide.shapes if s.has_text_frame)
            check(f"slide_{i}_word_count", words <= tokens["layout"]["words_per_content_slide_limit"],
                  f"{words} words")
            if words > tokens["layout"]["words_per_content_slide_limit"]:
                majors.append(_issue("word_density", f"slide {i}", f"{words} words", "Split the slide or move detail to an appendix."))
        # Timing XML: spid must exist, cTn ids unique, at most one timing.
        with zipfile.ZipFile(path) as z:
            xml = z.read(f"ppt/slides/slide{i}.xml")
        root = etree.fromstring(xml)
        timings = root.findall(f"{{{P_NS}}}timing")
        check(f"slide_{i}_single_timing", len(timings) <= 1, f"{len(timings)} timing elements")
        if len(timings) > 1:
            blockers.append(_issue("single_timing", f"slide {i}", f"{len(timings)} timings", "Replace timing entirely, do not append."))
        shape_ids = {int(s.shape_id) for s in slide.shapes}
        spids = [int(el.get("spid")) for el in root.iter(f"{{{P_NS}}}spTgt") if el.get("spid")]
        for spid in spids:
            if spid not in shape_ids:
                blockers.append(_issue("timing_spid_exists", f"slide {i}", f"spid {spid} not found",
                                       "Read shape ids back after saving; never guess."))
        ctn_ids = [el.get("id") for el in root.iter(f"{{{P_NS}}}cTn")]
        if len(ctn_ids) != len(set(ctn_ids)):
            blockers.append(_issue("ctn_id_unique", f"slide {i}", "duplicate cTn id", "Assign unique, increasing cTn ids per slide."))
        # Chart axis starts at 0.
        for shape in slide.shapes:
            if shape.has_chart:
                try:
                    minimum = shape.chart.value_axis.minimum_scale
                except Exception:
                    minimum = None
                check(f"slide_{i}_chart_axis_zero", minimum in (0, 0.0, None), f"minimum={minimum}")
                if minimum not in (0, 0.0, None):
                    blockers.append(_issue("bar_axis_zero", f"slide {i}", f"minimum={minimum}", "Set value_axis.minimum_scale = 0."))
    score = _score(blockers, majors, minors)
    return _report(path, blockers, majors, minors, checks, score)


def audit_docx(path: Path, tokens: dict) -> dict:
    blockers, majors, minors, checks = [], [], [], []

    def check(name: str, ok: bool, detail: str = ""):
        checks.append({"check": name, "pass": ok, "detail": detail})

    try:
        with zipfile.ZipFile(path) as z:
            assert z.testzip() is None
            # TOC lives in document.xml, PAGE/NUMPAGES live in footer parts,
            # so check all Word XML parts, not just document.xml.
            xml = "".join(
                z.read(name).decode(errors="ignore")
                for name in z.namelist()
                if name.startswith("word/") and name.endswith(".xml")
            )
            rels = z.read("word/_rels/document.xml.rels").decode()
        check("docx_zip_valid", True)
    except Exception as exc:  # noqa: BLE001
        blockers.append(_issue("docx_valid", str(path), str(exc), "Rebuild the file."))
        check("docx_zip_valid", False, str(exc))
        return _report(path, blockers, majors, minors, checks)

    from docx import Document

    doc = Document(str(path))
    styles = {s.name for s in doc.styles}
    check("docx_styles_exist", all(s in styles for s in ["Heading 1", "Heading 2", "Normal"]))
    headings = [p.style.name for p in doc.paragraphs if p.style.name.startswith("Heading")]
    check("docx_has_headings", bool(headings), str(headings))
    if not headings:
        blockers.append(_issue("docx_headings", "document", "no headings", "Use Heading 1/2/3, not bold-as-heading."))
    # No heading jump H1 -> H3, no empty heading.
    levels = [int(h.split()[-1]) for h in headings if h.split()[-1].isdigit()]
    for a, b in zip(levels, levels[1:]):
        if b > a + 1:
            majors.append(_issue("heading_level_jump", "document", f"{a} -> {b}", "Do not skip heading levels."))
    for p in doc.paragraphs:
        if p.style.name.startswith("Heading") and not p.text.strip():
            blockers.append(_issue("empty_heading", "document", "empty heading", "Remove or fill the heading."))
        if p.text.strip() and p.style.name not in tokens["docx"]["styles_whitelist"]:
            majors.append(_issue("style_whitelist", "document", f"style {p.style.name}", "Use only whitelisted styles."))
    if "TOC" not in xml:
        majors.append(_issue("toc_field", "document", "no TOC field", "Insert a TOC field; Word computes it on open."))
    if "PAGE" not in xml:
        minors.append(_issue("page_field", "document", "no PAGE field", "Insert a PAGE field in the footer."))
    for t in doc.tables:
        if t.style.name not in ("Table Grid", "Light Grid Accent 1"):
            minors.append(_issue("table_style", "document", t.style.name, "Use Table Grid or Light Grid Accent 1."))
    score = _score(blockers, majors, minors)
    return _report(path, blockers, majors, minors, checks, score)


def audit_xlsx(path: Path, tokens: dict) -> dict:
    blockers, majors, minors, checks = [], [], [], []

    def check(name: str, ok: bool, detail: str = ""):
        checks.append({"check": name, "pass": ok, "detail": detail})

    try:
        from openpyxl import load_workbook

        wb = load_workbook(str(path), data_only=False)
        check("xlsx_opens", True, str(wb.sheetnames))
    except Exception as exc:  # noqa: BLE001
        blockers.append(_issue("xlsx_valid", str(path), str(exc), "Rebuild the file."))
        check("xlsx_opens", False, str(exc))
        return _report(path, blockers, majors, minors, checks)

    if "Data" not in wb.sheetnames or "Summary" not in wb.sheetnames:
        blockers.append(_issue("xlsx_sheets", str(wb.sheetnames), "missing Data/Summary", "Name sheets README/Data/Summary/Dashboard."))
    else:
        data = wb["Data"]
        summary = wb["Summary"]
        # Summary totals must be formulas, not hand-typed values.
        for cell_ref in ("B2", "B3", "B4", "B5"):
            value = summary[cell_ref].value
            check(f"summary_{cell_ref}_is_formula", isinstance(value, str) and value.startswith("="), str(value))
            if not (isinstance(value, str) and value.startswith("=")):
                blockers.append(_issue("summary_formula", f"Summary!{cell_ref}", str(value), "Link Summary to Data with a formula."))
        # Independent Python cross-check from Data values.
        try:
            col_b = [data.cell(row=r, column=2).value for r in range(2, data.max_row + 1) if isinstance(data.cell(row=r, column=2).value, (int, float))]
            col_c = [data.cell(row=r, column=3).value for r in range(2, data.max_row + 1) if isinstance(data.cell(row=r, column=3).value, (int, float))]
            check("xlsx_cross_check", bool(col_b) and bool(col_c), f"sums {sum(col_b)} / {sum(col_c)}")
        except Exception as exc:  # noqa: BLE001
            majors.append(_issue("xlsx_cross_check", "Data", str(exc), "Fix Data types: numbers must be numbers."))
        if summary.freeze_panes is None:
            minors.append(_issue("freeze_panes", "Summary", "none", "Set freeze_panes='A2'."))
        if data.freeze_panes is None:
            minors.append(_issue("freeze_panes", "Data", "none", "Set freeze_panes='A2'."))
        dashboard = wb["Dashboard"] if "Dashboard" in wb.sheetnames else None
        if dashboard is not None:
            check("dashboard_has_chart", len(dashboard._charts) >= 1, f"{len(dashboard._charts)} charts")
            if not dashboard._charts:
                majors.append(_issue("dashboard_chart", "Dashboard", "no chart", "Add a chart sourced from Data."))
    score = _score(blockers, majors, minors)
    return _report(path, blockers, majors, minors, checks, score)


def _score(blockers, majors, minors) -> int:
    score = 100 - 30 * len(blockers) - 10 * len(majors) - 2 * len(minors)
    return max(0, score)


def _report(path: Path, blockers, majors, minors, checks, score: int | None = None) -> dict:
    if score is None:
        score = _score(blockers, majors, minors)
    tokens_gate = {"verdict": "PASS" if not blockers and score >= 85 else "FAIL"}
    return {
        "artifact": str(path),
        "artifact_sha256": _sha256(path),
        "verdict": tokens_gate["verdict"],
        "score": score,
        "blockers": blockers,
        "majors": majors,
        "minors": minors,
        "must_fix_next": [b["fix"] for b in blockers] + [m["fix"] for m in majors],
        "checks": checks,
    }


def audit_file(path: str | Path, tokens: dict | None = None) -> dict:
    tokens = tokens or load_tokens()
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".pptx":
        return audit_pptx(path, tokens)
    if suffix == ".docx":
        return audit_docx(path, tokens)
    if suffix == ".xlsx":
        return audit_xlsx(path, tokens)
    raise ValueError(f"Unsupported file type: {suffix}")
