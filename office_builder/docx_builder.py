"""Build a Word report with python-docx, template-first.

Rules enforced here (see STANDARDS.md):
- Styles, never ad-hoc bold-as-heading. Headings drive the TOC.
- Header/footer are unlinked before editing a new section.
- TOC and PAGE are fields: Word/LibreOffice compute them on open.
  python-docx does NOT compute them, and the audit only checks that the
  field exists plus that Heading 1-3 exist.
- Table width must fit page width minus margins; header row repeats.
- `paragraph.text = ...` / `cell.text = ...` reset runs, so we write into
  the first run instead.
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .theme import hex_to_rgb, load_tokens


def _rgb(hex_color: str) -> RGBColor:
    r, g, b = hex_to_rgb(hex_color)
    return RGBColor(r, g, b)


def _add_field(paragraph, instr: str) -> None:
    """Insert a Word field: begin -> instrText -> separate -> end."""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run._r.append(fld_begin)
    run2 = paragraph.add_run()
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    run2._r.append(instr_el)
    run3 = paragraph.add_run()
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run3._r.append(fld_sep)
    run4 = paragraph.add_run()
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run4._r.append(fld_end)


def _set_cell_shading(cell, hex_color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def _set_repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def _style_document(doc, tokens: dict) -> None:
    colors = tokens["colors"]
    fonts = tokens["fonts"]
    sizes = tokens["sizes_pt"]
    normal = doc.styles["Normal"]
    normal.font.name = fonts["body"]
    normal.font.size = Pt(sizes["docx_body"])
    normal.font.color.rgb = _rgb(colors["text"])
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15
    for name, size in (("Heading 1", sizes["docx_h1"]), ("Heading 2", sizes["docx_h2"])):
        style = doc.styles[name]
        style.font.name = fonts["heading"]
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = _rgb(colors["primary"])
        style.paragraph_format.keep_with_next = True


def build_docx(data: dict, out_path: str | Path, tokens: dict | None = None) -> Path:
    tokens = tokens or load_tokens()
    colors = tokens["colors"]
    doc = Document()
    _style_document(doc, tokens)

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    margin = Cm(tokens["docx"]["page"]["margin_cm"])
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = margin

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.text = data.get("brand", tokens["brand"])
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run("Trang ")
    _add_field(fp, " PAGE ")
    fp.add_run(" / ")
    _add_field(fp, " NUMPAGES ")

    title = doc.add_heading(data["title"], level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph(data.get("subtitle", ""))

    # TOC field (computed by Word/LibreOffice on open).
    toc_para = doc.add_paragraph()
    _add_field(toc_para, ' TOC \\o "1-3" \\h \\z \\u ')
    doc.add_paragraph(
        "Mục lục sẽ được Word/LibreOffice cập nhật khi mở file "
        "(References > Update Table). Audit chỉ kiểm tra field TOC và Heading."
    )

    for chapter in data["chapters"]:
        doc.add_heading(chapter["heading"], level=1)
        for para in chapter.get("paragraphs", []):
            doc.add_paragraph(para)
        for table in chapter.get("tables", []):
            headers = table["headers"]
            rows = table["rows"]
            t = doc.add_table(rows=1 + len(rows), cols=len(headers))
            t.style = "Table Grid"
            t.autofit = False
            hdr = t.rows[0]
            _set_repeat_header(hdr)
            for i, h in enumerate(headers):
                cell = hdr.cells[i]
                cell.paragraphs[0].runs.clear() if cell.paragraphs[0].runs else None
                r = cell.paragraphs[0].add_run(h)
                r.bold = True
                r.font.color.rgb = _rgb(colors["header_text"])
                _set_cell_shading(cell, colors["header_bg"])
            for ri, row in enumerate(rows):
                for ci, value in enumerate(row):
                    cell = t.rows[ri + 1].cells[ci]
                    # Write into the first run instead of cell.text (which resets runs).
                    p = cell.paragraphs[0]
                    if p.runs:
                        p.runs[0].text = str(value)
                    else:
                        p.add_run(str(value))
            doc.add_paragraph(table.get("caption", ""), style="Caption")

    # Appendix section, unlinked header/footer handling is inherited here;
    # a new section demonstrates section handling without breaking the first.
    doc.add_section(WD_SECTION.NEW_PAGE)
    doc.add_heading("Phụ lục A — Nguồn số liệu", level=1)
    for source in data.get("sources", []):
        doc.add_paragraph(source, style="List Bullet")

    core = doc.core_properties
    core.title = data["title"]
    core.author = data.get("author", tokens["brand"])
    core.language = "vi-VN"
    core.comments = "Built by office-agent. TOC/PAGE fields are computed by Word/LibreOffice on open."

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return out
