"""Build a PowerPoint deck with python-pptx, template-first.

Rules enforced here (see STANDARDS.md):
- Start from the default template's layouts; do not invent a master from
  scratch in code. Read layout/placeholder names, do not hard-code layout
  indexes blindly.
- 16:9, one idea per slide, assertion-style titles, <=45 words per
  content slide (good range), body >=18pt.
- Bar charts start at 0 by construction: values are non-negative and the
  value axis minimum is set to 0.
- Coordinates are EMU; python-pptx does NOT measure rendered text, so
  word counts and slide coverage are checked by audit.py and by a
  LibreOffice render (tools/render_check.py).
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

from .theme import hex_to_rgb, load_tokens


def _rgb(hex_color: str):
    from pptx.dml.color import RGBColor

    r, g, b = hex_to_rgb(hex_color)
    return RGBColor(r, g, b)


def _set_text(shape, text: str, size: int, bold: bool, color: str, alignment=PP_ALIGN.LEFT):
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = _rgb(color)
    p.font.name = "Calibri"
    p.alignment = alignment
    return tf


def _add_slide(prs, layout_name: str):
    for layout in prs.slide_layouts:
        if layout.name == layout_name:
            return prs.slides.add_slide(layout)
    return prs.slides.add_slide(prs.slide_layouts[6])  # Blank fallback


def build_pptx(data: dict, out_path: str | Path, tokens: dict | None = None) -> Path:
    tokens = tokens or load_tokens()
    colors = tokens["colors"]
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = data["title"]
    prs.core_properties.author = data.get("author", tokens["brand"])
    prs.core_properties.language = "vi-VN"

    # Title slide: use the Title Slide layout if present, else Blank + shapes.
    title_slide = _add_slide(prs, "Title Slide")
    if title_slide.shapes.title:
        _set_text(title_slide.shapes.title, data["title"], 40, True, colors["primary"])
    else:
        box = title_slide.shapes.add_textbox(Inches(0.7), Inches(2.2), Inches(11.9), Inches(1.2))
        _set_text(box, data["title"], 40, True, colors["primary"])
    if len(title_slide.placeholders) > 1:
        _set_text(title_slide.placeholders[1], data.get("subtitle", ""), 20, False, colors["muted"])
    else:
        box = title_slide.shapes.add_textbox(Inches(0.7), Inches(3.6), Inches(11.9), Inches(0.8))
        _set_text(box, data.get("subtitle", ""), 20, False, colors["muted"])

    for slide_data in data["slides"]:
        slide = _add_slide(prs, "Title and Content")
        if slide.shapes.title:
            _set_text(slide.shapes.title, slide_data["title"], 32, True, colors["primary"])
        body_text = "\n".join(f"• {line}" for line in slide_data.get("bullets", []))
        if len(slide.placeholders) > 1:
            _set_text(slide.placeholders[1], body_text, 20, False, colors["text"])
        else:
            box = slide.shapes.add_textbox(Inches(0.7), Inches(1.6), Inches(11.9), Inches(5.2))
            _set_text(box, body_text, 20, False, colors["text"])

        chart_data = slide_data.get("chart")
        if chart_data:
            # Replace the body box area with a chart on the right half when
            # a chart is present, keeping bullets on the left. Simpler and
            # safer: add the chart below the bullets in its own slide area.
            # Here we add it as a second visual on the same slide.
            cd = CategoryChartData()
            cd.categories = chart_data["categories"]
            cd.add_series(chart_data["series_name"], chart_data["values"])
            x, y, cx, cy = Inches(7.0), Inches(1.8), Inches(5.6), Inches(4.8)
            graphic = slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, cd
            )
            chart = graphic.chart
            chart.has_legend = False
            chart.value_axis.minimum_scale = 0  # Bar axis must start at 0.
            chart.value_axis.has_title = True
            chart.value_axis.axis_title.text_frame.text = chart_data.get("value_label", "Giá trị")
            chart.category_axis.has_title = False
            # Colour the first series with the brand primary.
            series = chart.series[0]
            fill = series.format.fill
            fill.solid()
            fill.fore_color.rgb = _rgb(colors["primary"])

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    return out
