# PPTX reference (python-pptx)

- Layouts: `prs.slide_layouts[0..8]` follow the standard theme order
  (Title, Title and Content, Section Header, Two Content, Comparison,
  Title Only, Blank, Content with Caption, Picture with Caption) but
  other templates differ — read layout names and placeholders.
- Master decides the whole deck, layout decides a slide type. Edit the
  master/layout/template for consistency, not slide by slide.
- Theme is `ppt/theme/theme1.xml`: `a:clrScheme` colours,
  `a:fontScheme` fonts. Chart series take Accent 1-6 in order.
- Fonts: set per run, including Latin, East Asian and Complex Script
  for Vietnamese, and check the font exists on the presenting machine.
- Pictures: `add_picture`, keep ratio, measure real size/DPI with Pillow
  first. python-pptx does not compress images.
- Tables: `add_table` returns a GraphicFrame; the table is `.table`.
  Create the right row/column count up front.
- Charts: `add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, ...)` with
  `CategoryChartData` / `XyChartData` / `BubbleChartData`. Pie has one
  series. Colour individual columns via `series.points[i]`.
- Coordinates are EMU: 914400 EMU = 1 inch.
- Text: python-pptx does NOT render text. `word_wrap` and `auto_size`
  do not guarantee no overflow. Limit words, font size and width, then
  render-check with LibreOffice/PowerPoint.
