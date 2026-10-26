# DOCX reference (python-docx)

- Styles: `doc.styles['Normal']`, `['Heading 1'..'Heading 3']`, `['Title']`,
  `['Caption']`, `['List Bullet'/'List Number']`. Set font for ascii +
  hAnsi + eastAsia for Vietnamese text. A missing style raises KeyError:
  assert it exists in the template first.
- Sections: `doc.sections[0]`, `add_section(WD_SECTION.NEW_PAGE)`.
  Changing orientation swaps width/height. Each section has its own
  header/footer.
- Header/footer: set `is_linked_to_previous=False` before editing.
- Page numbers: no built-in API. Insert a field
  `w:fldChar begin -> w:instrText PAGE -> w:fldChar end` (same for NUMPAGES).
- Tables: `add_table`, `table.style='Table Grid'`, `autofit=False`, set
  `tblW` + `gridCol` so the total fits the page minus margins, header row
  `tblHeader`. `cell.text = ...` resets runs: write into the first run.
- TOC: no built-in API. Insert field `TOC \o "1-3" \h \z \u` and set
  `updateFields` so Word asks to update on open. Audit can only check
  field + headings, not TOC page numbers before a render.
- Numbering: default List Bullet/Number numbering; restart/multilevel
  needs custom `numbering.xml` or python-docx-ng (do not install it
  alongside python-docx — same namespace).
- Images: `add_picture(path, width=...)` only (keeps ratio). Caption is
  a `Caption` paragraph + `SEQ Figure` field via lxml.
- Template-first is preferred: build a `.docx` in Word, then
  `Document('template.docx')` or docxtpl (Jinja tags must stay in one run).
