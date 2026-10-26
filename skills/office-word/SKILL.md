---
name: office-word
description: Build a professional Word report with python-docx, template-first, with TOC and page-number fields that Word computes on open.
---

# Office Word

Use this skill when the task is a `.docx` report.

Workflow (short; details live in references/):

1. Read `design_tokens.json` and `STANDARDS.md`. Use only whitelisted styles.
2. Start from a template `.docx` that already has styles/header/footer,
   or from `office_builder/docx_builder.py`'s defaults.
3. Headings drive the TOC: Heading 1/2/3, never bold-as-heading.
4. Header/footer: set `is_linked_to_previous=False` before editing a new section.
5. TOC and PAGE are fields. python-docx does NOT compute them. Tell the
   user to Update Fields in Word, and audit only for field + headings.
6. Tables: total column width <= page width minus margins, header row
   repeats (`tblHeader`), no merges in the data area.
7. Run `office_builder/audit.py` on the real file before delivering.

See `references/docx.md` for the full API pitfalls and checks.
