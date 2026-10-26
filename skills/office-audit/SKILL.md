---
name: office-audit
description: Audit a real Office file with code, independently of the builder, and produce a structured audit_report.json.
---

# Office Audit

Use this skill when checking any `.docx/.xlsx/.pptx` before delivery.

Workflow:

1. Load `rubric.md` and `STANDARDS.md`. Do NOT load the Builder's
   self-assessment.
2. Run `office_builder/audit.py` (or `tools/audit.py`) on the real files.
3. Pre-mortem: list at least 3 ways this version could be wrong before
   scoring.
4. Report blockers/majors/minors with rule, location, measured value,
   threshold and fix. No vague praise.
5. If code FAILs, verdict is FAIL, even if the deck "looks good".
6. For important files, also run `tools/render_check.py` (LibreOffice
   headless -> PDF) and, for Excel, a recalc + Python cross-check.
