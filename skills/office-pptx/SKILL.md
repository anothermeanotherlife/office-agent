---
name: office-pptx
description: Build a professional PowerPoint deck with python-pptx, 16:9, one idea per slide, assertion titles, and charts that start at zero.
---

# Office PPTX

Use this skill when the task is a `.pptx` deck (without deep animation;
for animation, also load office-pptx-motion).

Workflow:

1. Read `design_tokens.json` and `STANDARDS.md`.
2. Start from a designed `.pptx` template; read layout/placeholder names,
   do not hard-code layout indexes blindly.
3. 16:9, one idea per slide, assertion title, body >=18pt, <=45 words
   per content slide (good range, 80 is the limit).
4. Bar chart value axis minimum = 0. One chart = one message.
5. python-pptx does NOT measure rendered text. Check word count and
   bounding boxes in code, then render with LibreOffice.
6. Run `office_builder/audit.py` on the real file before delivering.

See `references/pptx.md`.
