---
name: office-pptx-motion
description: Add native, editable PowerPoint animation and Morph with click beats, using safe timing-XML injection. Inspired by togcoder/skill-for-pptx, with an honest evidence ladder.
---

# Office PPTX Motion

Use this skill when animation or Morph is requested. Load office-pptx
first for the deck itself.

Workflow:

1. Read `design_tokens.json` (motion durations, stagger, easing) and
   `STANDARDS.md` §6. Generator and critic must use the same numbers.
2. Model the rhythm first: Slide -> Click Beat -> Stage -> Effects.
   A click is a presenter boundary, not a long delay.
3. Do NOT target the title. Max 4 moving objects per click, <=6 clicks
   per slide (good), <=8 (limit), <=8 effects per click.
4. Animation is XML: python-pptx has no animation API. Read shape ids
   back after saving, never guess them. If a slide already has
   `<p:timing>`, replace it entirely. Keep schema order:
   cSld, clrMapOvr?, transition, timing, extLst. cTn ids unique.
5. Prefer a "golden file": build one slide in PowerPoint with the exact
   effect, unzip it, copy its timing fragment, and only change spid,
   dur, delay and preset.
6. Morph is a transition on the DESTINATION slide, with paired shapes
   sharing a `!!` name. LibreOffice has no Morph — Morph needs a
   PowerPoint check.
7. Triggers live in `interactiveSeq` with `evt="onClick"` on the trigger
   shape. CreateVideo cannot click, so triggers only count as verified
   after real interactive playback.
8. Evidence ladder — claim only the rung you actually reached:
   structure -> simulation -> LibreOffice -> PowerPoint read/render ->
   interactive click playback.
9. Run `office_builder/audit.py` (timing spid, cTn uniqueness, single
   timing, overflow) and `tools/render_check.py` before delivering.

See `references/animation-xml.md` for the timing tree, presets, and
the three quarantined scripts from the reference repo that must NOT be
reused (morph_playbook, cognitive_pacing, interactive_triggers — they
wrote invalid schema XML).
