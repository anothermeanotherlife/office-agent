# Animation XML reference

Timing tree (simplified):

```xml
<p:timing><p:tnLst><p:par>
  <p:cTn id="1" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
    <p:par><p:cTn id="2" fill="hold">
      <p:stCondLst><p:cond delay="indefinite" evt="onClick"/></p:stCondLst>
      <p:childTnLst><!-- effect --></p:childTnLst>
    </p:cTn></p:par>
  </p:childTnLst></p:cTn>
</p:par></p:tnLst></p:timing>
```

Effect kinds: `p:animEffect` (preset entrance/exit), `p:anim` (numeric
attribute), `p:animClr` (colour), `p:animRot`, `p:animScale`,
`p:animMotion` (path, using `#ppt_x`, `#ppt_y`, `#ppt_w`, `#ppt_h`),
`p:set`. Target: `p:tgtEl > p:spTgt spid="<real shape id>"`, or with
`p:txEl > p:paraRg` for paragraph ranges.

Durations (design_tokens.json): entrance 250-700ms (default 500),
exit 200-500ms (350), emphasis 250-600ms (450), move 400-1200ms (700),
Morph 1000-1600ms (1400), stagger 60-150ms (120), ambient loop 6-60s (12).
Easing in PowerPoint is only accel/decel: enter 0/0.90, exit 0.35/0,
move/emphasis 0.15/0.70.

Common failures: duplicate cTn id, spid pointing nowhere after a
python-pptx re-save, two `<p:timing>` elements, wrong schema order,
presetID/presetClass/presetSubtype mismatch (PowerPoint shows "Custom"
or drops the effect), delay instead of click, Morph names not matching,
text overflow because python-pptx never renders.

Do NOT reuse the reference repo's quarantined scripts
(morph_playbook.py, cognitive_pacing.py, interactive_triggers.py):
they wrote invalid schema XML or failed on real decks.
