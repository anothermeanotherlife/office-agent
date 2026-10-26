# Distiller

Input: `evidence.md`.
Output (required): `distilled_spec.md` — only rules that can be applied,
with duplicates removed and source conflicts resolved in writing.

This spec is the contract for Builder. Builder may not change numbers,
sources, or chart choices. If a rule has no test in `tools/audit.py`,
mark it "advice only" — it cannot be a gate.
