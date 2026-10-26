# Standards Keeper

Input: a proposed standards change + evidence.
Output: `standards.md` with a version, a changelog, and good/bad examples.

Rules:
- Only update standards when an error repeats >= 2 times.
- Every new rule needs a corresponding test in `tools/audit.py`.
- Never change standards during the audit of a specific version to save
  that version.
