"""Render check via LibreOffice headless -> PDF.

This is the second rung of the evidence ladder. It does NOT prove
PowerPoint playback, and LibreOffice has no Morph. It catches broken
pages, overflow and conversion failures that code-only audit misses.

Usage: python -m tools.render_check out.pdf file.pptx file.docx
Requires: soffice / libreoffice on PATH.
"""
import subprocess
import sys
from pathlib import Path


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print("Usage: python -m tools.render_check <outdir> <file> [file ...]")
        return 2
    outdir = Path(argv[0])
    outdir.mkdir(parents=True, exist_ok=True)
    for arg in argv[1:]:
        result = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), arg],
            capture_output=True, text=True, timeout=180,
        )
        print(arg, "->", result.returncode)
        if result.returncode != 0:
            print(result.stdout)
            print(result.stderr, file=sys.stderr)
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
