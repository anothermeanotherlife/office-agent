"""Visual QA: render each file to its own PDF, then to PNGs for viewing.

Usage: python -m tools.visual_qa <outdir> <file> [file ...]
Unlike the old render_check (exit code only), this produces one PNG per
page/slide so a human or a vision agent can actually LOOK at the file.
It does not claim the PNGs were viewed — viewing is a separate step,
documented in agents/visual_auditor.md.
"""
import subprocess
import sys
from pathlib import Path


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print("Usage: python -m tools.visual_qa <outdir> <file> [file ...]")
        return 2
    outdir = Path(argv[0]); outdir.mkdir(parents=True, exist_ok=True)
    for arg in argv[1:]:
        stem = Path(arg).stem + "-" + Path(arg).suffix.lstrip(".")
        sub = outdir / stem; sub.mkdir(parents=True, exist_ok=True)
        r = subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(sub), arg],
                           capture_output=True, text=True, timeout=180)
        if r.returncode != 0:
            print(arg, "convert FAILED"); print(r.stdout); print(r.stderr, file=sys.stderr); return 1
        pdf = sub / (Path(arg).stem + ".pdf")
        r = subprocess.run(["pdftoppm", "-png", "-r", "130", str(pdf), str(sub / "page")],
                           capture_output=True, text=True, timeout=180)
        if r.returncode != 0:
            print(arg, "pdftoppm FAILED"); return 1
        print(arg, "->", sorted(p.name for p in sub.glob("page-*.png")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
