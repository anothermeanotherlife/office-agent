"""Build the example Q3 report through the orchestrator.

This is the reference run: it builds docx/xlsx/pptx, adds Fade click
beats to the PPTX, audits all three with code, and keeps the best
version. Run from the repo root:

    python examples/build_example.py
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from agents.orchestrator import run_task  # noqa: E402

ROOT = Path(__file__).resolve().parent
DATA = json.loads((ROOT / "data.json").read_text(encoding="utf-8"))

if __name__ == "__main__":
    task = {"title": DATA["title"], "data": DATA}
    result = run_task(task, ROOT / "output" / ".." / "runs" / "example-q3", )
    # Also copy the best version to examples/output for easy inspection.
    out = ROOT / "output"
    out.mkdir(parents=True, exist_ok=True)
    for f in result.best_dir.iterdir():
        (out / f.name).write_bytes(f.read_bytes())
    print(f"verdict={result.verdict} best_score={result.best_score} rounds={result.rounds}")
    print(f"best_dir={result.best_dir} output={out}")
    for line in result.decision_log:
        print(line)
    for report in result.reports:
        print(report["artifact"], report["verdict"], report["score"],
              "blockers:", len(report["blockers"]), "majors:", len(report["majors"]))
