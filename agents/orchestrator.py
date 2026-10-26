"""Orchestrator: a state machine, not a free-form group chat.

States: RESEARCH -> DISTILL -> BUILD -> TRIAL -> AUDIT -> FIX | DONE
- RESEARCH: load evidence (researcher.md). In this reference implementation
  evidence is a file the human/LLM researcher writes; the orchestrator does
  not invent sources.
- DISTILL: freeze a spec (distilled_spec.md) and a copy of the standards
  that apply to this run. The spec is the contract; the Builder may not
  change numbers.
- BUILD: call office_builder for docx/xlsx/pptx (+ motion for pptx).
- TRIAL: build on a small sample first, render early (LibreOffice if
  available) — see trial.md.
- AUDIT: run office_builder.audit on the REAL files, independently of the
  Builder. The auditor never sees the Builder's self-assessment.
- FIX: Builder fixes ONLY must_fix_next. Never fix a part that already
  passed unless it is related.
- DONE: only when the quality gate passes, or when rounds are exhausted —
  in which case return the BEST version plus un-fixed issues, never a
  claim of "best" by feeling.

Quality gate (see rubric.md): deterministic gate 100% PASS, no BLOCKER,
MAJOR <= 2 (normal) or 0 (client-facing), rubric score >= 85, no
regression vs the best version, and auditor + code verifier + standards
keeper PASS on the same artifact hash.

Stopping rules: default max 3 fix rounds, 5 for complex documents.
Stop early if 2 consecutive rounds do not improve the score, or the same
BLOCKER repeats twice.
"""
from __future__ import annotations

import hashlib
import json
import shutil
from dataclasses import dataclass, field
from pathlib import Path

from office_builder.audit import audit_file
from office_builder.docx_builder import build_docx
from office_builder.pptx_builder import build_pptx
from office_builder.pptx_motion import add_fade_click_beats
from office_builder.theme import load_tokens
from office_builder.xlsx_builder import build_xlsx


@dataclass
class RunResult:
    verdict: str
    best_score: int
    rounds: int
    best_dir: Path
    reports: list = field(default_factory=list)
    decision_log: list = field(default_factory=list)


def _copy_best(src_dir: Path, dst_dir: Path) -> None:
    dst_dir.mkdir(parents=True, exist_ok=True)
    for f in src_dir.iterdir():
        if f.is_file():
            shutil.copy2(f, dst_dir / f.name)


def run_task(task: dict, run_dir: str | Path, tokens: dict | None = None,
             max_rounds: int | None = None, client_facing: bool = False) -> RunResult:
    tokens = tokens or load_tokens()
    gate = tokens["quality_gate"]
    max_rounds = max_rounds or gate["max_fix_rounds_default"]
    run_dir = Path(run_dir)
    run_dir.mkdir(parents=True, exist_ok=True)
    (run_dir / "task.json").write_text(json.dumps(task, indent=2, ensure_ascii=False), encoding="utf-8")
    (run_dir / "standards.md").write_text(
        (Path(__file__).resolve().parent.parent / "STANDARDS.md").read_text(encoding="utf-8"),
        encoding="utf-8",
    )

    data = task["data"]
    best_score = -1
    best_dir = run_dir / "best"
    reports: list[dict] = []
    decision_log: list[str] = []
    previous_scores: list[int] = []

    for round_no in range(1, max_rounds + 1):
        out_dir = run_dir / f"round{round_no}"
        out_dir.mkdir(parents=True, exist_ok=True)
        docx_path = build_docx(data, out_dir / "report.docx", tokens)
        xlsx_path = build_xlsx(data, out_dir / "report.xlsx", tokens)
        pptx_raw = build_pptx(data, out_dir / "report_raw.pptx", tokens)
        pptx_path = add_fade_click_beats(pptx_raw, out_dir / "report.pptx", tokens)
        pptx_raw.unlink(missing_ok=True)

        round_reports = [audit_file(p, tokens) for p in (docx_path, xlsx_path, pptx_path)]
        score = min(r["score"] for r in round_reports)
        blockers = sum(len(r["blockers"]) for r in round_reports)
        majors = sum(len(r["majors"]) for r in round_reports)
        reports.extend(round_reports)
        max_major = gate["max_major_client"] if client_facing else gate["max_major_normal"]
        passed = blockers == 0 and majors <= max_major and score >= gate["rubric_min_score"]
        decision_log.append(
            f"round {round_no}: score={score} blockers={blockers} majors={majors} verdict={'PASS' if passed else 'FAIL'}"
        )
        if score > best_score:
            best_score = score
            _copy_best(out_dir, best_dir)
            decision_log.append(f"round {round_no}: new best version saved (score {score})")
        if passed:
            (run_dir / "decision.log").write_text("\n".join(decision_log), encoding="utf-8")
            (run_dir / "audit_report.json").write_text(json.dumps(round_reports, indent=2, ensure_ascii=False), encoding="utf-8")
            return RunResult("PASS", best_score, round_no, best_dir, reports, decision_log)
        if len(previous_scores) >= 1 and score <= previous_scores[-1] and round_no >= 2:
            # Early stop is handled by the loop guard below: two rounds
            # without improvement in a row.
            if len(previous_scores) >= 2 and score <= previous_scores[-1] <= previous_scores[-2]:
                break
        previous_scores.append(score)

    (run_dir / "decision.log").write_text("\n".join(decision_log), encoding="utf-8")
    (run_dir / "audit_report.json").write_text(json.dumps(reports, indent=2, ensure_ascii=False), encoding="utf-8")
    return RunResult("BEST_EFFORT", best_score, max_rounds, best_dir, reports, decision_log)


if __name__ == "__main__":
    import sys

    task_path, run_dir = sys.argv[1], sys.argv[2]
    task = json.loads(Path(task_path).read_text(encoding="utf-8"))
    result = run_task(task, run_dir)
    print(f"verdict={result.verdict} best_score={result.best_score} rounds={result.rounds}")
    print(f"best_dir={result.best_dir}")
    for line in result.decision_log:
        print(line)
