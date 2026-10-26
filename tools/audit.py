"""CLI wrapper for office_builder.audit.

Usage: python -m tools.audit file1.pptx file2.docx file3.xlsx
Exit code 0 only if every file passes.
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from office_builder.audit import audit_file  # noqa: E402


def main(argv: list[str]) -> int:
    if not argv:
        print("Usage: python -m tools.audit <file> [file ...]")
        return 2
    failed = False
    for arg in argv:
        report = audit_file(arg)
        print(json.dumps(report, indent=2, ensure_ascii=False))
        if report["verdict"] != "PASS":
            failed = True
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
