from __future__ import annotations

import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
ALLOWED_PRIVATE_FILE = ROOT / "data" / "private" / "README.md"
TEXT_SUFFIXES = {".py", ".md", ".txt", ".csv", ".json", ".yml", ".yaml", ".cff"}
FORBIDDEN_PATTERNS = {
    "absolute Windows user path": re.compile(r"[A-Za-z]:\\Users\\", re.IGNORECASE),
    "email address": re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.IGNORECASE),
    "linkage secret": re.compile(r"\b[A-Z0-9_-]+-v1\|", re.IGNORECASE),
}


def main() -> int:
    errors: list[str] = []

    for path in (ROOT / "data" / "private").rglob("*"):
        if path.is_file() and path.resolve() != ALLOWED_PRIVATE_FILE.resolve():
            errors.append(f"Restricted file present: {path.relative_to(ROOT)}")

    for path in ROOT.rglob("*"):
        if not path.is_file() or ".git" in path.parts or path.suffix.lower() not in TEXT_SUFFIXES:
            continue
        text = path.read_text(encoding="utf-8", errors="ignore")
        for label, pattern in FORBIDDEN_PATTERNS.items():
            if pattern.search(text):
                errors.append(f"{label} found in {path.relative_to(ROOT)}")

    if errors:
        print("Public-release verification failed:")
        for error in errors:
            print(f"- {error}")
        return 1

    print("Public-release verification passed: no restricted files or flagged text patterns found.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
