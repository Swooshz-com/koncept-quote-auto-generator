#!/usr/bin/env python3
"""Guard against source-code pricing-reference semantic hardcoding."""

from __future__ import annotations

import re
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]

PRODUCTION_GLOBS = (
    "scripts/*.py",
    "webapp/*.py",
    "webapp/static/*.js",
)

EXCLUDED_FILES = {
    Path("scripts/validate_dynamic_pricing_reference_rules.py"),
}

BANNED_LITERAL_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("hardcoded semantic group constant", re.compile(r"\bSEMANTIC_GROUPS\b")),
    ("hardcoded family expansion constant", re.compile(r"\bEXPANDED_TERMS_BY_FAMILY\b")),
    ("hardcoded catalog object-family token map", re.compile(r"\bCATALOG_OBJECT_FAMILY_TOKENS\b")),
    ("hardcoded broad catalog-family set", re.compile(r"\bCATALOG_BROAD_MATCH_FAMILIES\b")),
    ("hardcoded structural quantity term set", re.compile(r"\bSUSPICIOUS_LINEAR_STRUCTURE_TERMS\b")),
    ("hardcoded object preference branch", re.compile(r"\bobject_preferences\b")),
    ("hardcoded attribute preference branch", re.compile(r"\battribute_preferences\b")),
    ("runtime family-overlap scoring", re.compile(r"\bfamily_overlap\b")),
    ("runtime family-context scoring", re.compile(r"\bfamily_context_match\b")),
    ("runtime text-to-family classifier", re.compile(r"\bcatalog_object_families\b")),
    ("family-specific tied catalog resolver", re.compile(r"\bresolve_tied_catalog_family_item\b")),
    (
        "hardcoded pricing synonym alias",
        re.compile(
            r"""["'](?:canop(?:y|ies)|plinths?|screens?|monitors?|displays?|tvs?|televisions?)["']\s*:\s*["'](?:fascia|counter|display)["']""",
            re.IGNORECASE,
        ),
    ),
)


def iter_production_files() -> list[Path]:
    files: set[Path] = set()
    for pattern in PRODUCTION_GLOBS:
        for path in PROJECT_ROOT.glob(pattern):
            if not path.is_file():
                continue
            relative = path.relative_to(PROJECT_ROOT)
            if relative in EXCLUDED_FILES:
                continue
            if path.name.startswith("test_"):
                continue
            files.add(path)
    return sorted(files)


def main() -> int:
    errors: list[str] = []
    for path in iter_production_files():
        relative = path.relative_to(PROJECT_ROOT)
        text = path.read_text(encoding="utf-8", errors="replace")
        for line_number, line in enumerate(text.splitlines(), start=1):
            for label, pattern in BANNED_LITERAL_PATTERNS:
                if pattern.search(line):
                    errors.append(f"{relative}:{line_number}: {label}")

    if errors:
        print("Pricing reference matching must stay data-driven; remove source-code semantic hardcoding:")
        for error in errors:
            print(f"- {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
