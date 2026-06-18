#!/usr/bin/env python3
"""Data-backed import cleanup rules for pricing reference text."""

from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]
IMPORT_CLEANUP_RULES_PATH = PROJECT_ROOT / "_pricing-references" / "import-cleanup-rules.json"


def normalized_import_text(value: Any) -> str:
    text = str(value or "")
    text = text.replace("\xa0", " ")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    return re.sub(r"\s+", " ", text).strip()


@lru_cache(maxsize=1)
def import_cleanup_rules() -> dict[str, Any]:
    try:
        data = json.loads(IMPORT_CLEANUP_RULES_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def import_word_replacements() -> dict[str, str]:
    values = import_cleanup_rules().get("word_replacements")
    if not isinstance(values, dict):
        return {}
    return {
        normalized_import_text(source).lower(): normalized_import_text(target)
        for source, target in values.items()
        if normalized_import_text(source) and normalized_import_text(target)
    }


def normalize_word_slash_spacing(value: Any) -> str:
    return re.sub(r"(?<=[A-Za-z])\s*/\s*(?=[A-Za-z])", " / ", normalized_import_text(value))


def apply_import_text_cleanup(value: Any) -> str:
    text = normalize_word_slash_spacing(value)
    replacements = import_word_replacements()

    def replace_word(match: re.Match[str]) -> str:
        replacement = replacements[match.group(0).lower()]
        return replacement[:1].upper() + replacement[1:] if match.group(0)[:1].isupper() else replacement

    for source in replacements:
        text = re.sub(rf"\b{re.escape(source)}\b", replace_word, text, flags=re.IGNORECASE)

    unit_replacements = import_cleanup_rules().get("leading_unit_replacements")
    if isinstance(unit_replacements, dict):
        for source, target in unit_replacements.items():
            source_text = normalized_import_text(source)
            target_text = normalized_import_text(target)
            if source_text and target_text:
                text = re.sub(rf"(?i)^{re.escape(source_text)}\b", target_text, text)

    label_spacing = import_cleanup_rules().get("label_spacing")
    if isinstance(label_spacing, list):
        for label in label_spacing:
            label_text = normalized_import_text(label)
            if label_text:
                text = re.sub(rf"(?i)\b({re.escape(label_text)}):\s*", r"\1 ", text)

    return normalized_import_text(text)
