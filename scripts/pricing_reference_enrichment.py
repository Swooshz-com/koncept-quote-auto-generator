#!/usr/bin/env python3
"""Import-time matching metadata for pricing reference catalog rows."""

from __future__ import annotations

import re
from typing import Any

import pricing_reference_cleanup

TEXT_FIXES = pricing_reference_cleanup.import_word_replacements()

STOP_TERMS = {
    "and",
    "as",
    "at",
    "by",
    "for",
    "from",
    "height",
    "including",
    "of",
    "or",
    "per",
    "proposal",
    "ready",
    "rental",
    "the",
    "with",
}

UNIT_PREFIX_RE = re.compile(
    r"^(?:m2|sqm|m\.?\s*length|m\.?\s*run|nos?\.?|no\.|sets?|lot\.?)\s+(?:of\s+|rental\s+of\s+)?",
    flags=re.IGNORECASE,
)

SEMANTIC_GROUPS: tuple[tuple[str, set[str], set[str]], ...] = (
    ("display", {"display", "monitor", "monitors", "screen", "screens", "tv", "television", "televisions"}, {"display", "monitor", "screen", "tv", "television"}),
    ("graphics", {"graphic", "graphics", "logo", "logos", "lettering", "print", "printed", "poster", "posters", "sign", "signage", "vinyl"}, {"graphic", "graphics", "logo", "lettering", "printed", "print", "signage", "vinyl"}),
    ("partition", {"partition", "partitions", "wall", "walls", "backwall", "backwalls", "enclosure", "enclosures"}, {"partition", "wall", "backwall", "enclosure", "room build"}),
    ("fascia", {"fascia", "fascias", "header", "headers", "canopy", "canopies"}, {"fascia", "header", "canopy"}),
    ("pillar", {"pillar", "pillars", "column", "columns", "support", "supports"}, {"pillar", "column", "support"}),
    ("counter", {"counter", "counters", "cabinet", "cabinets", "reception"}, {"counter", "cabinet", "reception"}),
    ("table", {"table", "tables", "desk", "desks"}, {"table", "desk"}),
    ("chair", {"chair", "chairs", "stool", "stools"}, {"chair", "stool"}),
    ("sofa", {"sofa", "sofas", "lounge"}, {"sofa", "lounge"}),
    ("plant", {"plant", "plants", "planter", "planters", "greenery"}, {"plant", "planter", "greenery"}),
    ("water", {"water", "sink", "sinks", "plumbing", "inlet", "outlet", "drainage", "tap", "taps"}, {"water", "sink", "plumbing", "inlet", "outlet", "drainage", "tap"}),
    ("beverage_service", {"coffee", "tea", "beverage", "beverages", "drink", "drinks", "barista", "consumable", "consumables"}, {"coffee", "tea", "beverage", "drink", "barista", "consumable", "supplies", "service", "package"}),
    ("socket", {"socket", "sockets", "power", "13amp"}, {"socket", "power", "13amp", "outlet"}),
    ("lighting", {"light", "lights", "lighting", "downlight", "downlights", "led", "spotlight", "spotlights"}, {"light", "lighting", "downlight", "led", "spotlight"}),
    ("rigging", {"rigging", "truss", "hoist", "hoists", "lift", "endorsement", "engineer"}, {"rigging", "truss", "hoist", "lift", "endorsement", "engineer"}),
    ("floor", {"floor", "flooring", "carpet", "platform", "platforms"}, {"floor", "flooring", "carpet", "platform"}),
)

EXPANDED_TERMS_BY_FAMILY = {
    "beverage_service": {"beverage", "drink", "service", "package"},
    "display": {"display", "monitor", "screen", "tv", "television"},
    "fascia": {"fascia", "header", "canopy"},
    "partition": {"partition", "wall", "backwall", "enclosure", "room build"},
    "pillar": {"pillar", "column", "support"},
}

TOKEN_ALIASES = {
    "lcd": "led",
    "m2": "sqm",
    "prints": "printed",
    "printing": "printed",
    **TEXT_FIXES,
}


def normalized_text(value: Any) -> str:
    text = str(value or "")
    text = text.replace("\xa0", " ")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    return re.sub(r"\s+", " ", text).strip()


def token(value: str) -> str:
    normalized = TOKEN_ALIASES.get(value.lower(), value.lower())
    if normalized.endswith("ies") and len(normalized) > 4:
        normalized = f"{normalized[:-3]}y"
    elif normalized.endswith("s") and len(normalized) > 3 and not normalized.endswith("ss"):
        normalized = normalized[:-1]
    return TOKEN_ALIASES.get(normalized, normalized)


def tokens_for_text(value: Any) -> set[str]:
    tokens: set[str] = set()
    for raw in re.findall(r"[a-z0-9]+", normalized_text(value).lower()):
        next_token = token(raw)
        if (len(next_token) <= 2 and next_token != "tv") or next_token in STOP_TERMS:
            continue
        tokens.add(next_token)
    return tokens


def add_unique(values: list[str], value: Any, *, limit: int | None = None) -> None:
    text = normalized_text(value).lower()
    if not text:
        return
    if limit is not None and len(text) > limit:
        return
    if text not in values:
        values.append(text)


def stripped_unit_phrase(value: Any) -> str:
    return normalized_text(UNIT_PREFIX_RE.sub("", normalized_text(value)))


def phrase_candidates(value: Any) -> list[str]:
    source = stripped_unit_phrase(value)
    if not source:
        return []
    phrases: list[str] = []
    for part in re.split(r"[;/,]", source):
        cleaned = stripped_unit_phrase(part)
        lowered = cleaned.lower()
        if len(tokens_for_text(lowered)) >= 2:
            add_unique(phrases, lowered, limit=80)
    lowered_source = source.lower()
    if len(tokens_for_text(lowered_source)) >= 2:
        add_unique(phrases, lowered_source, limit=100)
    return phrases


def families_for_tokens(tokens: set[str], section_tokens: set[str]) -> list[str]:
    families: list[str] = []
    all_tokens = tokens | section_tokens
    for family, source_tokens, _expanded in SEMANTIC_GROUPS:
        if all_tokens & {token(value) for value in source_tokens}:
            add_unique(families, family)
    if "led" in tokens and "display" not in families and not (section_tokens & {"av", "audio", "visual"}):
        # LED is lighting unless the AV/display context says otherwise.
        pass
    if "graphics" in families:
        for family in ("partition", "fascia"):
            if family in families:
                families.remove(family)
    if "display" in families and "lighting" in families:
        families.remove("lighting")
    return families


def terms_for_family_tokens(tokens: set[str], families: list[str]) -> list[str]:
    terms: list[str] = []
    for family in families:
        add_unique(terms, family)
        for value in sorted(EXPANDED_TERMS_BY_FAMILY.get(family, set())):
            add_unique(terms, value)
    for value in sorted(tokens):
        if value not in STOP_TERMS and not value.isdigit():
            add_unique(terms, value)
    return terms


def clean_string_list(value: Any) -> list[str]:
    if isinstance(value, list):
        raw_values = value
    else:
        raw_values = re.split(r"[|;\n]", str(value or ""))
    result: list[str] = []
    for raw in raw_values:
        add_unique(result, raw)
    return result


def enrich_pricing_reference_item(item: dict[str, Any]) -> dict[str, Any]:
    section = normalized_text(item.get("reference_section") or item.get("section"))
    section_key = section.lower()
    aliases = [
        alias
        for alias in clean_string_list(item.get("aliases"))
        if not alias.lower().startswith(section_key)
    ]
    source_values: list[Any] = [
        item.get("description"),
        item.get("unit_hint"),
        *clean_string_list(item.get("remarks")),
        *aliases,
    ]
    content_tokens: set[str] = set()
    section_tokens = tokens_for_text(section)
    for value in source_values:
        content_tokens.update(tokens_for_text(value))
    content_families = families_for_tokens(content_tokens, set())
    families = families_for_tokens(content_tokens, section_tokens)
    match_terms: list[str] = []
    for value in source_values:
        for phrase in phrase_candidates(value):
            add_unique(match_terms, phrase, limit=100)
    for term in terms_for_family_tokens(content_tokens, content_families):
        add_unique(match_terms, term, limit=80)
    if "water" in families:
        add_unique(match_terms, "water connection")
        if content_tokens & {"inlet", "outlet"}:
            for term in ("water inlet", "water outlet"):
                add_unique(match_terms, term)
        if "sink" in content_tokens:
            add_unique(match_terms, "sink connection")
        for term in ("plumbing", "drainage", "tap"):
            if term in content_tokens:
                add_unique(match_terms, term)
    if "graphics" in families:
        if content_tokens & {"printed", "print"}:
            for term in ("brand graphics", "wall graphics", "graphic panels", "printed graphics"):
                add_unique(match_terms, term)
        if "vinyl" in content_tokens:
            add_unique(match_terms, "vinyl graphics")
        if "logo" in content_tokens:
            add_unique(match_terms, "logo graphics")
        if "lettering" in content_tokens:
            add_unique(match_terms, "lettering graphics")
        if "counter" in content_tokens:
            add_unique(match_terms, "counter graphics")
    if "beverage_service" in content_families:
        for term in ("coffee tea service", "coffee tea package", "coffee tea supplies", "beverage service"):
            add_unique(match_terms, term)
    if "display" in content_families:
        for term in ("tv display", "led tv", "monitor display", "screen display", "wall mounted display"):
            add_unique(match_terms, term)

    item["match_terms"] = match_terms[:36]
    item["object_families"] = families[:12]
    return item


def enrich_pricing_reference_items(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    for item in items:
        if isinstance(item, dict):
            enrich_pricing_reference_item(item)
    return items
