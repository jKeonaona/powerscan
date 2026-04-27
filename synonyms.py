"""Domain synonym dictionary for Workspace keyword expansion.

expand_query_terms() takes a raw query string and returns two sets of lowercase
search terms: direct terms (from the query itself) and synonym terms (expanded via
this dictionary).  Direct matches score higher than synonym matches during context
assembly.
"""

import re

CONSTRUCTION_SYNONYMS = {
    "bonding": ["bond", "surety", "payment bond", "performance bond", "warranty bond", "bid bond"],
    "abrasive_blast": ["sandblast", "blast cleaning", "dry blast", "abrasive blasting", "grit blast"],
    "wet_blast": ["water blast", "hydroblast", "pressure wash", "high pressure water washing", "high-pressure washing"],
    "coating": ["paint", "primer", "intermediate coat", "topcoat", "finish coat", "coating system"],
    "containment": ["enclosure", "shrouding", "encapsulation", "containment system"],
    "lead": ["LBP", "lead-bearing paint", "lead-based paint", "lead paint"],
    "lead_abatement": ["lead removal", "lead mitigation", "lead remediation", "lead abatement"],
    "DBE": ["disadvantaged business enterprise", "DBE participation", "DBE requirement"],
    "DVBE": ["disabled veteran business", "disabled veteran-owned business", "disabled veteran"],
    "SB": ["small business", "small business participation", "small business requirement"],
    "DIR": ["department of industrial relations", "prevailing wage", "DIR compliance"],
    "prevailing_wage": ["prevailing wage rate", "wage determination", "DIR rate", "PWR"],
    "OSHA": ["occupational safety", "workplace safety", "safety requirements"],
    "Cal_OSHA": ["California OSHA", "Cal/OSHA", "California occupational safety", "Title 8"],
    "SWPPP": ["stormwater pollution prevention plan", "stormwater plan", "stormwater"],
    "BMP": ["best management practice", "best management practices", "stormwater BMP"],
    "NPDES": ["national pollutant discharge", "NPDES permit", "stormwater permit"],
    "takeoff": ["take-off", "quantity takeoff", "quantity survey", "material takeoff"],
    "estimate": ["bid", "proposal", "quotation", "cost estimate"],
    "manhour": ["man-hour", "labor hour", "labor hours", "work hours"],
    "SFHR": ["square feet per hour", "production rate", "SF/HR", "SF per hour", "sq ft per hour"],
    "specification": ["spec", "spec book", "specifications", "project specs", "technical spec"],
    "as_built": ["as-built", "asbuilt", "record drawing", "final drawing"],
    "submittal": ["submittals", "shop drawing", "product data", "submittal package"],
    "insurance": ["general liability", "workers comp", "workers compensation", "liability insurance", "certificate of insurance", "COI"],
}


def _build_lookup() -> dict[str, set[str]]:
    """Build a flat dict mapping every term (lowercase) → set of all related terms."""
    lookup: dict[str, set[str]] = {}
    for key, synonyms in CONSTRUCTION_SYNONYMS.items():
        canonical = key.replace("_", " ").lower()
        group: set[str] = {canonical} | {s.lower() for s in synonyms}
        for term in group:
            lookup[term] = group
    return lookup


_LOOKUP = _build_lookup()


def expand_query_terms(query: str) -> tuple[set[str], set[str]]:
    """Return (direct_terms, synonym_terms) derived from the query.

    direct_terms  — individual words and bi-grams from the query (filtered to len >= 3).
    synonym_terms — terms from the synonym dictionary that relate to any direct term,
                    minus those already in direct_terms.

    Matching rules (conservative — avoids generic words like "requirements" pulling in
    unrelated synonym groups):
      1. Exact match: query term == lookup key
      2. Single-word lookup key only: query term is contained in the key OR vice versa
         (e.g. "blast" hits "abrasive_blast" group via key "abrasive blast")
      Multi-word keys require exact match only, preventing partial hits.
    """
    raw_words = [w for w in re.split(r"\W+", query.lower()) if len(w) >= 3]
    bigrams = [f"{raw_words[i]} {raw_words[i + 1]}" for i in range(len(raw_words) - 1)]
    direct: set[str] = set(raw_words) | set(bigrams)

    synonyms: set[str] = set()
    for term in direct:
        if term in _LOOKUP:
            synonyms |= _LOOKUP[term]
            continue
        # Substring match only against single-word lookup keys to avoid false positives
        for key, group in _LOOKUP.items():
            if " " not in key and (term in key or key in term):
                synonyms |= group
                break

    return direct, synonyms - direct


STOPWORDS: set[str] = {
    "a", "an", "the", "is", "are", "was", "were", "be", "been", "being",
    "have", "has", "had", "do", "does", "did", "will", "would", "should",
    "could", "what", "who", "where", "when", "why", "how", "of", "in",
    "on", "for", "to", "and", "or", "but", "by", "with", "from", "at",
    "as", "this", "that", "these", "those", "i", "you", "he", "she", "it",
    "we", "they", "me", "him", "her", "us", "them",
}
