"""
engine_selector.py
Automatically selects the most appropriate engine based on query/context type.
"""

from typing import Optional

# ── Keyword → engine mapping ─────────────────────────────────────────────────
_ROUTING_TABLE = {
    # Drug / ADMET lookups
    "drug": "master_drug_atlas",
    "admet": "aether_engine",
    "absorption": "aether_engine",
    "distribution": "aether_engine",
    "metabolism": "omnipotent_engine",
    "excretion": "aether_engine",
    "toxicity": "aether_engine",
    "bbb": "aether_engine",
    "blood-brain": "aether_engine",
    # Reaction / reactivity
    "reaction": "omnipotent_engine",
    "reactivity": "omnipotent_engine",
    "metabolic": "omnipotent_engine",
    "ddi": "omnipotent_engine",
    # Deep / quantum computation
    "quantum": "quantum_accuracy_engine",
    "accuracy": "quantum_accuracy_engine",
    "similarity": "quantum_accuracy_engine",
    "fingerprint": "quantum_accuracy_engine",
    # Chemical intelligence
    "chemical": "chemical_intelligence_db",
    "scaffold": "chemical_intelligence_db",
    "filter": "chemical_intelligence_db",
    # Specific engines by capability
    "logp": "quantum_accuracy_engine",
    "mw": "omega_engine",
    "molecular weight": "omega_engine",
    "bioavailability": "nova_engine",
    "herg": "nova_engine",
    "cytotoxicity": "xenon_engine",
    "mutagenicity": "xenon_engine",
    "celestial": "celestial_engine",
    "tissue": "celestial_engine",
}

# Priority order when multiple engines match
_ENGINE_PRIORITY = [
    "quantum_accuracy_engine",
    "aether_engine",
    "omnipotent_engine",
    "nova_engine",
    "xenon_engine",
    "celestial_engine",
    "omega_engine",
    "chemical_intelligence_db",
    "master_drug_atlas",
]

# Default fallback engine
DEFAULT_ENGINE = "quantum_accuracy_engine"


def select_engine(query: str, context: Optional[str] = None) -> str:
    """
    Given a free-text query (and optional context tag), return the best engine name.
    Returns engine name string.
    """
    text = (query + " " + (context or "")).lower()
    matched = set()
    for keyword, engine in _ROUTING_TABLE.items():
        if keyword in text:
            matched.add(engine)
    if not matched:
        return DEFAULT_ENGINE
    # Pick highest-priority match
    for engine in _ENGINE_PRIORITY:
        if engine in matched:
            return engine
    return next(iter(matched))


def select_engines_multi(query: str, top_n: int = 3) -> list:
    """
    Return up to top_n engines relevant to the query, in priority order.
    Useful for running ensemble analysis.
    """
    text = query.lower()
    matched = set()
    for keyword, engine in _ROUTING_TABLE.items():
        if keyword in text:
            matched.add(engine)
    if not matched:
        matched = {DEFAULT_ENGINE}
    ordered = [e for e in _ENGINE_PRIORITY if e in matched]
    # Append any not in priority list
    for e in matched:
        if e not in ordered:
            ordered.append(e)
    return ordered[:top_n]


def get_routing_table() -> dict:
    """Expose routing table for UI/debug display."""
    return dict(_ROUTING_TABLE)
