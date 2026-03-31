"""
api_reliability.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · Fault-Tolerant API Reliability Layer
────────────────────────────────────────────────────────────────────────────

DESIGN PRINCIPLE
    APIs are OPTIONAL enhancements. The system NEVER depends on them.
    Every call is wrapped, timed, retried, cached, and logged.
    If all APIs go down → system runs entirely on RDKit local computation.

PARTS IMPLEMENTED
    Part 1  — Safe API Wrapper (safe_api_call)
    Part 2  — Timeout + Retry Logic (with_retry)
    Part 3  — Fallback System (per-API fallback data)
    Part 4  — 24h Caching Layer (st.cache_data TTL)
    Part 5  — API Health Tracker (dynamic status monitoring)
    Part 6  — Non-Blocking Execution (button-gated calls)
    Part 7  — Partial Success Handling (collect + show what works)
    Part 8  — Rate Limit Protection (per-domain spacing)
    Part 9  — Unified Response Format
    Part 10 — Logging System (session-state audit log)
────────────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

import time
import hashlib
import threading
import logging
from typing import Any, Callable, Optional

import streamlit as st

# ─── Optional deps ────────────────────────────────────────────────────────────
try:
    import requests as _req
    _REQ_OK = True
except ImportError:
    _req = None          # type: ignore[assignment]
    _REQ_OK = False

# ─── Internal logger (never surfaces raw errors to user) ──────────────────────
_LOG = logging.getLogger("chemofilter.api")
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s")

# ─────────────────────────────────────────────────────────────────────────────
# PART 9 — UNIFIED RESPONSE FORMAT
# Every public function in this module returns this shape.
# ─────────────────────────────────────────────────────────────────────────────

def _ok(source: str, data: Any) -> dict:
    return {"status": "success", "data": data, "source": source,
            "error": None, "ts": int(time.time())}

def _fail(source: str, error: str, data: Any = None) -> dict:
    return {"status": "failed", "data": data, "source": source,
            "error": str(error)[:300], "ts": int(time.time())}

def _cached(source: str, data: Any) -> dict:
    return {"status": "success", "data": data, "source": source,
            "error": None, "ts": int(time.time()), "_from_cache": True}


# ─────────────────────────────────────────────────────────────────────────────
# PART 10 — LOGGING SYSTEM
# Writes to session_state._api_log list (never shown raw to user).
# ─────────────────────────────────────────────────────────────────────────────

def _audit(source: str, event: str, detail: str = ""):
    """Append structured event to session-state audit log."""
    if "_api_log" not in st.session_state:
        st.session_state["_api_log"] = []
    entry = {"ts": time.strftime("%H:%M:%S"), "api": source,
             "event": event, "detail": detail[:200]}
    st.session_state["_api_log"].append(entry)
    _LOG.info("[%s] %s — %s", source, event, detail[:120])


# ─────────────────────────────────────────────────────────────────────────────
# PART 5 — API HEALTH TRACKER
# Updated dynamically on each call. Never blocks UI.
# ─────────────────────────────────────────────────────────────────────────────

# Threshold (seconds) beyond which a response is considered "slow"
_SLOW_THRESHOLD = 2.5

KNOWN_APIS = ["pubchem", "chembl", "pdb", "openfda", "uniprot",
              "kegg", "europe_pmc", "clinicaltrials", "semantic_scholar",
              "nci_cactus", "gtopdb", "unichem"]

def _init_health():
    if "_api_health" not in st.session_state:
        st.session_state["_api_health"] = {
            k: {"status": "unknown", "latency": None, "last_check": 0,
                "fail_count": 0, "success_count": 0}
            for k in KNOWN_APIS
        }

def _update_health(source: str, status: str, latency: float | None = None):
    _init_health()
    h = st.session_state["_api_health"].setdefault(
        source, {"status": "unknown", "latency": None, "last_check": 0,
                 "fail_count": 0, "success_count": 0})
    h["last_check"] = int(time.time())
    h["latency"] = latency
    if status == "success":
        if latency and latency > _SLOW_THRESHOLD:
            h["status"] = "slow"
        else:
            h["status"] = "healthy"
        h["success_count"] = h.get("success_count", 0) + 1
        h["fail_count"] = 0          # reset on success
    elif status == "failed":
        h["fail_count"] = h.get("fail_count", 0) + 1
        h["status"] = "down" if h["fail_count"] >= 2 else "degraded"
    elif status == "timeout":
        h["status"] = "slow"
        h["fail_count"] = h.get("fail_count", 0) + 1

def get_health(source: str) -> dict:
    _init_health()
    return st.session_state["_api_health"].get(
        source, {"status": "unknown", "latency": None, "fail_count": 0})

def health_badge(source: str) -> str:
    """Return an emoji badge string for UI display."""
    s = get_health(source)["status"]
    return {"healthy": "🟢 Healthy", "slow": "🟡 Slow",
            "degraded": "🟠 Degraded", "down": "🔴 Down",
            "unknown": "⚪ Not checked"}.get(s, "⚪ Unknown")

def render_api_health_panel():
    """Render the compact API health dashboard panel."""
    _init_health()
    health = st.session_state.get("_api_health", {})
    if not health:
        st.caption("No API health data yet.")
        return
    cols = st.columns(4)
    api_list = [k for k in KNOWN_APIS if k in health]
    for i, key in enumerate(api_list):
        h = health[key]
        badge = health_badge(key)
        lat = f"{h['latency']:.2f}s" if h.get("latency") else "—"
        with cols[i % 4]:
            st.markdown(
                f'<div style="background:rgba(0,0,0,0.2);border-radius:8px;'
                f'padding:8px 10px;margin:3px 0;font-size:.72rem;border:1px solid rgba(255,255,255,.06)">'
                f'<b style="color:#c8deff">{key.upper()}</b><br>'
                f'{badge} · {lat}</div>',
                unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# PART 8 — RATE LIMIT PROTECTION
# Per-domain last-request timestamp prevents hammering.
# ─────────────────────────────────────────────────────────────────────────────

_DOMAIN_LOCK = threading.Lock()
_DOMAIN_LAST: dict[str, float] = {}
# Minimum gap (seconds) between calls to same domain
_DOMAIN_GAP = {"pubchem.ncbi.nlm.nih.gov": 0.5,
               "www.ebi.ac.uk": 0.3,
               "api.fda.gov": 0.5,
               "clinicaltrials.gov": 0.3,
               "rest.uniprot.org": 0.3,
               "cactus.nci.nih.gov": 0.5}

def _rate_limit_wait(url: str):
    """Block the minimum required gap before hitting a domain."""
    for domain, gap in _DOMAIN_GAP.items():
        if domain in url:
            with _DOMAIN_LOCK:
                last = _DOMAIN_LAST.get(domain, 0.0)
                wait = gap - (time.time() - last)
                if wait > 0:
                    time.sleep(wait)
                _DOMAIN_LAST[domain] = time.time()
            break


# ─────────────────────────────────────────────────────────────────────────────
# PART 2 — TIMEOUT + RETRY LOGIC
# ─────────────────────────────────────────────────────────────────────────────

def with_retry(fn: Callable, *args,
               retries: int = 2,
               timeout: int = 5,
               source: str = "api",
               **kwargs) -> dict:
    """
    Execute fn(*args, **kwargs) up to `retries` times.
    fn must accept a `timeout` kwarg.
    Returns unified response dict.
    """
    last_err = "Unknown error"
    for attempt in range(retries):
        t0 = time.perf_counter()
        try:
            result = fn(*args, timeout=timeout, **kwargs)
            latency = time.perf_counter() - t0
            _update_health(source, "success", latency)
            if latency > _SLOW_THRESHOLD:
                _audit(source, "slow", f"{latency:.2f}s on attempt {attempt+1}")
            return result
        except Exception as exc:
            latency = time.perf_counter() - t0
            last_err = str(exc)[:200]
            if "Timeout" in last_err or "timed out" in last_err.lower():
                _update_health(source, "timeout", latency)
                _audit(source, "timeout", f"attempt {attempt+1}/{retries} — {latency:.2f}s")
            else:
                _audit(source, "error", f"attempt {attempt+1}/{retries} — {last_err}")
            if attempt < retries - 1:
                time.sleep(0.3)   # brief back-off between retries
    _update_health(source, "failed")
    _audit(source, "failed", last_err)
    return _fail(source, last_err)


# ─────────────────────────────────────────────────────────────────────────────
# PART 1 — SAFE API WRAPPER
# Top-level shield: wraps ANY callable. Never raises.
# ─────────────────────────────────────────────────────────────────────────────

def safe_api_call(fn: Callable, *args,
                  source: str = "api",
                  fallback_data: Any = None,
                  **kwargs) -> dict:
    """
    Universal safe wrapper. Any exception → returns failure with fallback.
    Never propagates exceptions to the caller.
    """
    try:
        result = fn(*args, **kwargs)
        if isinstance(result, dict) and result.get("status") in ("success", "ok"):
            return result
        if isinstance(result, dict):
            return result
        return _ok(source, result)
    except Exception as exc:
        _audit(source, "exception", str(exc)[:200])
        _update_health(source, "failed")
        return _fail(source, str(exc), data=fallback_data)


# ─────────────────────────────────────────────────────────────────────────────
# PART 4 — STRICT CACHING LAYER
# Two-level cache: st.cache_data (cross-session) + session_state (in-session).
# ─────────────────────────────────────────────────────────────────────────────

def _cache_key(source: str, *args) -> str:
    raw = source + "".join(str(a) for a in args)
    return "_rc_" + hashlib.md5(raw.encode(), usedforsecurity=False).hexdigest()[:14]

def _session_get(key: str) -> dict | None:
    return st.session_state.get(key)

def _session_put(key: str, value: dict):
    st.session_state[key] = value

@st.cache_data(ttl=86400, show_spinner=False)
def _disk_cached_get(url: str, timeout: int = 5) -> dict:
    """Cached HTTP GET. Only calls network on cache miss."""
    if not _REQ_OK:
        return _fail("http", "requests library unavailable")
    _rate_limit_wait(url)
    try:
        t0 = time.perf_counter()
        r = _req.get(url, timeout=timeout)   # type: ignore[union-attr]
        r.raise_for_status()
        data = r.json()
        latency = time.perf_counter() - t0
        return _ok(url, {"json": data, "latency": latency})
    except Exception as exc:
        return _fail(url, str(exc))

@st.cache_data(ttl=86400, show_spinner=False)
def _disk_cached_post(url: str, body_json: str, timeout: int = 5) -> dict:
    """Cached HTTP POST (body serialised as string for hashability)."""
    if not _REQ_OK:
        return _fail("http", "requests library unavailable")
    import json as _json
    try:
        body = _json.loads(body_json)
        _rate_limit_wait(url)
        t0 = time.perf_counter()
        r = _req.post(url, json=body, timeout=timeout)  # type: ignore[union-attr]
        r.raise_for_status()
        data = r.json()
        latency = time.perf_counter() - t0
        return _ok(url, {"json": data, "latency": latency})
    except Exception as exc:
        return _fail(url, str(exc))


# ─────────────────────────────────────────────────────────────────────────────
# SAFE HTTP HELPERS (public, used by other modules)
# ─────────────────────────────────────────────────────────────────────────────

def safe_get(url: str, params: dict | None = None,
             timeout: int = 5, source: str = "http") -> dict:
    """
    Safe, cached GET. Returns unified response dict.
    Handles timeouts, connection errors, and bad status codes.
    """
    full_url = url
    if params:
        import urllib.parse
        full_url = url + "?" + urllib.parse.urlencode(params)
    # Check session cache first (fastest)
    ck = _cache_key(source, full_url)
    hit = _session_get(ck)
    if hit and hit.get("status") == "success":
        return hit
    result = _disk_cached_get(full_url, timeout=timeout)
    if result.get("status") == "success":
        data = result["data"].get("json", {})
        latency = result["data"].get("latency", 0)
        _update_health(source, "success", latency)
        out = _ok(source, data)
        _session_put(ck, out)
        return out
    _update_health(source, "failed")
    _audit(source, "get_failed", result.get("error", ""))
    return _fail(source, result.get("error", "Request failed"))


def safe_post(url: str, body: dict,
              timeout: int = 5, source: str = "http") -> dict:
    """Safe, cached POST."""
    import json as _json
    body_str = _json.dumps(body, sort_keys=True)
    ck = _cache_key(source, url, body_str)
    hit = _session_get(ck)
    if hit and hit.get("status") == "success":
        return hit
    result = _disk_cached_post(url, body_str, timeout=timeout)
    if result.get("status") == "success":
        data = result["data"].get("json", {})
        latency = result["data"].get("latency", 0)
        _update_health(source, "success", latency)
        out = _ok(source, data)
        _session_put(ck, out)
        return out
    _update_health(source, "failed")
    _audit(source, "post_failed", result.get("error", ""))
    return _fail(source, result.get("error", "Request failed"))


# ─────────────────────────────────────────────────────────────────────────────
# PART 3 — FALLBACK SYSTEM
# Local RDKit-based fallbacks for every core API.
# ─────────────────────────────────────────────────────────────────────────────

def _rdkit_fallback(smiles: str) -> dict:
    """Compute basic properties locally when PubChem is down."""
    try:
        from rdkit import Chem
        from rdkit.Chem import Descriptors, rdMolDescriptors
        mol = Chem.MolFromSmiles(smiles)
        if mol is None:
            return {}
        return {
            "mw":               round(Descriptors.MolWt(mol), 2),
            "logp":             round(Descriptors.MolLogP(mol), 2),
            "tpsa":             round(Descriptors.TPSA(mol), 2),
            "hbd":              rdMolDescriptors.CalcNumHBD(mol),
            "hba":              rdMolDescriptors.CalcNumHBA(mol),
            "rotatable_bonds":  rdMolDescriptors.CalcNumRotatableBonds(mol),
            "source":           "RDKit (local fallback — PubChem unavailable)",
        }
    except Exception:
        return {}

FALLBACKS: dict[str, Callable] = {
    "pubchem":          lambda smiles="", **_: _ok("pubchem_fallback",
                            {**_rdkit_fallback(smiles),
                             "note": "PubChem unavailable — showing local RDKit properties"}),
    "chembl":           lambda **_: _ok("chembl_fallback",
                            {"note": "No bioactivity data available — ChEMBL offline",
                             "bioactivities": [], "assay_count": 0}),
    "pdb":              lambda **_: _ok("pdb_fallback",
                            {"note": "No 3D structure found — PDB offline", "entries": []}),
    "openfda":          lambda **_: _ok("openfda_fallback",
                            {"note": "FDA label data unavailable offline"}),
    "uniprot":          lambda **_: _ok("uniprot_fallback",
                            {"note": "Protein data unavailable — UniProt offline",
                             "entries": []}),
    "kegg":             lambda **_: _ok("kegg_fallback",
                            {"note": "Pathway data unavailable — KEGG offline",
                             "compounds": []}),
    "europe_pmc":       lambda **_: _ok("europe_pmc_fallback",
                            {"note": "Literature data unavailable — Europe PMC offline",
                             "papers": [], "hit_count": 0}),
    "clinicaltrials":   lambda **_: _ok("clinicaltrials_fallback",
                            {"note": "Clinical trial data unavailable — offline",
                             "trials": []}),
    "semantic_scholar": lambda **_: _ok("semantic_scholar_fallback",
                            {"note": "Academic papers unavailable — offline",
                             "papers": []}),
}

def get_fallback(api_key: str, smiles: str = "", **kwargs) -> dict:
    """Return the registered fallback for an API. Always succeeds."""
    fn = FALLBACKS.get(api_key)
    if fn:
        try:
            return fn(smiles=smiles, **kwargs)
        except Exception:
            pass
    return _ok(f"{api_key}_fallback",
               {"note": f"{api_key} unavailable — no data to display"})


# ─────────────────────────────────────────────────────────────────────────────
# PART 7 — PARTIAL SUCCESS HANDLING
# Collect results from multiple APIs. Show what worked, fall back on rest.
# ─────────────────────────────────────────────────────────────────────────────

def fetch_multiple(api_calls: dict[str, Callable],
                   smiles: str = "",
                   show_status: bool = True) -> dict[str, dict]:
    """
    Run multiple API fetchers. 
    api_calls: {api_key: callable_returning_unified_dict}

    Returns dict of results. Failed APIs replaced by fallbacks.
    NEVER raises. Shows partial results.
    """
    results: dict[str, dict] = {}
    succeeded, failed = [], []

    for key, fn in api_calls.items():
        try:
            res = safe_api_call(fn, source=key, fallback_data=None)
            if res.get("status") == "success":
                results[key] = res
                succeeded.append(key)
            else:
                results[key] = get_fallback(key, smiles=smiles)
                failed.append(key)
        except Exception as exc:
            _audit(key, "fetch_multiple_error", str(exc))
            results[key] = get_fallback(key, smiles=smiles)
            failed.append(key)

    if show_status and (succeeded or failed):
        _render_partial_status(succeeded, failed)

    return results


def _render_partial_status(succeeded: list[str], failed: list[str]):
    """Render compact partial-success summary in UI."""
    parts = []
    if succeeded:
        parts.append(
            f'<span style="color:#4ade80;font-size:.72rem">'
            f'🟢 {len(succeeded)} API{"s" if len(succeeded)!=1 else ""} OK: '
            f'{", ".join(succeeded)}</span>')
    if failed:
        parts.append(
            f'<span style="color:#f87171;font-size:.72rem">'
            f'🔴 {len(failed)} using fallback: {", ".join(failed)}</span>')
    st.markdown(
        f'<div style="display:flex;gap:16px;flex-wrap:wrap;padding:6px 0">'
        + " ".join(parts) + "</div>",
        unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# PART 6 — NON-BLOCKING EXECUTION HELPER
# Gate expensive API calls behind a st.button so page load stays fast.
# ─────────────────────────────────────────────────────────────────────────────

def api_fetch_button(label: str,
                     api_key: str,
                     fetch_fn: Callable,
                     smiles: str = "",
                     query: str = "",
                     key_suffix: str = "") -> dict | None:
    """
    Renders a button. Only calls fetch_fn on click.
    Returns result dict or None if not yet clicked.
    UI stays fully responsive regardless.
    """
    btn_key = f"_apibtn_{api_key}_{key_suffix}"
    result_key = f"_apiresult_{api_key}_{key_suffix}"

    # Already fetched this session → show result without button
    cached = st.session_state.get(result_key)
    if cached:
        return cached

    if st.button(label, key=btn_key, use_container_width=False):
        with st.spinner(f"Fetching from {api_key}…"):
            result = safe_api_call(fetch_fn, source=api_key)
            if result.get("status") != "success":
                result = get_fallback(api_key, smiles=smiles)
            st.session_state[result_key] = result
        return result
    return None


# ─────────────────────────────────────────────────────────────────────────────
# PUBLIC AUDIT LOG VIEWER (for debug / admin use in UI)
# ─────────────────────────────────────────────────────────────────────────────

def render_audit_log(max_entries: int = 30):
    """Render the last N API audit log entries."""
    log = st.session_state.get("_api_log", [])
    if not log:
        st.caption("No API calls logged yet.")
        return
    st.markdown(
        '<div style="font-family:monospace;font-size:.68rem;'
        'background:rgba(0,0,0,.3);border-radius:8px;padding:10px;'
        'max-height:260px;overflow-y:auto">',
        unsafe_allow_html=True)
    for entry in reversed(log[-max_entries:]):
        color = {"failed": "#f87171", "timeout": "#fbbf24",
                 "slow": "#fbbf24", "error": "#f87171",
                 "exception": "#f87171"}.get(entry["event"], "#94a3b8")
        st.markdown(
            f'<div style="padding:1px 0;color:{color}">'
            f'[{entry["ts"]}] <b>{entry["api"]}</b> '
            f'{entry["event"]} — {entry.get("detail","")}'
            f'</div>',
            unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# CONVENIENCE: Enhanced versions of the api_manager fetch functions
# These replace direct calls and route through the reliability layer.
# ─────────────────────────────────────────────────────────────────────────────

def reliable_pubchem(smiles: str) -> dict:
    """PubChem via reliability layer. Falls back to local RDKit on failure."""
    try:
        from api_manager import _fetch_pubchem
        result = safe_api_call(_fetch_pubchem, smiles, source="pubchem")
        if result.get("status") in ("success", "ok"):
            return result
    except Exception:
        pass
    _audit("pubchem", "fallback_used", "returning local RDKit properties")
    return get_fallback("pubchem", smiles=smiles)

def reliable_chembl(smiles: str) -> dict:
    """ChEMBL via reliability layer."""
    try:
        from api_manager import _fetch_chembl
        result = safe_api_call(_fetch_chembl, smiles, source="chembl")
        if result.get("status") in ("success", "ok"):
            return result
    except Exception:
        pass
    return get_fallback("chembl", smiles=smiles)

def reliable_pdb(query: str) -> dict:
    """PDB via reliability layer."""
    try:
        from api_manager import _fetch_pdb
        result = safe_api_call(_fetch_pdb, query, source="pdb")
        if result.get("status") in ("success", "ok"):
            return result
    except Exception:
        pass
    return get_fallback("pdb")

def reliable_fetch(api_key: str, smiles: str = "", query: str = "") -> dict:
    """
    Generic reliable fetch — looks up api_manager dispatch table,
    wraps in reliability layer, falls back gracefully.
    """
    try:
        from api_manager import fetch_api as _orig_fetch
        result = safe_api_call(_orig_fetch, api_key, smiles=smiles,
                                query=query, source=api_key)
        if result.get("status") in ("success", "ok"):
            return result
    except Exception as exc:
        _audit(api_key, "reliable_fetch_error", str(exc))
    return get_fallback(api_key, smiles=smiles)


# ─────────────────────────────────────────────────────────────────────────────
# TESTING HELPERS — simulate failure scenarios
# ─────────────────────────────────────────────────────────────────────────────

def simulate_api_success(source: str = "pubchem") -> dict:
    """Simulate a successful API call for testing."""
    _update_health(source, "success", 0.4)
    return _ok(source, {"test": "success", "mw": 180.16, "logp": 1.19})

def simulate_api_timeout(source: str = "chembl") -> dict:
    """Simulate a timeout for testing."""
    _update_health(source, "timeout", 5.01)
    _audit(source, "timeout", "simulated timeout after 5.01s")
    return _fail(source, f"Timeout after 5s — {source} unavailable")

def simulate_api_failure(source: str = "pdb") -> dict:
    """Simulate a hard failure for testing."""
    _update_health(source, "failed")
    _audit(source, "failed", "simulated connection error")
    return get_fallback(source)

def simulate_no_internet() -> dict[str, dict]:
    """Simulate complete offline mode. Returns fallbacks for all APIs."""
    results = {}
    for key in KNOWN_APIS:
        _update_health(key, "failed")
        results[key] = get_fallback(key)
    _audit("system", "offline_mode", "all APIs returning local fallbacks")
    return results
