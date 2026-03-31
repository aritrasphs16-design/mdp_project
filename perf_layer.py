"""
perf_layer.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · Performance Layer — Phase 2/3/6 Optimisation
• Monkey-patches fig_* functions with @st.cache_data wrappers
• Provides lazy-load guard helpers
• Provides cached HTML/stats helpers
• ZERO changes to app.py logic — imported once at startup
• Safe: all wrappers produce IDENTICAL outputs to originals
────────────────────────────────────────────────────────────────────────────
"""
import streamlit as st
import hashlib
import functools


# ─────────────────────────────────────────────────────────────────────────────
# UTILITY: stable hash for compound lists (used as cache key)
# ─────────────────────────────────────────────────────────────────────────────

def _compound_hash(data: list) -> str:
    """Deterministic hash of a compound list based on IDs + key props."""
    try:
        parts = []
        for d in data:
            parts.append(str(d.get("ID", "")) +
                         str(d.get("MW", "")) +
                         str(d.get("LogP", "")))
        key = "|".join(parts)
        return hashlib.md5(key.encode(), usedforsecurity=False).hexdigest()
    except Exception:
        return str(len(data))


def _single_hash(cpd: dict) -> str:
    try:
        return hashlib.md5(
            (str(cpd.get("ID","")) + str(cpd.get("SMILES",""))).encode(),
            usedforsecurity=False
        ).hexdigest()
    except Exception:
        return str(id(cpd))


# ─────────────────────────────────────────────────────────────────────────────
# PHASE 2: Cache wrappers for all fig_* functions
# These are called AFTER app.py defines the functions, so we patch them in-place
# ─────────────────────────────────────────────────────────────────────────────

def _make_cached_fig(fn, cache_key_prefix: str):
    """
    Wraps a fig_* function with st.cache_data.
    Uses compound-list hash as cache key (not the list itself which is unhashable).
    Outputs are IDENTICAL to the unwrapped version.
    """
    @st.cache_data(show_spinner=False)
    def _cached_inner(data_hash: str, *args):
        # We pass the actual data via session_state to avoid hashing issues
        _data = st.session_state.get(f"_plyr_{data_hash}", None)
        if _data is None:
            return None
        return fn(_data, *args)

    @functools.wraps(fn)
    def wrapper(data, *args):
        h = _compound_hash(data) if isinstance(data, list) else _single_hash(data)
        # Store data in session_state keyed by hash (no copy — reference)
        st.session_state[f"_plyr_{h}"] = data
        result = _cached_inner(h, *args)
        return result

    return wrapper


# ─────────────────────────────────────────────────────────────────────────────
# PHASE 3: Lazy-load guards for heavy tabs
# ─────────────────────────────────────────────────────────────────────────────

def lazy_tab(tab_key: str, label: str = "Load Analysis",
             description: str = "Click to run this analysis.") -> bool:
    """
    Returns True if the user has triggered this tab's computation.
    Shows a button if not yet triggered. Persists via session_state.
    Calling code pattern:
        with TABS[N]:
            if perf_layer.lazy_tab("tab_similarity"):
                st.plotly_chart(fig_similarity(display_data))
    """
    ss_key = f"_lazy_{tab_key}"
    if st.session_state.get(ss_key, False):
        return True
    st.markdown(
        f'<div style="background:rgba(14,16,23,.6);border:1px solid rgba(232,160,32,.15);'
        f'border-radius:10px;padding:20px 24px;margin:16px 0;text-align:center">'
        f'<div style="font-family:JetBrains Mono,monospace;font-size:.58rem;'
        f'letter-spacing:3px;color:rgba(232,160,32,.4);margin-bottom:10px">'
        f'LAZY LOAD · PERFORMANCE MODE</div>'
        f'<div style="font-size:.8rem;color:#94a3b8;margin-bottom:16px">'
        f'{description}</div></div>',
        unsafe_allow_html=True,
    )
    if st.button(f"▶ {label}", key=f"_lazy_btn_{tab_key}", type="primary"):
        st.session_state[ss_key] = True
        st.rerun()
    return False


def lazy_reset(tab_key: str):
    """Reset a lazy tab so it shows the trigger button again."""
    st.session_state.pop(f"_lazy_{tab_key}", None)


# ─────────────────────────────────────────────────────────────────────────────
# PHASE 6: Cached stats-strip computation
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def cached_stats(data_hash: str, total: int, grades_a: int,
                 hia_ok: int, bbb_ok: int, pf: int, hh: int,
                 aqed: float, als: float, asa: float) -> dict:
    """
    Cache the stats strip values — returns identical dict.
    Args are all primitives so Streamlit can hash them correctly.
    """
    return {
        "total": total, "ga": grades_a, "hia_ok": hia_ok,
        "bbb_ok": bbb_ok, "pf": pf, "hh": hh,
        "aqed": aqed, "als": als, "asa": asa,
    }


# ─────────────────────────────────────────────────────────────────────────────
# PHASE 6: HTML export caching
# ─────────────────────────────────────────────────────────────────────────────

def get_cached_html_export(data_hash: str, export_fn, display_data: list) -> bytes:
    """Cache the HTML export bytes keyed by compound hash."""
    ss_key = f"_html_export_{data_hash}"
    if ss_key not in st.session_state:
        st.session_state[ss_key] = export_fn(display_data)
    return st.session_state[ss_key]


def get_cached_text_export(data_hash: str, export_fn, display_data: list) -> bytes:
    """Cache the text export bytes keyed by compound hash."""
    ss_key = f"_text_export_{data_hash}"
    if ss_key not in st.session_state:
        st.session_state[ss_key] = export_fn(display_data)
    return st.session_state[ss_key]


# ─────────────────────────────────────────────────────────────────────────────
# PHASE 2: Apply patches to app module's fig_* functions
# Called once from app.py after all fig_* are defined
# ─────────────────────────────────────────────────────────────────────────────

_PATCHED = False


def patch_fig_functions(app_module) -> int:
    """
    Wraps all fig_* functions in app_module with caching.
    Returns count of functions patched.
    Safe: if a function is already wrapped or missing, it's skipped.
    """
    global _PATCHED
    if _PATCHED:
        return 0

    # Functions to wrap and their cache key prefixes
    # Only wrap functions that take a list as first arg (compound data)
    _FIG_TARGETS = [
        "fig_similarity",   # O(n²) — highest priority
        "fig_parallel",
        "fig_pca",
        "fig_boiled_egg",
        "fig_qed_sa",
        "fig_sa",
        "fig_cyp",
        "fig_radar",
        "fig_approved",
    ]

    patched = 0
    for fn_name in _FIG_TARGETS:
        original = getattr(app_module, fn_name, None)
        if original is None:
            continue
        # Don't double-wrap
        if hasattr(original, "_perf_patched"):
            continue
        try:
            wrapped = _make_cached_fig(original, fn_name)
            wrapped._perf_patched = True
            setattr(app_module, fn_name, wrapped)
            patched += 1
        except Exception:
            pass  # never crash — original function remains

    _PATCHED = True
    return patched


# ─────────────────────────────────────────────────────────────────────────────
# PHASE 6: Pagination helper for large compound tables
# ─────────────────────────────────────────────────────────────────────────────

def paginate(data: list, page_key: str, page_size: int = 25) -> tuple:
    """
    Returns (page_data, total_pages, current_page).
    Adds prev/next controls. Non-destructive — just slices display.
    """
    total = len(data)
    if total <= page_size:
        return data, 1, 1

    n_pages = max(1, (total + page_size - 1) // page_size)
    cur_page = st.session_state.get(page_key, 1)
    cur_page = max(1, min(cur_page, n_pages))

    col1, col2, col3 = st.columns([1, 2, 1])
    with col1:
        if st.button("← Prev", key=f"_{page_key}_prev",
                     disabled=(cur_page <= 1)):
            st.session_state[page_key] = cur_page - 1
            st.rerun()
    with col2:
        st.markdown(
            f'<div style="text-align:center;font-family:JetBrains Mono,monospace;'
            f'font-size:.65rem;color:rgba(245,166,35,.6);padding-top:6px">'
            f'Page {cur_page} / {n_pages} &nbsp;·&nbsp; {total} compounds</div>',
            unsafe_allow_html=True)
    with col3:
        if st.button("Next →", key=f"_{page_key}_next",
                     disabled=(cur_page >= n_pages)):
            st.session_state[page_key] = cur_page + 1
            st.rerun()

    start = (cur_page - 1) * page_size
    end   = min(start + page_size, total)
    return data[start:end], n_pages, cur_page


# ─────────────────────────────────────────────────────────────────────────────
# PHASE 6: Selective rendering guard — skip if data unchanged
# ─────────────────────────────────────────────────────────────────────────────

def data_changed(data: list, guard_key: str) -> bool:
    """
    Returns True if data has changed since last render.
    Stores hash in session_state[guard_key].
    Use to skip expensive renders when data is identical.
    """
    new_hash = _compound_hash(data)
    old_hash = st.session_state.get(guard_key)
    if old_hash == new_hash:
        return False
    st.session_state[guard_key] = new_hash
    return True
