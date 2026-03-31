"""
ai_explainer_tab.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · AI Explainer — Tab 41 (PHASE 3)
• Uses Anthropic Claude API to generate plain-language explanations
• Explains drug-likeness scores, ADMET properties, and safety flags
• Uses same _get_api_key() pattern as rest of app
• Graceful fallback if API key unavailable
• Fully isolated — additive only
────────────────────────────────────────────────────────────────────────────
"""

import streamlit as st
import json

try:
    import anthropic
    _ANTHRO_OK = True
except Exception:
    _ANTHRO_OK = False

# ── Prompt builder ────────────────────────────────────────────────────────

def _build_prompt(compound: dict, mode: str) -> str:
    props = {
        "ID":           compound.get("ID", "Unknown"),
        "MW":           compound.get("MW", "?"),
        "LogP":         compound.get("LogP", "?"),
        "TPSA":         compound.get("tPSA", compound.get("TPSA", "?")),
        "QED":          compound.get("QED", compound.get("_qed", "?")),
        "HBD":          compound.get("HBD", "?"),
        "HBA":          compound.get("HBA", "?"),
        "LeadScore":    compound.get("LeadScore", "?"),
        "SA_Score":     compound.get("SA_Score", compound.get("_sa", "?")),
        "PAINS":        "Flagged" if compound.get("_pains") else "Clear",
        "hERG":         compound.get("_herg", "?"),
        "Grade":        compound.get("Grade", "?"),
        "BBB":          "Penetrates" if compound.get("_bbb") else "Does not penetrate",
        "HIA":          "Good" if compound.get("_hia") else "Poor",
    }
    props_text = "\n".join(f"  {k}: {v}" for k, v in props.items())

    if mode == "Compound Overview & Drug-Likeness Assessment":
        return (
            f"You are a medicinal chemistry expert explaining drug analysis results "
            f"to a pharmaceutical researcher. Given these ADMET properties for compound {props['ID']}:\n\n"
            f"{props_text}\n\n"
            f"Provide a concise 3-4 paragraph plain-language summary covering: "
            f"(1) overall drug-likeness assessment, "
            f"(2) key strengths of this compound, "
            f"(3) potential liabilities or concerns, "
            f"(4) recommended next steps in drug development."
        )
    elif mode == "Toxicity & Safety Liability Analysis":
        return (
            f"You are a toxicology expert. Analyse the safety profile of compound {props['ID']}:\n\n"
            f"{props_text}\n\n"
            f"Focus on: PAINS/BRENK flags, hERG risk, TPSA implications for absorption, "
            f"and any structural concerns. Be specific and actionable. 2-3 paragraphs."
        )
    elif mode == "Lead Optimisation Recommendations":
        return (
            f"You are a drug design expert providing lead optimisation guidance for compound {props['ID']}:\n\n"
            f"{props_text}\n\n"
            f"Suggest 3-5 specific medicinal chemistry modifications to improve drug-likeness, "
            f"reduce liabilities, or improve ADMET properties. "
            f"Format as a numbered list with brief rationale for each suggestion."
        )
    else:  # Property Deep Dive
        return (
            f"Explain each of these ADMET properties for compound {props['ID']} "
            f"in plain language suitable for a chemistry student:\n\n"
            f"{props_text}\n\n"
            f"For each property, explain what it means, what the value indicates, "
            f"and whether it's favourable or concerning. Use bullet points."
        )


def _call_claude(prompt: str, api_key: str, max_tokens: int = 600) -> str:
    if not _ANTHRO_OK:
        return "Anthropic library not installed."
    try:
        client = anthropic.Anthropic(api_key=api_key)
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}],
        )
        return msg.content[0].text if msg.content else "No response."
    except Exception as e:
        return f"API error: {e}"


# ── Main render function ──────────────────────────────────────────────────

def render_tab(res: list, api_key: str = ""):
    st.markdown(
        '<div style="font-family:\'JetBrains Mono\',monospace;font-size:.6rem;'
        'letter-spacing:3px;color:rgba(232,160,32,.5);text-transform:uppercase;'
        'margin-bottom:12px">⬡ AI Explainer — Claude-Powered Analysis Narrative</div>',
        unsafe_allow_html=True,
    )

    if not res:
        st.warning("No compounds loaded.")
        return

    if not api_key:
        st.info(
            "🔑 **AI Explainer requires an Anthropic API key.**  \n"
            "Add `ANTHROPIC_API_KEY` to your Streamlit Cloud secrets "
            "(App Settings → Secrets) to enable AI explanations."
        )
        return

    # Compound picker
    ids = [c.get("ID", f"Cpd-{i+1}") for i, c in enumerate(res)]
    sel_id = st.selectbox("Select Compound for AI Explanation", ids, key="_aiexp_sel")
    compound = next((c for c in res if c.get("ID", "") == sel_id), res[0])

    # Analysis mode
    mode = st.radio(
        "Explanation Mode",
        ["Compound Overview & Drug-Likeness Assessment", "Toxicity & Safety Liability Analysis", "Lead Optimisation Recommendations", "In-Depth Physicochemical Property Explanation"],
        horizontal=True,
        key="_aiexp_mode",
    )

    # Token budget slider
    max_tok = st.slider("AI Response Detail Level (Tokens)", 200, 1000, 500, 100, key="_aiexp_tokens")

    # Quick properties preview
    with st.expander("Compound Physicochemical Profile"):
        pcols = st.columns(4)
        for i, (k, v) in enumerate([
            ("ID", compound.get("ID")), ("MW", compound.get("MW")),
            ("LogP", compound.get("LogP")), ("QED", compound.get("QED", compound.get("_qed"))),
            ("Grade", compound.get("Grade")), ("LeadScore", compound.get("LeadScore")),
            ("SA_Score", compound.get("SA_Score")), ("PAINS", "Yes" if compound.get("_pains") else "No"),
        ]):
            pcols[i % 4].metric(k, v or "–")

    # Generate button
    if st.button("✨ Generate Scientific AI Explanation", key="_aiexp_run", type="primary"):
        prompt = _build_prompt(compound, mode)
        with st.spinner("Claude is thinking..."):
            response = _call_claude(prompt, api_key, max_tok)

        # Render response in styled box
        st.markdown(
            f'<div style="background:rgba(232,160,32,.04);border:1px solid rgba(232,160,32,.2);'
            f'border-radius:10px;padding:18px 22px;margin-top:10px;'
            f'font-size:.85rem;line-height:1.7;color:#c8deff">'
            f'{response.replace(chr(10), "<br>")}'
            f'</div>',
            unsafe_allow_html=True,
        )

        # Download option
        st.download_button(
            "↓ Download AI Explanation Report",
            data=f"ChemoFilter AI Explainer\nCompound: {sel_id}\nMode: {mode}\n\n{response}",
            file_name=f"ai_explanation_{sel_id}.txt",
            mime="text/plain",
            key="_aiexp_dl",
        )

    # ── Batch summary mode ────────────────────────────────────────────────
    st.divider()
    with st.expander("Batch AI Summary — Top 3 Lead Candidates"):
        if st.button("Generate Batch Scientific Summary", key="_aiexp_batch"):
            top3 = sorted(res, key=lambda c: float(c.get("LeadScore", 0)), reverse=True)[:3]
            for cpd in top3:
                with st.spinner(f"Explaining {cpd.get('ID', '–')}..."):
                    p = _build_prompt(cpd, "Compound Overview & Drug-Likeness Assessment")
                    r = _call_claude(p, api_key, 300)
                st.markdown(f"### {cpd.get('ID', '–')}")
                st.markdown(r)
                st.divider()
