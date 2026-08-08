"""
miRNA–Target Gene Pipeline — Streamlit Web Interface

Usage: streamlit run app.py
"""

import sys
import os
import tempfile
import time
from io import BytesIO

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

import streamlit as st
import extract_mirna_genes as ext
import llm_review
import normalize_names as norm

st.set_page_config(
    page_title="Plant miRNA–Target Gene Pipeline",
    page_icon="🧬",
    layout="wide",
)

# ============================================================
# Sidebar
# ============================================================
st.sidebar.title("🧬 Plant miRNA–Target Gene Pipeline")
st.sidebar.caption("Extract · Review · Normalize · Deduplicate")

api_key = st.sidebar.text_input(
    "API Key",
    type="password",
    value=os.environ.get("LLM_API_KEY", os.environ.get("DEEPSEEK_API_KEY", "")),
    help="Your provider's API key",
)

PROVIDERS = {
    "DeepSeek": ("https://api.deepseek.com/v1/chat/completions", ["deepseek-chat", "deepseek-reasoner"]),
    "OpenAI": ("https://api.openai.com/v1/chat/completions", ["gpt-4o", "gpt-4o-mini", "gpt-4.1"]),
    "Groq": ("https://api.groq.com/openai/v1/chat/completions", ["llama-3.1-70b-versatile", "mixtral-8x7b-32768"]),
    "Together AI": ("https://api.together.xyz/v1/chat/completions", ["meta-llama/Llama-3.1-70B-Instruct"]),
    "OpenRouter": ("https://openrouter.ai/api/v1/chat/completions", ["openai/gpt-4o", "anthropic/claude-sonnet-4"]),
    "SiliconFlow": ("https://api.siliconflow.cn/v1/chat/completions", ["deepseek-ai/DeepSeek-V3"]),
    "Moonshot": ("https://api.moonshot.cn/v1/chat/completions", ["moonshot-v1-8k"]),
    "Zhipu": ("https://open.bigmodel.cn/api/paas/v4/chat/completions", ["glm-4"]),
    "Custom": ("", []),
}

with st.sidebar.expander("⚙️ Settings"):
    provider = st.selectbox("Provider", list(PROVIDERS.keys()), index=0)
    base_url, default_models = PROVIDERS[provider]
    if base_url:
        llm_review.LLM_API_URL = base_url
    else:
        llm_review.LLM_API_URL = st.text_input("Custom API URL",
                                                value="https://api.deepseek.com/v1/chat/completions",
                                                help="Any OpenAI-compatible endpoint")

    if default_models:
        llm_model = st.selectbox("Model", default_models, index=0)
    else:
        llm_model = st.text_input("Model name", value="deepseek-chat",
                                  help="Enter the model ID for your provider")

    entries_per_batch = st.number_input("Entries per batch", 1, 20, 5,
                                        help="More = faster, but may reduce accuracy")

llm_review.LLM_MODEL = llm_model
llm_review.ENTRIES_PER_BATCH = entries_per_batch

# ============================================================
# Session state
# ============================================================
for key in ["extracted_pairs", "extract_stats", "extract_excel_path", "extract_excel_bytes",
            "reviewed_excel_path", "reviewed_excel_bytes", "review_stats",
            "final_excel_path", "final_excel_bytes", "final_pairs", "dedup_stats"]:
    if key not in st.session_state:
        st.session_state[key] = None


# ============================================================
# Helpers
# ============================================================
def run_llm_review(excel_path):
    """Run LLM review on an extracted Excel file. Returns stats dict."""
    os.environ["DEEPSEEK_API_KEY"] = api_key
    llm_review.INPUT_XLSX = excel_path
    llm_review.OUTPUT_XLSX = os.path.join(SCRIPT_DIR, "miRNA-Target_Gene_Pairs_Reviewed.xlsx")
    if os.path.exists(llm_review.OUTPUT_XLSX):
        os.remove(llm_review.OUTPUT_XLSX)
    llm_review.PROGRESS_FILE = os.path.join(SCRIPT_DIR, "llm_review_progress.json")

    entries = llm_review.load_uncertain_entries()
    pending = list(entries)
    reviewed = {}
    total_batches = (len(pending) + entries_per_batch - 1) // entries_per_batch
    prog = st.progress(0, f"LLM review — {len(pending)} entries in {total_batches} batches")

    # Quick connectivity check on first batch
    if pending:
        test_prompt = llm_review.build_batch_prompt(pending[:1])
        test_result = llm_review.call_llm_batch(test_prompt, api_key)
        if test_result is None:
            st.error("API call failed — check your API key, provider, and model. "
                     "Common issues: wrong provider URL for your key, or model not available.")
            return None

    for batch_idx in range(0, len(pending), entries_per_batch):
        batch = pending[batch_idx:batch_idx + entries_per_batch]
        prompt = llm_review.build_batch_prompt(batch)
        results = llm_review.call_llm_batch(prompt, api_key)
        if results and len(results) == len(batch):
            for entry, result in zip(batch, results):
                reviewed[entry["idx"]] = result
        else:
            for entry in batch:
                result = llm_review.call_llm(llm_review.build_prompt(entry), api_key)
                if result:
                    reviewed[entry["idx"]] = result
        prog.progress(min((batch_idx + entries_per_batch) / len(pending), 1.0),
                      f"LLM review {min(batch_idx + entries_per_batch, len(pending))}/{len(pending)}")
        time.sleep(0.1)

    prog.progress(1.0, "Done!")
    rstats = llm_review.merge_results(entries, reviewed)
    if os.path.exists(llm_review.PROGRESS_FILE):
        os.remove(llm_review.PROGRESS_FILE)
    st.session_state.reviewed_excel_path = llm_review.OUTPUT_XLSX
    with open(llm_review.OUTPUT_XLSX, "rb") as f:
        st.session_state.reviewed_excel_bytes = f.read()
    st.session_state.review_stats = rstats
    return rstats


# ============================================================
# Main UI
# ============================================================
st.title("Plant miRNA–Target Gene Relationship Extraction")
st.caption("Upload a batch of PubMed abstracts and extract miRNA–gene target pairs "
           "with regex + LLM review + normalization.")

# ── Step 1: Upload & Extract ──
st.header("1. Upload & Extract")
uploaded_file = st.file_uploader(
    "Choose an abstracts file (.txt)",
    type=["txt"],
    help="Each record should start with 'N. JournalName ...'",
)

if uploaded_file:
    col1, col2 = st.columns([1, 3])
    with col1:
        run_extract = st.button("🔍 Run Extraction", type="primary", use_container_width=True)
    with col2:
        auto_review = st.checkbox("Auto-run LLM review after extraction", value=False)

    if run_extract:
        with st.spinner("Extracting miRNA–gene pairs..."):
            raw_bytes = uploaded_file.getvalue()
            with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as tmp:
                tmp.write(raw_bytes)
                tmp_path = tmp.name

            blocks = ext.parse_abstracts_file(tmp_path)
            os.unlink(tmp_path)

            all_pairs = []
            stats = {"total": len(blocks), "has_mirna": 0, "has_gene": 0,
                     "has_both": 0, "has_pair": 0, "total_mirna": 0, "total_gene": 0}

            progress = st.progress(0, "Scanning abstracts...")
            for i, block in enumerate(blocks):
                pmid, doi, title = ext.extract_metadata(block)
                full_text = block
                mirnas = ext.extract_mirnas(full_text)
                genes = ext.extract_genes(full_text)
                if mirnas:
                    stats["has_mirna"] += 1
                    stats["total_mirna"] += len(mirnas)
                if genes:
                    stats["has_gene"] += 1
                    stats["total_gene"] += len(genes)
                if mirnas and genes:
                    stats["has_both"] += 1
                    species = ext.detect_species(
                        mirnas, [(g[0], g[1], g[2], g[3]) for g in genes], full_text, title or "")
                    pairs = ext.pair_mirna_gene(mirnas, genes, full_text, ext.get_abstract_body(block))
                    if pairs:
                        stats["has_pair"] += 1
                        for p in pairs:
                            p["pmid"] = pmid
                            p["doi"] = doi
                            p["title"] = title
                            p["abstract_index"] = i + 1
                            p["species"] = species
                        all_pairs.extend(pairs)

                if (i + 1) % 200 == 0:
                    progress.progress((i + 1) / len(blocks),
                                      f"Processed {i+1}/{len(blocks)} — {len(all_pairs)} pairs found")

            progress.progress(1.0, "Done!")
            excel_path = os.path.join(SCRIPT_DIR, "miRNA-Target_Gene_Pairs.xlsx")
            ext.build_excel(all_pairs, excel_path)
            with open(excel_path, "rb") as f:
                st.session_state.extract_excel_bytes = f.read()

            st.session_state.extracted_pairs = all_pairs
            st.session_state.extract_stats = stats
            st.session_state.extract_excel_path = excel_path

        st.success(f"Extraction complete — {len(all_pairs)} pairs found in {stats['total']} abstracts")

        if auto_review:
            if api_key:
                with st.spinner("Running LLM review..."):
                    rstats = run_llm_review(excel_path)
                if rstats:
                    st.success(f"LLM review done — {rstats['Confirmed']} confirmed, "
                               f"{rstats.get('Excluded (non-target)', 0) + rstats.get('Excluded (non-gene)', 0)} excluded")
            else:
                st.warning("Auto-review skipped — API key not set in sidebar")

    if st.session_state.extract_stats:
        stats = st.session_state.extract_stats
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Abstracts", stats["total"])
        c2.metric("With miRNA", stats["has_mirna"])
        c3.metric("With gene", stats["has_gene"])
        c4.metric("Both", stats["has_both"])
        c5.metric("With pairs", stats["has_pair"])
        c6.metric("Pairs found", len(st.session_state.extracted_pairs or []))

    if st.session_state.extract_excel_bytes:
        st.download_button(
            "📥 Download Extraction Results",
            st.session_state.extract_excel_bytes,
            file_name="miRNA-Target_Gene_Pairs.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ── Step 2: LLM Review ──
st.header("2. LLM Review")
st.caption("Send uncertain pairs to the LLM for verification. Requires API key in the sidebar.")

col_r1, _ = st.columns([1, 3])
with col_r1:
    run_review = st.button("🤖 Run LLM Review", type="primary", use_container_width=True,
                           disabled=not (api_key and st.session_state.extract_excel_path))

if run_review:
    with st.spinner("Running LLM review..."):
        rstats = run_llm_review(st.session_state.extract_excel_path)
    if rstats:
        st.success("LLM review complete!")

if st.session_state.review_stats:
    rstats = st.session_state.review_stats
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Confirmed", rstats["Confirmed"])
    c2.metric("Weak association", rstats["Weak association"])
    c3.metric("Excluded (non-target)", rstats.get("Excluded (non-target)", 0))
    c4.metric("Excluded (non-gene)", rstats.get("Excluded (non-gene)", 0))
    c5.metric("Retained", rstats["Confirmed"] + rstats["Weak association"])

if st.session_state.reviewed_excel_bytes:
        st.download_button(
            "📥 Download Reviewed Results",
            st.session_state.reviewed_excel_bytes,
            file_name="miRNA-Target_Gene_Pairs_Reviewed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ── Step 3: Normalize & Dedup ──
st.header("3. Normalize & Deduplicate")
st.caption("Standardize miRNA names and merge duplicate pairs across studies.")

norm_source = st.radio(
    "Input source",
    ["Use LLM-reviewed result from Step 2", "Upload a manually-reviewed Excel file"],
    horizontal=True,
)

input_for_norm = None
uploaded_reviewed = None
if norm_source.startswith("Use"):
    input_for_norm = st.session_state.reviewed_excel_path or st.session_state.extract_excel_path
    if not input_for_norm:
        st.info("No reviewed file available — run Step 2 first or upload manually.")
else:
    uploaded_reviewed = st.file_uploader(
        "Choose a reviewed Excel file (.xlsx)",
        type=["xlsx"],
        key="norm_upload",
        help="Upload the Reviewed Excel (or your manually-corrected version of it).",
    )
    if uploaded_reviewed:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(uploaded_reviewed.getvalue())
            input_for_norm = tmp.name

col_n1, _ = st.columns([1, 3])
with col_n1:
    run_norm = st.button("📊 Normalize & Dedup", type="primary", use_container_width=True,
                         disabled=not input_for_norm)

if run_norm and input_for_norm:
    with st.spinner("Normalizing names and deduplicating..."):
        pairs = norm.load_pairs(input_for_norm)
        m_changed = 0
        for p in pairs:
            sp = p.get("species", "unknown")
            p["mirna_norm"] = norm.normalize_mirna(p["mirna_raw"], species=sp)
            p["gene_norm"] = norm.normalize_gene(p["gene_raw"], species=sp)
            if p["mirna_norm"] != p["mirna_raw"]:
                m_changed += 1

        deduped = norm.deduplicate(pairs)

    with st.spinner("Resolving gene IDs (ricedata.cn / NCBI)..."):
        gene_map = {}
        gene_resolved = 0
        gene_total = 0
        try:
            gene_map = norm.resolve_gene_ids(deduped)
            gene_total = len(gene_map)
            gene_resolved = sum(1 for v in gene_map.values() if v.get("RAP") or v.get("NCBI"))
        except Exception:
            st.warning("Gene ID resolution failed (network issue). IDs will be empty.")

    output_path = os.path.join(SCRIPT_DIR, "miRNA-Target_Gene_Pairs_Final.xlsx")
    norm.export_excel(pairs, deduped, gene_map, gene_resolved, gene_total, input_for_norm, output_path)
    with open(output_path, "rb") as f:
        st.session_state.final_excel_bytes = f.read()

    if uploaded_reviewed:
        os.unlink(input_for_norm)

    st.session_state.final_excel_path = output_path
    st.session_state.final_pairs = deduped
    st.session_state.dedup_stats = {
        "before": len(pairs), "after": len(deduped),
        "miRNA_changed": m_changed,
        "multi_pmid": sum(1 for d in deduped if d["pmid_count"] >= 2),
    }

    st.success(f"Done — {len(pairs)} entries → {len(deduped)} groups "
               f"({m_changed} miRNA names normalized)")

if st.session_state.dedup_stats:
    ds = st.session_state.dedup_stats
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Before dedup", ds["before"])
    c2.metric("After dedup", ds["after"])
    c3.metric("miRNA names changed", ds["miRNA_changed"])
    c4.metric("≥2 PMID support", ds["multi_pmid"])

if st.session_state.final_excel_bytes:
        st.download_button(
            "📥 Download Final Results",
            st.session_state.final_excel_bytes,
            file_name="miRNA-Target_Gene_Pairs_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.divider()
