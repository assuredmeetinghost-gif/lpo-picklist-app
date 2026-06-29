#!/usr/bin/env python3
"""
LPO → Sales Order Pick List — Streamlit Web App
------------------------------------------------
Wraps lpo_to_picklist.py + picklist_builder.py (both untouched).
Deploy on Streamlit Community Cloud (share.streamlit.io) — free, one-click from GitHub.
"""
import os
import io
import json
import pathlib
import sys
import tempfile

import streamlit as st

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="LPO → Pick List",
    page_icon="📋",
    layout="centered",
)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")

    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        placeholder="sk-ant-api03-...",
        help="Used only for this session. Never stored or logged.",
    )

    model = st.selectbox(
        "AI Model",
        options=["claude-opus-4-8", "claude-sonnet-4-6"],
        index=0,
        help=(
            "**Opus 4** — most accurate, especially on messy handwritten forms (slower, costs more).\n\n"
            "**Sonnet 4** — faster and cheaper; good for clean printed LPOs."
        ),
    )

    st.divider()
    st.caption(
        "Powered by **Claude Vision** · PyMuPDF · openpyxl\n\n"
        "Handles printed, scanned, and handwritten purchase orders."
    )

# ── Main ──────────────────────────────────────────────────────────────────────
st.title("📋 LPO → Sales Order Pick List")
st.caption(
    "Upload a Purchase Order PDF (printed, scanned, or handwritten). "
    "Claude reads every page and produces a warehouse-ready Excel Pick List."
)

uploaded = st.file_uploader(
    "Drop your LPO / Purchase Order here",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=False,
    label_visibility="collapsed",
)

if not uploaded:
    st.markdown(
        """
**How it works**

1. Enter your Anthropic API key in the sidebar (free to get at [console.anthropic.com](https://console.anthropic.com))
2. Upload any LPO or Purchase Order — printed PDF, phone scan, or handwritten form
3. Click **Generate Pick List** — Claude Vision reads every page, even rotated/two-column layouts
4. Preview extracted line items and download the Excel Pick List

> ℹ️ **Tip:** Set `ANTHROPIC_API_KEY` as an environment variable on your deployment and leave the
> sidebar field blank — the app will use it automatically.
        """
    )
    st.stop()

# File is uploaded — show info and action button
st.info(f"📄 **{uploaded.name}** · {uploaded.size // 1024} KB")

run = st.button("▶ Generate Pick List", type="primary", use_container_width=True)

if not run:
    st.stop()

# ── Validate inputs ───────────────────────────────────────────────────────────
effective_key = api_key or os.environ.get("ANTHROPIC_API_KEY", "")
if not effective_key:
    st.error(
        "No API key found. "
        "Enter your Anthropic API key in the sidebar, "
        "or set the `ANTHROPIC_API_KEY` environment variable on your deployment."
    )
    st.stop()

# Inject key so lpo_to_picklist.get_client() picks it up
os.environ["ANTHROPIC_API_KEY"] = effective_key

# ── Import core modules (after env is set) ────────────────────────────────────
# We add the app directory to sys.path so Python finds lpo_to_picklist & picklist_builder
# even if the working directory differs.
APP_DIR = str(pathlib.Path(__file__).resolve().parent)
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

import lpo_to_picklist as lpo  # noqa: E402  (import after env + path setup)

# ── Process ───────────────────────────────────────────────────────────────────
with tempfile.TemporaryDirectory() as tmp_str:
    tmp = pathlib.Path(tmp_str)

    # Override module-level CACHE_DIR so it writes to our temp space (cloud = no persistent disk)
    lpo.CACHE_DIR = tmp / "cache"
    lpo.CACHE_DIR.mkdir(parents=True, exist_ok=True)

    # Write uploaded bytes to a real file path (fitz.open needs a path, not a buffer)
    lpo_path = tmp / uploaded.name
    lpo_path.write_bytes(uploaded.getvalue())

    out_dir = tmp / "output"

    data = None
    xlsx_bytes = None

    with st.status("Reading your LPO with Claude Vision…", expanded=True) as status:

        def progress(msg: str):
            """Callback so lpo_to_picklist can push log lines into the Streamlit status box."""
            st.write(msg)

        try:
            # extract_data: render PDF pages → send to Claude Vision → structured JSON
            data = lpo.extract_data(
                str(lpo_path),
                model=model,
                use_cache=False,   # temp dir → cache won't persist on cloud anyway
                progress=progress,
            )
        except SystemExit:
            # lpo_to_picklist calls sys.exit() if no API key — we handle it above,
            # but guard here just in case.
            status.update(label="❌ API key error", state="error")
            st.error("API key was rejected. Check that it is valid and has available credits.")
            st.stop()
        except Exception as exc:
            status.update(label=f"❌ Extraction failed", state="error")
            st.exception(exc)
            st.stop()

        if not data or not data.get("items"):
            status.update(label="⚠️ No line items found", state="error")
            st.warning(
                "Claude Vision did not detect any ordered items in this document.\n\n"
                "**Try:**\n"
                "- A higher-quality scan\n"
                "- Switching to the Opus 4 model (more accurate on handwritten forms)\n"
                "- Confirming the PDF actually contains order quantities"
            )
            st.stop()

        try:
            name = pathlib.Path(uploaded.name).stem
            # build_outputs writes .xlsx (and optionally .pdf — we skip PDF on cloud)
            xlsx_path, _ = lpo.build_outputs(data, out_dir, name, make_pdf=False)
            xlsx_bytes = pathlib.Path(xlsx_path).read_bytes()
        except Exception as exc:
            status.update(label="❌ Excel build failed", state="error")
            st.exception(exc)
            st.stop()

        status.update(label="✅ Pick List ready!", state="complete")

# ── Results ───────────────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
col1.metric("Line Items", data["total_lines"])
col2.metric("Total Qty", int(data["total_qty"]))
col3.metric("Customer", (data.get("customer") or "—")[:20])

with st.expander("📊 Preview line items", expanded=True):
    import pandas as pd  # noqa: E402

    df = pd.DataFrame(data["items"])[["gtin", "desc", "it_code", "uom", "qty"]]
    df.columns = ["GTIN / Code", "Description", "IT Code", "UOM", "Qty"]
    st.dataframe(df, use_container_width=True, hide_index=True)

st.download_button(
    label="⬇️  Download Pick List  (.xlsx)",
    data=xlsx_bytes,
    file_name=f"PickList_{name}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)

with st.expander("🗂 Raw extracted JSON"):
    st.json(data)
