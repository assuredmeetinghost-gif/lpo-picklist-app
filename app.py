#!/usr/bin/env python3
"""LPO -> Sales Order Pick List  .  Streamlit Web App"""
import os, pathlib, sys, tempfile
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="LPO -> Pick List", page_icon="clipboard", layout="wide")

CSS = """
<style>
#MainMenu,footer,header,[data-testid="stToolbar"],[data-testid="stDecoration"],
[data-testid="stStatusWidget"],.stDeployButton,[data-testid="manage-app-button"],
.viewerBadge_container__r5tak { display:none!important; }

[data-testid="stAppViewContainer"] { background:#EEF2FF; }

[data-testid="stSidebar"] { background:#1B3A6B!important; }
[data-testid="stSidebar"] p,[data-testid="stSidebar"] label,
[data-testid="stSidebar"] span,[data-testid="stSidebar"] small
  { color:rgba(255,255,255,.9)!important; }
[data-testid="stSidebar"] h2 { color:#fff!important; font-size:17px!important; }
[data-testid="stSidebar"] hr { border-color:rgba(255,255,255,.15)!important; }
[data-testid="stSidebar"] input {
  background:rgba(255,255,255,.10)!important;
  border:1px solid rgba(255,255,255,.22)!important;
  border-radius:7px!important; color:white!important; }
[data-testid="stSidebar"] input::placeholder { color:rgba(255,255,255,.4)!important; }
[data-testid="stSidebar"] [data-baseweb="select"]>div:first-child {
  background:rgba(255,255,255,.10)!important;
  border:1px solid rgba(255,255,255,.22)!important; border-radius:7px!important; }
[data-testid="stSidebar"] [data-baseweb="select"] svg { fill:white!important; }

[data-testid="block-container"],.block-container
  { max-width:880px!important; padding:1.8rem 1.5rem 2rem!important; }

[data-testid="stFileUploader"] {
  background:white!important; border:2px dashed #CBD5E1!important;
  border-radius:12px!important; transition:border-color .2s; }
[data-testid="stFileUploader"]:hover { border-color:#1B3A6B!important; }

.stButton>button[kind="primary"] {
  background:#1B3A6B!important; border:none!important; border-radius:8px!important;
  font-size:15px!important; font-weight:600!important;
  box-shadow:0 2px 8px rgba(27,58,107,.30)!important; transition:all .15s!important; }
.stButton>button[kind="primary"]:hover {
  background:#152F5A!important; transform:translateY(-1px)!important;
  box-shadow:0 4px 14px rgba(27,58,107,.40)!important; }

.stDownloadButton>button {
  background:white!important; color:#1B3A6B!important;
  border:2px solid #1B3A6B!important; border-radius:8px!important; font-weight:600!important; }
.stDownloadButton>button:hover { background:#EEF2FF!important; }

[data-testid="stStatus"],[data-testid="stExpander"] { border-radius:10px!important; }
[data-testid="stExpander"] { background:white; }

/* Model radio styled as segmented control */
div[data-testid="stRadio"] > label { display:none!important; }
div[data-testid="stRadio"] > div {
  background:white; border-radius:10px; padding:10px 14px;
  box-shadow:0 1px 4px rgba(0,0,0,.06); display:flex; gap:8px; }
div[data-testid="stRadio"] > div > label {
  display:flex!important; align-items:center; gap:8px;
  border:1.5px solid #CBD5E1; border-radius:8px; padding:9px 20px;
  font-size:13px; font-weight:500; cursor:pointer;
  transition:all .15s; color:#374151!important; background:white; flex:1;
  justify-content:center; }
div[data-testid="stRadio"] > div > label:has(input:checked) {
  border-color:#1B3A6B!important; background:#EEF2FF!important;
  color:#1B3A6B!important; font-weight:700!important; }
div[data-testid="stRadio"] > div > label > div:first-child { display:none!important; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

# Sidebar
LBL = '<p style="font-size:10px;font-weight:700;letter-spacing:.8px;text-transform:uppercase;color:rgba(255,255,255,.5);margin:{} 0 4px">{}</p>'
with st.sidebar:
    st.markdown("## Settings")
    st.divider()
    st.markdown(LBL.format("0","Anthropic API Key"), unsafe_allow_html=True)
    api_key = st.text_input("k", type="password", placeholder="sk-ant-api03-...", label_visibility="collapsed")
    st.divider()
    st.markdown('<p style="font-size:10px;color:rgba(255,255,255,.35);line-height:1.6">Powered by Claude Vision - PyMuPDF - openpyxl<br>Supports printed, scanned and handwritten orders.</p>', unsafe_allow_html=True)

# Header
st.markdown("""
<div style="background:white;border-radius:14px;padding:22px 28px;
box-shadow:0 1px 6px rgba(27,58,107,.10);border-left:5px solid #1B3A6B;
margin-bottom:20px;display:flex;align-items:center;gap:18px;">
<div style="font-size:40px;line-height:1">&#128203;</div>
<div><h1 style="color:#1B3A6B;font-size:22px;font-weight:700;margin:0">
LPO to Sales Order Pick List</h1>
<p style="color:#6B7280;font-size:13px;margin:5px 0 0">
Upload a Purchase Order PDF (printed, scanned, or handwritten).
Claude Vision reads every page and builds a warehouse-ready Pick List.</p></div></div>
""", unsafe_allow_html=True)

# Model selector (always visible)
model = st.radio(
    "Model",
    options=["claude-opus-4-8", "claude-sonnet-4-6"],
    format_func=lambda x: (
        "Opus 4  -  Most accurate  (best for handwritten / scanned forms)"
        if x == "claude-opus-4-8"
        else "Sonnet 4  -  Faster & cheaper  (great for clean printed LPOs)"
    ),
    horizontal=True,
    label_visibility="collapsed",
)

# Upload
uploaded = st.file_uploader("Upload LPO", type=["pdf","png","jpg","jpeg"], accept_multiple_files=False, label_visibility="collapsed")

if not uploaded:
    st.markdown("""
<div style="background:white;border-radius:12px;padding:26px 30px;box-shadow:0 1px 4px rgba(0,0,0,.06);margin-top:6px;">
<p style="font-weight:600;color:#1B3A6B;font-size:15px;margin:0 0 14px">How it works</p>
<ol style="color:#6B7280;font-size:14px;line-height:2.1;padding-left:22px;margin:0">
<li>Enter your Anthropic API key in the sidebar</li>
<li>Select the AI model above</li>
<li>Upload any LPO / Purchase Order PDF</li>
<li>Click <strong style="color:#1B3A6B">Generate</strong> - Claude reads every page</li>
<li>Preview the line items, then print or download as Excel</li>
</ol></div>""", unsafe_allow_html=True)
    st.stop()

c_info, c_btn = st.columns([4, 1])
with c_info:
    st.markdown(f'<div style="background:white;border-radius:10px;padding:13px 18px;box-shadow:0 1px 4px rgba(0,0,0,.06);font-size:14px;">&#128196; <strong>{uploaded.name}</strong> &nbsp;-&nbsp; <span style="color:#6B7280">{uploaded.size//1024} KB</span></div>', unsafe_allow_html=True)
with c_btn:
    run = st.button("Generate", type="primary", use_container_width=True)

if not run:
    st.stop()

effective_key = (api_key or "").strip() or os.environ.get("ANTHROPIC_API_KEY","")
if not effective_key:
    st.error("No API key. Enter your Anthropic API key in the sidebar.")
    st.stop()

os.environ["ANTHROPIC_API_KEY"] = effective_key
APP_DIR = str(pathlib.Path(__file__).resolve().parent)
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)
import lpo_to_picklist as lpo

# Extract + build
data = None; xlsx_bytes = None
name = pathlib.Path(uploaded.name).stem

with tempfile.TemporaryDirectory() as _tmp:
    tmp = pathlib.Path(_tmp)
    lpo.CACHE_DIR = tmp/"cache"; lpo.CACHE_DIR.mkdir(parents=True, exist_ok=True)
    lpo_path = tmp/uploaded.name
    lpo_path.write_bytes(uploaded.getvalue())
    out_dir = tmp/"output"

    with st.status("Reading LPO with Claude Vision...", expanded=True) as status:
        def progress(msg):
            st.write(msg)
        try:
            data = lpo.extract_data(str(lpo_path), model=model, use_cache=False, progress=progress)
        except SystemExit:
            status.update(label="API key rejected", state="error")
            st.error("API key rejected - verify it has credits."); st.stop()
        except Exception as exc:
            status.update(label="Extraction failed", state="error")
            st.exception(exc); st.stop()

        if not data or not data.get("items"):
            status.update(label="No ordered items found", state="error")
            st.warning("No items detected. Try Opus 4 or a clearer scan."); st.stop()

        try:
            xlsx_path, _ = lpo.build_outputs(data, out_dir, name, make_pdf=False)
            xlsx_bytes = pathlib.Path(xlsx_path).read_bytes()
        except Exception as exc:
            status.update(label="Excel build failed", state="error")
            st.exception(exc); st.stop()

        status.update(label="Pick List ready!", state="complete")

# Metrics
def _metric(label, value):
    return (f'<div style="background:white;border-radius:10px;padding:16px 18px;text-align:center;'
            f'box-shadow:0 1px 4px rgba(0,0,0,.06);border-top:3px solid #1B3A6B;">'
            f'<div style="font-size:10px;font-weight:700;letter-spacing:.7px;text-transform:uppercase;color:#9CA3AF;margin-bottom:6px">{label}</div>'
            f'<div style="font-size:28px;font-weight:700;color:#1B3A6B">{value}</div></div>')

st.markdown(
    f'<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:20px 0 16px">'
    + _metric("Line Items", data["total_lines"])
    + _metric("Total Qty", int(data["total_qty"]))
    + f'<div style="background:white;border-radius:10px;padding:16px 18px;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,.06);border-top:3px solid #1B3A6B;"><div style="font-size:10px;font-weight:700;letter-spacing:.7px;text-transform:uppercase;color:#9CA3AF;margin-bottom:6px">Customer</div><div style="font-size:15px;font-weight:700;color:#1B3A6B;padding-top:6px">{(data.get("customer") or "-")[:26]}</div></div>'
    + '</div>',
    unsafe_allow_html=True
)

# Print HTML builder
def _build_print_html(d):
    rows = "".join(
        f'<tr><td class="c">{i}</td><td class="c code">{it.get("gtin","") or "-"}</td>'
        f'<td class="l desc">{it.get("desc","")}</td><td class="c">{it.get("uom","NOS")}</td>'
        f'<td class="c"><input type="checkbox"></td>'
        f'<td class="c qty">{it.get("qty","")}</td><td class="c"></td></tr>'
        for i, it in enumerate(d["items"], 1)
    )
    m = {
        "customer": d.get("customer","") or "-",
        "order_no": d.get("order_no","") or "-",
        "order_date": d.get("order_date","") or "-",
        "source_file": d.get("source_file",""),
        "order_type": d.get("order_type","") or "-",
        "currency": d.get("currency","OMR"),
        "total_qty": int(d["total_qty"]),
    }
    return f"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<title>Pick List</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:Arial,sans-serif;font-size:12px;background:#EEF2FF;padding:16px;color:#111827}}
.wrap{{background:#fff;max-width:800px;margin:0 auto;border-radius:12px;padding:24px 28px;box-shadow:0 2px 12px rgba(27,58,107,.12)}}
.toolbar{{margin-bottom:18px}}
.btn{{background:#1B3A6B;color:#fff;border:none;border-radius:7px;padding:9px 20px;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 2px 6px rgba(27,58,107,.30);transition:background .15s}}
.btn:hover{{background:#152F5A}}
.doc-title{{font-size:16px;font-weight:700;color:#1B3A6B;border-bottom:2px solid #1B3A6B;padding-bottom:7px;margin-bottom:14px}}
.meta{{display:grid;grid-template-columns:1fr 1fr;gap:4px 28px;margin-bottom:18px;font-size:11px;color:#4B5563;line-height:1.7}}
.meta b{{color:#1B3A6B}}
table{{width:100%;border-collapse:collapse;margin-bottom:14px}}
thead th{{background:#1B3A6B;color:#fff;padding:7px 5px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;text-align:center}}
thead th.l{{text-align:left;padding-left:8px}}
tbody td{{padding:6px 5px;border-bottom:1px solid #E5E7EB;vertical-align:middle}}
tbody tr:nth-child(even){{background:#F8FAFF}}
td.c{{text-align:center}} td.l{{text-align:left;padding-left:8px}}
td.desc{{word-break:break-word;max-width:270px}}
td.code{{font-family:monospace;font-size:11px;color:#374151}}
td.qty{{font-weight:700;font-size:13px;color:#1B3A6B}}
input[type="checkbox"]{{width:15px;height:15px;accent-color:#1B3A6B;cursor:pointer}}
.tot td{{font-weight:700;border-top:2px solid #1B3A6B;background:#F0F4FF;padding:8px 5px;color:#1B3A6B}}
.sigs{{display:grid;grid-template-columns:repeat(3,1fr);gap:18px;margin-top:26px}}
.sig{{border-top:1px solid #9CA3AF;padding-top:6px;font-size:10px;color:#6B7280}}
@media print{{
  body{{background:white;padding:0}}
  .wrap{{box-shadow:none;border-radius:0;max-width:100%;padding:8px 12px}}
  .toolbar{{display:none!important}}
  tbody tr:nth-child(even){{background:white}}
  thead th,.tot td{{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
  input[type="checkbox"]{{appearance:none;-webkit-appearance:none;width:12px;height:12px;border:1.5px solid #1B3A6B;border-radius:2px;display:inline-block}}
}}
</style></head><body>
<div class="wrap">
<div class="toolbar"><button class="btn" onclick="window.print()">Print Pick List</button></div>
<div class="doc-title">SALES ORDER - PICK LIST</div>
<div class="meta">
<div><b>Customer:</b> {m['customer']}</div><div><b>Order No:</b> {m['order_no']}</div>
<div><b>Order Date:</b> {m['order_date']}</div><div><b>Source File:</b> {m['source_file']}</div>
<div><b>Order Type:</b> {m['order_type']}</div><div><b>Currency:</b> {m['currency']}</div>
</div>
<table><thead><tr>
<th style="width:32px">SI</th><th style="width:100px">GTIN/Code</th>
<th class="l">Description</th><th style="width:44px">UOM</th>
<th style="width:40px">Check</th><th style="width:46px">Qty</th><th style="width:52px">Picked</th>
</tr></thead>
<tbody>{rows}
<tr class="tot"><td colspan="5" style="text-align:right;padding-right:10px">Total Quantity:</td>
<td class="c">{m['total_qty']}</td><td></td></tr>
</tbody></table>
<div class="sigs">
<div class="sig">Prepared By: ____________________</div>
<div class="sig">Checked By: ____________________</div>
<div class="sig">Picked By: ____________________</div>
</div></div></body></html>"""

# Render pick list preview
iframe_h = max(500, min(1100, 400 + len(data["items"]) * 30))
st.markdown('<div style="background:white;border-radius:12px;padding:18px 20px 6px;box-shadow:0 1px 6px rgba(0,0,0,.07);"><p style="font-weight:700;color:#1B3A6B;font-size:14px;margin:0 0 12px">Pick List Preview and Print</p>', unsafe_allow_html=True)
components.html(_build_print_html(data), height=iframe_h, scrolling=True)
st.markdown("</div>", unsafe_allow_html=True)

# Download Excel
st.markdown("<div style='margin-top:14px'>", unsafe_allow_html=True)
st.download_button(
    label="Download Pick List (.xlsx)",
    data=xlsx_bytes,
    file_name=f"PickList_{name}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
st.markdown("</div>", unsafe_allow_html=True)

with st.expander("Raw extracted JSON"):
    st.json(data)
