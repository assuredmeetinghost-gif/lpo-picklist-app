#!/usr/bin/env python3
"""
LPO -> Pick List   (Solution 5: Custom App + Gen AI Vision)
===========================================================
Drop any LPO / purchase order / handwritten order form (PDF or image) in and get a
Sales Order Pick List (Excel + PDF) out.

Pipeline (per the Solution Report):
  1. Load the PDF/image.
  2. Render each page to an image (auto-rotates scanned/handwritten forms).
  3. Send each page to Claude Vision -> structured JSON (gtin, description, qty...).
  4. Validate + merge across pages.
  5. Fill the Sales Order Pick List template -> .xlsx and .pdf.

Usage:
  python lpo_to_picklist.py "path/to/LPO.pdf"
  python lpo_to_picklist.py "path/to/LPO.pdf" --only-with-qty   (default for order forms)
  python lpo_to_picklist.py --watch          (watch the input/ folder, auto-process new files)

API key:  set environment variable  ANTHROPIC_API_KEY  (or put it in config.json -> "api_key")
"""
import os, sys, io, json, base64, time, argparse, re, pathlib, hashlib

import fitz                      # PyMuPDF
from PIL import Image
import numpy as np
try:
    from scipy import ndimage
    HAVE_SCIPY = True
except Exception:
    HAVE_SCIPY = False
import anthropic

import picklist_builder

HERE = pathlib.Path(__file__).resolve().parent
CONFIG = {}
_cfg = HERE / "config.json"
if _cfg.exists():
    CONFIG = json.loads(_cfg.read_text(encoding="utf-8"))

DEFAULT_MODEL = CONFIG.get("model", "claude-opus-4-8")
DPI = int(CONFIG.get("dpi", 220))
PROMPT_VERSION = "v3-tiled"    # bump to invalidate cache when the prompt/method changes
CACHE_DIR = HERE / "cache"

EXTRACT_PROMPT = """You are an expert data-extraction engine for a food-distribution warehouse.
You are looking at ONE page of a purchase order / LPO / vendor order form. It may be a clean
printed invoice OR a scanned, possibly rotated, handwritten order sheet (a printed product
catalogue where the customer hand-writes an order quantity next to the items they want).

Extract the ORDER LINE ITEMS as JSON. Rules:
- The page may be rotated 90/180 degrees - read it in whatever orientation makes the text legible.
- Columns are typically: SL.NO, GTIN/BARCODE, DESCRIPTION, IT CODE, QTY/CTN (pack size), QTY (order qty).
- ONLY include a row if it has an ORDER QUANTITY filled in (handwritten or typed in the order/QTY
  column). Skip every catalogue row that has a blank order-quantity cell.  (If this is a normal
  printed invoice where every listed line is part of the order, include them all.)
- Do NOT confuse the printed "QTY/CTN" (pack/carton size) with the ordered quantity. The ordered
  quantity is the value the customer entered; on handwritten forms it is the pen/ink number.
- Read handwritten digits carefully. Distinguish 6 vs 0, 4 vs 9, 5 vs 6, and 2-digit numbers (e.g. 60 vs 6).
- IGNORE summary/footer rows such as "Total", "Total Quantity", "Total Amount", "Grand Total".
  Never output a line whose quantity is actually a column total. A per-line ordered qty is small.
- The page may be rotated; the QTY column is the LAST data column. Match each qty to the SAME row's description.

Return STRICT JSON only, no prose, no markdown fences:
{
  "header": {"customer": "", "vendor": "", "order_no": "", "order_date": "", "order_type": "", "currency": ""},
  "items": [
    {"gtin": "", "desc": "", "it_code": "", "qty": <number>, "uom": ""}
  ]
}
Use "" for unknown header fields. desc = full product description. qty = ordered quantity as a number.
If the page has NO order lines, return {"header":{...}, "items": []}.
"""


def log(m): print(m, flush=True)


def get_client():
    key = os.environ.get("ANTHROPIC_API_KEY") or CONFIG.get("api_key")
    if not key:
        sys.exit("ERROR: set ANTHROPIC_API_KEY env var (or add \"api_key\" to config.json).")
    return anthropic.Anthropic(api_key=key)


def render_pages(path):
    """Return list of PIL.Image, one per page (no rotation here)."""
    ext = pathlib.Path(path).suffix.lower()
    imgs = []
    if ext == ".pdf":
        doc = fitz.open(path)
        for page in doc:
            pix = page.get_pixmap(dpi=DPI)
            imgs.append(Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB"))
    else:
        imgs.append(Image.open(path).convert("RGB"))
    return imgs


def to_b64(img, max_px=2000):
    w, h = img.size
    if max(w, h) > max_px:
        s = max_px / max(w, h)
        img = img.resize((int(w * s), int(h * s)))
    buf = io.BytesIO(); img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()


LAYOUT_PROMPT = """Look at this scanned order form / invoice and reply with STRICT JSON only:
{"rotation": <0|90|180|270>, "columns": <1|2>}
- rotation = degrees to rotate the image COUNTER-CLOCKWISE so the printed text becomes upright/readable.
- columns = 2 if the page is laid out as TWO separate side-by-side tables (each with its own
  GTIN/DESCRIPTION/QTY header and its own rows). Otherwise 1.
No prose, JSON only."""


def detect_layout(client, model, img):
    thumb = img.copy(); thumb.thumbnail((1100, 1100))
    try:
        msg = client.messages.create(
            model=model, max_tokens=200,
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                             "data": to_b64(thumb, 1100)}},
                {"type": "text", "text": LAYOUT_PROMPT}]}])
        t = "".join(b.text for b in msg.content if b.type == "text")
        t = re.sub(r"^```(json)?|```$", "", t.strip(), flags=re.M).strip()
        d = json.loads(t)
        rot = int(d.get("rotation", 0)) % 360
        cols = 2 if int(d.get("columns", 1)) >= 2 else 1
        return rot, cols
    except Exception:
        return 0, 1


def make_tiles(img, columns, bands=3, overlap=0.07):
    """Split into 1-or-2 vertical columns, each into <bands> overlapping horizontal strips."""
    w, h = img.size
    if columns >= 2:
        ov = int(w * 0.03); cut = w // 2
        col_imgs = [img.crop((0, 0, cut + ov, h)), img.crop((cut - ov, 0, w, h))]
    else:
        col_imgs = [img]
    # fewer bands for short pages
    tiles = []
    for col in col_imgs:
        cw, ch = col.size
        nb = bands if ch > 1300 else (2 if ch > 700 else 1)
        bh = ch // nb; ov = int(ch * overlap)
        for b in range(nb):
            y0 = max(0, b * bh - ov)
            y1 = ch if b == nb - 1 else min(ch, (b + 1) * bh + ov)
            tiles.append(col.crop((0, y0, cw, y1)))
    return tiles


def extract_image(client, img, model, tag):
    b64 = to_b64(img)
    for attempt in range(3):
        try:
            msg = client.messages.create(
                model=model, max_tokens=8000,
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {"type": "base64",
                        "media_type": "image/png", "data": b64}},
                    {"type": "text", "text": EXTRACT_PROMPT},
                ]}])
            text = "".join(b.text for b in msg.content if b.type == "text").strip()
            text = re.sub(r"^```(json)?|```$", "", text, flags=re.M).strip()
            return json.loads(text)
        except json.JSONDecodeError:
            pass
        except Exception as e:
            log(f"   {tag}: API error ({e}); retry {attempt+1}")
            time.sleep(2 * (attempt + 1))
    log(f"   {tag}: FAILED")
    return {"header": {}, "items": []}


def extract_page_tiled(client, img, model, page_no, progress=None):
    """Orient, split into tiles, extract each tile, return merged page result."""
    rot, cols = detect_layout(client, model, img)
    if rot:
        img = img.rotate(rot, expand=True)
    tiles = make_tiles(img, cols)
    msg = f"   page {page_no}: orientation {rot}deg, {cols} column(s), {len(tiles)} tiles"
    log(msg)
    if progress: progress(msg)
    header, items = {}, []
    for ti, t in enumerate(tiles, 1):
        d = extract_image(client, t, model, f"p{page_no}.t{ti}")
        for k, v in (d.get("header") or {}).items():
            if v and not header.get(k):
                header[k] = v
        items += d.get("items", [])
        if progress: progress(f"   page {page_no}: tile {ti}/{len(tiles)}")
    log(f"   page {page_no}: {len(items)} raw line(s) before de-dup")
    return {"header": header, "items": items}


def _norm(s):
    return re.sub(r"\s+", " ", str(s)).strip().upper()


def merge(results, src_name):
    header = {}
    for r in results:
        for k, v in (r.get("header") or {}).items():
            if v and not header.get(k):
                header[k] = v
    items, seen, conflicts = [], {}, 0
    for pi, r in enumerate(results, 1):
        for it in r.get("items", []):
            try:
                q = float(it.get("qty"))
            except (TypeError, ValueError):
                continue
            if q <= 0:
                continue
            q = int(q) if q == int(q) else q
            desc = str(it.get("desc", "")).strip()
            key = _norm(desc) + "|" + _norm(it.get("it_code", "") or it.get("gtin", ""))
            if key in seen:                       # same row seen in an overlapping tile
                prev = items[seen[key]]
                if prev["qty"] != q:
                    conflicts += 1
                    prev.setdefault("_alt_qty", []).append(q)
                continue
            seen[key] = len(items)
            items.append({"gtin": str(it.get("gtin", "")).strip(),
                          "desc": desc,
                          "it_code": str(it.get("it_code", "")).strip(),
                          "uom": it.get("uom") or "NOS",
                          "qty": q,
                          "page": pi})
    if conflicts:
        log(f"   note: {conflicts} qty conflict(s) across overlapping tiles - review those lines")
    return {
        "source_file": src_name,
        "customer": header.get("customer", ""),
        "vendor": header.get("vendor", ""),
        "order_no": header.get("order_no", ""),
        "order_date": header.get("order_date", ""),
        "order_type": header.get("order_type", ""),
        "currency": header.get("currency", "OMR"),
        "items": items,
        "total_lines": len(items),
        "total_qty": sum(i["qty"] for i in items),
    }


def export_pdf(xlsx_path):
    xlsx_path = str(pathlib.Path(xlsx_path).resolve())
    pdf_path = str(pathlib.Path(xlsx_path).with_suffix(".pdf"))
    # 1) Excel COM (Windows + Excel installed)
    try:
        import win32com.client as win32
        xl = win32.Dispatch("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(xlsx_path); xl.CalculateFull(); wb.Save()
        wb.ExportAsFixedFormat(0, pdf_path); wb.Close(False); xl.Quit()
        return pdf_path
    except Exception:
        pass
    # 2) Excel via PowerShell (Windows with Excel but no pywin32)
    import shutil, subprocess
    if os.name == "nt":
        ps = ('$xl=New-Object -ComObject Excel.Application;$xl.Visible=$false;'
              '$xl.DisplayAlerts=$false;$wb=$xl.Workbooks.Open(\'%s\');$xl.CalculateFull();'
              '$wb.Save();$wb.ExportAsFixedFormat(0,\'%s\');$wb.Close($false);$xl.Quit()'
              % (xlsx_path.replace("'", "''"), pdf_path.replace("'", "''")))
        try:
            subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                           check=True, timeout=120,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if pathlib.Path(pdf_path).exists():
                return pdf_path
        except Exception:
            pass
    # 3) LibreOffice
    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if soffice:
        try:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                            str(pathlib.Path(xlsx_path).parent), xlsx_path],
                           check=True, timeout=120)
            return pdf_path
        except Exception:
            pass
    log("   (PDF export skipped - open the .xlsx and Save As PDF, or install Excel/LibreOffice)")
    return None


# ----------------------------------------------------------------------------- caching
def file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    h.update(f"|{DPI}|{PROMPT_VERSION}".encode())   # bound to render + prompt version
    return h.hexdigest()[:16]


def cache_file(path, model):
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return CACHE_DIR / f"{file_hash(path)}_{model}.json"


def extract_data(path, model=DEFAULT_MODEL, use_cache=True, progress=None):
    """Render -> Vision -> merged data dict. Cached by file content + model.
    progress(msg) optional callback for UI. Returns the merged-data dict (with 'from_cache')."""
    path = str(path)
    def say(m):
        log(m)
        if progress: progress(m)
    cf = cache_file(path, model)
    if use_cache and cf.exists():
        say("   cache hit - reusing previous extraction (no API call)")
        data = json.loads(cf.read_text(encoding="utf-8"))
        data["from_cache"] = True
        return data
    client = get_client()
    pages = render_pages(path)
    say(f"   {len(pages)} page(s) @ {DPI} dpi - reading with {model} ...")
    results = []
    for i, im in enumerate(pages):
        results.append(extract_page_tiled(client, im, model, i + 1, progress))
    data = merge(results, pathlib.Path(path).name)
    data["from_cache"] = False
    cf.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    return data


def build_outputs(data, out_dir, name, make_pdf=True):
    out_dir = pathlib.Path(out_dir); out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / f"extracted_{name}.json").write_text(
        json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    xlsx_path = str(out_dir / f"PickList_{name}.xlsx")
    picklist_builder.build(data, xlsx_path)
    pdf_path = export_pdf(xlsx_path) if make_pdf else None
    return xlsx_path, pdf_path


def process(path, out_dir, model=DEFAULT_MODEL, use_cache=True):
    path = str(path)
    name = pathlib.Path(path).stem
    log(f"-> {pathlib.Path(path).name}")
    data = extract_data(path, model, use_cache)
    if not data["items"]:
        log("   no ordered items found - nothing to build."); return None
    xlsx_path, pdf_path = build_outputs(data, out_dir, name)
    tag = " (from cache)" if data.get("from_cache") else ""
    log(f"   DONE{tag}: {data['total_lines']} items, total qty {data['total_qty']}")
    log(f"   xlsx: {xlsx_path}")
    if pdf_path: log(f"   pdf : {pdf_path}")
    return xlsx_path


def watch(in_dir, out_dir, model):
    in_dir = pathlib.Path(in_dir); seen = set()
    log(f"Watching {in_dir} ... drop LPO PDFs/images here (Ctrl+C to stop).")
    while True:
        for f in in_dir.iterdir():
            if f.suffix.lower() in (".pdf", ".png", ".jpg", ".jpeg") and f.name not in seen:
                seen.add(f.name)
                try:
                    process(f, out_dir, model)
                except Exception as e:
                    log(f"   ERROR on {f.name}: {e}")
        time.sleep(3)


def main():
    ap = argparse.ArgumentParser(description="LPO -> Sales Order Pick List (Gen AI Vision)")
    ap.add_argument("file", nargs="?", help="LPO PDF or image")
    ap.add_argument("--out", default=str(HERE / "output"), help="output folder")
    ap.add_argument("--model", default=DEFAULT_MODEL)
    ap.add_argument("--watch", action="store_true", help="watch the input/ folder")
    ap.add_argument("--no-cache", action="store_true", help="ignore cached extraction, re-read")
    args = ap.parse_args()
    if args.watch:
        watch(HERE / "input", args.out, args.model)
    elif args.file:
        process(args.file, args.out, args.model, use_cache=not args.no_cache)
    else:
        ap.print_help()


if __name__ == "__main__":
    main()
