# Update the Streamlit app to support DOCX uploads (auto-convert to PDF via docx2pdf or LibreOffice fallback),
# add a sidebar "Clear all / Start fresh" button using st.rerun(), and keep existing features.
# Saves to /mnt/data/tpt_preview_app.py for download.

import io
import os
import math
import zipfile
import tempfile
import shutil
import subprocess
from pathlib import Path
from dataclasses import dataclass
from typing import Dict, List, Tuple

import streamlit as st

# Deps:
#   pip install streamlit pymupdf pillow docx2pdf
import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont

# ---------------------------- DOCX -> PDF Conversion ----------------------------

def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes:
    """
    Convert DOCX bytes to PDF bytes.
    Tries docx2pdf (MS Word on Windows/macOS), falls back to LibreOffice (if installed).
    Raises RuntimeError if both methods fail.
    """
    tmpdir = Path(tempfile.mkdtemp(prefix="tpt_prev_"))
    in_path = tmpdir / "input.docx"
    out_path = tmpdir / "output.pdf"
    with open(in_path, "wb") as f:
        f.write(docx_bytes)

    # Try docx2pdf first
    try:
        from docx2pdf import convert as docx2pdf_convert  # type: ignore
        try:
            # docx2pdf can take (input, output)
            docx2pdf_convert(str(in_path), str(out_path))
        except Exception:
            # Some versions expect output directory instead of full path
            docx2pdf_convert(str(in_path), str(tmpdir))
        if out_path.exists():
            with open(out_path, "rb") as fpdf:
                data = fpdf.read()
            shutil.rmtree(tmpdir, ignore_errors=True)
            return data
    except Exception as e:
        # docx2pdf not available or failed — try LibreOffice
        pass

    # Fallback: LibreOffice (soffice) headless conversion
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(tmpdir), str(in_path)],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
            # LibreOffice names output as input.pdf
            lo_out = in_path.with_suffix(".pdf")
            if lo_out.exists():
                with open(lo_out, "rb") as fpdf:
                    data = fpdf.read()
                shutil.rmtree(tmpdir, ignore_errors=True)
                return data
        except Exception as e:
            pass

    # If we reach here, conversion failed
    shutil.rmtree(tmpdir, ignore_errors=True)
    raise RuntimeError("Unable to convert DOCX to PDF. Install Microsoft Word (for docx2pdf) or LibreOffice (soffice).")

# ---------------------------- Rendering & Watermarking ----------------------------

def load_font(size: int) -> ImageFont.FreeTypeFont:
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        "/System/Library/Fonts/SFNS.ttf",
        "C:\\Windows\\Fonts\\arialbd.ttf",
        "C:\\Windows\\Fonts\\ARIALBD.TTF",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size=size)
        except Exception:
            continue
    return ImageFont.load_default()

def pil_from_pix(pix: "fitz.Pixmap") -> Image.Image:
    mode = "RGBA" if pix.alpha else "RGB"
    return Image.frombytes(mode, [pix.width, pix.height], pix.samples)

def render_thumbnails(pdf_bytes: bytes, thumb_width: int = 220) -> List[bytes]:
    thumbs: List[bytes] = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        for page in doc:
            zoom = thumb_width / page.rect.width
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = pil_from_pix(pix)
            buf = io.BytesIO()
            img.save(buf, format="PNG", optimize=True)
            thumbs.append(buf.getvalue())
    finally:
        doc.close()
    return thumbs

def tile_watermark(img: Image.Image,
                   text: str,
                   opacity: float = 0.45,
                   angle_deg: float = 45.0,
                   font_size: int = 80,
                   coverage: float = 0.70) -> Image.Image:
    if not text.strip():
        return img
    w, h = img.size
    font = load_font(font_size)

    # Measure text
    tmp_img = Image.new("RGBA", (1, 1), (0, 0, 0, 0))
    tmp_draw = ImageDraw.Draw(tmp_img)
    bbox = tmp_draw.textbbox((0, 0), text, font=font)
    text_w, text_h = bbox[2] - bbox[0], bbox[3] - bbox[1]

    # Tile containing rotated text
    tile_side = int(math.hypot(text_w, text_h)) + 10
    tile = Image.new("RGBA", (tile_side, tile_side), (0, 0, 0, 0))
    tdraw = ImageDraw.Draw(tile)
    tx = (tile_side - text_w) // 2
    ty = (tile_side - text_h) // 2
    fill = (0, 0, 0, int(255 * opacity))
    tdraw.text((tx, ty), text, font=font, fill=fill)
    tile = tile.rotate(angle_deg, expand=True)

    # Spacing based on coverage
    step = int(max(40, min(tile.size) * (1.0 - coverage * 0.85)))
    step_x = max(40, tile.size[0] - step)
    step_y = max(40, tile.size[1] - step)

    overlay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    for y in range(-tile.size[1], h + tile.size[1], step_y):
        for x in range(-tile.size[0], w + tile.size[0], step_x):
            overlay.alpha_composite(tile, dest=(x, y))

    result = Image.alpha_composite(img.convert("RGBA"), overlay)
    return result.convert("RGB")

def rasterize_pages_with_watermark(pdf_bytes: bytes,
                                   pages_to_keep: List[int],
                                   dpi: int,
                                   wm_text: str,
                                   wm_opacity: float,
                                   wm_angle: float,
                                   wm_font_size: int,
                                   wm_coverage: float) -> List[Image.Image]:
    images: List[Image.Image] = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        if not pages_to_keep:
            return images
        for pnum in sorted(set(pages_to_keep)):
            if pnum < 1 or pnum > len(doc):
                continue
            page = doc[pnum - 1]
            zoom = dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = pil_from_pix(pix)
            img = tile_watermark(
                img,
                text=wm_text,
                opacity=wm_opacity,
                angle_deg=wm_angle,
                font_size=wm_font_size,
                coverage=wm_coverage,
            )
            images.append(img.convert("RGB"))
    finally:
        doc.close()
    return images

def images_to_pdf_bytes(images: List[Image.Image]) -> bytes:
    buf = io.BytesIO()
    if not images:
        return buf.getvalue()
    images[0].save(buf, format="PDF", save_all=True, append_images=images[1:])
    return buf.getvalue()

# ---------------------------- Data Model ----------------------------

@dataclass
class FileEntry:
    name: str              # original uploaded name (may be .pdf or .docx)
    size: int              # original upload size
    bytes_pdf: bytes       # converted PDF bytes (for .docx) or original PDF bytes
    page_count: int
    thumbs: List[bytes]
    selected_pages: List[int]

def make_key(name: str, data: bytes) -> str:
    return f"{name}::{len(data)}"

def ensure_state():
    if "files" not in st.session_state:
        st.session_state.files = {}

# ---------------------------- App ----------------------------

st.set_page_config(page_title="TPT Preview Maker", page_icon="🧲", layout="wide")
st.title("🛡️ TPT Preview Maker (Rasterized, DOCX/PDF, Heavily Watermarked)")

with st.expander("About this tool", expanded=False):
    st.markdown("""
- **Imports:** PDF or **Word (.docx)** — DOCX auto-converts to PDF (requires Microsoft Word or LibreOffice).
- **Visual selection:** Pick pages using thumbnails (you can also type ranges). Selecting none is allowed.
- **Dense tiled watermarks:** Covers most of the page; tweak **coverage**, **opacity**, **angle**, and **font size**.
- **Outputs:** Keeps original file names, adds `_preview.pdf`, can bundle into ZIP, and can **merge into a single PDF**.
    """)

ensure_state()

# Sidebar session controls
st.sidebar.subheader("⚙️ Session Controls")
if st.sidebar.button("🗑️ Clear all uploaded files / Start fresh"):
    st.session_state.files = {}
    st.rerun()

uploaded = st.file_uploader("Upload one or more DOCX/PDF files", type=["pdf", "docx"], accept_multiple_files=True)

# Ingest uploads
if uploaded:
    for uf in uploaded:
        data = uf.getvalue()
        key = make_key(uf.name, data)
        if key not in st.session_state.files:
            # Convert DOCX to PDF if needed
            ext = Path(uf.name).suffix.lower()
            if ext == ".docx":
                try:
                    pdf_bytes = convert_docx_to_pdf_bytes(data)
                except Exception as e:
                    st.error(f"Failed to convert {uf.name} to PDF: {e}")
                    continue
            else:
                pdf_bytes = data

            # Analyze converted/original PDF
            try:
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                page_count = len(doc)
                doc.close()
                thumbs = render_thumbnails(pdf_bytes, thumb_width=220)
            except Exception as e:
                st.error(f"Unable to read {uf.name} as PDF after conversion: {e}")
                continue

            st.session_state.files[key] = FileEntry(
                name=uf.name,
                size=len(data),
                bytes_pdf=pdf_bytes,
                page_count=page_count,
                thumbs=thumbs,
                selected_pages=[],
            )

if not st.session_state.files:
    st.info("Upload PDFs or DOCX files to begin.")
    st.stop()

# Global Controls
st.subheader("Global selection helpers")
left, right = st.columns(2)
with left:
    apply_first_n = st.checkbox("Apply 'First N' selection to all files", value=False)
    N = st.number_input("First N pages (if applied)", min_value=1, value=3, step=1)
with right:
    clear_all = st.checkbox("Clear selection for all files", value=False)
    select_all_pages = st.checkbox("Select all pages for all files", value=False)

if apply_first_n:
    for k, entry in st.session_state.files.items():
        entry.selected_pages = list(range(1, min(entry.page_count, int(N)) + 1))
if clear_all:
    for k, entry in st.session_state.files.items():
        entry.selected_pages = []
if select_all_pages:
    for k, entry in st.session_state.files.items():
        entry.selected_pages = list(range(1, entry.page_count + 1))

# Per-file UI
st.subheader("Per-file page selection")
for key, entry in list(st.session_state.files.items()):
    with st.container(border=True):
        st.markdown(f"### {entry.name} — {entry.page_count} pages")
        c1, c2, c3, c4, c5 = st.columns([1,1,1,2,1])
        with c1:
            if st.button("All", key=f"all_{key}"):
                entry.selected_pages = list(range(1, entry.page_count + 1))
        with c2:
            if st.button("1st 3", key=f"first3_{key}"):
                entry.selected_pages = list(range(1, min(3, entry.page_count) + 1))
        with c3:
            if st.button("None", key=f"none_{key}"):
                entry.selected_pages = []
        with c4:
            rng = st.text_input("Or type ranges (e.g., 1-3,6,9-10)", key=f"ranges_{key}")
            if rng:
                parts = [p.strip() for p in rng.split(",") if p.strip()]
                sel = set(entry.selected_pages)
                for part in parts:
                    if "-" in part:
                        try:
                            a, b = part.split("-", 1)
                            a, b = int(a), int(b)
                            for i in range(a, b + 1):
                                if 1 <= i <= entry.page_count:
                                    sel.add(i)
                        except Exception:
                            pass
                    else:
                        try:
                            i = int(part)
                            if 1 <= i <= entry.page_count:
                                sel.add(i)
                        except Exception:
                            pass
                entry.selected_pages = sorted(sel)
        with c5:
            if st.button("Remove file", key=f"rm_{key}"):
                del st.session_state.files[key]
                st.rerun()

        cols = st.columns(5)
        for i in range(entry.page_count):
            col = cols[i % 5]
            with col:
                if i < len(entry.thumbs):
                    st.image(entry.thumbs[i], use_container_width=True)
                label = f"Pg {i+1}"
                key_cb = f"cb_{key}_{i+1}"
                current = (i + 1) in entry.selected_pages
                new_val = st.checkbox(label, value=current, key=key_cb)
                if new_val and (i + 1) not in entry.selected_pages:
                    entry.selected_pages.append(i + 1)
                elif not new_val and (i + 1) in entry.selected_pages:
                    entry.selected_pages.remove(i + 1)
        entry.selected_pages = sorted(set(entry.selected_pages))

# Watermark settings
st.subheader("Watermark")
wm_cols = st.columns(6)
with wm_cols[0]:
    wm_text = st.text_input("Text", value="Preview - TrueNorthTeachingTools")
with wm_cols[1]:
    wm_opacity = st.slider("Opacity", 0.10, 0.90, 0.50, 0.01)
with wm_cols[2]:
    wm_angle = st.number_input("Angle", -180, 180, 45, 1)
with wm_cols[3]:
    wm_font = st.number_input("Font size", 24, 200, 92, 2)
with wm_cols[4]:
    wm_coverage = st.slider("Coverage (dense ↔︎ light)", 0.20, 1.00, 0.80, 0.01)
with wm_cols[5]:
    dpi = st.selectbox("Raster DPI", options=[100, 120, 144, 150, 180, 200, 240, 300], index=5, help="Higher = crisper & larger file")

# Export settings
st.subheader("Export")
suffix = st.text_input("Output suffix", value="_preview")
make_zip = st.checkbox("Also create a ZIP with all outputs", value=True)
skip_empty = st.checkbox("Skip files with no selected pages", value=True)
merge_all = st.checkbox("🔗 Also merge all generated previews into a single PDF", value=True)
merge_order = st.selectbox("Merged order", options=["By file name (A→Z)", "By file size (small→large)"], index=0)

if st.button("🚀 Generate Previews", type="primary"):
    out_list: List[Tuple[str, bytes]] = []
    pages_per_file: Dict[str, List[Image.Image]] = []
    pages_map: Dict[str, List[Image.Image]] = {}

    zip_buf = io.BytesIO()
    zf = zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) if make_zip else None

    progress_files = st.progress(0, text="Processing files...")
    total_files = len(st.session_state.files)
    processed = 0

    for key, entry in st.session_state.files.items():
        processed += 1
        progress_files.progress(processed / total_files, text=f"Processing {entry.name} ({processed}/{total_files})")

        if not entry.selected_pages:
            if skip_empty:
                st.warning(f"Skipped (no pages): {entry.name}")
                continue

        pages = rasterize_pages_with_watermark(
            entry.bytes_pdf,
            pages_to_keep=entry.selected_pages,
            dpi=int(dpi),
            wm_text=wm_text,
            wm_opacity=float(wm_opacity),
            wm_angle=float(wm_angle),
            wm_font_size=int(wm_font),
            wm_coverage=float(wm_coverage),
        )
        if not pages:
            st.info(f"No pages generated for {entry.name}.")
            continue

        pdf_bytes = images_to_pdf_bytes(pages)
        base = Path(entry.name).stem  # strip .pdf or .docx
        out_name = f"{base}{suffix}.pdf"
        out_list.append((out_name, pdf_bytes))
        pages_map[out_name] = pages
        if zf:
            zf.writestr(out_name, pdf_bytes)

    if zf:
        zf.close()
        zip_buf.seek(0)

    if make_zip and out_list:
        st.download_button("⬇️ Download all (ZIP)", data=zip_buf, file_name="tpt_previews.zip", mime="application/zip")

    for name, data in out_list:
        st.download_button(f"Download {name}", data=data, file_name=name, mime="application/pdf")

    # Build merged PDF if requested
    if merge_all and pages_map:
        items = list(pages_map.items())
        if merge_order == "By file name (A→Z)":
            items.sort(key=lambda x: x[0].lower())
        else:
            # by size of original file name stem
            name_to_size = {}
            for k, entry in st.session_state.files.items():
                base = Path(entry.name).stem
                name_to_size[base] = entry.size
            def size_key(item):
                fname = item[0]
                base = Path(fname).stem.replace(suffix, "")
                return name_to_size.get(base, float('inf'))
            items.sort(key=size_key)

        merged_images: List[Image.Image] = []
        for fname, pages in items:
            merged_images.extend(pages)
        merged_bytes = images_to_pdf_bytes(merged_images)
        st.download_button("🔗 Download merged_previews.pdf", data=merged_bytes, file_name="merged_previews.pdf", mime="application/pdf")

    if not out_list and not (merge_all and pages_map):
        st.warning("No previews generated. Check your selections.")
