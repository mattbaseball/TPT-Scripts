import streamlit as st
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches, Pt
import tempfile
import os
from PyPDF2 import PdfReader

st.set_page_config(page_title="PDF → PowerPoint Converter", layout="centered")

st.title("📑➡️📊 PDF to PowerPoint Converter")
st.write("Upload one or more PDFs. Each PDF page will become a slide with the **same proportions as the PDF**.")

uploaded_files = st.file_uploader(
    "Choose PDF files",
    type=["pdf"],
    accept_multiple_files=True
)

dpi = st.slider("Image DPI (higher = better quality, bigger file)", 100, 400, 200)

def pdf_to_pptx(pdf_path, dpi=200):
    # Get PDF page size (points = 1/72 inch)
    reader = PdfReader(pdf_path)
    first_page = reader.pages[0]
    width_pt = float(first_page.mediabox.width)
    height_pt = float(first_page.mediabox.height)
    
    # Convert to Inches (pptx uses EMU internally)
    pdf_width_in = width_pt / 72
    pdf_height_in = height_pt / 72

    # Convert PDF pages to images
    pages = convert_from_path(pdf_path, dpi=dpi)

    # Create PowerPoint
    prs = Presentation()
    prs.slide_width = Inches(pdf_width_in)
    prs.slide_height = Inches(pdf_height_in)

    for page in pages:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
            page.save(tmp_img.name, "PNG")
            img_path = tmp_img.name

        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(img_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
        os.remove(img_path)

    # Save to temp file
    out_pptx = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(out_pptx.name)
    return out_pptx.name

if uploaded_files:
    for uploaded_file in uploaded_files:
        with st.spinner(f"Processing {uploaded_file.name}..."):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf.write(uploaded_file.read())
                tmp_pdf_path = tmp_pdf.name

            pptx_path = pdf_to_pptx(tmp_pdf_path, dpi=dpi)

            st.success(f"Finished: {uploaded_file.name}")
            st.download_button(
                label=f"⬇️ Download {uploaded_file.name.replace('.pdf','.pptx')}",
                data=open(pptx_path, "rb").read(),
                file_name=uploaded_file.name.replace(".pdf", ".pptx"),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

            os.remove(tmp_pdf_path)
            os.remove(pptx_path)
