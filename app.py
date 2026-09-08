"""
app.py – IEC/ISO PDF Ultimate Processor
Alur:
  1. Trim PDF (engine11.py) -> Potong bahasa Prancis
  2. Clean Footer (engine12.py) -> Hapus nomor halaman & copyright IEC
  3. Convert to DOCX (engine13.py) -> Siap diterjemahkan
"""

import os
import streamlit as st
import logging
import fitz
from engine11 import PDFTrimmerEngine
from engine12 import TextCleanerEngine
from engine13 import PDFConverterEngine
from engine14 import SNIEnglishFormatterEngine
from engine15 import IECFinalPolishEngine
from engine16 import IECStyleIdentificationEngine

# ── Setup Logging ─────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO)

# ── Konfigurasi Halaman ────────────────────────────────────────────────────
st.set_page_config(page_title="IEC to ISO Converter", page_icon="📄", layout="centered")

# ── Inisialisasi Engine ────────────────────────────────────────────────────
trimmer = PDFTrimmerEngine()
cleaner = TextCleanerEngine(bg_color=(1, 1, 1))

DEFAULT_TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
converter = PDFConverterEngine(tesseract_path=DEFAULT_TESSERACT_PATH)
formatter = SNIEnglishFormatterEngine()
polisher = IECFinalPolishEngine()
style_identifier = IECStyleIdentificationEngine()


def _first_page_contains_introduction(pdf_path):
    """Verifikasi pengaman agar Introduction tidak hilang antar-engine."""
    doc = fitz.open(pdf_path)
    try:
        if len(doc) == 0:
            return False
        lines = doc[0].get_text().splitlines()
        return any(line.strip().upper() == "INTRODUCTION" for line in lines)
    finally:
        doc.close()

# ── Buat Folder Temp ──────────────────────────────────────────────────────
if not os.path.exists("temp"):
    os.makedirs("temp")

# ── UI Utama ───────────────────────────────────────────────────────────────
st.markdown(
    "<h1 style='text-align: center;'>IEC to ISO Converter</h1>",
    unsafe_allow_html=True
)

st.divider()

# ── Tombol Remove Watermark ───────────────────────────────────────────────
st.link_button(
    label="📘 Generator RSNI", 
    url="https://generator-rsni.streamlit.app", 
    use_container_width=True
)

with st.form("iec_converter_form", clear_on_submit=False):
    uploaded_file = st.file_uploader("Upload Dokumen IEC", type=["pdf"])

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        process_btn = st.form_submit_button(
            "🚀 Proses",
            type="primary",
            use_container_width=True,
        )

if process_btn:
    if uploaded_file is None:
        st.warning("Silakan upload dokumen IEC terlebih dahulu.")
    else:
        base_filename = os.path.splitext(uploaded_file.name)[0]
        
        input_pdf = os.path.join("temp", uploaded_file.name)
        trimmed_pdf = os.path.join("temp", f"{base_filename}_1_trimmed.pdf")
        cleaned_pdf = os.path.join("temp", f"{base_filename}_2_cleaned.pdf")
        engine13_docx = os.path.join("temp", f"{base_filename}_3_engine13.docx")
        engine14_docx = os.path.join("temp", f"{base_filename}_4_engine14.docx")
        engine15_docx = os.path.join("temp", f"{base_filename}_5_engine15.docx")
        output_docx = os.path.join("temp", f"{base_filename}_6_engine16.docx")

        with open(input_pdf, "wb") as f:
            f.write(uploaded_file.getbuffer())

        with st.status("Memproses...", expanded=False) as status:
            
            try:
                status.update(label="Tahap 1/6...")
                success_trim, res_trim, info_trim = trimmer.trim(input_path=input_pdf, output_path=trimmed_pdf)
                
                if not success_trim:
                    st.error(res_trim)
                    status.update(label="Gagal", state="error")
                else:
                    if info_trim.get("introduction_found") and not _first_page_contains_introduction(trimmed_pdf):
                        raise RuntimeError(
                            "Verifikasi gagal: halaman Introduction hilang setelah trim."
                        )

                    status.update(label="Tahap 2/6...")
                    success_clean, res_clean, info_clean = cleaner.clean_iec_footer(input_path=trimmed_pdf, output_path=cleaned_pdf)
                    
                    if not success_clean:
                        st.error(res_clean)
                        status.update(label="Gagal", state="error")
                    else:
                        if not info_clean.get("verified_no_bsn_notice", False):
                            raise RuntimeError(
                                "Validasi gagal: notice/watermark BSN masih terdeteksi pada PDF hasil pembersihan."
                            )
                        if not info_clean.get("verified_no_bsn_logo", False):
                            raise RuntimeError(
                                "Validasi gagal: logo BSN masih terdeteksi pada PDF hasil pembersihan."
                            )
                        if not info_clean.get("verified_no_iec_running_header", False):
                            raise RuntimeError(
                                "Validasi gagal: nomor halaman/header copyright IEC masih tersisa."
                            )
                        if info_trim.get("introduction_found") and not _first_page_contains_introduction(cleaned_pdf):
                            raise RuntimeError(
                                "Verifikasi gagal: halaman Introduction hilang setelah pembersihan."
                            )

                        status.update(label="Tahap 3/6 — konversi PDF → DOCX, layout dipertahankan...")
                        success_conv, res_conv, mode_conv = converter.convert(cleaned_pdf, engine13_docx)
                        
                        if not success_conv:
                            st.error(res_conv)
                            status.update(label="Gagal", state="error")
                        else:
                            status.update(label="Tahap 4/6 — validasi struktur DOCX untuk Generator RSNI...")
                            success_fmt, res_fmt, info_fmt = formatter.process(
                                engine13_docx, engine14_docx
                            )
                            if not success_fmt:
                                st.error(res_fmt)
                                status.update(label="Gagal", state="error")
                            else:
                                # Jangan izinkan output diteruskan bila engine14
                                # melaporkan artefak BSN masih tersisa.
                                if info_fmt.get("bsn_residual_verified") is not False:
                                    raise RuntimeError(
                                        "Validasi akhir gagal: status pembersihan watermark/notice BSN tidak valid."
                                    )
                                if not info_fmt.get("iec_layout_preserve", False):
                                    raise RuntimeError(
                                        "Validasi akhir gagal: dokumen tidak memiliki penanda IEC fidelity untuk Generator RSNI."
                                    )

                                status.update(label="Tahap 5/6 — perbaikan judul dan heading...")
                                success_polish, res_polish, info_polish = polisher.process(
                                    engine14_docx, engine15_docx, source_pdf=cleaned_pdf
                                )
                                if not success_polish:
                                    raise RuntimeError(f"Engine15 Error: {res_polish}")
                                if not info_polish.get("title_single_paragraph", False):
                                    raise RuntimeError("Engine15: judul utama belum menjadi satu paragraf.")

                                status.update(label="Tahap 6/6 — identifikasi style Judul/Pasal tanpa TOC...")
                                success_style, res_style, info_style = style_identifier.process(
                                    engine15_docx, output_docx
                                )
                                if not success_style:
                                    raise RuntimeError(res_style)
                                if not info_style.get("generator_rsni_ready", False):
                                    raise RuntimeError("Engine16: output belum tervalidasi kompatibel dengan Generator RSNI.")
                                if info_style.get("toc_created") is not False:
                                    raise RuntimeError("Engine16: validasi gagal karena TOC tidak boleh dibuat.")

                                with open(output_docx, "rb") as f:
                                    st.session_state["docx_bytes"] = f.read()

                                st.session_state["download_filename"] = f"{base_filename}_RESULT.docx"
                                st.success(
                                    "DOCX siap untuk Generator RSNI. Layout PDF dipertahankan "
                                    "sedapat mungkin dan watermark/notice BSN telah diverifikasi bersih."
                                )
                                with st.expander("🔎 Hasil validasi pembersihan", expanded=False):
                                    c1, c2, c3 = st.columns(3)
                                    c1.metric("Logo BSN dihapus", info_clean.get("bsn_logos", 0))
                                    c2.metric("Notice BSN dihapus", info_clean.get("bsn_notice_blocks", 0))
                                    c3.metric("Sisa watermark", "0" if info_clean.get("verified_no_bsn_notice") and info_clean.get("verified_no_bsn_logo") and info_fmt.get("bsn_residual_verified") is False else "GAGAL")
                                    st.caption(
                                        "Mode layout: ukuran halaman, tabel, dan gambar dipertahankan; page break artefak PDF dinormalisasi dan hanya Annex/section yang memaksa ganti halaman."
                                    )
                                status.update(label="✅ Selesai — output tervalidasi", state="complete")
                            
            except Exception as e:
                st.error(str(e))
                status.update(label="Error", state="error")
                
            finally:
                for path in [input_pdf, trimmed_pdf, cleaned_pdf, engine13_docx, engine14_docx, engine15_docx, output_docx]:
                    try:
                        if path and os.path.exists(path): os.remove(path)
                    except Exception:
                        pass

# ── Panel Download ─────────────────────────────────────────────────────────
if "docx_bytes" in st.session_state:
    st.divider()
    
    st.download_button(
        label="⬇️ Download .DOCX",
        data=st.session_state["docx_bytes"],
        file_name=st.session_state["download_filename"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
