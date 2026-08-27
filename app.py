import streamlit as st
import streamlit.components.v1 as _components
import os
import re
import time
import glob
import atexit
import threading
import uuid
import importlib.util
import shutil
import subprocess
import csv
import io
from io import BytesIO
from urllib.request import Request, urlopen

# ─────────────────────────────────────────────────────────────────────────────
# AUTO-CLEANUP — pembersihan file temporer otomatis
# ─────────────────────────────────────────────────────────────────────────────

# Database kamus ada di https://bit.ly/kamusSNI

# Pola file temporer yang dibuat oleh aplikasi
_TEMP_PATTERNS = ["temp_main_*", "converted_*", "engine1_*", "engine2_*", "engine3_*", "engine4_*", "engine5_*", "engine6_*", "engine7_*", "engine8_*", "engine9_*", "opt_*", "cover_*", "di_*", "pp_*", "ip_*", "ID_*"]
# Hapus file lebih lama dari N menit
_MAX_AGE_MINUTES = 30

def _cleanup_temp_files(max_age_minutes: int = _MAX_AGE_MINUTES, silent: bool = True):
    """Hapus semua file temporer yang lebih lama dari max_age_minutes."""
    deleted, freed = 0, 0
    cutoff = time.time() - (max_age_minutes * 60)
    for pattern in _TEMP_PATTERNS:
        for fpath in glob.glob(pattern):
            try:
                if os.path.isfile(fpath) and os.path.getmtime(fpath) < cutoff:
                    size = os.path.getsize(fpath)
                    os.remove(fpath)
                    deleted += 1
                    freed += size
            except Exception:
                pass
    if not silent and deleted > 0:
        mb = freed / (1024 * 1024)
        print(f"[AutoCleanup] {deleted} file dihapus, {mb:.2f} MB dibebaskan.")
    return deleted, freed

def _cleanup_session_files(session_state):
    """Hapus file milik sesi saat ini segera."""
    keys = ['_target_file', '_final_opt_file', '_final_tr_file']
    for k in keys:
        fpath = session_state.get(k)
        if fpath and os.path.isfile(fpath):
            try:
                os.remove(fpath)
            except Exception:
                pass

def _remember_engine_output(engine_no: int, path: str) -> None:
    """Simpan output terakhir yang benar untuk recovery UI."""
    if path and os.path.isfile(path):
        st.session_state['_last_engine'] = engine_no
        st.session_state['_last_output'] = path

def _reset_to_start() -> None:
    """Kembali ke form awal tanpa menghapus cache global aplikasi."""
    for key in (
        '_run_process', '_show_results', '_process_error', '_resume_engine',
        '_last_engine', '_last_output', '_target_file', '_final_opt_file',
        '_final_tr_file', '_final_time', '_final_engine', '_doc_sections',
        '_force_continue', '_forced_errors', '_forced_summary',
    ):
        st.session_state.pop(key, None)


def _convert_legacy_doc(input_doc: str, output_docx: str) -> str:
    """Konversi Word 97-2003 .doc menjadi .docx sebelum Engine 1."""
    input_abs = os.path.abspath(input_doc)
    output_abs = os.path.abspath(output_docx)
    errors = []

    # Pilihan utama pada komputer Windows pengguna: Microsoft Word COM.
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        document = None
        try:
            document = word.Documents.Open(input_abs, ReadOnly=True)
            document.SaveAs2(output_abs, FileFormat=16)  # wdFormatDocumentDefault
        finally:
            if document is not None:
                document.Close(False)
            word.Quit()
            pythoncom.CoUninitialize()
        if os.path.isfile(output_abs):
            return output_abs
    except Exception as exc:
        errors.append(f'Microsoft Word: {exc}')

    # Fallback lintas platform bila LibreOffice/soffice tersedia.
    office = shutil.which('soffice') or shutil.which('libreoffice')
    if office:
        try:
            completed = subprocess.run(
                [office, '--headless', '--convert-to', 'docx',
                 '--outdir', os.path.dirname(output_abs), input_abs],
                capture_output=True, text=True, timeout=300, check=False,
            )
            generated = os.path.join(
                os.path.dirname(output_abs),
                os.path.splitext(os.path.basename(input_abs))[0] + '.docx',
            )
            if os.path.isfile(generated):
                if os.path.abspath(generated) != output_abs:
                    os.replace(generated, output_abs)
                return output_abs
            errors.append(
                'LibreOffice: ' + (completed.stderr or completed.stdout or
                                    f'exit {completed.returncode}')[:500]
            )
        except Exception as exc:
            errors.append(f'LibreOffice: {exc}')

    raise RuntimeError(
        'Konversi file .doc ke .docx gagal. Pastikan Microsoft Word dan '
        'pywin32 tersedia, atau instal LibreOffice. ' + ' | '.join(errors)
    )

def _start_background_cleanup():
    """Jalankan cleanup berkala di background thread (tiap 15 menit)."""
    def _loop():
        while True:
            time.sleep(15 * 60)
            _cleanup_temp_files(silent=False)
    t = threading.Thread(target=_loop, daemon=True)
    t.start()

# Jalankan background cleanup sekali saat modul pertama kali diload
if 'bg_cleanup_started' not in st.session_state:
    _cleanup_temp_files(silent=True)   # bersihkan sisa sesi sebelumnya
    _start_background_cleanup()
    st.session_state['bg_cleanup_started'] = True

# Daftarkan cleanup saat proses Python berhenti (atexit)
atexit.register(_cleanup_temp_files, max_age_minutes=0, silent=False)

# --- IMPORT ENGINE ---
from engine1 import IntroductionContentTrimmerEngine

# Utamakan nama standar engine2.py. Jika pengguna masih menyimpan file hasil
# unduhan dengan nama seperti engine2(6).py, cari dan muat file tersebut secara
# otomatis selama di dalamnya tersedia kelas CoverPageEngine.
try:
    from engine2 import CoverPageEngine
except ImportError as _engine2_import_error:
    CoverPageEngine = None
    _app_dir = os.path.dirname(os.path.abspath(__file__))
    for _engine2_path in sorted(glob.glob(os.path.join(_app_dir, "engine2*.py"))):
        if os.path.basename(_engine2_path).lower() == "engine2.py":
            continue
        try:
            _spec = importlib.util.spec_from_file_location(
                f"_engine2_fallback_{abs(hash(_engine2_path))}",
                _engine2_path,
            )
            if _spec is None or _spec.loader is None:
                continue
            _module = importlib.util.module_from_spec(_spec)
            _spec.loader.exec_module(_module)
            _candidate = getattr(_module, "CoverPageEngine", None)
            if _candidate is not None:
                CoverPageEngine = _candidate
                break
        except Exception:
            continue
    if CoverPageEngine is None:
        raise ImportError(
            "CoverPageEngine tidak ditemukan. Pastikan file Engine 2 yang "
            "benar disimpan sebagai engine2.py di folder yang sama dengan app.py."
        ) from _engine2_import_error

from engine3 import DaftarIsiEngine
from engine4 import PrakataPendahuluanEngine
from engine5 import IntroductionContentDuplicatorEngine
from engine6 import InfoPendukungEngine
from engine7 import DaftarIsiBibliographyFormatterEngine
from engine8 import (
    SelectiveTranslationEngine,
    CustomDictionary,
    ItalicDictionary,
)
from engine9 import TableOfContentsEngine

# Seluruh Engine 1–9 dipakai oleh pipeline aktif. Kamus dashboard berasal
# langsung dari Engine 8; Engine 9 hanya menangani style final dan TOC.

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Generator RSNI",
    page_icon="📑",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- CSS CUSTOM ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

/* ══════════════════════════════════════════
   BASE
══════════════════════════════════════════ */
html, body, [class*="css"] {
    font-family: 'Outfit', sans-serif !important;
}
#MainMenu, footer, header { visibility: hidden; }

.stApp {
    background: #0d0f1a;
    background-image:
        radial-gradient(ellipse 80% 50% at 20% 10%, rgba(99,102,241,0.12) 0%, transparent 60%),
        radial-gradient(ellipse 60% 40% at 80% 80%, rgba(16,185,129,0.08) 0%, transparent 55%),
        radial-gradient(ellipse 50% 60% at 50% 50%, rgba(244,63,94,0.04) 0%, transparent 70%);
    min-height: 100vh;
}

/* ── Padding konten ── */
.block-container {
    padding-top: 2rem !important;
    padding-bottom: 3rem !important;
    max-width: 780px !important;
}

/* ══════════════════════════════════════════
   HEADER HERO
══════════════════════════════════════════ */
.app-header {
    position: relative;
    padding: 3rem 2rem 2.5rem;
    border-radius: 24px;
    text-align: center;
    margin-bottom: 1.8rem;
    overflow: hidden;
    background: linear-gradient(135deg, #1e1b4b 0%, #1e293b 50%, #0f172a 100%);
    border: 1px solid rgba(99,102,241,0.3);
    box-shadow: 0 0 60px rgba(99,102,241,0.15), 0 20px 40px rgba(0,0,0,0.4);
}
.app-header::before {
    content: '';
    position: absolute; inset: 0;
    background:
        radial-gradient(ellipse 60% 60% at 15% 40%, rgba(99,102,241,0.25) 0%, transparent 55%),
        radial-gradient(ellipse 40% 40% at 85% 20%, rgba(16,185,129,0.15) 0%, transparent 50%),
        radial-gradient(ellipse 30% 30% at 50% 90%, rgba(244,63,94,0.1) 0%, transparent 50%);
    pointer-events: none;
}
.app-header::after {
    content: '';
    position: absolute;
    top: -50%; left: -50%;
    width: 200%; height: 200%;
    background: repeating-linear-gradient(
        45deg,
        transparent,
        transparent 60px,
        rgba(255,255,255,0.012) 60px,
        rgba(255,255,255,0.012) 61px
    );
    pointer-events: none;
}
.app-header .badge {
    display: inline-block;
    background: rgba(99,102,241,0.2);
    border: 1px solid rgba(99,102,241,0.5);
    color: #a5b4fc;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 2px;
    text-transform: uppercase;
    padding: 0.3rem 0.9rem;
    border-radius: 99px;
    margin-bottom: 1rem;
    position: relative;
}
.app-header h1 {
    margin: 0 0 0.5rem;
    font-size: 2.6rem;
    font-weight: 800;
    letter-spacing: -1px;
    line-height: 1.1;
    position: relative;
    background: linear-gradient(135deg, #e2e8f0 0%, #a5b4fc 50%, #6ee7b7 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
.app-header p {
    margin: 0;
    color: rgba(255,255,255,0.45);
    font-size: 0.9rem;
    font-weight: 400;
    position: relative;
}
.app-header .stats-row {
    display: flex;
    justify-content: center;
    gap: 2rem;
    margin-top: 1.5rem;
    position: relative;
}
.app-header .stat-item {
    text-align: center;
}
.app-header .stat-num {
    font-size: 1.3rem;
    font-weight: 700;
    color: #a5b4fc;
    line-height: 1;
}
.app-header .stat-lbl {
    font-size: 0.68rem;
    color: rgba(255,255,255,0.35);
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-top: 0.2rem;
}
.app-header .stat-divider {
    width: 1px;
    background: rgba(255,255,255,0.1);
    align-self: stretch;
}

/* ══════════════════════════════════════════
   STATUS KAMUS — ganti st.success/warning
══════════════════════════════════════════ */

div[data-testid="stAlert"] {
    border-radius: 14px !important;
    border: none !important;
    font-size: 0.88rem !important;
    font-weight: 500 !important;
}
div[data-testid="stAlert"][data-baseweb="notification"] {
    background: rgba(16,185,129,0.12) !important;
    border-left: 3px solid #10b981 !important;
    color: #6ee7b7 !important;
}

/* ══════════════════════════════════════════
   SECTION LABEL
══════════════════════════════════════════ */
.section-label {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: rgba(165,180,252,0.6);
    margin: 1.6rem 0 0.8rem;
}

/* ══════════════════════════════════════════
   FILE UPLOADER
══════════════════════════════════════════ */
section[data-testid="stFileUploaderDropzone"] {
    background: rgba(30, 27, 75, 0.4) !important;
    border: 2px dashed rgba(99,102,241,0.35) !important;
    border-radius: 16px !important;
    transition: all 0.25s ease;
    backdrop-filter: blur(10px);
}
section[data-testid="stFileUploaderDropzone"]:hover {
    border-color: rgba(99,102,241,0.7) !important;
    background: rgba(99,102,241,0.08) !important;
    box-shadow: 0 0 20px rgba(99,102,241,0.15);
}
section[data-testid="stFileUploaderDropzone"] p,
section[data-testid="stFileUploaderDropzone"] span {
    color: rgba(255,255,255,0.5) !important;
}

/* ── Tombol Browse files ── */
section[data-testid="stFileUploaderDropzone"] button[data-testid="baseButton-secondary"],
section[data-testid="stFileUploaderDropzone"] button,
div[data-testid="stFileUploader"] button {
    background: rgba(99,102,241,0.12) !important;
    border: 1.5px solid rgba(99,102,241,0.35) !important;
    color: rgba(165,180,252,0.85) !important;
    border-radius: 10px !important;
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    font-family: 'Outfit', sans-serif !important;
    box-shadow: none !important;
    transition: all 0.2s ease !important;
    padding: 0.4rem 1rem !important;
}
section[data-testid="stFileUploaderDropzone"] button:hover,
div[data-testid="stFileUploader"] button:hover {
    background: rgba(99,102,241,0.22) !important;
    border-color: rgba(99,102,241,0.6) !important;
    color: #c7d2fe !important;
}

div[data-testid="stFileUploaderFile"] {
    background: rgba(99,102,241,0.1) !important;
    border: 1px solid rgba(99,102,241,0.3) !important;
    border-radius: 10px !important;
    color: #c7d2fe !important;
}

/* ══════════════════════════════════════════
   INPUT FIELDS
══════════════════════════════════════════ */
.stTextInput label, .stSelectbox label {
    color: rgba(255,255,255,0.55) !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.3px;
}
.stTextInput input {
    background: rgba(15,23,42,0.8) !important;
    border: 1.5px solid rgba(99,102,241,0.25) !important;
    border-radius: 12px !important;
    color: #e2e8f0 !important;
    font-family: 'Outfit', sans-serif !important;
    font-size: 0.9rem !important;
    transition: border-color 0.2s, box-shadow 0.2s;
}
.stTextInput input:focus {
    border-color: rgba(99,102,241,0.7) !important;
    box-shadow: 0 0 0 3px rgba(99,102,241,0.15) !important;
}
.stTextInput input::placeholder { color: rgba(255,255,255,0.2) !important; }

/* Selectbox */
.stSelectbox > div > div {
    background: rgba(15,23,42,0.8) !important;
    border: 1.5px solid rgba(99,102,241,0.25) !important;
    border-radius: 12px !important;
    color: #e2e8f0 !important;
}
.stSelectbox svg { color: rgba(165,180,252,0.6) !important; }

/* ══════════════════════════════════════════
   TOMBOL PROSES
══════════════════════════════════════════ */
.stButton > button {
    height: 3.2rem !important;
    border-radius: 14px !important;
    font-weight: 700 !important;
    font-size: 0.95rem !important;
    letter-spacing: 0.3px !important;
    border: none !important;
    background: linear-gradient(135deg, #6366f1 0%, #4f46e5 50%, #4338ca 100%) !important;
    color: white !important;
    box-shadow: 0 4px 20px rgba(99,102,241,0.4), 0 1px 0 rgba(255,255,255,0.1) inset !important;
    transition: all 0.2s ease !important;
    position: relative;
    overflow: hidden;
}
.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 28px rgba(99,102,241,0.5), 0 1px 0 rgba(255,255,255,0.1) inset !important;
    background: linear-gradient(135deg, #818cf8 0%, #6366f1 50%, #4f46e5 100%) !important;
}
.stButton > button:active {
    transform: translateY(0) !important;
    box-shadow: 0 2px 10px rgba(99,102,241,0.3) !important;
}

/* ══════════════════════════════════════════
   PROGRESS
══════════════════════════════════════════ */
.stProgress > div {
    background: rgba(255,255,255,0.07) !important;
    border-radius: 99px !important;
    height: 8px !important;
}
.stProgress > div > div {
    background: linear-gradient(90deg, #6366f1, #818cf8, #10b981) !important;
    border-radius: 99px !important;
    box-shadow: 0 0 10px rgba(99,102,241,0.5);
    transition: width 0.4s ease !important;
}
div[data-testid="stProgressText"] {
    color: rgba(165,180,252,0.8) !important;
    font-size: 0.82rem !important;
}

/* ══════════════════════════════════════════
   TIMER
══════════════════════════════════════════ */
.timer-text {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.82rem;
    color: rgba(110,231,183,0.7);
    text-align: center;
    letter-spacing: 1px;
    margin: 0.4rem 0;
}

/* ══════════════════════════════════════════
   SPINNER
══════════════════════════════════════════ */
div[data-testid="stSpinner"] > div {
    color: #a5b4fc !important;
}

/* ══════════════════════════════════════════
   RESULT CARD
══════════════════════════════════════════ */
.result-panel {
    background: linear-gradient(135deg, rgba(30,27,75,0.6) 0%, rgba(15,23,42,0.8) 100%);
    border: 1px solid rgba(99,102,241,0.25);
    border-radius: 20px;
    padding: 2rem;
    text-align: center;
    backdrop-filter: blur(12px);
    box-shadow: 0 20px 40px rgba(0,0,0,0.3), 0 0 0 1px rgba(255,255,255,0.04) inset;
    margin: 1rem 0;
}
.result-panel .check-icon {
    font-size: 2.5rem;
    display: block;
    margin-bottom: 0.5rem;
    filter: drop-shadow(0 0 10px rgba(16,185,129,0.6));
}
.result-panel h3 {
    color: #e2e8f0;
    font-size: 1.2rem;
    font-weight: 700;
    margin: 0 0 0.3rem;
}
.result-panel .time-badge {
    display: inline-block;
    background: rgba(16,185,129,0.15);
    border: 1px solid rgba(16,185,129,0.3);
    color: #6ee7b7;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.8rem;
    padding: 0.25rem 0.8rem;
    border-radius: 99px;
    margin-bottom: 1.5rem;
}

/* ══════════════════════════════════════════
   DOWNLOAD BUTTONS
══════════════════════════════════════════ */
.stDownloadButton > button {
    border-radius: 12px !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    height: 3rem !important;
    transition: all 0.2s ease !important;
    border: 1.5px solid rgba(99,102,241,0.3) !important;
    background: rgba(30,27,75,0.6) !important;
    color: #c7d2fe !important;
    backdrop-filter: blur(8px);
    box-shadow: 0 2px 12px rgba(0,0,0,0.2) !important;
}
.stDownloadButton > button:hover {
    background: rgba(99,102,241,0.2) !important;
    border-color: rgba(99,102,241,0.6) !important;
    color: #e0e7ff !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(99,102,241,0.25) !important;
}

/* ══════════════════════════════════════════
   AI MENU EXPANDER — ungu kebiruan
══════════════════════════════════════════ */
div[data-testid="stExpander"] {
    border: 1.5px solid rgba(99,102,241,0.55) !important;
    border-radius: 14px !important;
    box-shadow: 0 4px 20px rgba(99,102,241,0.3), 0 1px 0 rgba(255,255,255,0.06) inset !important;
    overflow: hidden !important;
    background: rgba(20,18,50,0.6) !important;
}
div[data-testid="stExpander"]:hover {
    border-color: rgba(129,140,248,0.8) !important;
    box-shadow: 0 8px 28px rgba(99,102,241,0.45) !important;
}

/* Header / tombol expander */
div[data-testid="stExpander"] > details > summary,
div[data-testid="stExpander"] details summary {
    background: linear-gradient(135deg, #6366f1 0%, #4f46e5 55%, #4338ca 100%) !important;
    border-radius: 12px !important;
    padding: 0.85rem 1.2rem !important;
    cursor: pointer !important;
    box-shadow: 0 4px 15px rgba(99,102,241,0.4) !important;
    list-style: none !important;
}
div[data-testid="stExpander"] details summary:hover {
    background: linear-gradient(135deg, #818cf8 0%, #6366f1 55%, #4f46e5 100%) !important;
    box-shadow: 0 6px 22px rgba(99,102,241,0.55) !important;
}

/* Semua teks di dalam summary */
div[data-testid="stExpander"] details summary *,
div[data-testid="stExpander"] details summary p,
div[data-testid="stExpander"] details summary span,
div[data-testid="stExpander"] details summary div,
div[data-testid="stExpander"] details summary label {
    color: #ffffff !important;
    font-weight: 700 !important;
    font-size: 0.95rem !important;
}

/* Ikon panah */
div[data-testid="stExpander"] details summary svg {
    color: #ffffff !important;
    stroke: #ffffff !important;
    fill: #ffffff !important;
}

/* Konten dalam */
div[data-testid="stExpander"] details > div,
div[data-testid="stExpander"] .streamlit-expanderContent {
    background: transparent !important;
    padding-top: 0.6rem !important;
}

/* ══════════════════════════════════════════
   DIVIDER
══════════════════════════════════════════ */
hr {
    border: none !important;
    border-top: 1px solid rgba(255,255,255,0.07) !important;
    margin: 1.5rem 0 !important;
}

/* ══════════════════════════════════════════
   FOOTER
══════════════════════════════════════════ */
.footer {
    text-align: center;
    color: rgba(255,255,255,0.9);
    font-size: 0.75rem;
    padding: 0.7rem 0 0.65rem;
    letter-spacing: 0.5px;
    position: static;
    background: #0b0f1d;
    border-top: 1px solid rgba(99,102,241,0.12);
}
.footer span { color: rgba(255,255,255,0.9); }

/* ══════════════════════════════════════════
   STATUS PILL — di dalam header
══════════════════════════════════════════ */
.status-pill {
    display: inline-flex;
    align-items: center;
    gap: 0.45rem;
    font-size: 0.8rem;
    font-weight: 600;
    padding: 0.38rem 1.1rem;
    border-radius: 99px;
    margin-top: 1.2rem;
    position: relative;
    letter-spacing: 0.2px;
}
.status-ready {
    background: rgba(16,185,129,0.15);
    border: 1px solid rgba(16,185,129,0.45);
    color: #6ee7b7;
}
.status-warn {
    background: rgba(245,158,11,0.13);
    border: 1px solid rgba(245,158,11,0.4);
    color: #fcd34d;
}
.status-dot {
    width: 7px; height: 7px;
    border-radius: 50%;
    display: inline-block;
    flex-shrink: 0;
}
.status-ready .status-dot {
    background: #10b981;
    box-shadow: 0 0 6px rgba(16,185,129,0.8);
    animation: pulse-green 2s infinite;
}
.status-warn .status-dot {
    background: #f59e0b;
    box-shadow: 0 0 6px rgba(245,158,11,0.8);
}
@keyframes pulse-green {
    0%, 100% { opacity: 1; transform: scale(1); }
    50% { opacity: 0.5; transform: scale(1.35); }
}

/* ══════════════════════════════════════════
   GENERAL TEXT FIX
══════════════════════════════════════════ */
p, li, span, div { color: inherit; }
.stApp p, .stApp div, .stApp label { color: rgba(255,255,255,0.75); }

/* ══════════════════════════════════════════
   TOMBOL LOAD KAMUS — IDENTIK DENGAN PROSES
══════════════════════════════════════════ */
.sync-icon-wrap .stButton > button {
    height: 3.2rem !important;
    border-radius: 14px !important;
    font-weight: 700 !important;
    font-size: 0.95rem !important;
    letter-spacing: 0.3px !important;
    border: none !important;
    background: linear-gradient(135deg, #6366f1 0%, #4f46e5 50%, #4338ca 100%) !important;
    color: white !important;
    box-shadow: 0 4px 20px rgba(99,102,241,0.4), 0 1px 0 rgba(255,255,255,0.1) inset !important;
    transition: all 0.2s ease !important;
    width: 100% !important;
}
</style>
""", unsafe_allow_html=True)


# --- KONSTANTA STANDAR ISO ---
ISO_FONT_NAME = "Arial"
ISO_FONT_SIZE = 11
LANG_OPTIONS = {
    "auto": "🔍 Deteksi Otomatis", "en": "🇬🇧 Inggris",
    "fr": "🇫🇷 Prancis", "de": "🇩🇪 Jerman", "es": "🇪🇸 Spanyol",
    "it": "🇮🇹 Italia", "nl": "🇳🇱 Belanda", "pt": "🇵🇹 Portugis",
    "ru": "🇷🇺 Rusia", "ja": "🇯🇵 Jepang", "zh-CN": "🇨🇳 Mandarin",
    "ko": "🇰🇷 Korea", "ar": "🇸🇦 Arab",
}

_ENGINE_COUNT = 9
_SNI_SHEET_ID = "1BBPCMPwvbBk5LPdoDQwnjQzcPHv7_RDKENqeMsklF-8"
_FOREIGN_SHEET_ID = "1NZm1HjsjxmflxnZlzV_O2XF75ZlMUOu8VVofsKfp_FA"


@st.cache_data(ttl=30, show_spinner=False)
def _count_google_sheet_data_rows(sheet_id: str, gid: str = "0") -> int:
    """Hitung seluruh baris terisi pada CSV Google Sheets, minus header.

    Perhitungan ini sengaja terpisah dari parser CustomDictionary dan
    ItalicDictionary. Dengan demikian statistik dashboard benar-benar memakai
    nomor baris data: 49 baris menjadi 48 dan 5 baris menjadi 4.
    """
    url = (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/"
        f"export?format=csv&gid={gid}"
    )
    request = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(request, timeout=15) as response:
        text = response.read().decode("utf-8-sig", errors="replace")
    rows = [
        row for row in csv.reader(io.StringIO(text))
        if any((cell or "").strip() for cell in row)
    ]
    return max(len(rows) - 1, 0)

# --- HELPER FUNGSI ---
def _parse_doc_structure(docx_path: str) -> list:
    """Parsing dokumen menjadi daftar section: [{heading, level, paragraphs}]"""
    try:
        from docx import Document as _Doc
        doc = _Doc(docx_path)
        sections, current = [], {"heading": "Pembukaan", "level": 0, "paragraphs": []}
        for para in doc.paragraphs:
            txt = para.text.strip()
            if not txt:
                continue
            style = para.style.name.lower()
            if style.startswith("heading"):
                if current["paragraphs"]:
                    sections.append(current)
                try:
                    lvl = int(style.replace("heading", "").strip())
                except Exception:
                    lvl = 1
                current = {"heading": txt, "level": lvl, "paragraphs": []}
            else:
                current["paragraphs"].append(txt)
        if current["paragraphs"] or current["heading"] != "Pembukaan":
            sections.append(current)
        return sections
    except Exception:
        return []

def get_elapsed_str(start_time: float) -> str:
    elapsed = time.time() - start_time
    if elapsed < 60:
        return f"{elapsed:.1f} detik"
    else:
        mins = int(elapsed // 60)
        secs = elapsed % 60
        return f"{mins} menit {secs:.1f} detik"

def extract_titles_from_docx(docx_path: str):
    try:
        from docx import Document
        doc = Document(docx_path)
        candidates = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text or len(text) < 5: continue
            is_heading = para.style.name.lower().startswith('heading')
            max_size = 0
            is_italic = False
            for run in para.runs:
                if run.text.strip():
                    sz = run.font.size
                    if sz: max_size = max(max_size, sz.pt if hasattr(sz, 'pt') else sz / 12700)
                    if run.font.italic: is_italic = True
            if is_heading or max_size >= 12:
                candidates.append({'text': text, 'size': max_size})
            if len(candidates) >= 10: break
        if not candidates: return "", ""
        title_id = candidates[0]['text']
        title_en = title_id
        return title_id, title_en
    except Exception: return "", ""

# --- INISIALISASI ENGINE ---
@st.cache_resource
def load_engines():
    return (
        IntroductionContentTrimmerEngine(),
        CoverPageEngine(),
        DaftarIsiEngine(),
        PrakataPendahuluanEngine(),
        IntroductionContentDuplicatorEngine(),
        InfoPendukungEngine(),
        DaftarIsiBibliographyFormatterEngine(),
        SelectiveTranslationEngine(),
        TableOfContentsEngine(),
    )

engine1, engine2, engine3, engine4, engine5, engine6, engine7, engine8, engine9 = load_engines()

# ─────────────────────────────────────────────────────────────────────────────
# HEADER PERSISTEN
# Header dirender oleh app utama, bukan fragment terpisah, sehingga kotak info
# tetap ada ketika proses berjalan maupun ketika pengguna memproses file lagi.
# ─────────────────────────────────────────────────────────────────────────────

def _refresh_kamus_data(force: bool = False) -> None:
    """Ambil glosarium hanya saat awal sesi/refresh manual atau tombol Proses.

    Tidak ada timer/fragment berkala. Nilai terakhir yang berhasil dimuat tetap
    dipakai jika Google Sheet sementara tidak dapat dijangkau.
    """
    if not force and st.session_state.get('_kamus_initialized'):
        return
    if force:
        # Tombol Proses harus membaca nilai terkini, bukan cache 30 detik.
        _count_google_sheet_data_rows.clear()

    try:
        dictionary = CustomDictionary()
        dictionary.load_defaults()
        if len(dictionary) > 0:
            st.session_state['custom_dict'] = dictionary
    except Exception:
        pass
    try:
        italic_dictionary = ItalicDictionary()
        italic_dictionary.load_defaults()
        if len(italic_dictionary) > 0:
            st.session_state['italic_dict'] = italic_dictionary
    except Exception:
        pass
    try:
        st.session_state['kamus_count'] = _count_google_sheet_data_rows(
            _SNI_SHEET_ID
        )
    except Exception:
        pass
    try:
        st.session_state['italic_count'] = _count_google_sheet_data_rows(
            _FOREIGN_SHEET_ID
        )
    except Exception:
        pass
    st.session_state['_kamus_initialized'] = True
    st.session_state['_kamus_loaded_at'] = time.time()


def _render_header_with_live_kamus():
    """
    Render kotak informasi pada setiap rerun utama dan pertahankan nilai sukses
    terakhir jika Google Sheet sementara tidak dapat dijangkau.
    """
    _refresh_kamus_data(force=False)
    _d = st.session_state.get('custom_dict')

    _status_html = (
        f"""<div class="status-pill status-ready">
            <span class="status-dot"></span>
            Sistem Siap &nbsp;
        </div>"""
        if _d is not None and len(_d) > 0 else
        """<div class="status-pill status-warn">
            <span class="status-dot"></span>
            Kamus Tidak Aktif
        </div>"""
    )

    st.markdown(f"""
        <div class="app-header">
            <div class="badge">Generator RSNI · Dashboard v9</div>
            <h1>📑 ISO to RSNI Converter</h1>
            <p>Memformat & Menerjemahan Dokumen Standar ISO Menjadi Draft RSNI Secara Otomatis</p>
            <div class="stats-row">
                <div class="stat-item">
                    <div class="stat-num">{_ENGINE_COUNT}</div>
                    <div class="stat-lbl">Engine</div>
                </div>
                <div class="stat-divider"></div>
                <div class="stat-item">
                    <div class="stat-num">{st.session_state.get('kamus_count', '—')}</div>
                    <div class="stat-lbl">Kamus SNI</div>
                </div>
                <div class="stat-divider"></div>
                <div class="stat-item">
                    <div class="stat-num">{st.session_state.get('italic_count', '—')}</div>
                    <div class="stat-lbl">Kamus Istilah Asing</div>
                </div>
                <div class="stat-divider"></div>
                <div class="stat-item">
                    <div class="stat-num">13</div>
                    <div class="stat-lbl">Bahasa</div>
                </div>
            </div>
            {_status_html}
        </div>
    """, unsafe_allow_html=True)

# --- HALAMAN UTAMA ---

# Render header sebagai bagian permanen app utama.
_render_header_with_live_kamus()


# Ambil nilai kamus dari session_state untuk dipakai di bawah
_kamus = st.session_state.get('custom_dict')
_count = st.session_state.get('kamus_count', 0)
_italic_count = st.session_state.get('italic_count', 0)

# Footer fixed dirender sebelum blok proses yang dapat berjalan lama. Karena
# itu footer tetap berada di layar sejak awal rerun sampai proses selesai.
_FOOTER_HTML = """
<div class='footer'>
  <div style='margin-bottom:0.55rem;'>
    <a href='https://generator-sni.streamlit.app/' target='_blank'
       style='display:inline-block;padding:0.30rem 0.72rem;border:1px solid #6366f1;
              border-radius:7px;background:linear-gradient(135deg,#312e81,#3730a3);
              color:#ffffff;text-decoration:none;font-size:0.72rem;font-weight:700;
              line-height:1.1;box-shadow:0 0 10px rgba(99,102,241,0.22);'>
      Ganti Mode Terjemahan Cepat
    </a>
  </div>
  <span style='font-size:0.85rem;'>
    <a href='https://docs.google.com/spreadsheets/d/1BBPCMPwvbBk5LPdoDQwnjQzcPHv7_RDKENqeMsklF-8/edit?usp=sharing' target='_blank' style='color:#ffffff;text-decoration:none;'>📖 Kamus SNI</a>
    &nbsp;&nbsp;·&nbsp;&nbsp;
    <a href='https://docs.google.com/spreadsheets/d/1NZm1HjsjxmflxnZlzV_O2XF75ZlMUOu8VVofsKfp_FA/edit?usp=sharing' target='_blank' style='color:#ffffff;text-decoration:none;'>🌐 Kamus Istilah Asing</a><br>
    <a href='https://iec-to-iso.streamlit.app/' target='_blank' style='color:#ffffff;text-decoration:none;'>📄 IEC to ISO Converter</a>
  </span><br>
  <span style='font-size:0.8rem;color:#ffffff;'>© 2026 Generator RSNI · ISO to RSNI Converter · All rights reserved.</span><br>
  <span style='font-size:0.72rem;opacity:0.8;color:#ffffff;'>Developed by Denny Kusuma H.</span>
</div>
"""
_footer_rendered_this_run = False

def _render_footer_once():
    """Render footer satu kali dalam alur normal halaman saat ini."""
    global _footer_rendered_this_run
    if _footer_rendered_this_run:
        return
    st.markdown(_FOOTER_HTML, unsafe_allow_html=True)
    _footer_rendered_this_run = True

import datetime
_tahun = str(datetime.date.today().year)

# --- FORM INPUT ---
st.markdown('<div class="section-label">📂 Upload Dokumen ISO</div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "Upload file .doc/.docx di sini atau klik Browse",
    type=["doc", "docx"], key="upl_main", label_visibility="collapsed"
)

st.markdown('<div class="section-label">⚙️ Pengaturan</div>', unsafe_allow_html=True)
col_set1, col_set2 = st.columns([2, 3])
with col_set1:
    doc_title = st.text_input("📄 No. SNI", value="SNI ISO XXXXX-X:XXXX", key="title_main")
with col_set2:
    ics_number = st.text_input(
        "🔢 ICS",
        value="XX.XXX.XX",
        key="ics_main",
        help="Format Nomor ICS, contoh: 45.060.01"
    )

# --- TOMBOL PROSES ---
btn_process = st.button("🚀 Proses", key="btn_main", use_container_width=True)

# Saat error, ketiga tombol kecil tampil setelah tombol Proses dan tepat di
# atas tombol "Ganti Mode Terjemahan Cepat" yang berada di dalam footer.
if st.session_state.get('_process_error'):
    st.error(st.session_state['_process_error'])
    _last_path = st.session_state.get('_last_output')
    _last_engine = int(st.session_state.get('_last_engine', 0) or 0)
    _back_col, _download_col, _continue_col = st.columns(3)
    with _back_col:
        if st.button('↩ Kembali', key='error_back', use_container_width=True):
            _reset_to_start()
            st.rerun()
    with _download_col:
        if _last_path and os.path.isfile(_last_path):
            with open(_last_path, 'rb') as _last_file:
                st.download_button(
                    '⬇ Download', data=_last_file,
                    file_name=f'Hasil_Engine{_last_engine}.docx',
                    key='error_download', use_container_width=True,
                )
        else:
            st.button('⬇ Download', key='error_download_disabled',
                      disabled=True, use_container_width=True)
    with _continue_col:
        if st.button('▶ Lanjutkan', key='error_continue',
                     disabled=not (_last_path and os.path.isfile(_last_path)),
                     use_container_width=True):
            st.session_state['_target_file'] = _last_path
            st.session_state['_resume_engine'] = min(_last_engine + 1, 9)
            st.session_state['_force_continue'] = True
            st.session_state['_forced_errors'] = []
            st.session_state['_run_process'] = True
            st.session_state.pop('_process_error', None)
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# LOGIC EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

if btn_process:
    if uploaded_file:
        try:
            st.session_state.pop('_process_error', None)
            st.session_state.pop('_last_engine', None)
            st.session_state.pop('_last_output', None)
            st.session_state.pop('_force_continue', None)
            st.session_state.pop('_forced_errors', None)
            st.session_state.pop('_forced_summary', None)
            # Satu-satunya refresh otomatis setelah halaman terbuka: saat
            # pengguna secara eksplisit menekan Proses.
            _refresh_kamus_data(force=True)
            # ID unik per sesi browser — mencegah tabrakan nama file saat
            # beberapa pengguna mengakses aplikasi secara bersamaan.
            _sid = st.session_state.setdefault('_sid', uuid.uuid4().hex[:8])
            target_file = f"temp_main_{_sid}_{uploaded_file.name}"
            with open(target_file, "wb") as f:
                f.write(uploaded_file.getbuffer())

            if uploaded_file.name.lower().endswith('.doc'):
                converted_file = (
                    f"converted_{_sid}_{os.path.splitext(uploaded_file.name)[0]}.docx"
                )
                target_file = _convert_legacy_doc(target_file, converted_file)

            st.session_state['_run_process'] = True
            st.session_state['_target_file'] = target_file
            st.session_state['_doc_title'] = doc_title
            st.session_state['_ics_number'] = ics_number
            st.rerun()
        except Exception as upload_error:
            # Bahkan kegagalan sebelum Engine 1 tetap masuk ke UI recovery.
            st.session_state['_process_error'] = (
                f'❌ Error menyiapkan proses: {upload_error}'
            )
            st.session_state['_run_process'] = False
            st.rerun()
    else:
        st.warning("Silakan upload file terlebih dahulu.")


# Jalur "Lanjutkan": mulai persis dari engine setelah output terakhir dan
# teruskan sampai Engine 9. Dalam mode paksa, engine yang error dicatat lalu
# otomatis di-skip; input berikutnya tetap output terakhir yang valid.
if (st.session_state.get('_run_process') and
        st.session_state.get('_target_file') and
        int(st.session_state.get('_resume_engine', 1) or 1) > 1):
    _resume = int(st.session_state['_resume_engine'])
    _resume_input = st.session_state['_target_file']
    _sid = st.session_state.setdefault('_sid', uuid.uuid4().hex[:8])
    _base_name = os.path.basename(_resume_input)
    # Buang seluruh prefix engine terdahulu agar nama tetap pendek.
    _base_name = re.sub(r'^(?:engine\d+_[0-9a-f]{8}_)+', '', _base_name)
    _resume_output = f'engine{_resume}_{_sid}_{_base_name}'
    _sni = st.session_state.get('_doc_title', 'SNI ISO XXXXX-X:XXXX')
    _ics = st.session_state.get('_ics_number', 'XX.XXX.XX')
    try:
        if _resume == 2:
            _ok, _msg = engine2.process(
                input_docx=_resume_input, output_docx=_resume_output,
                sni_number=_sni, ics_number=_ics,
            )
        elif _resume == 3:
            _ok, _unused, _msg = engine3.process(_resume_input, _resume_output, _sni)
        elif _resume == 4:
            _ok, _unused, _msg = engine4.process(
                input_docx=_resume_input, output_docx=_resume_output,
                sni_number=_sni, bsn_year=_tahun,
            )
        elif _resume == 5:
            _ok, _unused, _msg = engine5.process(_resume_input, _resume_output)
        elif _resume == 6:
            _ok, _unused, _msg = engine6.process(_resume_input, _resume_output)
        elif _resume == 7:
            _ok, _unused, _msg = engine7.process(
                input_docx=_resume_input, output_docx=_resume_output,
                sni_number=_sni,
            )
        elif _resume == 8:
            _ok, _unused, _msg = engine8.process(_resume_input, _resume_output)
        elif _resume == 9:
            _ok, _unused, _msg = engine9.process(_resume_input, _resume_output)
        else:
            raise ValueError(f'Nomor engine lanjutan tidak valid: {_resume}')
        if not _ok:
            # Engine 8 sengaja menyimpan dokumen parsial yang valid sebelum
            # mengembalikan status gagal. Pakai file itu untuk Engine 9.
            if _resume == 8 and os.path.isfile(_resume_output):
                _remember_engine_output(8, _resume_output)
                st.session_state['_target_file'] = _resume_output
            raise RuntimeError(_msg)
        _remember_engine_output(_resume, _resume_output)
        st.session_state['_target_file'] = _resume_output
        if _resume < 9:
            st.session_state['_resume_engine'] = _resume + 1
            st.rerun()
        else:
            st.session_state['_final_opt_file'] = _resume_output
            st.session_state['_final_engine'] = 9
            st.session_state['_final_time'] = 'proses lanjutan'
            st.session_state['_show_results'] = True
            st.session_state['_run_process'] = False
            _skipped = st.session_state.get('_forced_errors', [])
            if _skipped:
                _numbers = ', '.join(str(item['engine']) for item in _skipped)
                st.session_state['_forced_summary'] = (
                    f'Mode paksa selesai. Engine yang dilewati karena error: '
                    f'{_numbers}.'
                )
            st.session_state.pop('_resume_engine', None)
            st.session_state.pop('_force_continue', None)
            st.rerun()
    except Exception as _resume_error:
        if st.session_state.get('_force_continue'):
            _errors = st.session_state.setdefault('_forced_errors', [])
            _errors.append({
                'engine': _resume,
                'message': str(_resume_error)[:1000],
            })
            st.session_state.pop('_process_error', None)
            if _resume < 9:
                # Jangan mengganti _target_file: engine sesudahnya memakai
                # output valid terakhir, bukan output engine yang gagal.
                st.session_state['_resume_engine'] = _resume + 1
                st.session_state['_run_process'] = True
                st.rerun()
            else:
                # Engine 9 juga gagal: mode paksa tetap selesai dan menyediakan
                # dokumen terakhir yang valid untuk diunduh.
                _last_valid = st.session_state.get('_last_output')
                _last_no = int(st.session_state.get('_last_engine', 0) or 0)
                if _last_valid and os.path.isfile(_last_valid):
                    st.session_state['_final_opt_file'] = _last_valid
                    st.session_state['_final_engine'] = _last_no
                    st.session_state['_final_time'] = 'mode paksa selesai'
                    st.session_state['_show_results'] = True
                    _numbers = ', '.join(
                        str(item['engine']) for item in _errors
                    )
                    st.session_state['_forced_summary'] = (
                        'Semua step lanjutan sudah dicoba. Engine yang '
                        f'dilewati karena error: {_numbers}. File yang tersedia '
                        f'adalah output terakhir Engine {_last_no}.'
                    )
                    st.session_state['_run_process'] = False
                    st.session_state.pop('_resume_engine', None)
                    st.session_state.pop('_force_continue', None)
                    st.rerun()
        # Fallback bila mode paksa tidak aktif atau tidak ada output valid.
        st.session_state['_process_error'] = (
            f'❌ Error Proses Engine {_resume}: {_resume_error}'
        )
        st.session_state['_run_process'] = False
        st.session_state.pop('_resume_engine', None)
        st.session_state.pop('_force_continue', None)
        st.rerun()
    st.stop()

if st.session_state.get('_run_process') and st.session_state.get('_target_file'):
    target_file = st.session_state['_target_file']
    doc_title_val = st.session_state['_doc_title']
    ics_number_val = st.session_state.get('_ics_number', 'XX.XXX.XX').strip() or 'XX.XXX.XX'
    
    # UI Progress — 3 elemen terpisah agar tidak saling tumpuk
    status_placeholder = st.empty()
    progress_bar = st.progress(0)
    start_time = time.time()

    # ── Live Timer: iframe via components.html agar <script> benar-benar jalan
    _TIMER_HTML = """
    <style>
      #sni-timer-wrap {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.82rem;
        color: rgba(110,231,183,0.85);
        text-align: center;
        letter-spacing: 1px;
        margin: 0;
        padding: 0;
        background: transparent;
      }
    </style>
    <div id="sni-timer-wrap">
      &#x23F1; <span id="sni-timer">0 detik</span>
    </div>
    <script>
      var start = Date.now();
      setInterval(function(){
        var sec = Math.floor((Date.now() - start) / 1000);
        var el = document.getElementById('sni-timer');
        if (!el) return;
        if (sec < 60) {
          el.textContent = sec + ' detik';
        } else {
          var m = Math.floor(sec / 60);
          var s = sec % 60;
          el.textContent = m + ' menit ' + s + ' detik';
        }
      }, 1000);
    </script>
    """
    # components.html() render ke iframe — script PASTI jalan, tidak disanitasi
    _components.html(_TIMER_HTML, height=36)
    # time_placeholder dipakai hanya untuk waktu final statis setelah selesai
    time_placeholder = st.empty()
    # Footer berada setelah progress/timer dalam alur normal, lalu proses berat
    # dimulai. Karena sudah dirender, footer tetap terlihat dan ikut ter-scroll.
    _render_footer_once()
    # ────────────────────────────────────────────────────────────────────────

    # Helper Update UI — TIDAK menyentuh timer iframe, hanya status & progress
    def update_ui(pct, msg, skip_progress=False):
        parts = msg.split("\n", 1)
        line1 = parts[0].strip()
        line2 = parts[1].strip() if len(parts) > 1 else ""
        if pct >= 100:
            dot_color, dot_glow, anim = "#10b981", "rgba(16,185,129,0.9)", ""
        elif pct >= 75:
            dot_color, dot_glow, anim = "#6366f1", "rgba(99,102,241,0.9)", "animation:_pd 0.9s infinite;"
        elif pct >= 50:
            dot_color, dot_glow, anim = "#818cf8", "rgba(129,140,248,0.8)", "animation:_pd 1s infinite;"
        else:
            dot_color, dot_glow, anim = "#a5b4fc", "rgba(165,180,252,0.7)", "animation:_pd 1.1s infinite;"
        detail_html = (
            f'<div style="font-size:0.78rem;color:rgba(199,210,254,0.72);margin-top:0.22rem;'
            f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-style:italic;">{line2}</div>'
        ) if line2 else ""
        status_placeholder.markdown(
            f'<div style="background:rgba(15,23,42,0.6);border:1px solid rgba(99,102,241,0.22);'
            f'border-radius:10px;padding:0.5rem 0.9rem;font-family:\'Outfit\',sans-serif;margin-bottom:0.3rem;">'
            f'<div style="display:flex;align-items:center;gap:0.5rem;">'
            f'<span style="display:inline-block;width:8px;height:8px;border-radius:50%;flex-shrink:0;'
            f'background:{dot_color};box-shadow:0 0 7px {dot_glow};{anim}"></span>'
            f'<span style="font-size:0.85rem;font-weight:600;color:rgba(165,180,252,0.92);'
            f'flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{line1}</span>'
            f'<span style="font-family:\'JetBrains Mono\',monospace;font-size:0.75rem;'
            f'color:rgba(110,231,183,0.65);flex-shrink:0;">{pct}%</span>'
            f'</div>{detail_html}</div>'
            f'<style>@keyframes _pd{{0%,100%{{opacity:1;transform:scale(1);}}50%{{opacity:.3;transform:scale(1.6);}}}}</style>',
            unsafe_allow_html=True
        )
        if not skip_progress:
            progress_bar.progress(pct)

    # Pipeline Optimasi
    # Pipeline lama sengaja tetap dipertahankan, tetapi tidak dipanggil selama
    # mode ENGINE1_ONLY aktif. Dengan demikian Engine 2/4/5/6/7 tidak dihapus.
    def run_legacy_optimization(input_file, doc_title):
        copyright_text = f"© BSN {_tahun}"
        cover_settings = {
            "sni_number": doc_title if doc_title else "SNI ISO XXXXX:20XX",
            "bsn_year": _tahun, "ics_number": "XX.XXX.XX", "ref_standard": "",
        }

        # Hitung total paragraf dokumen untuk info realtime
        try:
            from docx import Document as _DocCount
            _dc = _DocCount(input_file)
            _total_para = len(_dc.paragraphs)
            _total_tbl  = len(_dc.tables)
            del _dc
        except Exception:
            _total_para, _total_tbl = 0, 0
        _doc_info = f"{_total_para} paragraf, {_total_tbl} tabel"

        # Engine 1–5 berbagi 0–10% total progress
        # E2=1-2, E4=3-4, E5=5-6, E6=7-8, E7=9-10

        # 1. Engine 2 — Format Dasar
        update_ui(1, f"[1/5] Memuat & format dokumen... ({_doc_info})")
        output_file = f"opt_{os.path.basename(input_file)}"
        success, msg = engine2.process(input_file, output_file, ISO_FONT_NAME, ISO_FONT_SIZE, enable_headers=True, doc_title=doc_title, copyright_text=copyright_text)
        if not success: raise Exception(f"Engine 2: {msg}")
        update_ui(2, f"[1/5] Format dasar selesai ✓ ({_doc_info})")
        final_file = output_file
        auto_title_id, auto_title_en = extract_titles_from_docx(output_file)
        _title_info = auto_title_id[:30] + "..." if auto_title_id and len(auto_title_id) > 30 else (auto_title_id or "-")

        # 2. Engine 4 — Cover
        update_ui(3, f"[2/5] Membuat cover: {cover_settings['sni_number']}")
        cover_out = f"cover_{os.path.basename(final_file)}"
        ok4 = engine4.prepend_cover(input_docx=final_file, output_docx=cover_out, sni_number=cover_settings["sni_number"], bsn_year=cover_settings["bsn_year"], title_id=auto_title_id, title_en=auto_title_en, ref_standard=cover_settings["ref_standard"], ics_number=cover_settings["ics_number"])[0]
        if ok4:
            final_file = cover_out
            update_ui(4, f"[2/5] Cover selesai ✓ — {_title_info}")

        # 3. Engine 5 — Daftar Isi
        update_ui(5, f"[3/5] Membuat daftar isi dari {_total_para} paragraf...")
        di_out = f"di_{os.path.basename(final_file)}"
        ok5 = engine5.insert(input_docx=final_file, output_docx=di_out, doc_title=cover_settings["sni_number"], copyright_text=f"©BSN {cover_settings['bsn_year']}")[0]
        if ok5:
            final_file = di_out
            update_ui(6, f"[3/5] Daftar isi selesai ✓")

        # 4. Engine 6 — Prakata
        ref_std = re.sub(r'^SNI\s+', '', cover_settings["sni_number"]).strip()
        update_ui(7, f"[4/5] Menyisipkan prakata: {ref_std}")
        pp_out = f"pp_{os.path.basename(final_file)}"
        ok6 = engine6.insert(input_docx=final_file, output_docx=pp_out, sni_number=cover_settings["sni_number"], title_id=auto_title_id or 'Judul ID', title_en=auto_title_en or 'Title EN', ref_standard=ref_std, bsn_year=cover_settings["bsn_year"])[0]
        if ok6:
            final_file = pp_out
            update_ui(8, f"[4/5] Prakata selesai ✓")

        # 5. Engine 7 — Info Pendukung
        update_ui(9, f"[5/5] Menambahkan info pendukung BSN {_tahun}...")
        ip_out = f"ip_{os.path.basename(final_file)}"
        ok7 = engine7.append(input_docx=final_file, output_docx=ip_out)[0]
        if ok7:
            final_file = ip_out
            update_ui(10, f"[5/5] Formatting selesai ✓")

        return final_file

    try:
        # MODE ENGINE 1 → ENGINE 2 → ENGINE 3 → ENGINE 4 → ENGINE 5 →
        # ENGINE 6 → ENGINE 7 → ENGINE 8 → ENGINE 9.
        update_ui(1, "Membaca file dokumen...")
        _sid = st.session_state.setdefault('_sid', uuid.uuid4().hex[:8])
        original_name = os.path.basename(target_file)
        if original_name.startswith(f"temp_main_{_sid}_"):
            original_name = original_name[len(f"temp_main_{_sid}_"):]
        engine1_out = f"engine1_{_sid}_{original_name}"

        update_ui(1, "Menyiapkan struktur dokumen...")
        ok_e1, msg_e1 = engine1.process(target_file, engine1_out)
        if not ok_e1:
            raise Exception(f"Engine 1: {msg_e1}")
        _remember_engine_output(1, engine1_out)

        update_ui(2, f"Struktur awal selesai ✓\n{msg_e1}")

        engine2_out = f"engine2_{_sid}_{original_name}"
        update_ui(2, "Membuat cover dan copyright...")
        ok_e2, msg_e2 = engine2.process(
            input_docx=engine1_out,
            output_docx=engine2_out,
            sni_number=doc_title_val,
            ics_number=ics_number_val,
        )
        if not ok_e2:
            raise Exception(f"Engine 2: {msg_e2}")
        _remember_engine_output(2, engine2_out)

        update_ui(3, f"Cover dan copyright selesai ✓\n{msg_e2}")

        engine3_out = f"engine3_{_sid}_{original_name}"
        update_ui(3, "Menyisipkan daftar isi...")
        ok_e3, _path_e3, msg_e3 = engine3.process(
            input_docx=engine2_out,
            output_docx=engine3_out,
            sni_number=doc_title_val,
        )
        if not ok_e3:
            raise Exception(f"Engine 3: {msg_e3}")
        _remember_engine_output(3, engine3_out)

        update_ui(4, f"Daftar isi selesai ✓\n{msg_e3}")

        engine4_out = f"engine4_{_sid}_{original_name}"
        update_ui(4, "Menyisipkan prakata...")
        ok_e4, _path_e4, msg_e4 = engine4.process(
            input_docx=engine3_out,
            output_docx=engine4_out,
            sni_number=doc_title_val,
            bsn_year=_tahun,
        )
        if not ok_e4:
            raise Exception(f"Engine 4: {msg_e4}")
        _remember_engine_output(4, engine4_out)

        update_ui(5, f"Prakata selesai ✓\n{msg_e4}")

        engine5_out = f"engine5_{_sid}_{original_name}"
        update_ui(5, "Menyiapkan salinan Introduction dan Content...")
        ok_e5, _path_e5, msg_e5 = engine5.process(
            input_docx=engine4_out,
            output_docx=engine5_out,
        )
        if not ok_e5:
            raise Exception(f"Engine 5: {msg_e5}")
        _remember_engine_output(5, engine5_out)

        update_ui(7, f"Salinan dokumen selesai ✓\n{msg_e5}")

        engine6_out = f"engine6_{_sid}_{original_name}"
        update_ui(7, "Menambahkan informasi pendukung...")
        ok_e6, _path_e6, msg_e6 = engine6.process(
            input_docx=engine5_out,
            output_docx=engine6_out,
        )
        if not ok_e6:
            raise Exception(f"Engine 6: {msg_e6}")
        _remember_engine_output(6, engine6_out)

        update_ui(8, f"Informasi pendukung selesai ✓\n{msg_e6}")

        engine7_out = f"engine7_{_sid}_{original_name}"
        update_ui(8, "Menerapkan format dokumen dasar...")
        ok_e7, _path_e7, msg_e7 = engine7.process(
            input_docx=engine6_out,
            output_docx=engine7_out,
            sni_number=doc_title_val,
        )
        if not ok_e7:
            raise Exception(f"Engine 7: {msg_e7}")
        _remember_engine_output(7, engine7_out)

        update_ui(10, f"Format dasar selesai ✓\n{msg_e7}")

        engine8_out = f"engine8_{_sid}_{original_name}"
        update_ui(10, "Memulai penerjemahan dokumen...")
        _engine8_progress_state = {'pct': 10}

        def _engine8_progress(pct, msg):
            # Engine 8 mengirim antrean gabungan terjemahan normal + pemulihan
            # melalui tag [progres-total]. Counter fase seperti 319/352 tetap
            # ditampilkan kepada pengguna, tetapi progress bar memakai:
            # YYY normal + jumlah bagian yang perlu dipulihkan.
            match = re.search(
                r'\[progres-total\]\s*(\d+)\s*/\s*(\d+)', msg or '',
                re.IGNORECASE,
            )
            if match is None:
                match = re.search(
                    r'(?<!\d)(\d+)\s*/\s*(\d+)(?!\d)', msg or ''
                )
            if match and int(match.group(2)) > 0:
                ratio = min(int(match.group(1)) / int(match.group(2)), 1.0)
            else:
                ratio = max(0, min(pct, 100)) / 100.0
            calculated_pct = min(10 + int(ratio * 88), 98)
            # Jangan biarkan fase sinkronisasi/penyimpanan setelah 100% antrean
            # membuat progress bar mundur karena pesannya tidak punya xx/yyy.
            mapped_pct = max(_engine8_progress_state['pct'], calculated_pct)
            _engine8_progress_state['pct'] = mapped_pct
            display_msg = re.sub(
                r'\s*\|\s*\[progres-total\]\s*\d+\s*/\s*\d+\s*$',
                '', msg or '', flags=re.IGNORECASE,
            )
            update_ui(mapped_pct, f"Menerjemahkan dokumen...\n{display_msg}")

        ok_e8, _path_e8, msg_e8 = engine8.process(
            input_docx=engine7_out,
            output_docx=engine8_out,
            progress_callback=_engine8_progress,
        )
        if not ok_e8:
            _remember_engine_output(8, engine8_out)
            raise Exception(f"Engine 8: {msg_e8}")
        _remember_engine_output(8, engine8_out)

        update_ui(98, f"Penerjemahan selesai ✓\n{msg_e8}")

        engine9_out = f"engine9_{_sid}_{original_name}"
        update_ui(98, "Menerapkan style final dan daftar isi...")
        ok_e9, _path_e9, msg_e9 = engine9.process(
            input_docx=engine8_out,
            output_docx=engine9_out,
        )
        if not ok_e9:
            raise Exception(f"Engine 9: {msg_e9}")
        _remember_engine_output(9, engine9_out)

        update_ui(99, f"Menyimpan dokumen final...\n{msg_e9}")
        final_elapsed = get_elapsed_str(start_time)
        update_ui(100, "✅ Engine 1 sampai Engine 9 selesai!")
        time_placeholder.markdown(
            f'<div class="timer-text">⏱ {final_elapsed}</div>',
            unsafe_allow_html=True
        )
        st.session_state['_final_opt_file'] = engine9_out
        st.session_state['_final_engine'] = 9
        st.session_state.pop('_final_tr_file', None)
        st.session_state['_final_time'] = final_elapsed
        st.session_state['_show_results'] = True
        st.session_state['_doc_sections'] = _parse_doc_structure(engine9_out)

    except Exception as e:
        st.session_state['_process_error'] = f"❌ Error Proses: {e}"
        st.session_state['_run_process'] = False
        st.session_state['_show_results'] = False
        st.rerun()
    finally:
        if st.session_state.get('_show_results'):
            st.session_state['_run_process'] = False
            st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# TAMPILKAN HASIL AKHIR
# ─────────────────────────────────────────────────────────────────────────────

if st.session_state.get('_show_results'):
    st.divider()
    final_time = st.session_state.get('_final_time', '-')
    st.markdown(
        f"""<div class="result-panel">
            <span class="check-icon">✅</span>
            <h3>Dokumen Berhasil Diproses</h3>
            <div class="time-badge">⏱ {final_time}</div>
        </div>""",
        unsafe_allow_html=True
    )

    forced_summary = st.session_state.get('_forced_summary')
    if forced_summary:
        st.warning(forced_summary)

    opt_file = st.session_state.get('_final_opt_file')
    final_engine = int(st.session_state.get('_final_engine', 9) or 9)
    if opt_file and os.path.exists(opt_file):
        with open(opt_file, "rb") as f:
            st.download_button(
                label=f"📄 Download Hasil Engine {final_engine}",
                data=f,
                file_name=f"Hasil_Engine{final_engine}.docx",
                use_container_width=True
            )

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    if st.button("🔄 Proses File Baru", key="reset", use_container_width=True):
        # Hapus file sesi ini segera sebelum reset
        _cleanup_session_files(st.session_state)
        for k in ['_show_results', '_final_opt_file', '_final_tr_file', '_final_time',
                  '_final_engine', '_forced_summary', '_forced_errors', '_force_continue',
                  '_run_process', '_doc_text', '_chat_history', '_doc_sections', '_target_file']:
            if k in st.session_state: del st.session_state[k]
        st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# MESIN ANALISIS LOKAL — 100% offline, tidak ada data keluar
# ─────────────────────────────────────────────────────────────────────────────

def _local_answer(query: str, sections: list, history: list) -> str:
    """Mesin jawab lokal berbasis pencarian dan ekstraksi dari struktur dokumen."""
    import difflib

    q = query.lower().strip()
    words = re.findall(r'\w+', q)

    # ── 1. Deteksi intent ─────────────────────────────────────────────────────

    # Ringkasan
    is_summary = any(w in q for w in [
        'ringkas', 'rangkum', 'ringkasan', 'rangkuman', 'isi dokumen',
        'gambaran', 'overview', 'tentang apa', 'dokumen ini', 'keseluruhan'
    ])
    # Daftar heading
    is_list_sections = any(w in q for w in [
        'daftar bab', 'daftar bagian', 'bagian apa', 'bab apa', 'struktur',
        'apa saja bagian', 'apa saja bab', 'daftar isi', 'section'
    ])
    # Definisi / istilah
    is_definition = any(p in q for p in [
        'apa itu', 'apa yang dimaksud', 'definisi', 'pengertian', 'artinya',
        'maksudnya', 'jelaskan', 'explain'
    ])
    # Cari / di mana
    is_search = any(p in q for p in [
        'di mana', 'dimana', 'cari', 'temukan', 'ada di', 'letak',
        'sebutkan', 'mention', 'terdapat', 'berisi'
    ])
    # Referensi / standar
    is_ref = any(w in q for w in [
        'referensi', 'acuan', 'standar', 'normatif', 'bibliography',
        'pustaka', 'rujukan', 'iso', 'sni', 'iec'
    ])
    # Pasal / klausul tertentu
    clause_match = re.search(r'(?:bab|bagian|klausul|pasal|sub|annex|lampiran|section)\s*[\d\.]+', q)

    # ── 2. Bangun jawaban ─────────────────────────────────────────────────────

    def _section_text(s, max_chars=600):
        body = " ".join(s["paragraphs"])
        return body[:max_chars] + ("..." if len(body) > max_chars else "")

    def _score(s, kws):
        """Skor relevansi section berdasarkan kemunculan keyword."""
        txt = (s["heading"] + " " + " ".join(s["paragraphs"])).lower()
        return sum(2 if w in s["heading"].lower() else (1 if w in txt else 0) for w in kws)

    if not sections:
        return "⚠️ Dokumen tidak dapat dianalisis. Pastikan file .docx valid."

    # Ringkasan keseluruhan
    if is_summary:
        headings = [f"**{s['heading']}**" for s in sections if s['level'] <= 2][:12]
        total_para = sum(len(s['paragraphs']) for s in sections)
        intro = _section_text(sections[0], 400) if sections else ""
        return (
            f"📄 **Ringkasan Dokumen**\n\n"
            f"Dokumen ini terdiri dari **{len(sections)} bagian** dengan total **{total_para} paragraf**.\n\n"
            f"**Bagian utama:**\n" + "\n".join(f"- {h}" for h in headings) +
            (f"\n\n**Pembukaan:**\n{intro}" if intro else "")
        )

    # Daftar section/bab
    if is_list_sections:
        lines = []
        for s in sections:
            indent = "  " * max(0, s['level'] - 1)
            lines.append(f"{indent}{'#' * s['level']} {s['heading']}")
        return "📋 **Struktur Dokumen:**\n\n```\n" + "\n".join(lines[:40]) + "\n```"

    # Referensi & standar
    if is_ref:
        ref_secs = [s for s in sections if any(
            w in s['heading'].lower() for w in ['referensi', 'acuan', 'normatif', 'bibliography', 'pustaka', 'standar']
        )]
        if ref_secs:
            results = []
            for s in ref_secs:
                results.append(f"**{s['heading']}**\n{_section_text(s, 800)}")
            return "📚 **Referensi & Acuan Normatif:**\n\n" + "\n\n---\n\n".join(results)
        # Cari SNI/ISO/IEC dalam seluruh dokumen
        found = []
        for s in sections:
            for p in s['paragraphs']:
                if re.search(r'\b(ISO|SNI|IEC|IEEE)\s*[\d\-:]+', p):
                    found.append(f"- _{s['heading']}:_ {p[:200]}")
        if found:
            return "📚 **Referensi standar yang ditemukan:**\n\n" + "\n".join(found[:15])
        return "ℹ️ Tidak ditemukan bagian referensi atau acuan normatif dalam dokumen ini."

    # Klausul / bab spesifik
    if clause_match:
        target = clause_match.group(0).lower()
        kws = re.findall(r'\w+', target)
        scored = sorted(sections, key=lambda s: _score(s, kws), reverse=True)
        best = scored[:3]
        if best and _score(best[0], kws) > 0:
            results = []
            for s in best:
                if _score(s, kws) > 0:
                    results.append(f"**{s['heading']}**\n{_section_text(s, 700)}")
            return f"📖 **Hasil pencarian '{target}':**\n\n" + "\n\n---\n\n".join(results)

    # Definisi / jelaskan istilah
    if is_definition:
        stop = {'apa','itu','yang','dimaksud','definisi','pengertian','artinya',
                'maksudnya','jelaskan','dengan','dari','dan','di','ke','adalah'}
        kws = [w for w in words if w not in stop and len(w) > 2]
        if kws:
            scored = sorted(sections, key=lambda s: _score(s, kws), reverse=True)
            best = [s for s in scored if _score(s, kws) > 0][:3]
            if best:
                results = []
                for s in best:
                    relevant_paras = [p for p in s['paragraphs']
                                      if any(k in p.lower() for k in kws)][:3]
                    body = " ".join(relevant_paras) if relevant_paras else _section_text(s, 500)
                    results.append(f"**{s['heading']}**\n{body[:600]}")
                return f"🔍 **Penjelasan '{' '.join(kws)}':**\n\n" + "\n\n---\n\n".join(results)

    # Pencarian umum / cari kata kunci
    stop_general = {'apa','ada','di','ke','dan','atau','yang','adalah','ini','itu',
                    'dengan','untuk','dari','pada','dalam','tidak','bisa','cara',
                    'bagaimana','berapa','siapa','kapan','dimana','mana','sebutkan',
                    'cari','temukan','terdapat','berisi','tentang','mengenai'}
    kws = [w for w in words if w not in stop_general and len(w) > 2]

    if kws:
        scored = [(s, _score(s, kws)) for s in sections]
        scored = sorted(scored, key=lambda x: x[1], reverse=True)
        best = [(s, sc) for s, sc in scored if sc > 0][:4]

        if best:
            results = []
            for s, sc in best:
                relevant_paras = [p for p in s['paragraphs']
                                  if any(k in p.lower() for k in kws)][:2]
                body = " ".join(relevant_paras) if relevant_paras else _section_text(s, 400)
                results.append(f"**{s['heading']}**\n{body[:500]}")
            return (
                f"🔍 **Hasil pencarian '{query}'** — ditemukan di {len(best)} bagian:\n\n" +
                "\n\n---\n\n".join(results)
            )

    # Fallback — tidak ada yang cocok
    all_headings = [s['heading'] for s in sections[:10]]
    return (
        f"ℹ️ Tidak ditemukan informasi relevan untuk: **\"{query}\"**\n\n"
        f"**Bagian yang tersedia dalam dokumen:**\n" +
        "\n".join(f"- {h}" for h in all_headings) +
        "\n\n_Coba gunakan kata kunci yang lebih spesifik atau tanyakan tentang bagian di atas._"
    )


# ─────────────────────────────────────────────────────────────────────────────
# CLAUDE API CHAT — diskusi isi dokumen
# ─────────────────────────────────────────────────────────────────────────────

def _build_doc_context(sections: list, max_chars: int = 14000) -> str:
    """Bangun teks konteks dari sections, potong jika terlalu panjang."""
    lines, total = [], 0
    for s in sections:
        chunk = f"\n## {s['heading']}\n" + "\n".join(s['paragraphs']) + "\n"
        if total + len(chunk) > max_chars:
            sisa = max_chars - total
            if sisa > 100:
                lines.append(chunk[:sisa] + "\n[...dokumen dipotong...]")
            break
        lines.append(chunk)
        total += len(chunk)
    return "\n".join(lines)

def _claude_chat(system: str, messages: list) -> str:
    """Kirim chat ke Z.ai API (GLM), return teks jawaban."""
    import urllib.request, json

    ZAI_API_KEY = "bd7f64d4e11642599ca8d1772e89521c.imnp62IRucfcV4bA"

    all_messages = [{"role": "system", "content": system}] + messages

    payload = json.dumps({
        "model": "glm-4-flash",
        "messages": all_messages,
        "max_tokens": 1024,
        "temperature": 0.7,
    }).encode()

    req = urllib.request.Request(
        "https://api.z.ai/api/paas/v4/chat/completions",
        data=payload,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {ZAI_API_KEY}",
        },
        method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            result = json.loads(r.read())
            return result['choices'][0]['message']['content']
    except Exception as e:
        return f"❌ Z.ai Error: {e}"

# # ── TAMPILAN CHAT — selalu tampil di dashboard ────────────────────────────────

# _src = st.session_state.get('_final_tr_file') or st.session_state.get('_final_opt_file')
# _cached_sections = st.session_state.get('_doc_sections', [])

# if _src and os.path.exists(_src) and len(_cached_sections) == 0:
#     st.session_state['_doc_sections'] = _parse_doc_structure(_src)
# elif '_doc_sections' not in st.session_state:
#     st.session_state['_doc_sections'] = []

# sections = st.session_state.get('_doc_sections', [])
# n_sec    = len(sections)
# n_para   = sum(len(s['paragraphs']) for s in sections)
# has_doc  = n_sec > 0

# st.divider()

# import streamlit.components.v1 as _components

# _exp_label = (
#     f"🤖 Asisten AI  ·  {n_sec} bagian  ·  {n_para} paragraf"
#     if has_doc else
#     "🤖 Asisten AI  ·  Tanya seputar Draft RSNI"
# )

# st.markdown("""
# <style>
# div[data-testid="stExpander"] {
#     border: 1.5px solid rgba(99,102,241,0.6) !important;
#     border-radius: 14px !important;
#     overflow: hidden !important;
#     box-shadow: 0 4px 24px rgba(99,102,241,0.3) !important;
#     background: rgba(20,18,50,0.55) !important;
# }
# div[data-testid="stExpander"] summary {
#     background: linear-gradient(135deg, #6366f1 0%, #4f46e5 50%, #4338ca 100%) !important;
#     padding: 0.82rem 1.2rem !important;
#     border-radius: 12px !important;
#     box-shadow: 0 3px 14px rgba(99,102,241,0.45) !important;
# }
# div[data-testid="stExpander"] summary:hover {
#     background: linear-gradient(135deg, #818cf8 0%, #6366f1 50%, #4f46e5 100%) !important;
#     box-shadow: 0 6px 22px rgba(99,102,241,0.6) !important;
# }
# div[data-testid="stExpander"] summary span,
# div[data-testid="stExpander"] summary p,
# div[data-testid="stExpander"] summary div {
#     color: #fff !important;
#     font-weight: 700 !important;
# }
# div[data-testid="stExpander"] summary svg {
#     stroke: #fff !important;
#     color: #fff !important;
# }
# </style>
# """, unsafe_allow_html=True)

# with st.expander(_exp_label, expanded=True):

#     doc_badge = (
#         f"<span style='background:rgba(16,185,129,0.12);border:1px solid rgba(16,185,129,0.3);"
#         f"color:#6ee7b7;font-size:0.75rem;padding:0.2rem 0.7rem;border-radius:99px;'>"
#         f"📑 {n_sec} bagian · {n_para} paragraf</span>"
#         if has_doc else
#         f"<span style='background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);"
#         f"color:rgba(255,255,255,0.3);font-size:0.75rem;padding:0.2rem 0.7rem;border-radius:99px;'>"
#         f"📄 Belum ada dokumen — jawab hal umum RSNI</span>"
#     )
#     st.markdown(
#         f"<div style='display:flex;gap:0.7rem;margin-bottom:0.8rem;flex-wrap:wrap;'>"
#         f"{doc_badge}"
#         f"<span style='background:rgba(99,102,241,0.1);border:1px solid rgba(99,102,241,0.25);"
#         f"color:#a5b4fc;font-size:0.75rem;padding:0.2rem 0.7rem;border-radius:99px;'>"
#         f"✨ Z.ai GLM</span></div>",
#         unsafe_allow_html=True
#     )

#     # ── Pemilih Suara TTS ────────────────────────────────────────────────────
#     _components.html("""
# <div id="tts-voice-bar" style="margin-bottom:8px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
#   <span style="color:rgba(165,180,252,0.7);font-size:0.73rem;font-family:sans-serif;">🎙️ Suara:</span>
#   <select id="tts-voice-select"
#     style="background:rgba(15,23,42,0.85);border:1.5px solid rgba(99,102,241,0.35);
#            color:#c7d2fe;font-size:0.73rem;padding:3px 8px;border-radius:8px;
#            font-family:sans-serif;cursor:pointer;outline:none;max-width:260px;">
#     <option value="">⏳ Memuat daftar suara...</option>
#   </select>
#   <select id="tts-rate-select"
#     style="background:rgba(15,23,42,0.85);border:1.5px solid rgba(99,102,241,0.35);
#            color:#c7d2fe;font-size:0.73rem;padding:3px 8px;border-radius:8px;
#            font-family:sans-serif;cursor:pointer;outline:none;">
#     <option value="0.7">🐢 Lambat</option>
#     <option value="0.92" selected>▶️ Normal</option>
#     <option value="1.2">⚡ Cepat</option>
#     <option value="1.5">🚀 Sangat Cepat</option>
#   </select>
# </div>
# <script>
# (function(){
#   var sel = document.getElementById('tts-voice-select');

#   function populateVoices(){
#     var voices = window.speechSynthesis.getVoices();
#     if(!voices.length) return;
#     sel.innerHTML = '';

#     var id_voices = voices.filter(function(v){
#       return v.lang.startsWith('id') || v.lang.startsWith('ms');
#     });
#     var en_voices = voices.filter(function(v){
#       return v.lang.startsWith('en');
#     }).slice(0, 10);

#     if(id_voices.length){
#       var og = document.createElement('optgroup');
#       og.label = '🇮🇩 Indonesia / Melayu';
#       id_voices.forEach(function(v){
#         var o = document.createElement('option');
#         o.value = v.name;
#         o.textContent = v.name + ' (' + v.lang + ')';
#         og.appendChild(o);
#       });
#       sel.appendChild(og);
#     }

#     if(en_voices.length){
#       var og2 = document.createElement('optgroup');
#       og2.label = '🌐 English (fallback)';
#       en_voices.forEach(function(v){
#         var o = document.createElement('option');
#         o.value = v.name;
#         o.textContent = v.name + ' (' + v.lang + ')';
#         og2.appendChild(o);
#       });
#       sel.appendChild(og2);
#     }

#     sel.onchange = function(){
#       try{ localStorage.setItem('tts_voice', sel.value); }catch(e){}
#     };
#     try{
#       var saved = localStorage.getItem('tts_voice');
#       if(saved && sel.querySelector('option[value="'+saved+'"]')) sel.value = saved;
#     }catch(e){}
#   }

#   if(window.speechSynthesis.getVoices().length > 0){
#     populateVoices();
#   } else {
#     window.speechSynthesis.onvoiceschanged = populateVoices;
#     setTimeout(function(){ populateVoices(); }, 500);
#   }
# })();
# </script>
# """, height=52)

#     # TTS helper
#     def _tts_html(text: str, btn_id: str) -> str:
#         import base64
#         b64 = base64.b64encode(text[:3000].encode('utf-8')).decode('ascii')
#         return f"""
# <div style="margin-top:5px;display:flex;gap:6px;flex-wrap:wrap;">
#   <span id="tts_data_{btn_id}" style="display:none">{b64}</span>
#   <button id="btn_speak_{btn_id}"
#     onclick="(function(){{
#       if(!('speechSynthesis' in window)){{alert('Browser tidak mendukung TTS');return;}}
#       window.speechSynthesis.cancel();

#       var raw = document.getElementById('tts_data_{btn_id}').textContent;
#       var txt = decodeURIComponent(escape(atob(raw)));
#       var btn = document.getElementById('btn_speak_{btn_id}');

#       var voiceName = '';
#       var rate = 0.92;
#       try {{
#         var topDoc = window.top.document;
#         var vSel = topDoc.getElementById('tts-voice-select');
#         var rSel = topDoc.getElementById('tts-rate-select');
#         if(vSel) voiceName = vSel.value;
#         if(rSel) rate = parseFloat(rSel.value) || 0.92;
#       }} catch(e) {{
#         try{{ voiceName = localStorage.getItem('tts_voice') || ''; }}catch(e2){{}}
#       }}

#       var voices = window.speechSynthesis.getVoices();
#       var chosenVoice = null;
#       if(voiceName) chosenVoice = voices.find(function(v){{return v.name===voiceName;}});
#       if(!chosenVoice) chosenVoice = voices.find(function(v){{return v.lang.startsWith('id');}});
#       if(!chosenVoice) chosenVoice = voices.find(function(v){{return v.lang.startsWith('ms');}});
#       if(!chosenVoice) chosenVoice = voices.find(function(v){{return v.lang.startsWith('en');}});

#       var sentences = txt.match(/[^.!?\\n]{{1,220}}[.!?\\n]?/g) || [txt];
#       btn.textContent = '⏳ Membaca...';
#       btn.style.borderColor = 'rgba(16,185,129,0.6)';
#       btn.style.color = '#6ee7b7';

#       function speakChunk(i){{
#         if(i >= sentences.length){{
#           btn.textContent='🔊 Bacakan';
#           btn.style.borderColor='rgba(99,102,241,0.4)';
#           btn.style.color='#a5b4fc';
#           return;
#         }}
#         var u = new SpeechSynthesisUtterance(sentences[i]);
#         u.lang = 'id-ID';
#         u.rate = rate;
#         u.pitch = 1.0;
#         if(chosenVoice) u.voice = chosenVoice;
#         u.onend  = function(){{ speakChunk(i+1); }};
#         u.onerror = function(){{
#           btn.textContent='🔊 Bacakan';
#           btn.style.borderColor='rgba(99,102,241,0.4)';
#           btn.style.color='#a5b4fc';
#         }};
#         window.speechSynthesis.speak(u);
#       }}

#       function startSpeak(){{
#         if(window.speechSynthesis.getVoices().length===0){{
#           window.speechSynthesis.onvoiceschanged=function(){{ speakChunk(0); }};
#         }} else {{
#           speakChunk(0);
#         }}
#       }}
#       startSpeak();
#     }})()"
#     style="background:rgba(99,102,241,0.15);border:1px solid rgba(99,102,241,0.4);
#            color:#a5b4fc;font-size:0.72rem;padding:4px 12px;border-radius:99px;
#            cursor:pointer;font-family:sans-serif;transition:all 0.2s;">
#     🔊 Bacakan
#   </button>
#   <button onclick="(function(){{
#       window.speechSynthesis.cancel();
#       var b=document.getElementById('btn_speak_{btn_id}');
#       if(b){{b.textContent='🔊 Bacakan';b.style.borderColor='rgba(99,102,241,0.4)';b.style.color='#a5b4fc';}}
#     }})()"
#     style="background:rgba(239,68,68,0.1);border:1px solid rgba(239,68,68,0.3);
#            color:#fca5a5;font-size:0.72rem;padding:4px 12px;border-radius:99px;
#            cursor:pointer;font-family:sans-serif;transition:all 0.2s;">
#     ⏹ Stop
#   </button>
# </div>
# """

#     # Inisialisasi riwayat
#     if '_chat_history' not in st.session_state:
#         st.session_state['_chat_history'] = []

#     # Tampilkan riwayat
#     for idx, msg in enumerate(st.session_state['_chat_history']):
#         with st.chat_message(msg['role'], avatar='🧑' if msg['role'] == 'user' else '🤖'):
#             st.markdown(msg['content'])
#             if msg['role'] == 'assistant':
#                 _clean = re.sub(r'[*_`#>\-]+', '', msg['content'])
#                 _clean = re.sub(r'\s+', ' ', _clean).strip()
#                 _components.html(_tts_html(_clean, f"hist_{idx}"), height=44)

#     # ── Voice Input ─────────────────────────────────────────────────────────
#     _components.html("""
# <style>
# #mic-fixed-btn {
#   display: none;
#   position: fixed;
#   bottom: 14px;
#   right: calc(50% - 375px);
#   z-index: 99999;
#   background: linear-gradient(135deg, #6366f1, #4f46e5);
#   border: none; color: #fff;
#   font-size: 0.78rem; font-weight: 700;
#   padding: 8px 15px; border-radius: 99px; cursor: pointer;
#   font-family: 'Outfit', sans-serif;
#   box-shadow: 0 4px 16px rgba(99,102,241,0.5);
#   transition: all 0.2s; white-space: nowrap;
#   align-items: center; gap: 5px;
# }
# #mic-fixed-btn.listening {
#   background: linear-gradient(135deg,#ef4444,#dc2626) !important;
#   animation: mic-pulse 1s infinite;
# }
# @keyframes mic-pulse {
#   0%   { box-shadow: 0 0 0 0 rgba(239,68,68,0.6); }
#   70%  { box-shadow: 0 0 0 10px rgba(239,68,68,0); }
#   100% { box-shadow: 0 0 0 0 rgba(239,68,68,0); }
# }
# #mic-toast-local {
#   display: none;
#   position: fixed; bottom: 60px; left: 50%;
#   transform: translateX(-50%);
#   background: rgba(15,23,42,0.96);
#   border: 1px solid rgba(99,102,241,0.4);
#   border-radius: 12px; padding: 7px 16px;
#   font-size: 0.78rem; color: #c7d2fe;
#   font-family: sans-serif; z-index: 99999;
#   box-shadow: 0 4px 20px rgba(0,0,0,0.4);
#   white-space: nowrap; pointer-events: none;
#   max-width: 90vw; overflow: hidden; text-overflow: ellipsis;
# }
# </style>

# <button id="mic-fixed-btn" onclick="toggleMic()">🎤 Bicara</button>
# <div   id="mic-toast-local"></div>

# <script>
# (function(){
#   if (!window._mic) window._mic = { rec: null, listening: false, timer: null };
#   var M = window._mic;

#   var fbBtn   = document.getElementById('mic-fixed-btn');
#   var fbToast = document.getElementById('mic-toast-local');
#   var activeBtn   = fbBtn;
#   var activeToast = fbToast;

#   function showToast(msg, ms) {
#     activeToast.textContent = msg;
#     activeToast.style.display = 'block';
#     clearTimeout(M.timer);
#     if (ms) M.timer = setTimeout(function(){ activeToast.style.display='none'; }, ms);
#   }

#   function resetBtn() {
#     activeBtn.innerHTML = '🎤 Bicara';
#     activeBtn.classList.remove('listening');
#   }
#   function forceStop() {
#     if (M.rec) { try { M.rec.abort(); } catch(e){} M.rec = null; }
#     M.listening = false; resetBtn();
#   }

#   function sendToChat(txt) {
#     var sent = false;
#     var targets = [];
#     try { targets.push(window); } catch(e){}
#     try { if (window.parent !== window) targets.push(window.parent); } catch(e){}
#     try { if (window.top !== window && window.top !== window.parent) targets.push(window.top); } catch(e){}
#     for (var i = 0; i < targets.length && !sent; i++) {
#       try {
#         var w = targets[i];
#         var ta = w.document.querySelector('div[data-testid="stChatInput"] textarea');
#         if (!ta) continue;
#         var setter = Object.getOwnPropertyDescriptor(w.HTMLTextAreaElement.prototype, 'value').set;
#         setter.call(ta, txt);
#         ta.dispatchEvent(new Event('input',  {bubbles:true}));
#         ta.dispatchEvent(new Event('change', {bubbles:true}));
#         setTimeout(function(){
#           ['keydown','keypress','keyup'].forEach(function(ev){
#             ta.dispatchEvent(new w.KeyboardEvent(ev, {key:'Enter',code:'Enter',keyCode:13,which:13,bubbles:true,cancelable:true}));
#           });
#           activeToast.style.display = 'none';
#         }, 150);
#         sent = true;
#       } catch(e) {}
#     }
#     if (!sent) {
#       try { navigator.clipboard.writeText(txt).then(function(){ showToast('📋 Tersalin — paste Ctrl+V ke chat', 4000); }); }
#       catch(e) { showToast('💬 ' + txt.substring(0,60), 5000); }
#     }
#   }

#   window.toggleMic = function() {
#     if (M.listening) { forceStop(); showToast('⏹ Dihentikan', 1500); return; }
#     var SR = window.SpeechRecognition || window.webkitSpeechRecognition;
#     if (!SR) { showToast('❌ Gunakan Chrome/Edge untuk voice input', 3000); return; }
#     forceStop();
#     var rec = new SR();
#     rec.lang = 'id-ID'; rec.continuous = false; rec.interimResults = true; rec.maxAlternatives = 1;
#     M.rec = rec;
#     var lastTxt = '';
#     rec.onstart = function() {
#       M.listening = true;
#       activeBtn.innerHTML = '🔴 Stop'; activeBtn.classList.add('listening');
#       showToast('🎤 Sedang mendengarkan...');
#     };
#     rec.onresult = function(e) {
#       var interim='', final_t='';
#       for (var i=e.resultIndex; i<e.results.length; i++) {
#         if (e.results[i].isFinal) final_t += e.results[i][0].transcript;
#         else interim += e.results[i][0].transcript;
#       }
#       lastTxt = final_t || interim;
#       showToast('💬 ' + lastTxt);
#     };
#     rec.onend = function() {
#       M.rec = null; M.listening = false; resetBtn();
#       var txt = lastTxt.trim(); lastTxt = '';
#       if (!txt) { showToast('⚠️ Tidak terdeteksi — coba lagi', 2500); return; }
#       sendToChat(txt);
#     };
#     rec.onerror = function(e) {
#       M.rec = null; M.listening = false; resetBtn();
#       if (e.error === 'aborted') return;
#       var msgs = {'no-speech':'⚠️ Tidak ada suara','not-allowed':'❌ Izin mikrofon ditolak','audio-capture':'❌ Mikrofon tidak ada','network':'❌ Gangguan jaringan'};
#       showToast(msgs[e.error] || ('❌ Error: ' + e.error), 3000);
#     };
#     rec.start();
#   };

#   function setupParentBtn() {
#     try {
#       var pw = window.parent;
#       if (pw === window) throw new Error('no parent');

#       if (!pw.document.getElementById('_mic_style')) {
#         var s = pw.document.createElement('style');
#         s.id = '_mic_style';
#         s.textContent =
#           'div[data-testid="stChatInput"]{position:relative !important;}' +
#           'div[data-testid="stChatInput"] textarea{padding-right:128px !important;}' +
#           '#_mic_btn{' +
#             'position:absolute;right:52px;top:50%;transform:translateY(-50%);' +
#             'background:linear-gradient(135deg,#6366f1,#4f46e5);' +
#             'border:none;color:#fff;font-size:0.78rem;font-weight:700;' +
#             'padding:7px 14px;border-radius:99px;cursor:pointer;' +
#             'font-family:Outfit,sans-serif;' +
#             'box-shadow:0 3px 12px rgba(99,102,241,0.5);' +
#             'transition:all 0.2s;white-space:nowrap;' +
#             'display:inline-flex;align-items:center;gap:5px;z-index:200;' +
#           '}' +
#           '#_mic_btn:hover{background:linear-gradient(135deg,#818cf8,#6366f1);box-shadow:0 5px 18px rgba(99,102,241,0.65);}' +
#           '#_mic_btn.listening{background:linear-gradient(135deg,#ef4444,#dc2626)!important;animation:_mpulse 1s infinite;}' +
#           '@keyframes _mpulse{0%{box-shadow:0 0 0 0 rgba(239,68,68,0.6)}70%{box-shadow:0 0 0 8px rgba(239,68,68,0)}100%{box-shadow:0 0 0 0 rgba(239,68,68,0)}}' +
#           '#_mic_toast{display:none;position:fixed;bottom:72px;left:50%;transform:translateX(-50%);' +
#             'background:rgba(15,23,42,0.96);border:1px solid rgba(99,102,241,0.4);' +
#             'border-radius:12px;padding:7px 16px;font-size:0.78rem;color:#c7d2fe;' +
#             'font-family:sans-serif;z-index:99999;box-shadow:0 4px 20px rgba(0,0,0,0.4);' +
#             'white-space:nowrap;pointer-events:none;max-width:90vw;overflow:hidden;text-overflow:ellipsis;}';
#         pw.document.head.appendChild(s);
#       }

#       if (!pw.document.getElementById('_mic_toast')) {
#         var t = pw.document.createElement('div');
#         t.id = '_mic_toast';
#         pw.document.body.appendChild(t);
#       }
#       activeToast = pw.document.getElementById('_mic_toast');

#       function doInject() {
#         var chatInput = pw.document.querySelector('div[data-testid="stChatInput"]');
#         if (!chatInput) return false;
#         var existing = pw.document.getElementById('_mic_btn');
#         if (existing && chatInput.contains(existing)) {
#           activeBtn = existing;
#           existing.onclick = function(e){ e.preventDefault(); window.toggleMic(); };
#           fbBtn.style.display = 'none';
#           return true;
#         }
#         if (existing) existing.remove();
#         var btn = pw.document.createElement('button');
#         btn.id   = '_mic_btn';
#         btn.type = 'button';
#         btn.innerHTML = '🎤 Bicara';
#         btn.onclick = function(e){ e.preventDefault(); window.toggleMic(); };
#         chatInput.appendChild(btn);
#         activeBtn   = btn;
#         fbBtn.style.display = 'none';
#         return true;
#       }

#       if (!doInject()) {
#         var obs = new pw.MutationObserver(function(){ doInject(); });
#         obs.observe(pw.document.body, { childList:true, subtree:true });
#       } else {
#         var obs2 = new pw.MutationObserver(function(){ doInject(); });
#         obs2.observe(pw.document.body, { childList:true, subtree:true });
#       }

#     } catch(e) {
#       fbBtn.style.display = 'flex';
#     }
#   }

#   setupParentBtn();
# })();
# </script>
# """, height=56)

#     # Input teks
#     user_input = st.chat_input("Tanya seputar isi dokumen RSNI...")

#     if user_input:
#         st.session_state['_chat_history'].append({'role': 'user', 'content': user_input})
#         with st.chat_message('user', avatar='🧑'):
#             st.markdown(user_input)

#         with st.chat_message('assistant', avatar='🤖'):
#             with st.spinner("asistant sedang menganalisis..."):
#                 if has_doc:
#                     doc_ctx = _build_doc_context(sections, max_chars=14000)
#                     system_prompt = (
#                         "Kamu adalah asisten ahli RSNI yang membantu pengguna memahami dokumen. "
#                         "Jawab HANYA berdasarkan isi dokumen berikut. "
#                         "Gunakan Bahasa Indonesia yang jelas, terstruktur, dan akurat. "
#                         "Jika informasi tidak ada dalam dokumen, katakan dengan jujur.\n\n"
#                         f"=== ISI DOKUMEN ===\n{doc_ctx}\n==================="
#                     )
#                 else:
#                     system_prompt = (
#                         "Kamu adalah asisten ahli RSNI dan dokumen teknis BSN. "
#                         "Jawab dalam Bahasa Indonesia dengan jelas dan akurat. "
#                         "Belum ada dokumen — jawab berdasarkan pengetahuan umum SNI, ISO, IEC, dan standar lainnya."
#                     )
#                 api_messages = [
#                     {"role": m['role'], "content": m['content']}
#                     for m in st.session_state['_chat_history']
#                 ]
#                 reply = _claude_chat(system_prompt, api_messages)

#             st.markdown(reply)
#             _clean = re.sub(r'[*_`#>\-]+', '', reply)
#             _clean = re.sub(r'\s+', ' ', _clean).strip()
#             _components.html(_tts_html(_clean, f"new_{int(time.time())}"), height=44)
#             st.session_state['_chat_history'].append({'role': 'assistant', 'content': reply})

#     if st.session_state.get('_chat_history'):
#         if st.button("🗑️ Hapus Riwayat", key="clear_chat"):
#             st.session_state['_chat_history'] = []
#             st.rerun()

# Kondisi tanpa proses aktif atau halaman hasil: footer berada di akhir alur
# halaman. Saat proses aktif fungsi ini no-op karena sudah dirender di bawah
# progress/timer sebelum pekerjaan berat dimulai.
_render_footer_once()
