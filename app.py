import streamlit as st
import streamlit.components.v1 as _components
import os
import re
import time
import glob
import atexit
import threading
from io import BytesIO

# ─────────────────────────────────────────────────────────────────────────────
# AUTO-CLEANUP — pembersihan file temporer otomatis
# ─────────────────────────────────────────────────────────────────────────────

# Database kamus ada di https://bit.ly/kamusSNI

# Pola file temporer yang dibuat oleh aplikasi
_TEMP_PATTERNS = ["temp_main_*", "opt_*", "cover_*", "di_*", "pp_*", "ip_*", "ID_*"]
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
from engine2 import DocxOptimizerEngine
from engine4 import CoverPageEngine
from engine5 import DaftarIsiEngine
from engine6 import PrakataPendahuluanEngine
from engine7 import InfoPendukungEngine
from engine9 import CustomDictionary, ItalicDictionary, DocxFinalTranslatorEngine

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
    color: rgba(255,255,255,0.2);
    font-size: 0.75rem;
    padding: 2rem 0 1rem;
    letter-spacing: 0.5px;
}
.footer span { color: rgba(99,102,241,0.5); }

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

# TTL kamus — berapa detik sebelum reload otomatis dari Google Sheet
# DIHAPUS: auto-refresh tiap 1 detik menyebabkan halaman berkedip terus.
# Kamus kini di-load SEKALI saat startup via @st.cache_resource.

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

# Pola "Bagian 1:" / "Part 1:" / "Bagian 12 :" dll. — dipakai untuk
# menormalkan spasi setelah titik dua jadi tepat satu spasi, SEKALIGUS
# mengkapitalisasi huruf pertama kata setelahnya, KHUSUS pada pola ini saja
# (bukan pada setiap titik dua di judul). Menangkap angka bagian di grup 1
# (tanpa titik dua) supaya spasi sebelum ":" ikut dirapikan.
_RE_BAGIAN_PART = re.compile(r'((?:Bagian|Part)\s+\d+)\s*:\s*([a-zA-Z])')

# Strip PANJANG (en dash "–" atau em dash "—") yang dipakai sebagai pemisah
# antar-klausa pada judul, mis. "Aplikasi perkeretaapian — Perlindungan
# kebakaran ... — Bagian 1: Umum". Sengaja TIDAK menyertakan hyphen pendek
# "-" karena hyphen pendek biasanya bagian dari kata majemuk (mis.
# "non-destructive") atau nomor standar (mis. "14001-1") yang tidak boleh
# diutak-atik spasinya.
_RE_LONG_DASH = re.compile(r'\s*[\u2013\u2014]\s*')
# Huruf pertama kata setelah strip panjang yang sudah dirapikan -> kapitalkan.
_RE_AFTER_LONG_DASH = re.compile(r'(\u2014[ ])([a-z])')

def _normalize_title_line(text: str) -> str:
    """
    Judul dari dokumen ISO sumber sering memuat:
    1. Line-break manual (<w:br/>) di tengah judul panjang, mis. "...— Bagian
       1:" lalu baris baru "umum". python-docx membaca <w:br/> sebagai
       karakter '\\n' literal di para.text. Jika dibiarkan, '\\n' itu akan
       tertulis ulang sebagai <w:br/> lagi saat judul ditempatkan di
       cover/body title, membuat judul terlihat terpecah jadi beberapa baris
       padahal seharusnya satu baris penuh.
    2. Strip pemisah klausa yang tidak konsisten — kadang en dash "–", kadang
       em dash "—", dengan spasi yang juga tidak konsisten (nempel di salah
       satu/kedua sisi, atau spasi ganda).
    3. Huruf kecil di awal klausa setelah strip panjang atau setelah
       "Bagian N:"/"Part N:", padahal konvensi judul SNI menulisnya kapital.

    Fungsi ini menghasilkan format baku, contoh:
    "Aplikasi perkeretaapian — Perlindungan kebakaran pada kendaraan kereta
    api — Bagian 1: Umum" / "Railway applications — Fire protection on
    railway vehicles — Part 1: General".

    Langkah:
    1. Menggabungkan semua baris jadi satu baris dengan spasi tunggal.
    2. Merapikan strip panjang (en dash/em dash) jadi " — " (spasi-em
       dash-spasi). Hyphen pendek "-" TIDAK disentuh sama sekali.
    3. Mengkapitalisasi huruf pertama kata setelah " — " hasil rapikan di
       atas, dan huruf pertama kata setelah "Bagian N:" / "Part N:".
    """
    if not text:
        return text
    # Gabungkan baris (dari \n, \r\n, atau spasi berlebih) jadi satu baris.
    one_line = re.sub(r'\s*[\r\n]+\s*', ' ', text)
    one_line = re.sub(r'\s{2,}', ' ', one_line).strip()
    # Rapikan strip panjang (en dash/em dash) jadi " — ", lalu kapitalisasi
    # huruf pertama kata sesudahnya.
    one_line = _RE_LONG_DASH.sub(' \u2014 ', one_line)
    one_line = _RE_AFTER_LONG_DASH.sub(lambda m: m.group(1) + m.group(2).upper(), one_line)
    # Kapitalisasi huruf pertama kata setelah "Bagian N:" / "Part N:", dan
    # rapikan spasi di sekitar titik duanya jadi tepat satu spasi.
    one_line = _RE_BAGIAN_PART.sub(lambda m: m.group(1) + ': ' + m.group(2).upper(), one_line)
    return one_line


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
        title_id = _normalize_title_line(candidates[0]['text'])
        title_en = title_id
        return title_id, title_en
    except Exception: return "", ""

# --- INISIALISASI ENGINE ---
@st.cache_resource
def load_engines():
    e2 = DocxOptimizerEngine()
    e4 = CoverPageEngine()
    e5 = DaftarIsiEngine()
    e6 = PrakataPendahuluanEngine()
    e7 = InfoPendukungEngine()
    return e2, e4, e5, e6, e7

engine2, engine4, engine5, engine6, engine7 = load_engines()

# ─────────────────────────────────────────────────────────────────────────────
# LOAD KAMUS — SELALU fetch langsung dari Google Sheet, TIDAK di-cache.
# Setiap perubahan di Google Sheet (kamus SNI maupun kamus istilah asing)
# otomatis terpakai pada render halaman berikutnya / saat proses konversi
# dijalankan — tanpa perlu tombol refresh manual.
# Tidak memakai @st.fragment(run_every=...) supaya tidak ada polling
# berkala yang membuat halaman "berkedip"; fetch hanya terjadi pada
# titik yang memang butuh data kamus (page load & sesaat sebelum translate).
# ─────────────────────────────────────────────────────────────────────────────

def _load_kamus_fresh():
    """Fetch kamus terbaru dari Google Sheet. Dipanggil ulang setiap dibutuhkan."""
    _d = CustomDictionary()
    _n = _d.load_defaults()
    _i = ItalicDictionary()
    _ni = _i.load_defaults()
    return _d, _n, _i, _ni


def _render_header(kamus_count: int, italic_count: int):
    """Render header statis — tidak ada network request, tidak ada re-run otomatis."""
    _status_html = (
        f"""<div class="status-pill status-ready">
            <span class="status-dot"></span>
            Sistem Siap &nbsp;
        </div>"""
        if kamus_count > 0 else
        """<div class="status-pill status-warn">
            <span class="status-dot"></span>
            Kamus Tidak Aktif
        </div>"""
    )
    st.markdown(f"""
        <div class="app-header">
            <div class="badge">Generator RSNI</div>
            <h1>📑 ISO to RSNI Converter</h1>
            <p>Memformat & Menerjemahan Dokumen Standar ISO Menjadi Draft RSNI Secara Otomatis</p>
            <div class="stats-row">
                <div class="stat-item">
                    <div class="stat-num">8</div>
                    <div class="stat-lbl">Engine</div>
                </div>
                <div class="stat-divider"></div>
                <div class="stat-item">
                    <div class="stat-num">{kamus_count if kamus_count > 0 else '—'}</div>
                    <div class="stat-lbl">Kamus SNI</div>
                </div>
                <div class="stat-divider"></div>
                <div class="stat-item">
                    <div class="stat-num">{italic_count if italic_count > 0 else '—'}</div>
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

# Fetch kamus terbaru — dijalankan setiap kali halaman di-render (page load,
# upload file, klik tombol apa pun) sehingga perubahan di Google Sheet
# langsung terpakai tanpa tombol refresh manual.
_kamus_obj, _count, _italic_obj, _italic_count = _load_kamus_fresh()

st.session_state['custom_dict']  = _kamus_obj
st.session_state['italic_dict']  = _italic_obj
st.session_state['kamus_count']  = _count
st.session_state['italic_count'] = _italic_count

# Render header — angka kamus selalu mencerminkan isi Google Sheet terkini
_render_header(_count, _italic_count)

import datetime
_tahun = str(datetime.date.today().year)

# --- FORM INPUT ---
st.markdown('<div class="section-label">📂 Upload Dokumen ISO</div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Upload file .docx di sini atau klik Browse", type=["docx"], key="upl_main", label_visibility="collapsed")

st.markdown('<div class="section-label">⚙️ Pengaturan</div>', unsafe_allow_html=True)
col_set1, col_set2 = st.columns([3, 2])
with col_set1:
    doc_title = st.text_input("📄 Masukan No. SNI", value="SNI ISO xxxxx-x:xxxx", key="title_main")
with col_set2:
    src_lang = st.selectbox(
        "🌐 Bahasa Sumber",
        options=list(LANG_OPTIONS.keys()),
        index=0,
        format_func=lambda x: LANG_OPTIONS[x],
        key="lang_main"
    )

# --- TOMBOL PROSES ---
btn_process = st.button("🚀 Proses", key="btn_main", use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# LOGIC EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

if btn_process:
    if uploaded_file:
        target_file = f"temp_main_{uploaded_file.name}"
        with open(target_file, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        st.session_state['_run_process'] = True
        st.session_state['_target_file'] = target_file
        st.session_state['_doc_title'] = doc_title
        st.session_state['_src_lang'] = src_lang
        st.rerun()
    else:
        st.warning("Silakan upload file terlebih dahulu.")


if st.session_state.get('_run_process') and st.session_state.get('_target_file'):
    target_file = st.session_state['_target_file']
    doc_title_val = st.session_state['_doc_title']
    src_lang_val = st.session_state['_src_lang']
    
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
    def run_optimization(input_file, doc_title):
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
        # --- LANGKAH 1: OPTIMASI (0–10%) ---
        update_ui(1, "Membaca file dokumen...")
        final_opt_file = run_optimization(target_file, doc_title_val)
        st.session_state['_final_opt_file'] = final_opt_file

        # --- LANGKAH 2: TERJEMAHAN (10–100%) ---
        update_ui(10, "[6/6] Memuat kamus terbaru & menginisialisasi mesin terjemahan...")
        tr_out = f"ID_{os.path.basename(final_opt_file)}"

        # Refresh kamus dari Google Sheet TEPAT sebelum translate, supaya
        # perubahan terbaru di spreadsheet (kamus SNI / istilah asing)
        # pasti terpakai meskipun halaman sudah dimuat sebelum sheet diedit.
        _kamus_for_run, _kc, _italic_for_run, _ic = _load_kamus_fresh()
        st.session_state['custom_dict']  = _kamus_for_run
        st.session_state['italic_dict']  = _italic_for_run
        st.session_state['kamus_count']  = _kc
        st.session_state['italic_count'] = _ic

        _engine9 = DocxFinalTranslatorEngine(
            source_lang=src_lang_val,
            target_lang='id',
            custom_dict=_kamus_for_run,
            italic_dict=_italic_for_run,
        )

        # ── Callback: parse format baru engine9, throttle ≤1x/detik ─────────
        # Format pesan dari engine9:
        #   "[tag] aksi\tdone/total\tn_trans\tn_skip\tn_tbl\tpreview"
        _cb_t0      = [time.time()]   # waktu update terakhir
        _cb_total   = [0]             # total item (diisi dari pesan pertama)

        _ICON = {
            'translate': '✏️', 'done': '✏️',
            'skip': '⏭', 'kosong': '⏭', 'cover-italic': '⏭',
            'heading': '⏭', 'toc': '⏭', 'header': '⏭', 'footer': '⏭',
            'tabel': '📊', 'annex': '📎', 'bibliografi': '📚',
        }

        def _cb_tr(pct, msg):
            # ── Fallback pct (dipakai saat fase init/post-proc, bukan fase elemen) ──
            # Fase elemen akan override ini dengan kalkulasi done/total di bawah
            final_pct = min(10 + int(pct * 0.90), 100)

            # ── Parse format tab-separated dari engine9 ──────────────────────
            # Format: "[tag] aksi\tdone/total\tn_trans\tn_skip\tn_tbl\tpreview"
            parts = msg.split('\t')
            if len(parts) >= 6:
                tag_aksi  = parts[0].strip()   # "[cover-italic] skip"
                frac      = parts[1].strip()   # "42/567"
                n_trans   = parts[2].strip()   # "12"
                n_skip    = parts[3].strip()   # "28"
                n_tbl     = parts[4].strip()   # "3"
                preview   = parts[5].strip()   # cuplikan teks

                # Ambil tag dalam kurung siku
                m_tag = re.match(r'\[([^\]]+)\]', tag_aksi)
                tag   = m_tag.group(1) if m_tag else "?"
                aksi  = tag_aksi[m_tag.end():].strip() if m_tag else tag_aksi

                # Parse done & total dari fraksi "42/567"
                done_int, total_int = 0, 0
                if '/' in frac:
                    try:
                        done_int  = int(frac.split('/')[0])
                        total_int = int(frac.split('/')[1])
                    except Exception:
                        pass

                # Simpan total sekali
                if _cb_total[0] == 0 and total_int > 0:
                    _cb_total[0] = total_int
                total_ref = _cb_total[0] if _cb_total[0] > 0 else total_int

                # ── Progress dihitung dari done/total elemen (fase loop) ──────
                # Progress RIIL: elemen done/total → 10%–100%
                # elemen 0/total = 10%,  elemen total/total = 100%
                # Tidak ada cap di 65% — progress benar-benar mencerminkan pekerjaan
                if total_ref > 0 and done_int >= 0:
                    elem_ratio = min(done_int / total_ref, 1.0)
                    final_pct  = min(10 + int(elem_ratio * 90), 100)

                icon = _ICON.get(tag, _ICON.get(aksi, '🔄'))

                # Baris atas: statistik + batch info
                stat_parts = []
                if n_trans and n_trans != '0': stat_parts.append(f"✏️ {n_trans} diterjemah")
                if n_skip  and n_skip  != '0': stat_parts.append(f"⏭ {n_skip} dilewati")
                if n_tbl   and n_tbl   != '0': stat_parts.append(f"📊 {n_tbl} tabel")
                stat_str = "  ·  ".join(stat_parts) if stat_parts else "memulai..."

                total_str = f"/{total_ref}" if total_ref else ""
                line1 = f"[6/6] Translate  ·  elemen {done_int}{total_str}  ·  {stat_str}"

                # Baris bawah: aksi + preview teks saat ini
                if preview and preview != '-':
                    line2 = f"{icon} [{tag}] {aksi}  —  \"{preview}\""
                else:
                    line2 = f"{icon} [{tag}] {aksi}"

            else:
                # Pesan non-tab: init (pct 2–5) dan post-processing (pct 66–100)
                # Progress langsung dari pct engine9 → 10–100%
                if "[pemulihan" in msg.casefold():
                    # Tampilkan fase pemulihan secara eksplisit di dashboard,
                    # bukan sebagai pesan Translate generik. Jangan dipotong 100
                    # karakter agar counter YY/ZZ, berhasil, dan gagal terlihat.
                    line1 = "[6/6] Translate  ·  Pemulihan 1 worker"
                    line2 = f"🛠️ {msg.strip()}"
                else:
                    line1 = f"[6/6] Translate"
                    line2 = f"🔄 {msg.strip()[:100]}"

            # ── Selalu update progress bar (tidak ikut throttle) ─────────────
            progress_bar.progress(final_pct)

            # Pemulihan adalah informasi penting dan harus SELALU terlihat.
            # Sebelumnya callback recovery dapat terbuang oleh throttle 1 detik,
            # lalu langsung tertimpa pesan Sinkronisasi/Saving pada 98%.
            is_recovery = "[pemulihan" in msg.casefold()
            now = time.time()
            penting = (pct <= 5 or pct >= 96 or is_recovery)
            if not penting and (now - _cb_t0[0]) < 1.0:
                return
            _cb_t0[0] = now

            # Update status teks saja (progress bar sudah diupdate di atas)
            update_ui(final_pct, f"{line1}\n{line2}", skip_progress=True)
        # ─────────────────────────────────────────────────────────────────────

        ok_tr, _ = _engine9.translate(input_docx=final_opt_file, output_docx=tr_out, progress_callback=_cb_tr, translate_headers=False)
        
        if ok_tr:
            update_ui(100, "✅ Selesai!")
            final_elapsed = get_elapsed_str(start_time)
            # Ganti JS timer dengan waktu final statis (berhenti otomatis)
            time_placeholder.markdown(
                f'<div class="timer-text">⏱ {final_elapsed}</div>',
                unsafe_allow_html=True
            )
            st.session_state['_final_tr_file'] = tr_out
            st.session_state['_final_time'] = final_elapsed
            st.session_state['_show_results'] = True
            # Parse dokumen langsung agar chat langsung siap setelah rerun
            st.session_state['_doc_sections'] = _parse_doc_structure(tr_out)
        else:
            raise Exception("Terjemahan gagal.")

    except Exception as e:
        st.error(f"❌ Error Proses: {e}")
        st.session_state['_run_process'] = False
        st.session_state['_show_results'] = False
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

    col_res1, col_res2 = st.columns(2)

    opt_file = st.session_state.get('_final_opt_file')
    if opt_file and os.path.exists(opt_file):
        with col_res1:
            with open(opt_file, "rb") as f:
                st.download_button(
                    label="📄 Download Hasil Formating",
                    data=f,
                    file_name="ISO_Fixed_Document.docx",
                    use_container_width=True
                )

    tr_file = st.session_state.get('_final_tr_file')
    if tr_file and os.path.exists(tr_file):
        with col_res2:
            with open(tr_file, "rb") as f:
                st.download_button(
                    label="🌐 Download Terjemahan ID",
                    data=f,
                    file_name=f"ID_{os.path.basename(opt_file) if opt_file else 'document.docx'}",
                    use_container_width=True
                )

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    if st.button("🔄 Proses File Baru", key="reset", use_container_width=True):
        # Hapus file sesi ini segera sebelum reset
        _cleanup_session_files(st.session_state)
        for k in ['_show_results', '_final_opt_file', '_final_tr_file', '_final_time', '_run_process',
                  '_doc_text', '_chat_history', '_doc_sections', '_target_file']:
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


# --- FOOTER ---
_now = time.strftime("%H:%M")
st.markdown(
    f"<div class='footer'>"
    f"<span style='font-size:0.85rem;'>"
    f"<a href='https://docs.google.com/spreadsheets/d/1BBPCMPwvbBk5LPdoDQwnjQzcPHv7_RDKENqeMsklF-8/edit?usp=sharing' target='_blank' style='color:#ffffff;text-decoration:none;'>📖 Kamus SNI</a>"
    f"&nbsp;&nbsp;·&nbsp;&nbsp;"
    f"<a href='https://docs.google.com/spreadsheets/d/1NZm1HjsjxmflxnZlzV_O2XF75ZlMUOu8VVofsKfp_FA/edit?usp=sharing' target='_blank' style='color:#ffffff;text-decoration:none;'>🌐 Kamus Istilah Asing</a>"
    f"<br>"
    f"<a href='https://iec-to-iso.streamlit.app/' target='_blank' style='color:#ffffff;text-decoration:none;'> 📄 IEC to ISO Converter</a>"
    f"</span>"
    f"<br>"
    f"<span style='font-size:0.8rem;color:#ffffff;'>© 2026 Generator RSNI · ISO to RSNI Converter · All rights reserved.</span>"
    f"<br>"
    f"<span style='font-size:0.72rem;opacity:0.8;color:#ffffff;'>Developed by Denny Kusuma H. & Ahmad Habibi</span>"
    f"</div>",
    unsafe_allow_html=True
)
