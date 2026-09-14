"""
Engine8: DocTranslator
=============================
Perilaku yang dinonaktifkan:
  - Tidak menyisipkan konten asli berbahasa Inggris sebelum Bibliografi.
  - Tidak menghapus atau mengubah hyperlink. Paragraf yang mengandung
    hyperlink dipertahankan utuh, termasuk teks, struktur, dan URL.

Fitur utama:
  - 2 SPREADSHEET TERPISAH:
    1. KAMUS_SPREADSHEET_URL → untuk terjemahan istilah
    2. ITALIC_SPREADSHEET_URL → daftar kata yang TIDAK diterjemahkan, OUTPUT MIRING
  - Dual Title Sync, Note/Catatan, Annex fix, dll.
"""

import re
import copy
import time
import uuid
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
import csv
import os

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import lxml.etree as etree


# ─────────────────────────────────────────────────────────────────────────────
# NAMESPACE & KONSTANTA
# ─────────────────────────────────────────────────────────────────────────────

_NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_W    = f'{{{_NS_W}}}'

_RE_PURE_NUMBER = re.compile(r'^[\d\s\.\,\:\;\-\(\)\[\]\/\\\+\=\*\%\&\^\$\#\@\!\"\'`~<>{}|_]+$')
_RE_COPYRIGHT = re.compile(r'©|BSN\s*\d{4}', re.IGNORECASE)
# Paragraf pendek berupa nama/singkatan resmi yang berdiri sendiri
# (mis. baris "BSN" saja pada kotak hak cipta) — tidak perlu, dan tidak
# boleh, dikirim ke Google Translate. Dicocokkan hanya jika SELURUH isi
# paragraf persis salah satu token ini (bukan sekadar mengandungnya).
_RE_STANDALONE_ACRONYM = re.compile(r'^(BSN|SNI|ISO|IEC)$', re.IGNORECASE)

_SKIP_STYLES = {
    'caption', 'header', 'footer',
    'toc 1', 'toc 2', 'toc 3', 'toc 4', 'toc 5',
    'table of figures', 'footnote text', 'endnote text', 'macro text',
}
_BIBLIO_TITLE_STYLES = {'biblioti', 'bibliotitle', 'bibliography title'}
_BIBLIO_KEYWORDS_EXACT = {
    'bibliografi', 'bibliography',
    'daftar acuan', 'daftar pustaka', 'daftar referensi',
}
_ANNEX_STYLE_IDS = {'ANNEX', 'Annex', 'annex'}
# Paragraf yang ditandai engine6 (Prakata/Pendahuluan) — sudah final berbahasa
# Indonesia, JANGAN diterjemahkan ulang di sini.
_NO_TRANSLATE_STYLE_IDS = {'BSNNoTranslate'}
_HEADING_STYLES_WITH_NUM = {
    'Heading1', 'Heading2', 'Heading3',
    'ANNEX', 'a2', 'a3',
    'Heading4', 'Heading5', 'Heading6',
}
_TRANSLATE_DELAY = 0.15
_EM_DASH = '—'

# ─────────────────────────────────────────────────────────────────────────────
# 2 URL SPREADSHEET TERPISAH
# ─────────────────────────────────────────────────────────────────────────────

KAMUS_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1BBPCMPwvbBk5LPdoDQwnjQzcPHv7_RDKENqeMsklF-8/edit?usp=sharing"
ITALIC_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1NZm1HjsjxmflxnZlzV_O2XF75ZlMUOu8VVofsKfp_FA/edit#gid=0"


# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM DICTIONARY (SPREADSHEET 1 - KAMUS TERJEMAHAN)
# ─────────────────────────────────────────────────────────────────────────────

class CustomDictionary:
    """Kamus istilah: source → target (AKAN diterjemahkan)."""
    
    def __init__(self):
        self._entries: dict[str, tuple[str, str]] = {}

    def add_term(self, source: str, target: str) -> None:
        s = source.strip()
        t = target.strip()
        if s and t:
            self._entries[s.lower()] = (s, t)

    def clear(self) -> None:
        self._entries.clear()

    def load_defaults(self) -> int:
        try:
            return self.load_from_google_sheet(KAMUS_SPREADSHEET_URL)
        except Exception:
            return 0

    def load_from_csv(self, filepath: str, src_col: str = 'source', 
                      tgt_col: str = 'target', delimiter: str = ',', 
                      encoding: str = 'utf-8-sig') -> int:
        if not os.path.isfile(filepath): 
            raise FileNotFoundError(f"File CSV tidak ditemukan: {filepath}")
        count = 0
        with open(filepath, newline='', encoding=encoding) as f:
            sample = f.read(1024); f.seek(0)
            has_header = csv.Sniffer().has_header(sample)
            reader = csv.DictReader(f, delimiter=delimiter) if has_header else csv.reader(f, delimiter=delimiter)
            for row in reader:
                if has_header:
                    src = row.get(src_col, row.get('source', '')).strip()
                    tgt = row.get(tgt_col, row.get('target', '')).strip()
                else:
                    row = list(row)
                    if len(row) < 2: continue
                    src, tgt = row[0].strip(), row[1].strip()
                if src and tgt:
                    self._entries[src.lower()] = (src, tgt)
                    count += 1
        return count

    def load_from_excel(self, filepath: str, sheet_name: str | int = 0, 
                        src_col: str = 'source', tgt_col: str = 'target') -> int:
        if not os.path.isfile(filepath): 
            raise FileNotFoundError(f"File Excel tidak ditemukan: {filepath}")
        try: import pandas as pd
        except ImportError: raise ImportError("Jalankan: pip install pandas openpyxl")
        df = pd.read_excel(filepath, sheet_name=sheet_name, dtype=str).fillna('')
        col_src = _find_col(df.columns.tolist(), [src_col, 'source', 'Inggris'])
        col_tgt = _find_col(df.columns.tolist(), [tgt_col, 'target', 'Indonesia'])
        if col_src is None: col_src = df.columns[0]
        if col_tgt is None and len(df.columns) >= 2: col_tgt = df.columns[1]
        if col_tgt is None: raise ValueError("Kolom target tidak ditemukan.")
        count = 0
        for _, row in df.iterrows():
            src = str(row[col_src]).strip()
            tgt = str(row[col_tgt]).strip()
            if src and tgt and src.lower() not in ('nan', '') and tgt.lower() not in ('nan', ''):
                self._entries[src.lower()] = (src, tgt); count += 1
        return count

    def load_from_google_sheet(self, url: str, src_col: str = 'source', 
                               tgt_col: str = 'target', timeout: int = 15) -> int:
        try: import urllib.request, io
        except ImportError: raise ImportError("urllib tidak tersedia.")
        csv_url = _google_sheet_to_csv_url(url)
        try:
            req = urllib.request.Request(csv_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=timeout) as resp: 
                raw = resp.read().decode('utf-8-sig')
        except Exception as e: raise ConnectionError(f"Gagal mengambil data: {e}")
        f = io.StringIO(raw); reader = csv.DictReader(f); fieldnames = reader.fieldnames or []
        col_src = _find_col(
            fieldnames, [src_col, 'source', 'Inggris', 'Bahasa Inggris', 'English']
        )
        col_tgt = _find_col(
            fieldnames, [tgt_col, 'target', 'Indonesia', 'Bahasa Indonesia']
        )
        if col_src is None and len(fieldnames) >= 1: col_src = fieldnames[0]
        if col_tgt is None and len(fieldnames) >= 2: col_tgt = fieldnames[1]
        if col_tgt is None: raise ValueError(f"Kolom target tidak ditemukan. Header: {fieldnames}")
        count = 0
        for row in reader:
            src = str(row.get(col_src, '')).strip()
            tgt = str(row.get(col_tgt, '')).strip()
            if src and tgt and src.lower() not in ('', 'nan') and tgt.lower() not in ('', 'nan'):
                self._entries[src.lower()] = (src, tgt); count += 1
        return count

    def __len__(self) -> int: return len(self._entries)
    def list_terms(self) -> list[tuple[str, str]]:
        return [(s, t) for _, (s, t) in sorted(self._entries.items())]

    def _apply_pre(self, text: str) -> tuple[str, dict]:
        if not self._entries: return text, {}
        token_map = {}; result = text.replace(' ', ' ')
        sorted_entries = sorted(self._entries.items(), key=lambda x: len(x[0]), reverse=True)
        for src_lower, (src_orig, tgt) in sorted_entries:
            pattern = re.compile(r'(?<![A-Za-z0-9])' + re.escape(src_lower) + r'(?![A-Za-z0-9])', re.IGNORECASE)
            if pattern.search(result):
                # Token huruf-angka tanpa simbol @/_ jauh lebih stabil saat
                # melewati Google Translate dan tetap mudah divalidasi.
                token = f'ZXQTK{uuid.uuid4().hex[:12].upper()}QXZ'
                token_map[token] = tgt
                result = pattern.sub(token, result)
        return result, token_map

    def _apply_post(self, translated: str, token_map: dict) -> str:
        result = translated
        for token, tgt in token_map.items(): 
            result = re.sub(re.escape(token), tgt, result, flags=re.IGNORECASE)
        return result


# ─────────────────────────────────────────────────────────────────────────────
# ITALIC DICTIONARY (SPREADSHEET 2 - KATA MIRING)
# ─────────────────────────────────────────────────────────────────────────────

class ItalicDictionary:
    """
    Daftar kata/frasa dari Spreadsheet 2.
    TIDAK diterjemahkan, otomatis bercetak MIRING di dokumen output.
    """

    def __init__(self):
        self._entries: dict[str, str] = {}

    def add_term(self, term: str) -> None:
        t = term.strip()
        if t:
            self._entries[t.lower()] = t

    def clear(self) -> None:
        self._entries.clear()

    def load_defaults(self) -> int:
        try:
            return self.load_from_google_sheet(ITALIC_SPREADSHEET_URL)
        except Exception:
            return 0

    def load_from_csv(self, filepath: str, term_col: str = 'term', 
                      delimiter: str = ',', encoding: str = 'utf-8-sig') -> int:
        if not os.path.isfile(filepath): 
            raise FileNotFoundError(f"File CSV tidak ditemukan: {filepath}")
        count = 0
        with open(filepath, newline='', encoding=encoding) as f:
            sample = f.read(1024); f.seek(0)
            has_header = csv.Sniffer().has_header(sample)
            reader = csv.DictReader(f, delimiter=delimiter) if has_header else csv.reader(f, delimiter=delimiter)
            skip_vals = {'nan', '', 'term', 'kata', 'italic', 'word', 'text'}
            for row in reader:
                if has_header:
                    term = ''
                    for col_name in [term_col, 'term', 'kata', 'italic', 'word', 'text']:
                        if col_name in row:
                            term = row.get(col_name, '').strip()
                            break
                    if not term and len(row) > 0:
                        term = list(row.values())[0].strip()
                else:
                    row = list(row)
                    if len(row) < 1: continue
                    term = row[0].strip()
                if term and term.lower() not in skip_vals:
                    self._entries[term.lower()] = term; count += 1
        return count

    def load_from_excel(self, filepath: str, sheet_name: str | int = 0, 
                        term_col: str = 'term') -> int:
        if not os.path.isfile(filepath): 
            raise FileNotFoundError(f"File Excel tidak ditemukan: {filepath}")
        try: import pandas as pd
        except ImportError: raise ImportError("Jalankan: pip install pandas openpyxl")
        df = pd.read_excel(filepath, sheet_name=sheet_name, dtype=str).fillna('')
        col_term = None
        for col_name in [term_col, 'term', 'kata', 'italic', 'word', 'text']:
            if col_name in df.columns:
                col_term = col_name; break
        if col_term is None: col_term = df.columns[0]
        count = 0
        skip_vals = {'nan', '', 'term', 'kata', 'italic'}
        for _, row in df.iterrows():
            term = str(row[col_term]).strip()
            if term and term.lower() not in skip_vals:
                self._entries[term.lower()] = term; count += 1
        return count

    def load_from_google_sheet(self, url: str, term_col: str = 'term', 
                               timeout: int = 15) -> int:
        try: import urllib.request, io
        except ImportError: raise ImportError("urllib tidak tersedia.")
        csv_url = _google_sheet_to_csv_url(url)
        try:
            req = urllib.request.Request(csv_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=timeout) as resp: 
                raw = resp.read().decode('utf-8-sig')
        except Exception as e: raise ConnectionError(f"Gagal mengambil data dari Spreadsheet Italic: {e}")
        f = io.StringIO(raw); reader = csv.DictReader(f); fieldnames = reader.fieldnames or []
        col_term = _find_col(
            fieldnames,
            [term_col, 'term', 'kata', 'italic', 'word', 'text',
             'istilah asing', 'bahasa asing'],
        )
        if col_term is None and len(fieldnames) >= 1:
            col_term = fieldnames[0]
        if col_term is None: raise ValueError(f"Tidak ada kolom. Header: {fieldnames}")
        count = 0
        skip_vals = {'nan', '', 'term', 'kata', 'italic', 'word', 'text'}
        for row in reader:
            term = str(row.get(col_term, '')).strip()
            if term and term.lower() not in skip_vals:
                self._entries[term.lower()] = term; count += 1
        return count

    def __len__(self) -> int: return len(self._entries)
    
    def list_terms(self) -> list[str]:
        return list(set(self._entries.values()))

    def _apply_pre(self, text: str) -> tuple[str, dict]:
        if not self._entries: return text, {}
        token_map = {}
        result = text
        sorted_entries = sorted(self._entries.items(), key=lambda x: len(x[0]), reverse=True)
        for term_lower, term_orig in sorted_entries:
            pattern = re.compile(r'(?<![A-Za-z0-9])' + re.escape(term_lower) + r'(?![A-Za-z0-9])', re.IGNORECASE)
            if pattern.search(result):
                token = f'ZXQIT{uuid.uuid4().hex[:12].upper()}QXZ'
                token_map[token] = term_orig
                result = pattern.sub(token, result)
        return result, token_map

    def _apply_post(self, translated: str, italic_map: dict) -> tuple[str, list[str]]:
        result = translated
        italic_terms_found = []
        for token, original in italic_map.items():
            if re.search(re.escape(token), result, re.IGNORECASE):
                result = re.sub(re.escape(token), original, result, flags=re.IGNORECASE)
                if original not in italic_terms_found:
                    italic_terms_found.append(original)
        return result, italic_terms_found


# ─────────────────────────────────────────────────────────────────────────────
# SHARED HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _find_col(columns: list, candidates: list) -> str | None:
    normalized = {
        re.sub(r'[^a-z0-9]+', '', str(col).casefold()): col
        for col in columns
    }
    for candidate in candidates:
        key = re.sub(r'[^a-z0-9]+', '', str(candidate).casefold())
        if key in normalized:
            return normalized[key]
    return None

def _google_sheet_to_csv_url(url: str) -> str:
    url = url.strip()
    if 'output=csv' in url or 'format=csv' in url: return url
    if 'docs.google.com/spreadsheets' not in url: return url
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if not m: raise ValueError(f"Tidak dapat mengekstrak Sheet ID: {url}")
    sheet_id = m.group(1)
    gid_match = re.search(r'[#&?]gid=(\d+)', url)
    gid_param = f'&gid={gid_match.group(1)}' if gid_match else ''
    return f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv{gid_param}'


# ─────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def _skip_text(text: str) -> bool:
    t = text.strip()
    if len(t) < 3: return True
    if _RE_PURE_NUMBER.fullmatch(t): return True
    if _RE_COPYRIGHT.search(t): return True
    if _RE_STANDALONE_ACRONYM.fullmatch(t): return True
    return False

def _skip_paragraph(para, past_bibliography: bool = False) -> bool:
    if past_bibliography: return True
    if not para.text.strip(): return True
    if _get_para_style_id(para) in _NO_TRANSLATE_STYLE_IDS: return True
    for tag in [f'{_W}drawing', f'{_W}pict']:
        if para._element.find('.//' + tag) is not None: return True
    style_name = (para.style.name or '').lower() if para.style is not None else ''
    if any(style_name.startswith(s) for s in _SKIP_STYLES): return True
    return False

def _is_biblio_title_para(para) -> bool:
    style_id = ''
    try:
        pStyle = para._element.find(f'{_W}pPr/{_W}pStyle')
        if pStyle is not None: style_id = pStyle.get(f'{_W}val', '').lower()
    except Exception: pass
    if style_id in _BIBLIO_TITLE_STYLES: return True
    txt = para.text.strip().lower()
    return bool(txt) and not txt[0].isdigit() and txt in _BIBLIO_KEYWORDS_EXACT

def _get_para_style_id(para) -> str:
    pStyle = para._element.find(f'{_W}pPr/{_W}pStyle')
    return pStyle.get(f'{_W}val', '') if pStyle is not None else ''

def _get_style_id(el) -> str:
    pStyle = el.find(f'{_W}pPr/{_W}pStyle')
    return pStyle.get(f'{_W}val', '') if pStyle is not None else 'Normal'

def _has_inline_sectpr(para) -> bool:
    pPr = para._element.find(f'{_W}pPr')
    return pPr is not None and pPr.find(f'{_W}sectPr') is not None

def _all_runs_italic(para) -> bool:
    text_runs = [r for r in para.runs if r.text.strip()]
    if not text_runs: return False
    para_style_italic = False
    try:
        if para.style and para.style.font and para.style.font.italic: para_style_italic = True
    except Exception: pass
    for run in text_runs:
        italic = False
        if run.font.italic is True: italic = True
        if not italic:
            rPr = run._element.find(f'{_W}rPr')
            if rPr is not None:
                i_el = rPr.find(f'{_W}i')
                if i_el is not None:
                    val = i_el.get(f'{_W}val', 'true')
                    if val.lower() not in ('false', '0'): italic = True
        if not italic and para_style_italic:
            rPr = run._element.find(f'{_W}rPr')
            if rPr is not None:
                i_el = rPr.find(f'{_W}i')
                if i_el is not None:
                    val = i_el.get(f'{_W}val', 'true')
                    if val.lower() not in ('false', '0'): italic = True
                else: italic = True
            else: italic = True
        if not italic: return False
    return True


# ─────────────────────────────────────────────────────────────────────────────
# PROTEKSI ISTILAH/JUDUL ASING YANG SUDAH MIRING DI SUMBER
# ─────────────────────────────────────────────────────────────────────────────
# Di dokumen ISO, judul dokumen acuan (mis. pada klausul "Acuan normatif")
# dan istilah asing lain biasanya SUDAH dicetak miring di file sumber.
# Sesuai konvensi SNI, bagian ini TIDAK diterjemahkan dan tetap dicetak
# miring pada dokumen hasil. Sebelumnya bagian ini "hilang" karena
# _translate_para menggabungkan semua run menjadi satu string polos
# sebelum diterjemahkan, sehingga info format miring per-run dibuang dan
# teksnya ikut diterjemahkan seperti teks biasa. Fungsi di bawah ini
# mendeteksi run miring tersebut dan melindunginya dengan mekanisme
# token yang sama seperti "Kamus Istilah Asing" (spreadsheet).

def _get_para_style_italic(para) -> bool:
    try:
        if para.style and para.style.font and para.style.font.italic: return True
    except Exception: pass
    return False

def _run_effective_italic(run, para_style_italic: bool) -> bool:
    if run.font.italic is True: return True
    if run.font.italic is False: return False
    rPr = run._element.find(f'{_W}rPr')
    if rPr is not None:
        i_el = rPr.find(f'{_W}i')
        if i_el is not None:
            val = i_el.get(f'{_W}val', 'true')
            return val.lower() not in ('false', '0')
        return para_style_italic
    return para_style_italic

def _extract_source_italic_map(text_runs, para_style_italic: bool) -> tuple[str, dict]:
    """
    Gabungkan run menjadi satu teks (seperti semula), TAPI bagian yang
    sudah diformat miring di sumber diganti token dulu agar tidak ikut
    diterjemahkan. token_map: token -> teks asli (verbatim, akan
    dikembalikan + dicetak miring setelah proses terjemahan selesai).
    """
    segments = []
    for _, r in text_runs:
        is_ital = _run_effective_italic(r, para_style_italic)
        t = r.text or ''
        if segments and segments[-1][1] == is_ital:
            segments[-1] = [segments[-1][0] + t, is_ital]
        else:
            segments.append([t, is_ital])

    token_map = {}
    out_parts = []
    for seg_text, is_ital in segments:
        stripped = seg_text.strip()
        if is_ital and len(stripped) >= 3 and not _RE_PURE_NUMBER.fullmatch(stripped):
            token = f'ZXQSRC{uuid.uuid4().hex[:12].upper()}QXZ'
            token_map[token] = seg_text
            out_parts.append(token)
        else:
            out_parts.append(seg_text)
    return ''.join(out_parts), token_map


# ─────────────────────────────────────────────────────────────────────────────
# PROTEKSI PANGKAT (SUPERSCRIPT) & INDEKS (SUBSCRIPT)
# ─────────────────────────────────────────────────────────────────────────────
# Dokumen ISO sering memuat notasi teknis dengan pangkat/indeks, misalnya
# "kWm⁻²" atau rumus kimia semacam "CO₂". Sebelumnya _translate_para
# menggabungkan seluruh run paragraf menjadi satu string polos sebelum
# dikirim ke mesin terjemahan, sehingga informasi w:vertAlign (superscript/
# subscript) per-run ikut hilang dan hasil akhirnya tampil rata (tidak lagi
# berupa pangkat/indeks) — baik pada dokumen hasil format (Inggris) maupun
# hasil terjemahan (Indonesia). Fungsi-fungsi berikut melindungi teks yang
# berformat superscript/subscript dengan mekanisme token yang sama seperti
# proteksi miring di atas, agar formatnya bisa dikembalikan persis setelah
# proses terjemahan selesai.

def _run_vertalign(run) -> str | None:
    """Kembalikan 'superscript', 'subscript', atau None sesuai format run."""
    try:
        if run.font.superscript:
            return 'superscript'
        if run.font.subscript:
            return 'subscript'
    except Exception:
        pass
    # Fallback baca langsung XML — beberapa run tidak selalu terbaca lewat
    # properti python-docx bila nilai w:vertAlign bukan True/False sederhana.
    rPr = run._element.find(f'{_W}rPr')
    if rPr is not None:
        va_el = rPr.find(f'{_W}vertAlign')
        if va_el is not None:
            val = va_el.get(f'{_W}val', '')
            if val == 'superscript': return 'superscript'
            if val == 'subscript': return 'subscript'
    return None


def _extract_source_format_map(text_runs, para_style_italic: bool) -> tuple[str, dict]:
    """
    Versi gabungan dari _extract_source_italic_map yang JUGA melindungi
    superscript/subscript. Menggabungkan run menjadi satu teks, tapi:
      - Segmen superscript/subscript SELALU dilindungi token (berapa pun
        panjangnya — bisa cuma 1 karakter seperti pangkat "2"), supaya
        teks & formatnya persis sama setelah diterjemahkan.
      - Segmen miring (tanpa superscript/subscript) tetap memakai aturan
        lama (istilah/judul asing yang sudah miring di sumber).
    token_map: token -> {'text': teks asli, 'italic': bool, 'vtype': str|None}
    """
    segments = []
    for _, r in text_runs:
        is_ital = _run_effective_italic(r, para_style_italic)
        vtype = _run_vertalign(r)
        key = (is_ital, vtype)
        t = r.text or ''
        if segments and segments[-1][1] == key:
            segments[-1][0] += t
        else:
            segments.append([t, key])

    token_map = {}
    out_parts = []
    for seg_text, (is_ital, vtype) in segments:
        stripped = seg_text.strip()
        if not stripped:
            out_parts.append(seg_text)
            continue
        protect = False
        if vtype is not None:
            protect = True
        elif is_ital and len(stripped) >= 3 and not _RE_PURE_NUMBER.fullmatch(stripped):
            protect = True
        if protect:
            token = f'ZXQSRC{uuid.uuid4().hex[:12].upper()}QXZ'
            token_map[token] = {'text': seg_text, 'italic': is_ital, 'vtype': vtype}
            out_parts.append(token)
        else:
            out_parts.append(seg_text)
    return ''.join(out_parts), token_map


def _detokenize_with_formatting(text: str, token_map: dict) -> tuple[str, list]:
    """
    Kembalikan token @@SRC_...@@ pada `text` menjadi teks aslinya (verbatim),
    sekaligus catat posisi (start, end, italic, vtype) di teks HASIL agar
    formatnya (miring/superscript/subscript) bisa diterapkan secara presisi
    saat membangun ulang run — tanpa perlu mencari-cari substring lagi.
    """
    if not token_map:
        return text, []
    pattern = re.compile(
        '|'.join(re.escape(t) for t in token_map.keys()), re.IGNORECASE
    )
    token_map_folded = {token.casefold(): info for token, info in token_map.items()}
    spans = []
    chunks = []
    cur_len = 0
    last_end = 0
    for m in pattern.finditer(text):
        plain = text[last_end:m.start()]
        chunks.append(plain)
        cur_len += len(plain)
        info = token_map_folded[m.group(0).casefold()]
        orig_text = info['text']
        spans.append((cur_len, cur_len + len(orig_text), info.get('italic', False), info.get('vtype')))
        chunks.append(orig_text)
        cur_len += len(orig_text)
        last_end = m.end()
    chunks.append(text[last_end:])
    return ''.join(chunks), spans


# ─────────────────────────────────────────────────────────────────────────────
# DETEKSI HYPERLINK
# ─────────────────────────────────────────────────────────────────────────────

def _has_hyperlinks(para) -> bool:
    """Cek apakah paragraf memiliki hyperlink."""
    return para._element.find(f'{_W}hyperlink') is not None

# ─────────────────────────────────────────────────────────────────────────────
# ITALIC FORMATTING
# ─────────────────────────────────────────────────────────────────────────────

def _apply_mixed_formatting_to_para(para, text: str, italic_terms: list[str],
                                     font_name: str = None, font_size: int = None,
                                     format_spans: list = None) -> None:
    """
    Terapkan formatting ke paragraf hasil terjemahan:
      - `format_spans`: posisi PRESISI (start, end, italic, vtype) hasil
        dari _detokenize_with_formatting — dipakai untuk mengembalikan
        superscript/subscript (dan miring sumber) persis seperti aslinya.
      - `italic_terms`: daftar istilah (dari Kamus Istilah Asing) yang perlu
        dicari lagi posisinya di teks (pendekatan lama, best-effort, hanya
        untuk MIRING, tidak pernah untuk superscript/subscript).
    """
    format_spans = format_spans or []

    # Posisi yang sudah dipakai oleh format_spans (presisi) — istilah kamus
    # tidak boleh menimpa area ini.
    covered = [(s, e) for s, e, _, _ in format_spans]

    def _overlaps_covered(s, e):
        for cs, ce in covered:
            if s < ce and e > cs:
                return True
        return False

    dict_positions = []
    if italic_terms:
        text_lower = text.lower()
        for term in italic_terms:
            term_lower = term.lower()
            start = 0
            while True:
                idx = text_lower.find(term_lower, start)
                if idx == -1: break
                dict_positions.append((idx, idx + len(term)))
                start = idx + 1
        dict_positions.sort(key=lambda x: x[0])
        filtered_dict = []
        last_end = -1
        for start, end in dict_positions:
            if start >= last_end and not _overlaps_covered(start, end):
                filtered_dict.append((start, end, True, None))
                last_end = end
        dict_positions = filtered_dict

    all_spans = sorted(list(format_spans) + dict_positions, key=lambda x: x[0])

    if not all_spans:
        if para.runs:
            para.runs[0].text = text
            for r in para.runs[1:]: r.text = ''
        else:
            para.add_run(text)
        return

    segments = []
    last_pos = 0
    for start, end, is_italic, vtype in all_spans:
        if start < last_pos:
            continue  # lewati span yang tumpang tindih (seharusnya sudah difilter)
        if start > last_pos:
            segments.append((text[last_pos:start], False, None))
        segments.append((text[start:end], is_italic, vtype))
        last_pos = end

    if last_pos < len(text):
        segments.append((text[last_pos:], False, None))

    # Salin direct run formatting dari Engine 7 (font, ukuran, bold, warna,
    # underline, language, dll.) sebagai basis semua run baru. Style paragraf,
    # numbering, indent, spacing, alignment, dan page-break tidak pernah diubah.
    source_runs = list(para.runs)
    base_rpr = None
    for source_run in source_runs:
        if source_run.text.strip() and not _run_effective_italic(
            source_run, _get_para_style_italic(para)
        ):
            if source_run._element.rPr is not None:
                base_rpr = copy.deepcopy(source_run._element.rPr)
            break
    if base_rpr is None:
        for source_run in source_runs:
            if source_run.text.strip() and source_run._element.rPr is not None:
                base_rpr = copy.deepcopy(source_run._element.rPr)
                break

    for run in source_runs:
        run._element.getparent().remove(run._element)

    for seg_text, is_italic, vtype in segments:
        if not seg_text: continue

        run = para.add_run(seg_text)
        if base_rpr is not None:
            if run._element.rPr is not None:
                run._element.remove(run._element.rPr)
            run._element.insert(0, copy.deepcopy(base_rpr))
        else:
            run.font.name = font_name or 'Arial'
            if font_size:
                run.font.size = Pt(font_size)

        run.italic = bool(is_italic)
        run.font.superscript = False
        run.font.subscript = False
        if vtype == 'superscript':
            run.font.superscript = True
        elif vtype == 'subscript':
            run.font.subscript = True


# ─────────────────────────────────────────────────────────────────────────────
# FITUR 1: REKONSTRUKSI ANNEX
# ─────────────────────────────────────────────────────────────────────────────

def _fix_annex_style_para(para, annex_letter: str = None) -> str:
    """Rekonstruksi paragraf judul Annex/Lampiran menjadi 'Lampiran X'.
    Mengembalikan huruf Lampiran (mis. 'A', 'B', 'C') yang benar-benar dipakai,
    supaya pemanggil bisa menyinkronkan penomoran sub-pasal (a2/a3) di bawahnya
    dengan huruf yang sama persis (lihat _fix_annex_sub_para)."""
    sid = _get_para_style_id(para)
    if sid not in _ANNEX_STYLE_IDS: return annex_letter
    full_text = ''.join(r.text for r in para.runs if r.text is not None).strip()
    if not full_text: return annex_letter
    tag_norm = None; title_part = full_text
    for t in ['(informatif)', '(normatif)', '(informative)', '(normative)', '(informasi)']:
        idx = full_text.lower().find(t)
        if idx != -1:
            tag_norm = '(normatif)' if 'norm' in t.lower() else '(informatif)'
            title_part = full_text[:idx] + " " + full_text[idx + len(t):]
            break
    annex_label = None
    resolved_letter = annex_letter
    m_annex = re.match(r'^(?:Annex|Lampiran)\s+([A-Z0-9\.]+)\s*', title_part, flags=re.IGNORECASE)
    if m_annex:
        resolved_letter = m_annex.group(1).upper()
        annex_label = f'Lampiran {resolved_letter}'
        title_part = title_part[m_annex.end():].strip()
    elif annex_letter:
        resolved_letter = annex_letter.upper()
        annex_label = f'Lampiran {resolved_letter}'
        title_part = re.sub(r'^(?:Annex|Lampiran)\s*\S*\s*', '', title_part, flags=re.IGNORECASE).strip()
    else:
        title_part = re.sub(r'^(?:Annex|Lampiran)\s*\S*\s*', '', title_part, flags=re.IGNORECASE).strip()
    title_part = title_part.lstrip('\n').strip()
    pPr = para._element.find(f'{_W}pPr')
    if pPr is None: pPr = etree.SubElement(para._element, f'{_W}pPr')
    old_numPr = pPr.find(f'{_W}numPr')
    if old_numPr is not None: pPr.remove(old_numPr)
    pPr.insert(0, etree.fromstring(f'<w:numPr xmlns:w="{_NS_W}"><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'))
    for child in list(para._element):
        if child is not pPr: para._element.remove(child)
    def mk(text, bold=False):
        esc = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        b = '<w:b/><w:bCs/>' if bold else ''
        return etree.fromstring(f'<w:r xmlns:w="{_NS_W}"><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>{b}<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">{esc}</w:t></w:r>')
    new_runs = []
    if annex_label: new_runs.extend([mk(annex_label, True), etree.fromstring(f'<w:r xmlns:w="{_NS_W}"><w:br/></w:r>')])
    else: new_runs.append(etree.fromstring(f'<w:r xmlns:w="{_NS_W}"><w:br/></w:r>'))
    if tag_norm: new_runs.append(mk(tag_norm, True))
    if title_part: new_runs.extend([etree.fromstring(f'<w:r xmlns:w="{_NS_W}"><w:br/></w:r>'), mk(title_part, True)])
    for run_el in new_runs: para._element.append(run_el)
    return resolved_letter


def _fix_annex_sub_para(para, number_prefix: str) -> None:
    """Sisipkan nomor pasal literal (mis. 'B.1    ') di depan judul sub-pasal
    Annex/Lampiran (style 'a2'/'a3') dan matikan numPr bawaan Word di paragraf
    tsb.

    BUG YANG DIPERBAIKI: paragraf judul Annex (style ANNEX) sengaja dimatikan
    penomoran otomatisnya oleh _fix_annex_style_para (numId di-set ke 0) agar
    teks 'Lampiran A/B/C' bisa ditulis manual. Tapi ini membuat counter level-0
    dari daftar bertingkat (multilevel list) milik style 'a2'/'a3' TIDAK PERNAH
    di-restart ketika masuk Lampiran baru -- akibatnya nomor pasal di Lampiran
    B, C, dst tetap numeric lanjut dari Lampiran A namun hurufnya nyangkut di
    'A' (mis. 'A.6', 'A.7' ... padahal seharusnya 'B.1', 'B.2', ...). Solusinya
    sama seperti judul Annex: matikan numPr paragraf ini dan tulis nomor pasal
    (huruf Lampiran + nomor urut yang di-reset di python) sebagai teks literal.
    """
    pPr = para._element.find(f'{_W}pPr')
    if pPr is None:
        pPr = etree.SubElement(para._element, f'{_W}pPr')
        para._element.insert(0, pPr)
    old_numPr = pPr.find(f'{_W}numPr')
    if old_numPr is not None: pPr.remove(old_numPr)
    pPr.insert(0, etree.fromstring(f'<w:numPr xmlns:w="{_NS_W}"><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'))

    # Jangan tambah dobel kalau paragraf ini sudah pernah diberi prefix (idempoten).
    first_run = None
    for r in para._element.findall(f'{_W}r'):
        t = r.find(f'{_W}t')
        if t is not None and t.text:
            first_run = r; break
    if first_run is not None:
        t_el = first_run.find(f'{_W}t')
        if t_el is not None and t_el.text and re.match(r'^[A-Z]\.\d+(\.\d+)?\s{2,}', t_el.text):
            return

    font_name = 'Arial'; sz_val = '22'
    for run in para.runs:
        if run.text and run.text.strip():
            if run.font.name: font_name = run.font.name
            if run.font.size: sz_val = str(int(run.font.size.pt * 2))
            break
    esc = number_prefix.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    new_run = etree.fromstring(
        f'<w:r xmlns:w="{_NS_W}"><w:rPr><w:rFonts w:ascii="{font_name}" w:hAnsi="{font_name}"/>'
        f'<w:sz w:val="{sz_val}"/><w:szCs w:val="{sz_val}"/></w:rPr>'
        f'<w:t xml:space="preserve">{esc}</w:t></w:r>')
    pPr.addnext(new_run)


# ─────────────────────────────────────────────────────────────────────────────
# FITUR 2: EM DASH TO BULLETS
# ─────────────────────────────────────────────────────────────────────────────

def _get_or_create_emdash_numid(doc: Document) -> str:
    try:
        np = doc.part.numbering_part
        if np is None: return None
        nxml = np._element; target_abstract_id = None
        for ab in nxml.findall(f'{_W}abstractNum'):
            for lvl in ab.findall(f'{_W}lvl'):
                txt_el = lvl.find(f'{_W}lvlText')
                if txt_el is not None and txt_el.get(f'{_W}val') == _EM_DASH:
                    target_abstract_id = ab.get(f'{_W}abstractNumId'); break
            if target_abstract_id: break
        if target_abstract_id:
            existing_nums = nxml.findall(f'{_W}num')
            max_num_id = max((int(n.get(f'{_W}numId', 0)) for n in existing_nums), default=0)
            new_num_id = str(max_num_id + 1)
            nxml.append(etree.fromstring(f'<w:num xmlns:w="{_NS_W}" w:numId="{new_num_id}"><w:abstractNumId w:val="{target_abstract_id}"/></w:num>'))
            return new_num_id
        existing_abstracts = nxml.findall(f'{_W}abstractNum')
        max_abstract_id = max((int(a.get(f'{_W}abstractNumId', 0)) for a in existing_abstracts), default=0)
        new_abstract_id = str(max_abstract_id + 1)
        existing_nums = nxml.findall(f'{_W}num')
        max_num_id = max((int(n.get(f'{_W}numId', 0)) for n in existing_nums), default=0)
        new_num_id = str(max_num_id + 1)
        nxml.insert(0, etree.fromstring(f'<w:abstractNum xmlns:w="{_NS_W}" w:abstractNumId="{new_abstract_id}"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="{_EM_DASH}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>'))
        nxml.append(etree.fromstring(f'<w:num xmlns:w="{_NS_W}" w:numId="{new_num_id}"><w:abstractNumId w:val="{new_abstract_id}"/></w:num>'))
        return new_num_id
    except Exception as e: print(f"[Engine9] Error numbering: {e}"); return None

def _convert_emdash_to_bullets(doc: Document) -> None:
    num_id = _get_or_create_emdash_numid(doc)
    if not num_id: return
    for para in doc.paragraphs:
        sid = _get_para_style_id(para)
        if sid in _ANNEX_STYLE_IDS or sid.startswith('Heading'): continue
        text = para.text.strip()
        if text.startswith(_EM_DASH):
            for r in para.runs:
                if _EM_DASH in r.text: r.text = r.text.replace(_EM_DASH, "", 1).lstrip(); break
            pPr = para._element.find(f'{_W}pPr')
            if pPr is None: pPr = etree.SubElement(para._element, f'{_W}pPr')
            old_num = pPr.find(f'{_W}numPr')
            if old_num is not None: pPr.remove(old_num)
            pPr.insert(0, etree.fromstring(f'<w:numPr xmlns:w="{_NS_W}"><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>'))
            old_ind = pPr.find(f'{_W}ind')
            if old_ind is not None: pPr.remove(old_ind)
            pPr.insert(1, etree.fromstring(f'<w:ind xmlns:w="{_NS_W}" w:left="360" w:hanging="360"/>'))


# ─────────────────────────────────────────────────────────────────────────────
# FITUR 3: FIX NOTE / CATATAN
# ─────────────────────────────────────────────────────────────────────────────

def _fix_note_para(para, force_upper: bool | None = None) -> None:
    raw_text = para.text or ''
    full_text = raw_text.strip()
    if not full_text: return
    txt_lower = full_text.lower()
    if not (txt_lower.startswith('note') or txt_lower.startswith('catatan')): return
    original_label = re.match(r'^(NOTE|Note|note|CATATAN|Catatan|catatan)', full_text)
    use_upper = force_upper if force_upper is not None else bool(
        original_label and original_label.group(1).isupper()
    )
    if txt_lower.startswith('note'):
        m = re.match(r'^(NOTE|Note|note)', full_text)
        if m: full_text = ('CATATAN' if use_upper else 'Catatan') + full_text[len(m.group(1)):]
    elif txt_lower.startswith('catatan'):
        m = re.match(r'^(CATATAN|Catatan|catatan)', full_text)
        if m: full_text = ('CATATAN' if use_upper else 'Catatan') + full_text[len(m.group(1)):]
    full_text = re.sub(r'((?:CATATAN|Catatan)(?:\s+\d+)?\s+)(untuk\s+masuk|untuk\s+dimasukkan|untuk\s+diterapkan)', lambda mo: mo.group(1) + 'untuk entri', full_text)
    m_colon = re.match(r'^((?:NOTE|CATATAN|Catatan)[^:]*:)\s*', full_text, re.IGNORECASE)
    if m_colon: bold_part, normal_part = m_colon.group(1).rstrip(), full_text[m_colon.end():]
    else:
        words = full_text.split(None, 1); bold_part = words[0]; normal_part = words[1] if len(words) > 1 else ''
    if not normal_part.strip(): return
    # Pertahankan rPr setiap run kalimat (termasuk w:vertAlign untuk pangkat
    # dan indeks), tetapi paksa kalimat setelah label menjadi tidak tebal.
    # Posisi kalimat dihitung dari teks sebelum label NOTE/CATATAN dinormalkan.
    stripped_before = raw_text.strip()
    leading = len(raw_text) - len(raw_text.lstrip())
    if m_colon:
        normal_start = leading + m_colon.end()
    else:
        first_word = re.match(r'^\S+\s*', stripped_before)
        normal_start = leading + (first_word.end() if first_word else 0)

    source_segments = []
    cursor = 0
    label_rpr = None
    for run in list(para.runs):
        run_text = run.text or ''
        run_start, run_end = cursor, cursor + len(run_text)
        if label_rpr is None and run_text.strip() and run._element.rPr is not None:
            label_rpr = copy.deepcopy(run._element.rPr)
        cut = max(normal_start, run_start)
        if cut < run_end:
            piece = run_text[cut - run_start:]
            if piece:
                rpr = copy.deepcopy(run._element.rPr) if run._element.rPr is not None else None
                source_segments.append((piece, rpr))
        cursor = run_end

    # Bila perubahan label membuat pemetaan posisi tidak cocok, teks kalimat
    # tetap aman; format basis diambil dari run pertama yang tersedia.
    preserved_text = ''.join(piece for piece, _ in source_segments).lstrip()
    if preserved_text != normal_part.lstrip():
        base = source_segments[0][1] if source_segments else label_rpr
        source_segments = [(normal_part.lstrip(), base)]
    else:
        while source_segments and not source_segments[0][0].strip():
            source_segments.pop(0)
        if source_segments:
            source_segments[0] = (source_segments[0][0].lstrip(), source_segments[0][1])

    for run in list(para.runs):
        run._element.getparent().remove(run._element)

    run_b = para.add_run(bold_part)
    if label_rpr is not None:
        if run_b._element.rPr is not None:
            run_b._element.remove(run_b._element.rPr)
        run_b._element.insert(0, copy.deepcopy(label_rpr))
    run_b.bold = True

    for index, (piece, rpr) in enumerate(source_segments):
        if not piece:
            continue
        run_n = para.add_run(('    ' if index == 0 else '') + piece)
        if rpr is not None:
            if run_n._element.rPr is not None:
                run_n._element.remove(run_n._element.rPr)
            run_n._element.insert(0, copy.deepcopy(rpr))
        run_n.bold = False

def _fix_all_notes(doc: Document) -> None:
    for para in doc.paragraphs: _fix_note_para(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs: _fix_note_para(para)


_INITIAL_CLAUSE_HEADINGS = {
    'scope': 'Ruang lingkup',
    'normative references': 'Acuan normatif',
    'terms and definitions': 'Istilah dan definisi',
}


def _normalize_initial_clause_heading(para, source_text: str) -> None:
    """Pastikan kapitalisasi tiga judul pasal awal mengikuti sentence case."""
    source = re.sub(r'\s+', ' ', source_text or '').strip()
    match = re.match(
        r'^(\d+(?:\.\d+)*)\s+(Scope|Normative references|Terms and definitions)$',
        source, re.IGNORECASE,
    )
    if not match:
        return
    target = f"{match.group(1)}    {_INITIAL_CLAUSE_HEADINGS[match.group(2).casefold()]}"
    if para.runs:
        para.runs[0].text = target
        for run in para.runs[1:]:
            run.text = ''
    else:
        para.add_run(target)


_CLAUSE_PREFIX_RE = re.compile(
    r'^(?P<number>(?:\d+(?:\.\d+)*|[A-Z]\.(?:\d+)(?:\.\d+)*))(?P<gap>[ \t]+)(?=\S)'
)


def _replace_run_text_range(para, start: int, end: int, replacement: str) -> None:
    """Ganti rentang teks tanpa membangun ulang run atau mengubah rPr."""
    cursor = 0
    inserted = False
    for run in para.runs:
        value = run.text or ''
        run_start, run_end = cursor, cursor + len(value)
        if run_end <= start or run_start >= end:
            cursor = run_end
            continue
        left = value[:max(0, start - run_start)]
        right = value[max(0, end - run_start):] if end < run_end else ''
        middle = replacement if not inserted else ''
        run.text = left + middle + right
        inserted = True
        cursor = run_end


def _normalize_clause_heading_spacing(para) -> bool:
    """Pastikan nomor pasal/subpasal/subsubpasal diikuti tepat 4 spasi.

    Berlaku untuk heading isi utama (mis. 4, 4.1, 4.1.2) dan Lampiran
    (mis. A.1, A.1.1), sambil mempertahankan seluruh format run Engine 7.
    """
    text = para.text or ''
    match = _CLAUSE_PREFIX_RE.match(text)
    if not match:
        return False
    style_id = _get_para_style_id(para)
    style_name = para.style.name if para.style is not None else ''
    style_key = f'{style_id} {style_name}'.casefold()
    looks_like_clause_style = (
        style_id in _HEADING_STYLES_WITH_NUM
        or any(token in style_key for token in ('pasal', 'clause', 'heading'))
        or style_id.casefold() in {'a2', 'a3'}
    )
    if not looks_like_clause_style:
        return False
    if match.group('gap') != '    ':
        _replace_run_text_range(para, match.start('gap'), match.end('gap'), '    ')
    return True


def _normalize_all_clause_heading_spacing(doc: Document) -> None:
    for para in doc.paragraphs:
        _normalize_clause_heading_spacing(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _normalize_clause_heading_spacing(para)


def _mark_untranslated_paragraph_red(para) -> None:
    """Warnai teks sumber yang gagal diterjemahkan tanpa merusak format run.

    Operasi dilakukan langsung pada ``w:rPr`` sehingga bold, italic,
    superscript, subscript, ukuran font, dan struktur tabel tetap utuh.
    """
    for run_element in para._element.iter(qn('w:r')):
        # Run tanpa teks (misalnya hanya drawing/field) tidak perlu diwarnai.
        if not any((node.text or '') for node in run_element.iter(qn('w:t'))):
            continue
        run_properties = run_element.find(qn('w:rPr'))
        if run_properties is None:
            run_properties = OxmlElement('w:rPr')
            run_element.insert(0, run_properties)
        for old_color in list(run_properties.findall(qn('w:color'))):
            run_properties.remove(old_color)
        color = OxmlElement('w:color')
        color.set(qn('w:val'), 'FF0000')
        run_properties.append(color)


# ─────────────────────────────────────────────────────────────────────────────
# BIBLIOGRAFI & AUTONUMBERING
# ─────────────────────────────────────────────────────────────────────────────

def _el_text(el) -> str: return ''.join(t.text or '' for t in el.findall(f'.//{_W}t')).strip()
def _is_bibliography_el(el) -> bool:
    if el.tag != f'{_W}p': return False
    sid = _get_style_id(el).lower()
    if sid in _BIBLIO_TITLE_STYLES: return True
    text = _el_text(el).lower().strip()
    if not text or text[0].isdigit(): return False
    return text in _BIBLIO_KEYWORDS_EXACT


def _notify(cb, pct: int, msg: str) -> None:
    if cb:
        try: cb(pct, msg)
        except Exception: pass


# ─────────────────────────────────────────────────────────────────────────────
# PROTEKSI: DETEKSI HASIL "TERJEMAHAN" YANG SEBENARNYA HALAMAN ERROR
# ─────────────────────────────────────────────────────────────────────────────
# deep_translator (Google Translate gratis) kadang TIDAK melempar exception
# saat diblokir/limit rate — ia malah mengembalikan teks halaman error mentah
# (mis. "500. That's an error. ... That's all we know.") seolah itu hasil
# terjemahan valid. Tanpa validasi, teks sampah ini langsung menimpa isi
# dokumen. Fungsi ini mendeteksi pola tersebut agar hasil semacam itu
# DIBUANG dan teks asli dipertahankan, bukan diganti dengan sampah.
_RE_ERROR_RESPONSE = re.compile(
    r"that'?s an error|that'?s all we know|server error|"
    r"unusual traffic|please try again later|"
    r"<html|<!doctype|\b50[0-9]\b.{0,15}error|\berror\b.{0,15}\b50[0-9]\b",
    re.IGNORECASE
)

def _looks_like_error_response(original: str, translated: str) -> bool:
    if not translated:
        return False
    if _RE_ERROR_RESPONSE.search(translated):
        return True
    # Hasil "terjemahan" yang tiba-tiba jauh lebih panjang dari teks aslinya
    # (mis. paragraf 3 kata jadi ratusan karakter) juga mencurigakan —
    # kemungkinan besar itu bukan terjemahan, tapi halaman/pesan lain.
    if len(original) <= 20 and len(translated) > 150:
        return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# TRANSLATOR WRAPPER
# ─────────────────────────────────────────────────────────────────────────────

_RE_PROTECTION_TOKEN = re.compile(
    r'ZXQ(?:TK|IT|SRC)[A-F0-9]{12}QXZ', re.IGNORECASE
)

_TITLE_SAFE_FALLBACKS = {
    'railway applications — fire protection on railway vehicles — part 1: general':
        'Aplikasi perkeretaapian — Perlindungan kebakaran pada sarana '
        'perkeretaapian — Bagian 1: Umum',
}


class TranslationFailedError(RuntimeError):
    """Mencegah dokumen setengah diterjemahkan tetap disimpan."""


class _Translator:
    def __init__(self, source: str = 'auto', target: str = 'id', 
                 custom_dict: CustomDictionary | None = None,
                 italic_dict: ItalicDictionary | None = None):
        try: from deep_translator import GoogleTranslator
        except ImportError: raise ImportError("Jalankan: pip install deep-translator")
        self._cls = GoogleTranslator
        self.source = source; self.target = target
        self.custom_dict = custom_dict; self.italic_dict = italic_dict
        self._client = self._cls(source=self.source, target=self.target)
        self.failed_texts: list[str] = []

    def translate_one(self, text: str, italic_map: dict = None) -> tuple[str, list[str]]:
        t = text.strip()
        if not t or _skip_text(t): return text, []
        token_map = {}
        if self.custom_dict and len(self.custom_dict) > 0:
            t, token_map = self.custom_dict._apply_pre(t)
        final_italic_map = italic_map or {}

        # Jika seluruh teks sudah dicakup Kamus SNI, hasil kamus adalah hasil
        # final. Ini berlaku untuk satu token ("Introduction") maupun beberapa
        # token yang hanya dipisahkan tanda baca, angka bagian, atau em dash
        # (judul Cover). Rangkaian token semacam itu tidak memiliki bahasa
        # alami untuk diterjemahkan dan sering ditolak Google Translate.
        uncovered = _RE_PROTECTION_TOKEN.sub('', t)
        uncovered_words = re.findall(r'[A-Za-zÀ-ÿ]{2,}', uncovered)
        if token_map and not uncovered_words:
            result = self.custom_dict._apply_post(t, token_map)
            italic_terms_found = []
            if final_italic_map:
                result, italic_terms_found = self.italic_dict._apply_post(
                    result, final_italic_map
                )
            return result, italic_terms_found

        expected_tokens = {
            token.casefold() for token in _RE_PROTECTION_TOKEN.findall(t)
        }
        result = None
        last_error = None
        for attempt in range(1, 5):
            try:
                candidate = self._client.translate(t)
                if not candidate or _looks_like_error_response(t, candidate):
                    raise ValueError('respons layanan terjemahan tidak valid')
                returned_tokens = {
                    token.casefold()
                    for token in _RE_PROTECTION_TOKEN.findall(candidate)
                }
                if returned_tokens != expected_tokens:
                    raise ValueError('token kamus/format berubah atau hilang')

                source_plain = _RE_PROTECTION_TOKEN.sub('', t)
                result_plain = _RE_PROTECTION_TOKEN.sub('', candidate)
                if (
                    re.search(r'[A-Za-z]{3}', source_plain)
                    and re.sub(r'\s+', ' ', source_plain).strip().casefold()
                    == re.sub(r'\s+', ' ', result_plain).strip().casefold()
                ):
                    raise ValueError('teks dikembalikan tanpa diterjemahkan')
                result = candidate
                break
            except Exception as exc:
                last_error = exc
                if attempt < 4:
                    time.sleep(0.8 * attempt)

        if result is None:
            # Judul standar lazim terdiri dari beberapa klausa yang dipisahkan
            # em dash. Jika layanan menolak judul panjang sebagai satu request,
            # coba setiap klausa secara mandiri lalu gabungkan kembali dengan
            # tanda pisah asli. Ini juga membuat bagian yang sudah dicakup
            # kamus langsung diselesaikan lokal per klausa.
            title_parts = re.split(r'(\s*[—–]\s*)', t)
            if len(title_parts) > 1:
                translated_parts = []
                try:
                    for part in title_parts:
                        if not part:
                            continue
                        if re.fullmatch(r'\s*[—–]\s*', part):
                            translated_parts.append(part)
                        else:
                            part_result, _ = self.translate_one(part, {})
                            translated_parts.append(part_result)
                    result = ''.join(translated_parts)
                except TranslationFailedError as split_error:
                    last_error = split_error

        if result is None:
            # Provider cadangan untuk kondisi Google Translate sedang menolak
            # request/rate-limited. Deep-translator menyertakan MyMemory;
            # kegagalannya tetap ditangani tanpa merusak dokumen.
            try:
                from deep_translator import MyMemoryTranslator
                fallback_source = (
                    'english' if self.source in ('auto', 'en') else self.source
                )
                fallback_target = (
                    'indonesian' if self.target == 'id' else self.target
                )
                fallback = MyMemoryTranslator(
                    source=fallback_source, target=fallback_target
                ).translate(t)
                if fallback and not _looks_like_error_response(t, fallback):
                    fallback_tokens = {
                        token.casefold()
                        for token in _RE_PROTECTION_TOKEN.findall(fallback)
                    }
                    if fallback_tokens == expected_tokens:
                        result = fallback
            except Exception as fallback_error:
                last_error = fallback_error

        if result is None:
            safe_title = _TITLE_SAFE_FALLBACKS.get(
                re.sub(r'\s+', ' ', text).strip().casefold()
            )
            if safe_title:
                result = safe_title

        if result is None:
            preview = re.sub(r'\s+', ' ', text).strip()[:100]
            # Jangan gagalkan seluruh pipeline karena satu request eksternal.
            # Simpan teks sumber (token kamus tetap dipulihkan di bawah) dan
            # laporkan sebagai peringatan pada ringkasan proses.
            self.failed_texts.append(preview)
            result = t
        if token_map: result = self.custom_dict._apply_post(result, token_map)
        italic_terms_found = []
        if final_italic_map:
            _idict = self.italic_dict if self.italic_dict is not None else ItalicDictionary()
            result, italic_terms_found = _idict._apply_post(result, final_italic_map)
        return result, italic_terms_found

def _match_capitalization(original: str, translated: str) -> str:
    orig = original.strip(); tran = translated.strip()
    if not orig or not tran: return translated
    letters = [c for c in orig if c.isalpha()]
    if not letters: return translated
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    if upper_ratio >= 0.8: return tran.upper()
    if orig[0].isupper(): return tran[0].upper() + tran[1:] if len(tran) > 1 else tran.upper()
    return tran


def _translate_para(para, tr, past_bibliography: bool = False) -> list[str]:
    """
    Terjemahkan paragraf biasa. Paragraf yang memiliki hyperlink sengaja
    dipertahankan utuh supaya teks, urutan run, relationship, dan URL asli
    tidak terhapus atau berubah.
    """
    if _skip_paragraph(para, past_bibliography): return []

    if _has_hyperlinks(para):
        return []
    
    # 2. Proses TEKS NORMAL (bukan hyperlink)
    text_runs = [(i, r) for i, r in enumerate(para.runs) if r.text and r.text.strip()]
    
    if not text_runs: return []
    
    combined_raw = ''.join(r.text for _, r in text_runs)
    combined = combined_raw.strip()
    if _skip_text(combined): return []
    
    # Get font
    font_name = 'Arial'; font_size = None
    for run in para.runs:
        if run.text.strip():
            if run.font.name: font_name = run.font.name
            if run.font.size: font_size = run.font.size.pt if run.font.size else None
            break
    
    # Teks asli (SEBELUM proteksi token apa pun) — dipakai untuk mencocokkan
    # kapitalisasi hasil terjemahan, agar token proteksi (huruf besar semua)
    # tidak ikut dihitung dan salah membuat seluruh paragraf jadi HURUF BESAR.
    original_for_case = combined
    source_note_upper = bool(re.match(r'^NOTE(?:\s|\t|:)', combined))

    # Proteksi 1: istilah/judul asing yang SUDAH miring di dokumen sumber
    # (mis. judul standar acuan pada klausul "Acuan normatif"), SERTA
    # seluruh teks berformat superscript/subscript (pangkat/indeks) —
    # keduanya tidak diterjemahkan dan formatnya dikembalikan persis
    # seperti sumber setelah proses terjemahan selesai.
    para_style_italic = _get_para_style_italic(para)
    combined_with_src_tokens, format_token_map = _extract_source_format_map(text_runs, para_style_italic)
    combined = combined_with_src_tokens.strip()

    # Proteksi 2: kamus istilah asing dari spreadsheet ("Kamus Istilah Asing")
    if tr.italic_dict and len(tr.italic_dict) > 0:
        combined, dict_italic_map = tr.italic_dict._apply_pre(combined)
    else:
        dict_italic_map = {}

    # Terjemahkan (token @@SRC_...@@ dari proteksi superscript/subscript
    # ikut terkirim dan diharapkan lolos utuh dari mesin terjemahan, sama
    # seperti token proteksi miring lain yang sudah terbukti aman).
    translated, italic_terms_found = tr.translate_one(combined, dict_italic_map)
    time.sleep(_TRANSLATE_DELAY)
    if not translated or translated == combined: return []
    translated = _match_capitalization(original_for_case, translated)

    # Kembalikan token superscript/subscript/miring-sumber menjadi teks asli,
    # sekaligus dapatkan posisi presisi untuk membangun ulang run.
    translated, format_spans = _detokenize_with_formatting(translated, format_token_map)

    # Apply formatting ke teks normal
    if italic_terms_found or format_spans:
        _apply_mixed_formatting_to_para(para, translated, italic_terms_found, font_name, font_size,
                                         format_spans=format_spans)
    else:
        if para.runs:
            para.runs[0].text = translated
            for r in para.runs[1:]: r.text = ''
        else:
            para.add_run(translated)

    _normalize_initial_clause_heading(para, original_for_case)
    if source_note_upper or re.match(
        r'^(?:NOTE|Note|note|CATATAN|Catatan|catatan)(?:\s|\t|:)',
        para.text.strip(),
    ):
        _fix_note_para(para, force_upper=source_note_upper)
    
    return italic_terms_found


def _translate_table(table, tr) -> None:
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs: _translate_para(para, tr)


def _is_translation_candidate(para) -> bool:
    """Seleksi lokal murni; tidak pernah memanggil layanan terjemahan."""
    if _skip_paragraph(para):
        return False
    if _has_hyperlinks(para):
        return False
    text_runs = [r for r in para.runs if r.text and r.text.strip()]
    if not text_runs:
        return False
    combined = ''.join(r.text for r in text_runs).strip()
    return bool(combined) and not _skip_text(combined)


def _build_translation_queue(doc: Document, para_targets: set,
                             table_targets: set) -> tuple[list, int, int]:
    """Hitung seluruh skip lebih dahulu dan hasilkan antrean riil terjemahan.

    Pemindaian hanya membaca XML DOCX di memori, sehingga tidak menunggu API,
    tidak melakukan sleep, dan biasanya selesai jauh di bawah 10 detik.
    Nilai total antrean inilah yang dipakai sebagai penyebut ``xx/yyy``.
    """
    queue = []
    seen = set()
    inspected = skipped = 0

    body = doc.element.body
    para_map = {p._element: p for p in doc.paragraphs}
    table_map = {t._element: t for t in doc.tables}

    for child in body:
        if child in para_map:
            inspected += 1
            para = para_map[child]
            if para._element in para_targets and _is_translation_candidate(para):
                if para._element not in seen:
                    queue.append(para)
                    seen.add(para._element)
            else:
                skipped += 1
        elif child in table_map:
            table = table_map[child]
            if table._element not in table_targets:
                # Tabel di zona skip dihitung sebagai satu unit struktur.
                inspected += 1
                skipped += 1
                continue
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if para._element in seen:
                            continue
                        seen.add(para._element)
                        inspected += 1
                        if _is_translation_candidate(para):
                            queue.append(para)
                        else:
                            skipped += 1

    return queue, inspected, skipped

def _translate_hf(hf_part, tr) -> None:
    if hf_part is None: return
    try:
        for para in hf_part.paragraphs:
            if _RE_COPYRIGHT.search(para.text or ''): continue
            _translate_para(para, tr)
        for table in hf_part.tables: _translate_table(table, tr)
    except Exception: pass


# ─────────────────────────────────────────────────────────────────────────────
# SINKRONISASI JUDUL
# ─────────────────────────────────────────────────────────────────────────────

def _extract_cover_titles(doc: Document) -> tuple[str, str]:
    id_title = ""; en_title = ""
    for para in doc.paragraphs:
        if _has_inline_sectpr(para): break
        text = para.text.strip()
        if not text: continue
        is_bold = False; is_italic = False; max_size = 0
        for run in para.runs:
            if run.text.strip():
                if run.font.bold: is_bold = True
                if run.font.italic: is_italic = True
                sz = run.font.size
                if sz and sz.pt > max_size: max_size = sz.pt
        if not id_title:
            if max_size >= 16 and is_bold and not is_italic: id_title = text; continue
        if id_title and not en_title:
            if max_size >= 14 and is_italic: en_title = text
    return id_title, en_title

def _sync_body_title(doc: Document, cover_id: str) -> bool:
    if not cover_id: return False
    _BODY_TITLE_STYLES = {'main title 2', 'main title2', 'maintitle2', 'boxedtitle', 'boxed title', 'body title'}
    _RE_H1 = re.compile(r'^\d+\s+\S')
    paras = doc.paragraphs; sect_idx = -1; h1_idx = -1
    for i, p in enumerate(paras):
        if sect_idx == -1 and _has_inline_sectpr(p): sect_idx = i
        txt = p.text.strip()
        style_name = (p.style.name or '').lower() if p.style is not None else ''
        if txt and _RE_H1.match(txt) and 'heading' in style_name: h1_idx = i; break
    start = sect_idx + 1 if sect_idx != -1 else 0
    end = h1_idx if h1_idx != -1 else start + 20
    if end > len(paras): end = len(paras)
    for i in range(start, end):
        style_name = (
            (paras[i].style.name or '').lower().strip()
            if paras[i].style is not None else ''
        )
        if style_name in _BODY_TITLE_STYLES and paras[i].text.strip():
            _replace_para_text(paras[i], cover_id); return True
    for i in range(start, end):
        txt = paras[i].text.strip()
        if txt and any(r.bold for r in paras[i].runs if r.text.strip()) and paras[i].alignment == WD_ALIGN_PARAGRAPH.CENTER:
            _replace_para_text(paras[i], cover_id); return True
    return False

def _replace_para_text(para, new_text: str) -> None:
    fn = 'Arial'; fs = None
    for r in para.runs:
        if r.text.strip():
            if r.font.name: fn = r.font.name
            if r.font.size: fs = r.font.size
            break
    for r in list(para.runs): r._element.getparent().remove(r._element)
    nr = para.add_run(new_text); nr.bold = True; nr.font.name = fn
    if fs: nr.font.size = fs
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER

def _sync_foreword_title(doc: Document, cover_id: str, cover_en: str, tr=None) -> None:
    """Ganti hanya kelompok italic pertama tanpa membangun ulang paragraf.

    Dengan cara ini seluruh run formatting hasil Engine 7 tetap utuh. Kelompok
    italic kedua (judul asli setelah referensi ISO) sengaja tidak diubah.
    """
    if not cover_id:
        return
    translated_title = cover_id
    if tr is not None and cover_en and cover_id.casefold() == cover_en.casefold():
        candidate, _ = tr.translate_one(cover_en)
        if candidate:
            translated_title = candidate
    for para in doc.paragraphs:
        text = re.sub(r'\s+', ' ', para.text or '').strip()
        if not re.match(r'^SNI\s+[^,]+,', text, re.IGNORECASE):
            continue
        in_first_group = False
        first_run = None
        for run in para.runs:
            has_text = bool(run.text.strip())
            is_italic = _run_effective_italic(
                run, _get_para_style_italic(para)
            )
            if has_text and is_italic and first_run is None:
                first_run = run
                first_run.text = translated_title
                in_first_group = True
                continue
            if in_first_group:
                if is_italic:
                    run.text = ''
                elif has_text:
                    return
        if first_run is not None:
            return


def _translation_targets(doc: Document) -> tuple[set, set]:
    """Tentukan paragraf dan tabel yang boleh diterjemahkan oleh Engine 8.

    Zona mengikuti kontrak keluaran Engine 5--7:
    - Cover: hanya judul Indonesia paling atas (bold, >= 16 pt, non-italic).
    - Section 3: heading dan seluruh isi Introduction jika bagian itu tersedia.
      Standar yang langsung dimulai dari Content (mis. CISPR 32) tetap valid.
    - Area Content asli: seluruh elemen sebelum bookmark duplikasi Engine 5,
      kecuali paragraf style ``RefNorm``.
    - Bibliography: hanya heading-nya.
    - Copyright, Daftar isi, isi Prakata, salinan Engine 5, entri Bibliography,
      dan section informasi perumus tidak menjadi target.
    """
    para_targets: set = set()
    table_targets: set = set()
    body = doc.element.body
    para_map = {p._element: p for p in doc.paragraphs}
    table_map = {t._element: t for t in doc.tables}

    section_no = 1
    in_introduction = False
    in_original_content = False
    duplicate_started = False
    cover_title_found = False

    for child in body:
        if child in para_map:
            para = para_map[child]
            text = re.sub(r'\s+', ' ', para.text or '').strip()
            folded = text.casefold()
            style_id = _get_para_style_id(para).casefold()
            bookmarks = {
                mark.get(qn('w:name'), '')
                for mark in para._element.xpath('.//w:bookmarkStart')
            }

            if 'Engine5DuplicateStart' in bookmarks:
                duplicate_started = True
                in_original_content = False

            if section_no == 1 and not cover_title_found and text:
                is_bold = any(r.bold for r in para.runs if r.text.strip())
                is_italic = any(r.italic for r in para.runs if r.text.strip())
                max_size = max(
                    (r.font.size.pt for r in para.runs
                     if r.text.strip() and r.font.size),
                    default=0,
                )
                if is_bold and not is_italic and max_size >= 16:
                    para_targets.add(para._element)
                    cover_title_found = True

            elif section_no == 3:
                if folded == 'introduction':
                    in_introduction = True
                if in_introduction:
                    para_targets.add(para._element)

            elif section_no >= 4 and not duplicate_started:
                # Semua section layout Content asli tetap masuk sampai marker
                # Engine 5. Section informasi perumus selalu berada sesudah
                # marker dan karena itu otomatis tidak pernah masuk target.
                in_original_content = True
                if style_id != 'refnorm':
                    para_targets.add(para._element)

            if duplicate_started and folded in _BIBLIO_KEYWORDS_EXACT:
                para_targets.add(para._element)

            if _has_inline_sectpr(para):
                section_no += 1
                if section_no >= 4:
                    in_introduction = False

        elif child in table_map:
            if (section_no == 3 and in_introduction) or (
                section_no >= 4 and in_original_content and not duplicate_started
            ):
                table_targets.add(table_map[child]._element)

    if not cover_title_found:
        raise ValueError('Judul atas Cover tidak ditemukan.')
    # Introduction memang opsional. Engine 1--4 secara sah menghasilkan front
    # matter tanpa Introduction untuk standar yang langsung dimulai dari pasal
    # Content, misalnya CISPR 32. Dalam kasus itu target section 3 cukup kosong
    # dan penerjemahan tetap dilanjutkan ke area Content pada section 4+.
    if not duplicate_started:
        raise ValueError(
            'Bookmark Engine5DuplicateStart tidak ditemukan; output Engine 5 '
            'diperlukan agar konten duplikat dapat di-skip dengan aman.'
        )
    return para_targets, table_targets


# ─────────────────────────────────────────────────────────────────────────────
# MAIN ENGINE (FIXED)
# ─────────────────────────────────────────────────────────────────────────────

class DocxFinalTranslatorEngine:
    """
    Engine gabungan dengan dua spreadsheet terpisah.
    Hyperlink dipertahankan utuh dan konten asli tidak disisipkan ulang.
    """
    
    def __init__(self, source_lang: str = 'auto', target_lang: str = 'id', 
                 custom_dict: CustomDictionary | None = None,
                 italic_dict: ItalicDictionary | None = None):
        self.source_lang = source_lang
        self.target_lang = target_lang
        self.custom_dict = custom_dict
        self.italic_dict = italic_dict
        self._custom_dict_provided = custom_dict is not None
        self._italic_dict_provided = italic_dict is not None

    def set_dictionary(self, d: CustomDictionary) -> None:
        self.custom_dict = d
        self._custom_dict_provided = True
    def set_italic_dictionary(self, d: ItalicDictionary) -> None:
        self.italic_dict = d
        self._italic_dict_provided = True
    def get_dictionary(self) -> CustomDictionary:
        if self.custom_dict is None: self.custom_dict = CustomDictionary()
        return self.custom_dict
    def get_italic_dictionary(self) -> ItalicDictionary:
        if self.italic_dict is None: self.italic_dict = ItalicDictionary()
        return self.italic_dict

    def translate(self, input_docx: str, output_docx: str, progress_callback=None, 
                  translate_headers: bool = False) -> tuple[bool, str]:
        try:
            # app.py membuat engine tanpa parameter kamus. Karena itu Engine 8
            # wajib mengambil kedua spreadsheet sendiri pada setiap proses,
            # sehingga perubahan Google Sheet langsung dipakai dan bukan hanya
            # ditampilkan sebagai angka pada dashboard.
            # Kamus online penting untuk konsistensi istilah, tetapi gangguan
            # DNS/Google Sheet sesaat tidak boleh menggagalkan seluruh Engine 8.
            # Muat ke objek kandidat terlebih dahulu. Kamus aktif baru diganti
            # setelah unduhan sukses; jika gagal, pakai cache proses sebelumnya.
            # Pada cold-start tanpa cache, gunakan minimal safety fallback untuk
            # istilah yang memang diwajibkan oleh pipeline.
            dictionary_warnings = []

            if not self._custom_dict_provided:
                previous_custom = self.custom_dict
                candidate_custom = CustomDictionary()
                try:
                    sni_count = candidate_custom.load_from_google_sheet(
                        KAMUS_SPREADSHEET_URL
                    )
                    self.custom_dict = candidate_custom
                except Exception as exc:
                    if previous_custom is not None and len(previous_custom) > 0:
                        self.custom_dict = previous_custom
                        sni_count = len(self.custom_dict)
                        dictionary_warnings.append(
                            'Google Sheet Kamus SNI tidak dapat diakses; '
                            'menggunakan cache kamus terakhir.'
                        )
                    else:
                        self.custom_dict = CustomDictionary()
                        self.custom_dict.add_term('Introduction', 'Pendahuluan')
                        sni_count = len(self.custom_dict)
                        dictionary_warnings.append(
                            'Google Sheet Kamus SNI tidak dapat diakses dan cache '
                            'belum tersedia; memakai kamus minimum bawaan.'
                        )
            else:
                if self.custom_dict is None:
                    self.custom_dict = CustomDictionary()
                sni_count = len(self.custom_dict)

            if not self._italic_dict_provided:
                previous_italic = self.italic_dict
                candidate_italic = ItalicDictionary()
                try:
                    italic_count_loaded = candidate_italic.load_from_google_sheet(
                        ITALIC_SPREADSHEET_URL
                    )
                    self.italic_dict = candidate_italic
                except Exception as exc:
                    if previous_italic is not None and len(previous_italic) > 0:
                        self.italic_dict = previous_italic
                        italic_count_loaded = len(self.italic_dict)
                        dictionary_warnings.append(
                            'Google Sheet istilah asing tidak dapat diakses; '
                            'menggunakan cache kamus terakhir.'
                        )
                    else:
                        self.italic_dict = ItalicDictionary()
                        self.italic_dict.add_term('ISO Online browsing platform')
                        self.italic_dict.add_term('IEC Electropedia')
                        italic_count_loaded = len(self.italic_dict)
                        dictionary_warnings.append(
                            'Google Sheet istilah asing tidak dapat diakses dan '
                            'cache belum tersedia; memakai daftar minimum bawaan.'
                        )
            else:
                if self.italic_dict is None:
                    self.italic_dict = ItalicDictionary()
                italic_count_loaded = len(self.italic_dict)

            # Dua entri ini adalah invariant pipeline. Jika spreadsheet diedit
            # keliru, pulihkan secara lokal alih-alih menghentikan dokumen.
            intro_entry = self.custom_dict._entries.get('introduction')
            if not intro_entry or intro_entry[1].strip().casefold() != 'pendahuluan':
                self.custom_dict.add_term('Introduction', 'Pendahuluan')
                sni_count = len(self.custom_dict)
                dictionary_warnings.append(
                    'Entri wajib Introduction → Pendahuluan dipulihkan otomatis.'
                )

            for required_term in (
                'ISO Online browsing platform',
                'IEC Electropedia',
            ):
                if required_term.casefold() not in self.italic_dict._entries:
                    self.italic_dict.add_term(required_term)
                    dictionary_warnings.append(
                        f'Istilah asing wajib "{required_term}" dipulihkan otomatis.'
                    )
            italic_count_loaded = len(self.italic_dict)

            info = f"{len(self.custom_dict) if self.custom_dict else 0} kamus"
            if self.italic_dict: info += f", {len(self.italic_dict)} miring"
            status_message = (
                f"Kamus siap: SNI={sni_count}, istilah asing={italic_count_loaded}"
            )
            if dictionary_warnings:
                status_message += ' | ' + ' '.join(dictionary_warnings)
            _notify(progress_callback, 2, status_message)
            
            _notify(progress_callback, 5, "Init translator...")
            # Satu instance translator tidak dibagi lintas thread. Setiap worker
            # memiliki client sendiri agar aman dan benar-benar berjalan paralel.
            tr = _Translator(self.source_lang, self.target_lang, self.custom_dict, self.italic_dict)
            doc = Document(input_docx)

            scan_started = time.perf_counter()
            para_targets, table_targets = _translation_targets(doc)
            translation_queue, inspected_count, skipped_count = (
                _build_translation_queue(doc, para_targets, table_targets)
            )
            scan_seconds = time.perf_counter() - scan_started
            if scan_seconds >= 10:
                raise RuntimeError(
                    f'Pra-pemindaian skip memerlukan {scan_seconds:.2f} detik; '
                    'batas maksimum adalah 10 detik.'
                )

            total = len(translation_queue)
            _notify(
                progress_callback, 5,
                f"[pra-scan] skip={skipped_count} dari {inspected_count} "
                f"({scan_seconds:.2f} detik) | antrean terjemahan=0/{total} "
                f"| [progres-total] 0/{total}",
            )

            done = translated_count = 0
            italic_count = 0
            failed_paras = []

            def _translate_job(index_para):
                index, para = index_para
                worker_tr = _Translator(
                    self.source_lang, self.target_lang,
                    self.custom_dict, self.italic_dict,
                )
                found = _translate_para(para, worker_tr)
                return index, para, found, bool(worker_tr.failed_texts)

            # Tahap utama: tepat 4 worker translate.
            with ThreadPoolExecutor(max_workers=4, thread_name_prefix='translate') as pool:
                futures = [pool.submit(_translate_job, item)
                           for item in enumerate(translation_queue)]
                for future in as_completed(futures):
                    index, para, found, failed = future.result()
                    done += 1
                    italic_count += len(found)
                    if failed:
                        failed_paras.append((index, para))
                    else:
                        translated_count += 1
                    # Translate normal memakai rentang 10%--90%, dihitung
                    # murni dari counter done/total (XXX/XXX).
                    pct = 10 + int(done / max(total, 1) * 80)
                    _notify(
                        progress_callback, pct,
                        f"[translate 4 worker] {done}/{total} | "
                        f"berhasil={translated_count} | "
                        f"tersisa={len(failed_paras)} | "
                        f"[progres-total] {done}/{total + len(failed_paras)}",
                    )

            # Hanya bagian yang belum berhasil diterjemahkan yang diulang,
            # secara berurutan dengan tepat 1 worker.
            recovery_total = len(failed_paras)
            still_failed = []
            recovery_success = 0
            recovery_tr = _Translator(
                self.source_lang, self.target_lang,
                self.custom_dict, self.italic_dict,
            )
            for recovery_done, (index, para) in enumerate(
                    sorted(failed_paras, key=lambda item: item[0]), start=1):
                before = len(recovery_tr.failed_texts)
                found = _translate_para(para, recovery_tr)
                failed = len(recovery_tr.failed_texts) > before
                if failed:
                    still_failed.append(para)
                else:
                    recovery_success += 1
                    translated_count += 1
                    italic_count += len(found)
                # Pemulihan memakai rentang 90%--98%, dihitung murni dari
                # counter recovery_done/recovery_total (YY/YY).
                pct = 90 + int(recovery_done / max(recovery_total, 1) * 8)
                _notify(
                    progress_callback, pct,
                    f"[pemulihan 1 worker] {recovery_done}/{recovery_total} | "
                    f"berhasil={recovery_success} | "
                    f"tersisa={recovery_total - recovery_done + len(still_failed)} "
                    f"| [progres-total] {total + recovery_done}/"
                    f"{total + recovery_total}",
                )

            # Dipakai ringkasan dan UI peringatan. Dokumen parsial tetap
            # disimpan agar dapat di-download atau dipaksa lanjut ke Engine 9.
            tr.failed_texts = [
                re.sub(r'\s+', ' ', para.text or '').strip()[:100]
                for para in still_failed
            ]
            for para in still_failed:
                _mark_untranslated_paragraph_red(para)

            # Tidak menjalankan formatting global di sini. Dengan demikian
            # Copyright, Daftar isi, salinan Engine 5, Bibliography entries,
            # dan informasi perumus benar-benar tidak tersentuh Engine 8.

            _notify(progress_callback, 98, "Sinkronisasi judul...")
            fid, fen = _extract_cover_titles(doc)
            if fid:
                _sync_foreword_title(doc, fid, fen, tr=tr)

            # Normalisasi terakhir untuk seluruh heading pasal pada konten dan
            # Lampiran. Operasi hanya mengganti karakter pemisah di run yang
            # sudah ada sehingga format Engine 7 tetap dipertahankan.
            _normalize_all_clause_heading_spacing(doc)

            _notify(progress_callback, 98, "Saving...")
            doc.save(output_docx)

            if still_failed:
                return False, (
                    f"{len(still_failed)} bagian masih belum berhasil "
                    "diterjemahkan setelah pemulihan 1 worker. Dokumen parsial "
                    "Engine 8 sudah disimpan; teks sumber yang gagal diberi "
                    "font merah dan dapat di-download atau "
                    "dipaksa lanjut ke Engine 9."
                )
            
            summary = f"✅ Done!"
            if italic_count > 0: summary += f" Miring: {italic_count}."
            summary += (
                f" Paragraf diterjemahkan: {translated_count}; "
                f"unit diperiksa: {inspected_count}; "
                f"zona/unit di-skip pra-scan: {skipped_count}."
            )
            if tr.failed_texts:
                summary += (
                    f" Peringatan: {len(tr.failed_texts)} bagian dipertahankan "
                    "karena seluruh layanan terjemahan sedang tidak merespons."
                )
            _notify(progress_callback, 100, summary)
            return True, output_docx

        except ImportError as e: return False, f"Dependensi: {e}"
        except Exception as e: return False, f"Engine8 Error: {str(e)}\n{traceback.format_exc()}"


class SelectiveTranslationEngine(DocxFinalTranslatorEngine):
    """Adaptor API Engine 8 yang digunakan oleh ``app.py``.

    Implementasi penerjemahan tetap berasal dari
    :class:`DocxFinalTranslatorEngine`. Kelas ini menyediakan nama kelas dan
    metode ``process()`` yang diharapkan pipeline aplikasi, termasuk bentuk
    nilai balik ``(berhasil, path_output, pesan)``.
    """

    def process(
        self,
        input_docx: str,
        output_docx: str,
        progress_callback=None,
        translate_headers: bool = False,
    ) -> tuple[bool, str, str]:
        success, result = self.translate(
            input_docx=input_docx,
            output_docx=output_docx,
            progress_callback=progress_callback,
            translate_headers=translate_headers,
        )
        if success:
            return True, result, "Penerjemahan selektif selesai."
        return False, output_docx, result
