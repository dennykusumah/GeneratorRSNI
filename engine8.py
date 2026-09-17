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

from pipeline_utils import validate_docx, atomic_save_docx
import re
import copy
import contextlib
import time
import uuid
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
import csv
import os
import json
import hashlib
import sqlite3
import tempfile
import threading
import random
from collections import OrderedDict, deque

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

ENGINE8_POLICY_VERSION = "2026.09-cache-v2-quality-gate"
ENGINE8_PROVIDER_VERSION = "deep-translator-google+mymemory-v1"
GEMINI_RECOVERY_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
GEMINI_RECOVERY_VERSION = "gemini-recovery-v3-dualkey-optimized"
GEMINI_RECOVERY_MAX_ITEMS = int(os.getenv("GEMINI_RECOVERY_MAX_ITEMS", "8"))
GEMINI_RECOVERY_TARGET_CHARS = int(os.getenv("GEMINI_RECOVERY_TARGET_CHARS", "6000"))
GEMINI_RECOVERY_HARD_CHARS = int(os.getenv("GEMINI_RECOVERY_HARD_CHARS", "10000"))
_ENGINE8_L1_MAX_ENTRIES = 5000
_ENGINE8_L2_MAX_ENTRIES = 100000
_ENGINE8_NEGATIVE_TTL = 60
_ENGINE8_NEGATIVE_CACHE = {}
_ENGINE8_NEGATIVE_LOCK = threading.RLock()
_ENGINE8_CACHE_PATH = os.getenv("ENGINE8_CACHE_PATH", os.path.join(tempfile.gettempdir(), "rsni_engine8_translation_cache.sqlite3"))


def _canonical_dictionary_fingerprint(custom_dict, italic_dict):
    custom = sorted((str(k), str(v[0]), str(v[1])) for k, v in (custom_dict._entries.items() if custom_dict else []))
    italic = sorted((str(k), str(v)) for k, v in (italic_dict._entries.items() if italic_dict else []))
    raw = json.dumps({"custom": custom, "italic": italic}, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _translation_quality_ok(source: str, translated: str) -> bool:
    """Conservative, language-agnostic gate before a result enters cache."""
    if not translated or _looks_like_error_response(source, translated):
        return False
    # Numeric values and explicit protection tokens must survive translation.
    nums_src = re.findall(r'(?<![A-Za-z])\d+(?:[.,]\d+)*(?![A-Za-z])', source)
    nums_dst = re.findall(r'(?<![A-Za-z])\d+(?:[.,]\d+)*(?![A-Za-z])', translated)
    if nums_src != nums_dst:
        return False
    toks_src = {x.casefold() for x in _RE_PROTECTION_TOKEN.findall(source)}
    toks_dst = {x.casefold() for x in _RE_PROTECTION_TOKEN.findall(translated)}
    if toks_src != toks_dst:
        return False
    # Reject obviously truncated/expanded provider garbage, but keep short headings.
    a, b = len(source.strip()), len(translated.strip())
    if a >= 30 and (b < max(5, int(a * .18)) or b > int(a * 5.0)):
        return False
    return True

def _negative_cache_active(key: str) -> bool:
    now = time.monotonic()
    with _ENGINE8_NEGATIVE_LOCK:
        until = _ENGINE8_NEGATIVE_CACHE.get(key, 0)
        if until <= now:
            _ENGINE8_NEGATIVE_CACHE.pop(key, None)
            return False
        return True

def _negative_cache_mark(key: str) -> None:
    with _ENGINE8_NEGATIVE_LOCK:
        _ENGINE8_NEGATIVE_CACHE[key] = time.monotonic() + _ENGINE8_NEGATIVE_TTL

def _stable_cache_source(text, italic_map=None):
    value = str(text or "")
    if italic_map:
        folded = {str(k).casefold(): str(v) for k, v in italic_map.items()}
        def _repl(m): return "[[ITALIC:" + folded.get(m.group(0).casefold(), "") + "]]"
        value = _RE_PROTECTION_TOKEN.sub(_repl, value)
    return re.sub(r"\s+", " ", value).strip()


class _Engine8TranslationCache:
    def __init__(self, path):
        self.path = path
        self._l1 = OrderedDict()
        self._lock = threading.RLock()
        self._inflight = {}
        self._init_db()

    def _connect(self):
        c = sqlite3.connect(self.path, timeout=5.0)
        c.execute("PRAGMA journal_mode=WAL")
        c.execute("PRAGMA synchronous=NORMAL")
        c.execute("PRAGMA temp_store=MEMORY")
        c.execute("PRAGMA busy_timeout=5000")
        return c

    def _init_db(self):
        try:
            os.makedirs(os.path.dirname(self.path) or ".", exist_ok=True)
            with self._connect() as c:
                c.execute("CREATE TABLE IF NOT EXISTS translation_cache (cache_key TEXT PRIMARY KEY, translated_text TEXT NOT NULL, italic_terms TEXT NOT NULL DEFAULT '[]', created_at INTEGER NOT NULL, last_used INTEGER NOT NULL, hits INTEGER NOT NULL DEFAULT 0)")
                c.execute("CREATE INDEX IF NOT EXISTS idx_translation_cache_last_used ON translation_cache(last_used)")
        except Exception:
            pass

    def make_key(self, text, source, target, fingerprint, italic_map=None):
        payload = "\x1f".join((_stable_cache_source(text, italic_map), str(source), str(target), fingerprint, ENGINE8_POLICY_VERSION, ENGINE8_PROVIDER_VERSION))
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    def get(self, key):
        with self._lock:
            if key in self._l1:
                item = self._l1[key]; self._l1.move_to_end(key); return item
        try:
            now = int(time.time())
            with self._connect() as c:
                row = c.execute("SELECT translated_text, italic_terms FROM translation_cache WHERE cache_key=?", (key,)).fetchone()
                if row: c.execute("UPDATE translation_cache SET last_used=?, hits=hits+1 WHERE cache_key=?", (now, key))
            if row:
                item = (row[0], json.loads(row[1] or "[]")); self._put_l1(key, item); return item
        except Exception:
            pass
        return None

    def _put_l1(self, key, item):
        with self._lock:
            self._l1[key] = item; self._l1.move_to_end(key)
            while len(self._l1) > _ENGINE8_L1_MAX_ENTRIES: self._l1.popitem(last=False)

    def put(self, key, text, italic_terms):
        item = (text, list(italic_terms or [])); self._put_l1(key, item)
        try:
            now = int(time.time())
            with self._connect() as c:
                c.execute("INSERT INTO translation_cache(cache_key,translated_text,italic_terms,created_at,last_used,hits) VALUES(?,?,?,?,?,0) ON CONFLICT(cache_key) DO UPDATE SET translated_text=excluded.translated_text, italic_terms=excluded.italic_terms, last_used=excluded.last_used", (key, text, json.dumps(item[1], ensure_ascii=False), now, now))
                count = c.execute("SELECT COUNT(*) FROM translation_cache").fetchone()[0]
                if count > _ENGINE8_L2_MAX_ENTRIES:
                    c.execute("DELETE FROM translation_cache WHERE cache_key IN (SELECT cache_key FROM translation_cache ORDER BY last_used ASC LIMIT ?)", (count-_ENGINE8_L2_MAX_ENTRIES,))
        except Exception:
            pass

    def acquire(self, key):
        with self._lock:
            ev = self._inflight.get(key)
            if ev is None:
                ev = threading.Event(); self._inflight[key] = ev; return True, ev
            return False, ev

    def finish(self, key):
        with self._lock:
            ev = self._inflight.pop(key, None)
            if ev: ev.set()

_ENGINE8_TRANSLATION_CACHE = _Engine8TranslationCache(_ENGINE8_CACHE_PATH)

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


def _segment_runs_by_format(text_runs, para_style_italic: bool) -> list:
    """Gabungkan run yang berdekatan dengan (miring, vertAlign) sama menjadi
    daftar segmen berurutan ``(teks, is_italic, vtype)``.

    Dipakai bersama oleh jalur lama (``_extract_source_format_map``, yang
    menyisipkan token proteksi ke dalam SATU string sebelum dikirim ke
    layanan terjemahan) dan jalur baru per-segmen
    (``_translate_paragraph_segments``, yang TIDAK PERNAH mengirim bagian
    superscript/subscript/miring-sumber ke layanan terjemahan sama sekali).
    """
    segments = []
    for _, r in text_runs:
        is_ital = _run_effective_italic(r, para_style_italic)
        vtype = _run_vertalign(r)
        t = r.text or ''
        if segments and (segments[-1][1], segments[-1][2]) == (is_ital, vtype):
            segments[-1][0] += t
        else:
            segments.append([t, is_ital, vtype])
    return [(t, i, v) for t, i, v in segments]


def _is_protected_segment(seg_text: str, is_ital: bool, vtype) -> bool:
    """True bila segmen ini harus dilindungi (tidak diterjemahkan) —
    superscript/subscript apa pun panjangnya, atau istilah/judul miring
    sumber yang cukup panjang untuk dianggap istilah asing, bukan tanda
    baca/angka semata."""
    stripped = seg_text.strip()
    if not stripped:
        return False
    if vtype is not None:
        return True
    if is_ital and len(stripped) >= 3 and not _RE_PURE_NUMBER.fullmatch(stripped):
        return True
    return False


def _translate_paragraph_segments(text_runs, tr, para_style_italic: bool):
    """Jalur terjemahan PER-SEGMEN untuk paragraf yang memiliki superscript/
    subscript atau istilah miring-sumber (mis. CATATAN/NOTE berisi notasi
    pangkat seperti "kWm⁻²" atau "0,7 m²").

    Masalah pada jalur lama: seluruh paragraf digabung jadi SATU string,
    dengan setiap segmen superscript/subscript disisipi token proteksi acak
    (mis. ``ZXQSRC...QXZ``), lalu SATU request itu dikirim ke Google
    Translate. Untuk CATATAN teknis yang memuat beberapa notasi pangkat
    berdekatan (beberapa token mirip dalam satu kalimat pendek), mesin
    terjemahan kerap gagal mengembalikan token-token itu persis utuh
    (tertukar/hilang/dianggap pengulangan lalu dipangkas), sehingga
    validasi round-trip pada ``_Translator.translate_one`` gagal berulang
    kali dan SELURUH paragraf berakhir tidak diterjemahkan (font merah) —
    padahal sebagian besar isinya teks biasa yang sebenarnya mudah
    diterjemahkan.

    Solusi: paragraf dipecah dulu secara lokal (murni baca run, tidak
    memanggil layanan apa pun) di setiap batas segmen terproteksi. Segmen
    "polos" yang berurutan digabung jadi satu blok teks dan diterjemahkan
    satu per satu lewat ``tr.translate_one`` (kamus SNI & Kamus Istilah
    Asing tetap berlaku per blok, sama seperti sebelumnya). Segmen
    terproteksi TIDAK PERNAH dikirim ke layanan terjemahan sama sekali —
    formatnya (dan teksnya) sehingga mustahil rusak, dan Google Translate
    tidak pernah perlu me-round-trip token aneh untuk bagian ini.

    Return ``(translated_text, format_spans, italic_terms_found)`` bila
    SEMUA blok polos berhasil diterjemahkan (atau memang trivial/di-skip
    karena terlalu pendek), atau ``None`` bila ADA satu saja blok yang
    gagal — dalam hal ini paragraf dianggap gagal secara UTUH dan TIDAK
    disentuh sama sekali (persis seperti kegagalan pada jalur lama),
    supaya tetap aman diulang di tahap pemulihan / retry.
    """
    segments = _segment_runs_by_format(text_runs, para_style_italic)

    blocks: list[tuple[str, object]] = []
    current_plain: list[str] = []
    for seg_text, is_ital, vtype in segments:
        if _is_protected_segment(seg_text, is_ital, vtype):
            if current_plain:
                blocks.append(('plain', ''.join(current_plain)))
                current_plain = []
            blocks.append(('protected', (seg_text, is_ital, vtype)))
        else:
            current_plain.append(seg_text)
    if current_plain:
        blocks.append(('plain', ''.join(current_plain)))

    out_parts: list[str] = []
    format_spans: list[tuple] = []
    italic_terms_found: list[str] = []
    cur_len = 0

    for kind, payload in blocks:
        if kind == 'protected':
            seg_text, is_ital, vtype = payload
            format_spans.append((cur_len, cur_len + len(seg_text), is_ital, vtype))
            out_parts.append(seg_text)
            cur_len += len(seg_text)
            continue

        block_text = payload
        if not block_text.strip():
            out_parts.append(block_text)
            cur_len += len(block_text)
            continue

        if tr.italic_dict and len(tr.italic_dict) > 0:
            pre_text, dict_italic_map = tr.italic_dict._apply_pre(block_text)
        else:
            pre_text, dict_italic_map = block_text, {}

        before_fail = len(tr.failed_texts)
        block_result, block_italic_terms = tr.translate_one(pre_text, dict_italic_map)
        block_result = _restore_plain_block_whitespace(block_text, block_result)
        if len(tr.failed_texts) > before_fail:
            # Satu blok gagal -> seluruh paragraf dianggap gagal ronde ini
            # (bukan hasil campuran sebagian bahasa Inggris/Indonesia),
            # supaya aman diulang utuh pada tahap pemulihan.
            return None
        out_parts.append(block_result)
        cur_len += len(block_result)
        for term in block_italic_terms:
            if term not in italic_terms_found:
                italic_terms_found.append(term)

    return ''.join(out_parts), format_spans, italic_terms_found


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

    # Empat spasi setelah CATATAN harus menjadi run baseline tersendiri.
    # Jangan gabungkan spasi ini ke run pertama kalimat karena run pertama
    # dapat mewarisi w:vertAlign=superscript/subscript dari notasi teknis.
    # Penggabungan tersebut sebelumnya dapat membuat spasi/huruf sesudah
    # pangkat ikut terangkat (contoh: kWm−2 hingga, m2 dengan ...).
    gap_run = para.add_run('    ')
    gap_run.bold = False
    gap_run.font.superscript = False
    gap_run.font.subscript = False

    for piece, rpr in source_segments:
        if not piece:
            continue
        run_n = para.add_run(piece)
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


def _google_error_code(exc: Exception):
    """Best-effort klasifikasi error Google Translate/deep-translator."""
    text = str(exc or "").casefold()
    m = re.search(r'(?<!\d)(400|403|408|409|429|500|502|503|504)(?!\d)', text)
    if m:
        return int(m.group(1))
    if any(x in text for x in ("too many request", "too many requests", "rate limit", "unusual traffic")):
        return 429
    if any(x in text for x in ("timeout", "timed out", "connection reset", "temporarily unavailable")):
        return 503
    return None


class _GoogleTranslateCoordinator:
    """Circuit breaker global untuk seluruh worker Google Translate.

    Saat satu worker melihat 429/error sementara, SEMUA worker berhenti menekan
    Google dan beralih ke Gemini. Setelah cooldown hanya satu worker diizinkan
    melakukan probe. Jika probe sukses, circuit ditutup dan semua worker kembali
    ke Google Translate. Ini mencegah thundering-herd/request bombing.
    """
    def __init__(self):
        self._lock = threading.Lock()
        self._open_until = 0.0
        self._probe_in_flight = False
        self._strike = 0
        self._base_cooldown = float(os.getenv("GOOGLE_TRANSLATE_COOLDOWN", "10"))
        self._max_cooldown = float(os.getenv("GOOGLE_TRANSLATE_MAX_COOLDOWN", "60"))
        self.last_code = None

    def permit_google(self) -> tuple[bool, bool]:
        now = time.monotonic()
        with self._lock:
            if now < self._open_until:
                return False, False
            if self._strike > 0:
                if self._probe_in_flight:
                    return False, False
                self._probe_in_flight = True
                return True, True
            return True, False

    def success(self, was_probe: bool = False):
        with self._lock:
            if was_probe or self._strike:
                self._strike = 0
                self._open_until = 0.0
            self._probe_in_flight = False
            self.last_code = None

    def failure(self, code=None, was_probe: bool = False):
        transient = code in (408, 409, 429, 500, 502, 503, 504) or code is None
        with self._lock:
            self._probe_in_flight = False
            self.last_code = code
            if transient:
                self._strike = min(self._strike + 1, 4)
                delay = min(self._max_cooldown, self._base_cooldown * (2 ** (self._strike - 1)))
                # Jitter kecil agar semua worker tidak bangun pada milidetik sama.
                self._open_until = max(self._open_until, time.monotonic() + delay + random.uniform(0.2, 0.9))

    def state(self) -> str:
        with self._lock:
            remain = max(0.0, self._open_until - time.monotonic())
            if remain > 0:
                return f"Gemini fallback (Google cooldown {remain:.1f}s; kode={self.last_code or 'network'})"
            if self._strike:
                return "probe Google Translate"
            return "Google Translate normal"


_GOOGLE_TRANSLATE_COORDINATOR = _GoogleTranslateCoordinator()

class _SharedFallbackCoordinator:
    """Circuit breaker ringan yang dipakai bersama seluruh worker per provider.

    Tujuannya bukan mempercepat provider, tetapi mencegah semua worker melakukan
    retry bersamaan saat provider sedang bermasalah. Setelah cooldown, hanya satu
    worker menjadi probe; worker lain langsung turun ke provider berikutnya.
    """
    def __init__(self, name: str, base_env: str, max_env: str,
                 base_default: float, max_default: float):
        self.name = name
        self._lock = threading.Lock()
        self._open_until = 0.0
        self._probe_in_flight = False
        self._strike = 0
        self._base = float(os.getenv(base_env, str(base_default)))
        self._max = float(os.getenv(max_env, str(max_default)))
        self.last_code = None

    def permit(self) -> tuple[bool, bool]:
        now = time.monotonic()
        with self._lock:
            if now < self._open_until:
                return False, False
            if self._strike:
                if self._probe_in_flight:
                    return False, False
                self._probe_in_flight = True
                return True, True
            return True, False

    def success(self, was_probe: bool = False):
        with self._lock:
            self._probe_in_flight = False
            self._strike = 0
            self._open_until = 0.0
            self.last_code = None

    def failure(self, code=None, was_probe: bool = False):
        with self._lock:
            self._probe_in_flight = False
            self.last_code = code
            self._strike = min(self._strike + 1, 5)
            delay = min(self._max, self._base * (2 ** (self._strike - 1)))
            self._open_until = max(
                self._open_until,
                time.monotonic() + delay + random.uniform(0.15, 0.75),
            )

    def state(self) -> str:
        with self._lock:
            remain = max(0.0, self._open_until - time.monotonic())
            if remain:
                return f"{self.name} cooldown {remain:.1f}s"
            if self._strike:
                return f"probe {self.name}"
            return f"{self.name} normal"


_GEMINI_COORDINATOR = _SharedFallbackCoordinator(
    'Gemini', 'GEMINI_FALLBACK_COOLDOWN', 'GEMINI_FALLBACK_MAX_COOLDOWN', 8, 60
)
_ARGOS_COORDINATOR = _SharedFallbackCoordinator(
    'Argos', 'ARGOS_FALLBACK_COOLDOWN', 'ARGOS_FALLBACK_MAX_COOLDOWN', 3, 20
)
_MYMEMORY_COORDINATOR = _SharedFallbackCoordinator(
    'MyMemory', 'MYMEMORY_FALLBACK_COOLDOWN', 'MYMEMORY_FALLBACK_MAX_COOLDOWN', 12, 90
)

_ARGOS_INIT_LOCK = threading.Lock()
_ARGOS_TRANSLATE_LOCK = threading.Lock()  # model lokal: hindari N inferensi berat serentak
_ARGOS_READY = False
_ARGOS_ERROR = None
_ARGOS_INIT_FAILURES = 0
_ARGOS_INIT_COOLDOWN_UNTIL = 0.0


def _argos_init_cooldown_seconds(failures: int) -> float:
    """Backoff inisialisasi Argos agar Streamlit tidak mengunduh model berulang-ulang."""
    base = max(5.0, float(os.getenv('ARGOS_INIT_COOLDOWN', '30')))
    maximum = max(base, float(os.getenv('ARGOS_INIT_MAX_COOLDOWN', '300')))
    # 30, 60, 120, 240, 300 ... (default), plus jitter kecil.
    return min(maximum, base * (2 ** max(0, failures - 1))) + random.uniform(0.25, 1.25)


def _argos_translate(text: str, source: str = 'en', target: str = 'id') -> str:
    """Terjemahan lokal Argos; init/download model thread-safe dengan failure cooldown."""
    global _ARGOS_READY, _ARGOS_ERROR, _ARGOS_INIT_FAILURES, _ARGOS_INIT_COOLDOWN_UNTIL
    try:
        import argostranslate.package as argos_package
        import argostranslate.translate as argos_translate
    except Exception as exc:
        raise RuntimeError('Argos Translate belum terpasang') from exc

    src = 'en' if source in ('auto', 'en') else source
    dst = 'id' if target in ('id', 'indonesian') else target

    if not _ARGOS_READY:
        # Fast-fail selama cooldown. Semua worker melihat state yang sama dan tidak
        # mencoba update index/download model secara berulang.
        remaining = _ARGOS_INIT_COOLDOWN_UNTIL - time.monotonic()
        if remaining > 0:
            detail = f': {_ARGOS_ERROR}' if _ARGOS_ERROR else ''
            raise RuntimeError(f'Inisialisasi Argos cooldown {remaining:.1f}s{detail}')

        with _ARGOS_INIT_LOCK:
            if not _ARGOS_READY:
                # Worker mungkin menunggu lock saat worker lain gagal; cek cooldown lagi.
                remaining = _ARGOS_INIT_COOLDOWN_UNTIL - time.monotonic()
                if remaining > 0:
                    detail = f': {_ARGOS_ERROR}' if _ARGOS_ERROR else ''
                    raise RuntimeError(f'Inisialisasi Argos cooldown {remaining:.1f}s{detail}')
                try:
                    installed = argos_translate.get_installed_languages()
                    ok = any(
                        a.code == src and any(t.code == dst for t in a.translations_to)
                        for a in installed
                    )
                    if not ok:
                        argos_package.update_package_index()
                        packages = argos_package.get_available_packages()
                        pkg = next((x for x in packages if x.from_code == src and x.to_code == dst), None)
                        if pkg is None:
                            raise RuntimeError(f'Paket Argos {src}->{dst} tidak tersedia')
                        download_path = pkg.download()
                        argos_package.install_from_path(download_path)

                    # Verifikasi ulang setelah instalasi; jangan menandai READY hanya
                    # karena proses install tidak melempar exception.
                    installed = argos_translate.get_installed_languages()
                    ok = any(
                        a.code == src and any(t.code == dst for t in a.translations_to)
                        for a in installed
                    )
                    if not ok:
                        raise RuntimeError(f'Model Argos {src}->{dst} belum aktif setelah instalasi')

                    _ARGOS_READY = True
                    _ARGOS_ERROR = None
                    _ARGOS_INIT_FAILURES = 0
                    _ARGOS_INIT_COOLDOWN_UNTIL = 0.0
                except Exception as exc:
                    _ARGOS_READY = False
                    _ARGOS_ERROR = exc
                    _ARGOS_INIT_FAILURES += 1
                    delay = _argos_init_cooldown_seconds(_ARGOS_INIT_FAILURES)
                    _ARGOS_INIT_COOLDOWN_UNTIL = time.monotonic() + delay
                    raise RuntimeError(
                        f'Inisialisasi Argos gagal; cooldown {delay:.1f}s '
                        f'(percobaan gagal #{_ARGOS_INIT_FAILURES}): {exc}'
                    ) from exc

    with _ARGOS_TRANSLATE_LOCK:
        return argos_translate.translate(text, src, dst)


def _fallback_candidate_ok(source: str, candidate: str, expected_tokens: set[str]) -> bool:
    if not candidate or not _translation_quality_ok(source, candidate):
        return False
    returned = {x.casefold() for x in _RE_PROTECTION_TOKEN.findall(candidate)}
    if returned != expected_tokens:
        return False
    src_plain = _RE_PROTECTION_TOKEN.sub('', source)
    dst_plain = _RE_PROTECTION_TOKEN.sub('', candidate)
    if (re.search(r'[A-Za-z]{3}', src_plain) and
        re.sub(r'\s+', ' ', src_plain).strip().casefold() ==
        re.sub(r'\s+', ' ', dst_plain).strip().casefold()):
        return False
    return True

_GEMINI_WORKER_LOCK = threading.Lock()
_GOOGLE_RETRY_STAGE_LOCK = threading.Lock()  # serialkan retry Google tahap 2/3 antar-worker
_GEMINI_WORKER_COUNTER = 0


def _next_gemini_worker_key_index(key_count: int) -> int:
    global _GEMINI_WORKER_COUNTER
    if key_count <= 0:
        return 0
    with _GEMINI_WORKER_LOCK:
        idx = _GEMINI_WORKER_COUNTER % key_count
        _GEMINI_WORKER_COUNTER += 1
        return idx


class _Translator:
    def __init__(self, source: str = 'auto', target: str = 'id', 
                 custom_dict: CustomDictionary | None = None,
                 italic_dict: ItalicDictionary | None = None,
                 dictionary_fingerprint: str = ''):
        try: from deep_translator import GoogleTranslator
        except ImportError: raise ImportError("Jalankan: pip install deep-translator")
        self._cls = GoogleTranslator
        self.source = source; self.target = target
        self.custom_dict = custom_dict; self.italic_dict = italic_dict
        self._client = self._cls(source=self.source, target=self.target)
        self.failed_texts: list[str] = []
        self.dictionary_fingerprint = dictionary_fingerprint or _canonical_dictionary_fingerprint(custom_dict, italic_dict)
        self._coordinator = _GOOGLE_TRANSLATE_COORDINATOR
        self._gemini_fallback = None
        self.last_provider = 'google'
        self.last_error_code = None

    def translate_one(self, text: str, italic_map: dict = None) -> tuple[str, list[str]]:
        t = text.strip()
        if not t or _skip_text(t): return text, []
        cache_key = _ENGINE8_TRANSLATION_CACHE.make_key(t, self.source, self.target, self.dictionary_fingerprint, italic_map)
        cached = _ENGINE8_TRANSLATION_CACHE.get(cache_key)
        if cached is not None:
            return cached
        owner, event = _ENGINE8_TRANSLATION_CACHE.acquire(cache_key)
        owns_flight = owner
        if not owner:
            event.wait(timeout=45.0)
            cached = _ENGINE8_TRANSLATION_CACHE.get(cache_key)
            if cached is not None:
                return cached
            # Owner gagal: jangan menunggu selamanya; request ini lanjut normal.
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
            _ENGINE8_TRANSLATION_CACHE.put(cache_key, result, italic_terms_found)
            if owns_flight:
                _ENGINE8_TRANSLATION_CACHE.finish(cache_key)
            return result, italic_terms_found

        expected_tokens = {
            token.casefold() for token in _RE_PROTECTION_TOKEN.findall(t)
        }
        result = None
        last_error = None

        # Translator utama memakai tiga gelombang terkontrol:
        # 1) Google -> seluruh pool Gemini
        # 2) Google -> seluruh pool Gemini -> Argos
        # 3) Google -> seluruh pool Gemini -> Argos -> MyMemory
        # Jika semuanya gagal, unit masuk tahap pemulihan.
        #
        # Gemini dibuat ulang pada setiap gelombang agar selalu mulai dari API key
        # awal milik worker lalu failover berurutan sampai seluruh pool dicoba.
        gemini_keys = _get_gemini_api_keys()
        if gemini_keys:
            start_idx = _next_gemini_worker_key_index(len(gemini_keys))
            worker_gemini_keys = gemini_keys[start_idx:] + gemini_keys[:start_idx]
        else:
            worker_gemini_keys = []

        def _google_stage(stage_no: int) -> bool:
            nonlocal result, last_error
            # Tahap pertama tetap menghormati circuit breaker global. Tahap 2/3
            # adalah retry eksplisit yang diminta pipeline, tetapi diserialkan
            # antar-worker agar Google tidak dibombardir secara paralel.
            if stage_no == 1:
                allowed, probe = self._coordinator.permit_google()
                if not allowed:
                    self.last_provider = 'gemini'
                    return False
                lock_ctx = contextlib.nullcontext()
            else:
                probe = False
                lock_ctx = _GOOGLE_RETRY_STAGE_LOCK

            with lock_ctx:
                try:
                    candidate = self._client.translate(t)
                    if not candidate or not _translation_quality_ok(t, candidate):
                        raise ValueError('respons layanan terjemahan tidak lolos quality gate')
                    returned_tokens = {
                        token.casefold() for token in _RE_PROTECTION_TOKEN.findall(candidate)
                    }
                    if returned_tokens != expected_tokens:
                        raise ValueError('token kamus/format berubah atau hilang')
                    source_plain = _RE_PROTECTION_TOKEN.sub('', t)
                    result_plain = _RE_PROTECTION_TOKEN.sub('', candidate)
                    if (re.search(r'[A-Za-z]{3}', source_plain) and
                        re.sub(r'\s+', ' ', source_plain).strip().casefold() ==
                        re.sub(r'\s+', ' ', result_plain).strip().casefold()):
                        raise ValueError('teks dikembalikan tanpa diterjemahkan')
                    result = candidate
                    self.last_provider = 'Google Translate'
                    self.last_error_code = None
                    self._coordinator.success(probe)
                    return True
                except Exception as exc:
                    last_error = exc
                    code = _google_error_code(exc)
                    self.last_error_code = code
                    self._coordinator.failure(code, probe)
                    return False

        def _gemini_stage(stage_no: int) -> tuple[bool, tuple[str, list[str]] | None]:
            nonlocal last_error
            if not worker_gemini_keys:
                last_error = RuntimeError('Gemini API key belum tersedia')
                return False, None
            try:
                # Instance baru = kembali ke API key awal worker pada setiap gelombang.
                gemini = _GeminiRecoveryTranslator(
                    self.source, self.target, self.custom_dict, self.italic_dict,
                    self.dictionary_fingerprint, api_keys=worker_gemini_keys,
                )
                gemini_result, gemini_terms = gemini.translate_one(text, italic_map or {})
                if not gemini_result or gemini_result == text:
                    raise RuntimeError('Gemini tidak menghasilkan terjemahan valid')
                self.last_provider = f'Gemini Flash API {gemini.active_key_number}'
                self.last_error_code = None
                return True, (gemini_result, gemini_terms)
            except Exception as exc:
                last_error = exc
                self.last_error_code = _gemini_status_code(exc)
                return False, None

        def _argos_stage(stage_no: int) -> bool:
            nonlocal result, last_error
            try:
                candidate = _argos_translate(t, self.source, self.target)
                if not _fallback_candidate_ok(t, candidate, expected_tokens):
                    raise RuntimeError('Argos tidak lolos quality gate')
                result = candidate
                self.last_provider = 'Argos Translate'
                self.last_error_code = None
                return True
            except Exception as exc:
                last_error = exc
                self.last_error_code = _google_error_code(exc)
                return False

        # GELOMBANG 1: Google -> Gemini seluruh key.
        _google_stage(1)
        if result is None:
            ok, gemini_payload = _gemini_stage(1)
            if ok and gemini_payload is not None:
                if owns_flight:
                    _ENGINE8_TRANSLATION_CACHE.finish(cache_key)
                return gemini_payload

        # GELOMBANG 2: Google lagi -> Gemini seluruh key -> Argos.
        if result is None:
            time.sleep(0.35 + random.uniform(0.05, 0.20))
            _google_stage(2)
        if result is None:
            ok, gemini_payload = _gemini_stage(2)
            if ok and gemini_payload is not None:
                if owns_flight:
                    _ENGINE8_TRANSLATION_CACHE.finish(cache_key)
                return gemini_payload
        if result is None:
            _argos_stage(2)

        # GELOMBANG 3: Google lagi -> Gemini seluruh key -> Argos -> MyMemory.
        if result is None:
            time.sleep(0.55 + random.uniform(0.05, 0.25))
            _google_stage(3)
        if result is None:
            ok, gemini_payload = _gemini_stage(3)
            if ok and gemini_payload is not None:
                if owns_flight:
                    _ENGINE8_TRANSLATION_CACHE.finish(cache_key)
                return gemini_payload
        if result is None:
            _argos_stage(3)

        if result is None:
            try:
                from deep_translator import MyMemoryTranslator
                mm_source = 'en-GB' if self.source in ('auto', 'en') else self.source
                mm_target = 'id-ID' if self.target == 'id' else self.target
                candidate = MyMemoryTranslator(source=mm_source, target=mm_target).translate(t)
                if not _fallback_candidate_ok(t, candidate, expected_tokens):
                    raise RuntimeError('MyMemory tidak lolos quality gate')
                result = candidate
                self.last_provider = 'MyMemory Translate'
                self.last_error_code = None
            except Exception as exc:
                last_error = exc
                self.last_error_code = _google_error_code(exc)

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
            safe_title = _TITLE_SAFE_FALLBACKS.get(
                re.sub(r'\s+', ' ', text).strip().casefold()
            )
            if safe_title:
                result = safe_title

        if result is None:
            preview = re.sub(r'\s+', ' ', text).strip()[:100]
            # Jangan gagalkan seluruh pipeline karena satu request eksternal.
            # Laporkan sebagai peringatan pada ringkasan proses, dan
            # KEMBALIKAN TEKS ASLI APA ADANYA (`text`, sebelum token kamus/
            # proteksi apa pun disisipkan) tanpa menerapkan `_apply_post`.
            #
            # Sebelumnya baris ini melanjutkan ke `_apply_post` di bawah
            # walau `result` gagal total (`result = t`, teks sumber yang
            # SUDAH disisipi token kamus). Karena `_apply_post` mengganti
            # token kamus dengan kata TARGET tanpa syarat, satu paragraf
            # yang gagal total bisa berakhir sebagai campuran aneh —
            # sebagian besar tetap bahasa sumber, tapi satu-dua kata umum
            # (mis. "can") ikut terganti ("dapat") seolah-olah paragraf itu
            # sudah diterjemahkan. Akibatnya pemanggil (`_translate_para`)
            # tidak mendeteksi ini sebagai "tidak ada perubahan" dan
            # menyimpan hasil setengah-campur itu ke dokumen. Dengan
            # langsung mengembalikan `text` asli di sini, pemanggil selalu
            # bisa mendeteksi kegagalan lewat kesamaan teks dan
            # membiarkan paragraf 100% utuh (termasuk semua run/format)
            # untuk diulang di tahap pemulihan.
            self.failed_texts.append(preview)
            _negative_cache_mark(cache_key)
            if owns_flight:
                _ENGINE8_TRANSLATION_CACHE.finish(cache_key)
            return text, []
        if token_map: result = self.custom_dict._apply_post(result, token_map)
        italic_terms_found = []
        if final_italic_map:
            _idict = self.italic_dict if self.italic_dict is not None else ItalicDictionary()
            result, italic_terms_found = _idict._apply_post(result, final_italic_map)
        _ENGINE8_TRANSLATION_CACHE.put(cache_key, result, italic_terms_found)
        if owns_flight:
            _ENGINE8_TRANSLATION_CACHE.finish(cache_key)
        return result, italic_terms_found



def _get_gemini_api_keys() -> list[str]:
    """Ambil maksimum 15 Gemini API key tanpa hard-code.

    Urutan prioritas:
      1) GEMINI_API_KEY_1
      2) GEMINI_API_KEY_2

    Hanya dua nama key tersebut yang didukung. Nilai duplikat dibuang.
    Environment variable dibaca lebih dahulu, lalu Streamlit Secrets untuk
    nama yang belum tersedia di environment.
    """
    names = tuple(f"GEMINI_API_KEY_{i}" for i in range(1, 16))
    values = []

    def _add(value):
        value = str(value or "").strip()
        if value and value not in values:
            values.append(value)

    for name in names:
        _add(os.getenv(name, ""))

    try:
        import streamlit as st
        for name in names:
            try:
                _add(st.secrets.get(name, ""))
            except Exception:
                pass
    except Exception:
        pass

    # Maksimum 15 key berbeda; worker memilih key awal secara round-robin.
    return values[:15]


def _gemini_auth_error(exc: Exception) -> bool:
    """True hanya untuk error autentikasi/otorisasi yang layak failover key.

    401/403 menandakan credential bermasalah. Error 429 ditangani terpisah
    oleh pool Gemini: request langsung mencoba key berikutnya sampai seluruh
    key yang tersedia sudah memperoleh satu kesempatan.
    """
    status = _gemini_status_code(exc)
    if status in (401, 403):
        return True

    text = str(exc).casefold()
    auth_markers = (
        "api key not valid", "api_key_invalid", "invalid api key",
        "invalid authentication", "authentication failed", "unauthenticated",
        "permission denied", "permission_denied", "forbidden",
        "credentials are invalid", "credential is invalid",
    )
    return any(marker in text for marker in auth_markers)


def _gemini_status_code(exc: Exception):
    """Best-effort HTTP/RPC status extraction without exposing credentials."""
    for obj in (exc, getattr(exc, "response", None)):
        if obj is None:
            continue
        for attr in ("status_code", "code"):
            value = getattr(obj, attr, None)
            if callable(value):
                try: value = value()
                except Exception: value = None
            try:
                if value is not None:
                    return int(value)
            except Exception:
                pass
    m = re.search(r'(?<!\d)(401|403|408|409|429|500|502|503|504)(?!\d)', str(exc))
    return int(m.group(1)) if m else None


def _gemini_transient_error(exc: Exception) -> bool:
    """Retry only failures likely to succeed later; permanent 4xx fail fast."""
    status = _gemini_status_code(exc)
    if status in (408, 409, 429, 500, 502, 503, 504):
        return True
    text = str(exc).casefold()
    markers = (
        "timeout", "timed out", "deadline exceeded", "temporarily unavailable",
        "service unavailable", "connection reset", "connection aborted",
        "connection error", "remote disconnected", "server disconnected",
        "internal server error", "bad gateway", "gateway timeout",
    )
    return any(x in text for x in markers)


class _GeminiRecoveryTranslator:
    """Satu tahap pemulihan Gemini 2.5 Flash yang hemat Free Tier.

    Recovery memakai adaptive micro-batch (maks. 8 unit / target 6000 karakter),
    concurrency=1, thinking=0, structured JSON, cache terpisah, serta seluruh
    proteksi Kamus SNI/Kamus Istilah Asing yang sudah dimiliki Engine 8.
    Tidak ada recovery AI kedua: item yang ditolak quality gate tetap gagal dan
    kemudian dipertahankan dalam bahasa Inggris/merah oleh pipeline.
    """

    def __init__(self, source: str = 'auto', target: str = 'id',
                 custom_dict: CustomDictionary | None = None,
                 italic_dict: ItalicDictionary | None = None,
                 dictionary_fingerprint: str = '', api_keys: list[str] | None = None):
        try:
            from google import genai
            from google.genai import types
        except ImportError as exc:
            raise ImportError("Jalankan: pip install -U google-genai") from exc
        api_keys = list(api_keys or _get_gemini_api_keys())
        if not api_keys:
            raise RuntimeError(
                "Gemini API key belum diset. Isi GEMINI_API_KEY_1 sampai "
                "GEMINI_API_KEY_15 di Streamlit Secrets/environment variable."
            )
        self.source = source; self.target = target
        self.custom_dict = custom_dict; self.italic_dict = italic_dict
        self.failed_texts: list[str] = []
        self.dictionary_fingerprint = dictionary_fingerprint or _canonical_dictionary_fingerprint(custom_dict, italic_dict)
        self.model = GEMINI_RECOVERY_MODEL
        self._genai = genai
        self._api_keys = api_keys
        self._active_key_index = 0
        self._client = genai.Client(api_key=self._api_keys[0])
        self._types = types
        # Key yang SUDAH mendapat satu kesempatan AI tetapi gagal validasi.
        # translate_one tidak akan mengirimnya lagi: tetap satu pemulihan.
        self._batch_failed_keys: set[str] = set()
        self.diagnostics: list[str] = []

    def _cache_key(self, text: str, italic_map=None) -> str:
        fp = "|".join((self.dictionary_fingerprint, GEMINI_RECOVERY_VERSION, self.model))
        return _ENGINE8_TRANSLATION_CACHE.make_key(text, self.source, self.target, fp, italic_map)

    def _prepare(self, text: str, italic_map=None):
        original = text
        t = text.strip()
        token_map = {}
        if self.custom_dict and len(self.custom_dict) > 0:
            t, token_map = self.custom_dict._apply_pre(t)
        return original, t, token_map, (italic_map or {})

    @staticmethod
    def _candidate_ok(source: str, candidate: str) -> bool:
        if not candidate or not _translation_quality_ok(source, candidate):
            return False
        expected = {x.casefold() for x in _RE_PROTECTION_TOKEN.findall(source)}
        returned = {x.casefold() for x in _RE_PROTECTION_TOKEN.findall(candidate)}
        if expected != returned:
            return False
        src_plain = _RE_PROTECTION_TOKEN.sub('', source)
        dst_plain = _RE_PROTECTION_TOKEN.sub('', candidate)
        if (re.search(r'[A-Za-z]{3}', src_plain) and
            re.sub(r'\s+', ' ', src_plain).strip().casefold() ==
            re.sub(r'\s+', ' ', dst_plain).strip().casefold()):
            return False
        return True

    def _config(self, schema=None):
        kw = dict(temperature=0, max_output_tokens=8192)
        # google-genai modern: thinking 0 = latency/token lebih rendah.
        try:
            kw['thinking_config'] = self._types.ThinkingConfig(thinking_budget=0)
        except Exception:
            pass
        if schema is not None:
            kw['response_mime_type'] = 'application/json'
            kw['response_schema'] = schema
        return self._types.GenerateContentConfig(**kw)

    @property
    def active_key_number(self) -> int:
        """Nomor key aktif dalam pool worker tanpa mengekspos secret."""
        return self._active_key_index + 1

    def _switch_to_backup_key(self) -> bool:
        """Pindah ke key berikutnya dalam pool bila autentikasi key aktif gagal."""
        next_index = self._active_key_index + 1
        if next_index >= len(self._api_keys):
            return False
        old_no = self.active_key_number
        self._active_key_index = next_index
        self._client = self._genai.Client(api_key=self._api_keys[next_index])
        self.diagnostics.append(f"Gemini key {old_no}: autentikasi gagal; failover ke key {self.active_key_number}.")
        return True

    def _generate_content_with_failover(self, *, prompt: str, config):
        """Generate dengan failover key berurutan.

        401/403/invalid credential dan 429 langsung memindahkan request ke key
        berikutnya. Setiap key hanya dicoba satu kali untuk request yang sama.
        Setelah seluruh pool habis, exception terakhir dilempar ke caller agar
        pipeline dapat berpindah ke provider berikutnya (Argos).
        """
        last_exc = None
        tried = 0
        while tried < len(self._api_keys):
            key_no = self.active_key_number
            try:
                return self._client.models.generate_content(
                    model=self.model, contents=prompt, config=config
                )
            except Exception as exc:
                last_exc = exc
                status = _gemini_status_code(exc)
                rotate = _gemini_auth_error(exc) or status == 429
                if not rotate:
                    raise
                tried += 1
                reason = f"HTTP {status}" if status else "credential invalid"
                if not self._switch_to_backup_key():
                    self.diagnostics.append(
                        f"Gemini key {key_no}: {reason}; seluruh {len(self._api_keys)} key telah dicoba."
                    )
                    raise
                self.diagnostics.append(
                    f"Gemini key {key_no}: {reason}; failover ke key {self.active_key_number}."
                )
        if last_exc is not None:
            raise last_exc
        raise RuntimeError('Gemini gagal tanpa exception yang dapat diidentifikasi')

    def _generate_batch(self, rows: list[dict]) -> dict[str, str]:
        payload = json.dumps(rows, ensure_ascii=False, separators=(',', ':'))
        prompt = (
            "Anda adalah penerjemah standar teknis ISO/IEC ke Bahasa Indonesia baku untuk rancangan SNI. "
            "Terjemahkan SETIAP item secara independen. Jangan menjelaskan, meringkas, menambah, atau mengurangi isi. "
            "Pertahankan semua angka, desimal, nomor pasal, simbol, satuan, rumus, tanda baca teknis, dan setiap token "
            "ZXQ...QXZ PERSIS. Token adalah placeholder dan dilarang diubah/dihapus/dipisah. "
            "Gunakan bahasa teknis formal; 'shall' normatif umumnya 'harus'. "
            "Kembalikan tepat satu hasil untuk setiap id dalam JSON terstruktur.\nITEM:\n" + payload
        )
        schema = {
            'type': 'object',
            'properties': {
                'translations': {
                    'type': 'array',
                    'items': {
                        'type': 'object',
                        'properties': {
                            'id': {'type': 'string'},
                            'text': {'type': 'string'},
                        },
                        'required': ['id', 'text'],
                    },
                },
            },
            'required': ['translations'],
        }
        response = self._generate_content_with_failover(
            prompt=prompt, config=self._config(schema)
        )
        raw = (getattr(response, 'text', None) or '').strip()
        data = json.loads(raw)
        out = {}
        for item in data.get('translations', []):
            iid = str(item.get('id', ''))
            if iid:
                out[iid] = str(item.get('text', '') or '').strip()
        return out

    def _chunks(self, jobs: list[dict]):
        batch=[]; chars=0
        max_items=max(1, min(GEMINI_RECOVERY_MAX_ITEMS, 8))
        target=max(1000, GEMINI_RECOVERY_TARGET_CHARS)
        hard=max(target, GEMINI_RECOVERY_HARD_CHARS)
        for job in jobs:
            n=len(job['protected'])
            # Teks sangat panjang tetap satu unit/request; jangan pecah semantik.
            if batch and (len(batch) >= max_items or chars+n > target or chars+n > hard):
                yield batch; batch=[]; chars=0
            batch.append(job); chars += n
            if chars >= target or chars >= hard or len(batch) >= max_items:
                yield batch; batch=[]; chars=0
        if batch: yield batch

    def prewarm(self, requests: list[tuple[str, dict]]) -> None:
        """Terjemahkan semua unit recovery dalam micro-batch dan isi cache.

        Satu unit hanya memperoleh satu hasil AI. Retry 429/5xx/timeout adalah
        retry transport dari request batch yang sama, bukan pemulihan kedua.
        """
        jobs=[]; seen=set()
        for text, italic_map in requests:
            t0=text.strip()
            if not t0 or _skip_text(t0):
                continue
            key=self._cache_key(t0, italic_map)
            if key in seen or _ENGINE8_TRANSLATION_CACHE.get(key) is not None:
                continue
            seen.add(key)
            original, protected, token_map, final_italic = self._prepare(t0, italic_map)
            uncovered = _RE_PROTECTION_TOKEN.sub('', protected)
            if token_map and not re.findall(r'[A-Za-zÀ-ÿ]{2,}', uncovered):
                result=self.custom_dict._apply_post(protected, token_map)
                terms=[]
                if final_italic:
                    idict=self.italic_dict if self.italic_dict is not None else ItalicDictionary()
                    result,terms=idict._apply_post(result, final_italic)
                _ENGINE8_TRANSLATION_CACHE.put(key,result,terms)
                continue
            jobs.append(dict(key=key, original=original, protected=protected,
                             token_map=token_map, italic_map=final_italic))

        seq=0
        for chunk in self._chunks(jobs):
            rows=[]; idmap={}
            for job in chunk:
                seq += 1; iid=f'R{seq:06d}'
                rows.append({'id':iid,'text':job['protected']}); idmap[iid]=job
            answers=None; last_exc=None
            # Hanya retry error request/transport. Output invalid TIDAK dikirim ulang.
            for attempt in range(3):
                try:
                    answers=self._generate_batch(rows); break
                except Exception as exc:
                    last_exc=exc
                    # Error permanen (mis. 400/404) tidak dibuang waktu dengan retry.
                    if _gemini_status_code(exc) == 429:
                        # Seluruh pool sudah dicoba oleh _generate_content_with_failover.
                        # Jangan mengulang pool/request yang sama; lanjut provider berikutnya.
                        self.diagnostics.append(
                            f"Gemini: seluruh key terkena 429 untuk batch ini; failover provider."
                        )
                        break
                    if not _gemini_transient_error(exc):
                        self.diagnostics.append(
                            f"Gemini key {self.active_key_number}: request gagal permanen; batch dilewati."
                        )
                        break
                    if attempt < 2:
                        # Backoff pendek: 1.0s lalu 2.0s + jitter deterministik kecil.
                        delay=(1.0 * (2 ** attempt)) + (0.15 * ((seq + attempt) % 4))
                        self.diagnostics.append(
                            f"Gemini key {self.active_key_number}: error sementara; retry {attempt+1}/2 setelah {delay:.2f}s."
                        )
                        time.sleep(delay)
            if answers is None:
                for job in chunk:
                    self._batch_failed_keys.add(job['key'])
                continue
            for iid, job in idmap.items():
                cand=answers.get(iid,'')
                if not self._candidate_ok(job['protected'], cand):
                    self._batch_failed_keys.add(job['key']); continue
                result=cand
                if job['token_map']:
                    result=self.custom_dict._apply_post(result, job['token_map'])
                terms=[]
                if job['italic_map']:
                    idict=self.italic_dict if self.italic_dict is not None else ItalicDictionary()
                    result,terms=idict._apply_post(result, job['italic_map'])
                if not result or _looks_like_error_response(job['original'], result):
                    self._batch_failed_keys.add(job['key']); continue
                _ENGINE8_TRANSLATION_CACHE.put(job['key'], result, terms)

    def translate_one(self, text: str, italic_map: dict = None) -> tuple[str, list[str]]:
        original=text; t=text.strip()
        if not t or _skip_text(t): return text, []
        key=self._cache_key(t, italic_map)
        cached=_ENGINE8_TRANSLATION_CACHE.get(key)
        if cached is not None: return cached
        # Setelah prewarm, key ini berarti sudah memperoleh SATU kesempatan Gemini.
        if key in self._batch_failed_keys:
            self.failed_texts.append(re.sub(r'\s+',' ',original).strip()[:100])
            return original, []
        # Jalur aman untuk pemanggilan di luar pipeline prewarm (mis. judul):
        # tetap satu unit/request, tanpa rescue kedua.
        self.prewarm([(t, italic_map or {})])
        cached=_ENGINE8_TRANSLATION_CACHE.get(key)
        if cached is not None: return cached
        self.failed_texts.append(re.sub(r'\s+',' ',original).strip()[:100])
        return original, []



def _get_deepl_api_key() -> str:
    """Ambil DeepL API key dari environment atau Streamlit Secrets."""
    value = str(os.getenv("DEEPL_API_KEY", "") or os.getenv("DEEPL_AUTH_KEY", "")).strip()
    if value:
        return value
    try:
        import streamlit as st
        for name in ("DEEPL_API_KEY", "DEEPL_AUTH_KEY"):
            try:
                value = str(st.secrets.get(name, "") or "").strip()
                if value:
                    return value
            except Exception:
                pass
    except Exception:
        pass
    return ""


def _deepl_translate(text: str, source: str = 'auto', target: str = 'id') -> str:
    """DeepL API Free/Pro. Endpoint dipilih otomatis dari suffix key :fx."""
    import requests
    key = _get_deepl_api_key()
    if not key:
        raise RuntimeError('DEEPL_API_KEY belum tersedia')
    endpoint = ('https://api-free.deepl.com/v2/translate'
                if key.endswith(':fx') else 'https://api.deepl.com/v2/translate')
    data = {'text': text, 'target_lang': 'ID'}
    if source not in ('auto', '', None):
        data['source_lang'] = 'EN' if source == 'en' else str(source).upper()
    response = requests.post(
        endpoint,
        headers={'Authorization': f'DeepL-Auth-Key {key}'},
        data=data,
        timeout=25,
    )
    if response.status_code >= 400:
        raise RuntimeError(f'DeepL HTTP {response.status_code}')
    payload = response.json()
    rows = payload.get('translations') or []
    if not rows or not rows[0].get('text'):
        raise RuntimeError('DeepL tidak mengembalikan hasil terjemahan')
    return str(rows[0]['text'])


class _RecoveryFallbackTranslator:
    """Pemulihan serial (1 worker): Google -> Gemini key 1..15 -> Argos ->
    MyMemory -> DeepL -> gagal.

    Setiap unit teks melewati provider secara berurutan. Gemini selalu mulai
    dari API key 1; 401/403/429 akan mencoba key berikutnya sampai pool habis.
    Microsoft Translator dan Yandex tidak digunakan.
    """
    def __init__(self, source='auto', target='id', custom_dict=None,
                 italic_dict=None, dictionary_fingerprint=''):
        try:
            from deep_translator import GoogleTranslator
        except ImportError as exc:
            raise ImportError('Jalankan: pip install deep-translator') from exc
        self.source = source; self.target = target
        self.custom_dict = custom_dict; self.italic_dict = italic_dict
        self.dictionary_fingerprint = dictionary_fingerprint or _canonical_dictionary_fingerprint(custom_dict, italic_dict)
        self.failed_texts = []
        self.last_provider = None
        self.last_error_code = None
        self.diagnostics = []
        self._google = GoogleTranslator(source=source, target=target)
        keys = _get_gemini_api_keys()
        self._gemini = (_GeminiRecoveryTranslator(
            source, target, custom_dict, italic_dict,
            self.dictionary_fingerprint, api_keys=keys
        ) if keys else None)

    def translate_one(self, text: str, italic_map: dict = None):
        original = text
        t = text.strip()
        if not t or _skip_text(t):
            return text, []
        italic_map = italic_map or {}
        cache_key = _ENGINE8_TRANSLATION_CACHE.make_key(
            t, self.source, self.target,
            self.dictionary_fingerprint + '|recovery-multilevel-v1', italic_map
        )
        cached = _ENGINE8_TRANSLATION_CACHE.get(cache_key)
        if cached is not None:
            return cached

        token_map = {}
        protected = t
        if self.custom_dict and len(self.custom_dict) > 0:
            protected, token_map = self.custom_dict._apply_pre(protected)
        uncovered = _RE_PROTECTION_TOKEN.sub('', protected)
        if token_map and not re.findall(r'[A-Za-zÀ-ÿ]{2,}', uncovered):
            result = self.custom_dict._apply_post(protected, token_map)
            terms = []
            if italic_map:
                idict = self.italic_dict if self.italic_dict is not None else ItalicDictionary()
                result, terms = idict._apply_post(result, italic_map)
            _ENGINE8_TRANSLATION_CACHE.put(cache_key, result, terms)
            return result, terms

        expected_tokens = {x.casefold() for x in _RE_PROTECTION_TOKEN.findall(protected)}
        result = None
        errors = []

        # 1) Google Translate -- tepat satu request pada tahap pemulihan.
        try:
            candidate = self._google.translate(protected)
            if not _fallback_candidate_ok(protected, candidate, expected_tokens):
                raise RuntimeError('Google tidak lolos quality gate')
            result = candidate; self.last_provider = 'Google Translate'; self.last_error_code = None
        except Exception as exc:
            errors.append(f'Google:{_google_error_code(exc) or type(exc).__name__}')
            self.last_error_code = _google_error_code(exc)

        # 2) Gemini -- mulai key 1 dan failover 429/auth sampai key 15/pool habis.
        if result is None and self._gemini is not None:
            try:
                # Setiap unit pemulihan selalu mulai lagi dari API key 1.
                self._gemini._active_key_index = 0
                self._gemini._client = self._gemini._genai.Client(
                    api_key=self._gemini._api_keys[0]
                )
                # Pakai protected text sehingga token Kamus SNI tetap identik.
                rows = [{'id': 'R000001', 'text': protected}]
                answers = self._gemini._generate_batch(rows)
                candidate = answers.get('R000001', '')
                if not _fallback_candidate_ok(protected, candidate, expected_tokens):
                    raise RuntimeError('Gemini tidak lolos quality gate')
                result = candidate; self.last_provider = f'Gemini Flash API {self._gemini.active_key_number}'; self.last_error_code = None
            except Exception as exc:
                errors.append(f'Gemini:{_gemini_status_code(exc) or type(exc).__name__}')
                self.last_error_code = _gemini_status_code(exc)

        # 3) Argos Translate.
        if result is None:
            try:
                candidate = _argos_translate(protected, self.source, self.target)
                if not _fallback_candidate_ok(protected, candidate, expected_tokens):
                    raise RuntimeError('Argos tidak lolos quality gate')
                result = candidate; self.last_provider = 'Argos Translate'; self.last_error_code = None
            except Exception as exc:
                errors.append(f'Argos:{type(exc).__name__}')

        # 4) MyMemory.
        if result is None:
            try:
                from deep_translator import MyMemoryTranslator
                mm_source = 'en-GB' if self.source in ('auto', 'en') else self.source
                mm_target = 'id-ID' if self.target == 'id' else self.target
                candidate = MyMemoryTranslator(source=mm_source, target=mm_target).translate(protected)
                if not _fallback_candidate_ok(protected, candidate, expected_tokens):
                    raise RuntimeError('MyMemory tidak lolos quality gate')
                result = candidate; self.last_provider = 'MyMemory Translate'; self.last_error_code = None
            except Exception as exc:
                errors.append(f'MyMemory:{_google_error_code(exc) or type(exc).__name__}')

        # 5) DeepL -- provider terakhir sebelum dinyatakan gagal.
        if result is None:
            try:
                candidate = _deepl_translate(protected, self.source, self.target)
                if not _fallback_candidate_ok(protected, candidate, expected_tokens):
                    raise RuntimeError('DeepL tidak lolos quality gate')
                result = candidate; self.last_provider = 'DeepL translator'; self.last_error_code = None
            except Exception as exc:
                errors.append(f'DeepL:{_google_error_code(exc) or type(exc).__name__}')

        if result is None:
            self.failed_texts.append(re.sub(r'\s+', ' ', original).strip()[:100])
            self.diagnostics.append(' -> '.join(errors))
            return original, []

        if token_map:
            result = self.custom_dict._apply_post(result, token_map)
        terms = []
        if italic_map:
            idict = self.italic_dict if self.italic_dict is not None else ItalicDictionary()
            result, terms = idict._apply_post(result, italic_map)
        _ENGINE8_TRANSLATION_CACHE.put(cache_key, result, terms)
        return result, terms

def _gemini_recovery_requests_for_para(para, tr) -> list[tuple[str, dict]]:
    """Mirror read-only dari input yang akan diminta _translate_para ke translator.
    Tidak mengubah DOCX; dipakai untuk micro-batch prewarm Gemini.
    """
    if _skip_paragraph(para) or _has_hyperlinks(para): return []
    text_runs=[(i,r) for i,r in enumerate(para.runs) if r.text and r.text.strip()]
    if not text_runs: return []
    combined=''.join(r.text for _,r in text_runs).strip()
    if not combined or _skip_text(combined): return []
    para_style_italic=_get_para_style_italic(para)
    segs=_segment_runs_by_format(text_runs, para_style_italic)
    if any(_is_protected_segment(t,i,v) for t,i,v in segs):
        req=[]; plain=[]
        def flush():
            if not plain: return
            block=''.join(plain); plain.clear()
            if not block.strip(): return
            if tr.italic_dict and len(tr.italic_dict)>0:
                pre,imap=tr.italic_dict._apply_pre(block)
            else: pre,imap=block,{}
            req.append((pre,imap))
        for seg_text,is_ital,vtype in segs:
            if _is_protected_segment(seg_text,is_ital,vtype): flush()
            else: plain.append(seg_text)
        flush(); return req
    # Tanpa protected segment, _extract_source_format_map tidak menghasilkan
    # token sehingga cukup mirror Kamus Istilah Asing sebelum translate_one.
    if tr.italic_dict and len(tr.italic_dict)>0:
        combined,imap=tr.italic_dict._apply_pre(combined)
    else: imap={}
    return [(combined,imap)]

def _match_capitalization(original: str, translated: str) -> str:
    orig = original.strip(); tran = translated.strip()
    if not orig or not tran: return translated
    letters = [c for c in orig if c.isalpha()]
    if not letters: return translated
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    if upper_ratio >= 0.8: return tran.upper()
    if orig[0].isupper(): return tran[0].upper() + tran[1:] if len(tran) > 1 else tran.upper()
    return tran


_SOURCE_CLAUSE_NUMBER_RE = re.compile(
    r'^(?P<number>(?:\d+(?:\.\d+)*|[A-Z]\.\d+(?:\.\d+)*))(?P<gap>[ \t\u00A0]+)(?P<body>.+)$'
)


def _restore_source_clause_number(source: str, translated: str) -> str:
    """Pertahankan nomor pasal sumber secara literal (termasuk A.1/A.2).

    Mesin terjemahan kadang mengubah ``A.1 General`` menjadi ``1 Umum``.
    Nomor pasal adalah metadata struktur, bukan materi yang boleh diterjemahkan.
    """
    m = _SOURCE_CLAUSE_NUMBER_RE.match(source or '')
    if not m or not translated:
        return translated
    number = m.group('number')
    body = translated.strip()
    # Buang nomor apa pun yang diciptakan/diubah provider di awal hasil.
    body = re.sub(r'^(?:[A-Z]\.)?\d+(?:\.\d+)*[ \t\u00A0]+', '', body).lstrip()
    return f'{number}    {body}' if body else number


_HYPHENATED_DUP_RE = re.compile(r'(?<!\w)([A-Za-zÀ-ÖØ-öø-ÿ]+)\s+-\s+([A-Za-zÀ-ÖØ-öø-ÿ]+)(?!\w)', re.IGNORECASE)


def _normalize_indonesian_reduplication_text(text: str) -> str:
    """Normalisasi kata ulang hasil MT: ``langkah - langkah`` -> ``langkah-langkah``."""
    if not text:
        return text
    return _HYPHENATED_DUP_RE.sub(lambda m: f'{m.group(1)}-{m.group(2)}' if m.group(1).casefold() == m.group(2).casefold() else m.group(0), text)


def _normalize_reduplication_in_para(para) -> None:
    """Hilangkan spasi di sekitar tanda hubung kata ulang tanpa meratakan format run."""
    # Kasus umum hasil Engine 8 berada dalam satu run.
    for run in para.runs:
        if run.text:
            run.text = _normalize_indonesian_reduplication_text(run.text)
    # Fallback bila pola terbelah antar-run: ganti hanya bila seluruh run tidak
    # membawa superscript/subscript sehingga aman direkonstruksi.
    full = para.text or ''
    fixed = _normalize_indonesian_reduplication_text(full)
    if fixed != full and not any(_run_vertalign(r) for r in para.runs):
        if para.runs:
            para.runs[0].text = fixed
            for r in para.runs[1:]:
                r.text = ''


def _restore_plain_block_whitespace(source: str, translated: str) -> str:
    """Pertahankan whitespace batas blok di sekitar run pangkat/indeks.

    Provider lazim melakukan strip pada request per-segmen; tanpa ini hasil
    seperti ``kWm⁻²hingga`` atau ``m²dengan`` dapat terbentuk.
    """
    if not translated:
        return translated
    lead = re.match(r'^[ \t\u00A0]+', source or '')
    trail = re.search(r'[ \t\u00A0]+$', source or '')
    core = translated.strip(' \t\u00A0')
    return (lead.group(0) if lead else '') + core + (trail.group(0) if trail else '')

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
    para_style_italic = _get_para_style_italic(para)

    # Paragraf yang memiliki superscript/subscript (pangkat/indeks) atau
    # istilah/judul miring-sumber (mis. CATATAN/NOTE berisi notasi teknis
    # seperti "kWm⁻²" atau "0,7 m²") memakai jalur PER-SEGMEN: bagian
    # terproteksi tidak pernah dikirim ke layanan terjemahan sama sekali,
    # jauh lebih tahan gagal dibanding menyisipkan token proteksi ke dalam
    # satu string besar. Lihat _translate_paragraph_segments untuk detail.
    preview_segments = _segment_runs_by_format(text_runs, para_style_italic)
    has_protected_segment = any(
        _is_protected_segment(t, i, v) for t, i, v in preview_segments
    )

    if has_protected_segment:
        segmented = _translate_paragraph_segments(text_runs, tr, para_style_italic)
        time.sleep(_TRANSLATE_DELAY)
        if segmented is None:
            return []
        translated, format_spans, italic_terms_found = segmented
        if not translated or translated == original_for_case:
            return []
        translated = _match_capitalization(original_for_case, translated)
    else:
        # Proteksi 1 (jalur lama): istilah/judul asing yang SUDAH miring di
        # dokumen sumber (mis. judul standar acuan pada klausul "Acuan
        # normatif") — tidak diterjemahkan, formatnya dikembalikan persis
        # seperti sumber setelah proses terjemahan selesai.
        combined_with_src_tokens, format_token_map = _extract_source_format_map(text_runs, para_style_italic)
        combined = combined_with_src_tokens.strip()

        # Proteksi 2: kamus istilah asing dari spreadsheet ("Kamus Istilah Asing")
        if tr.italic_dict and len(tr.italic_dict) > 0:
            combined, dict_italic_map = tr.italic_dict._apply_pre(combined)
        else:
            dict_italic_map = {}

        translated, italic_terms_found = tr.translate_one(combined, dict_italic_map)
        time.sleep(_TRANSLATE_DELAY)
        if not translated or translated == combined: return []
        translated = _match_capitalization(original_for_case, translated)

        # Kembalikan token miring-sumber menjadi teks asli, sekaligus
        # dapatkan posisi presisi untuk membangun ulang run.
        translated, format_spans = _detokenize_with_formatting(translated, format_token_map)

    # Nomor pasal/lampiran adalah struktur dan harus identik dengan sumber.
    translated = _restore_source_clause_number(original_for_case, translated)
    translated = _normalize_indonesian_reduplication_text(translated)

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
    _normalize_reduplication_in_para(para)
    
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
                  translate_headers: bool = False, worker_count: int = 2) -> tuple[bool, str]:
        try:
            validate_docx(input_docx)
            # app.py membuat engine tanpa parameter kamus. Karena itu Engine 8
            # wajib mengambil kedua spreadsheet sendiri pada setiap proses,
            # sehingga perubahan Google Sheet langsung dipakai dan bukan hanya
            # ditampilkan sebagai angka pada dashboard.
            if not self._custom_dict_provided:
                self.custom_dict = CustomDictionary()
                sni_count = self.custom_dict.load_from_google_sheet(
                    KAMUS_SPREADSHEET_URL
                )
            else:
                sni_count = len(self.custom_dict)
            if not self._italic_dict_provided:
                self.italic_dict = ItalicDictionary()
                italic_count_loaded = self.italic_dict.load_from_google_sheet(
                    ITALIC_SPREADSHEET_URL
                )
            else:
                italic_count_loaded = len(self.italic_dict)

            intro_entry = self.custom_dict._entries.get('introduction')
            if not intro_entry or intro_entry[1].strip().casefold() != 'pendahuluan':
                raise ValueError(
                    'Kamus SNI tidak memuat pasangan '
                    'Introduction → Pendahuluan.'
                )
            required_italic = {
                'iso online browsing platform', 'iec electropedia'
            }
            missing_italic = sorted(
                term for term in required_italic
                if term not in self.italic_dict._entries
            )
            if missing_italic:
                raise ValueError(
                    'Kamus istilah asing belum memuat: '
                    + ', '.join(missing_italic)
                )

            dictionary_fingerprint = _canonical_dictionary_fingerprint(self.custom_dict, self.italic_dict)

            info = f"{len(self.custom_dict) if self.custom_dict else 0} kamus"
            if self.italic_dict: info += f", {len(self.italic_dict)} miring"
            _notify(
                progress_callback, 2,
                f"Kamus siap: SNI={sni_count}, istilah asing={italic_count_loaded}"
            )
            
            _notify(progress_callback, 5, "Init translator...")
            # Satu instance translator tidak dibagi lintas thread. Setiap worker
            # memiliki client sendiri agar aman dan benar-benar berjalan paralel.
            tr = _Translator(self.source_lang, self.target_lang, self.custom_dict, self.italic_dict, dictionary_fingerprint)
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

            translated_count = 0
            italic_count = 0

            def _translate_job(index_para):
                index, para = index_para
                worker_tr = _Translator(
                    self.source_lang, self.target_lang,
                    self.custom_dict, self.italic_dict, dictionary_fingerprint,
                )
                found = _translate_para(para, worker_tr)
                return (index, para, found, bool(worker_tr.failed_texts),
                        worker_tr.last_provider or 'Google Translate')

            try:
                worker_count = int(worker_count)
            except (TypeError, ValueError):
                worker_count = 2
            worker_count = max(1, min(worker_count, 4))

            failure_times = deque()
            main_failed_attempts = 0

            def _wait_failure_slot():
                # Maksimum lima kegagalan yang dikomit ke dashboard per rolling
                # satu detik. Submission ikut tertahan agar provider tidak dibombardir.
                while True:
                    now = time.monotonic()
                    while failure_times and now - failure_times[0] >= 1.0:
                        failure_times.popleft()
                    if len(failure_times) < 5:
                        return
                    time.sleep(max(0.02, 1.0 - (now - failure_times[0]) + 0.01))

            def _provider_label(provider):
                provider = (provider or 'Google Translate').strip()
                if provider.lower().startswith('google'):
                    return f'Google Translate {worker_count} worker'
                if provider.lower().startswith('gemini flash api'):
                    return provider
                if provider.lower().startswith('argos'):
                    return 'Argos Translate'
                if provider.lower().startswith('mymemory'):
                    return 'MyMemory Translate'
                if provider.lower().startswith('deepl'):
                    return 'DeepL translator'
                return provider

            # -----------------------------------------------------------------
            # TRANSLATE UTAMA BERULANG
            # -----------------------------------------------------------------
            # Satu panggilan _Translator.translate_one() sudah berisi:
            #   Gelombang 1 -> Gelombang 2 -> Gelombang 3.
            # Hanya unit yang masih gagal sesudah Gelombang 3 yang dimasukkan
            # kembali ke putaran berikutnya. Dengan demikian putaran berikutnya
            # benar-benar mulai lagi dari Gelombang 1, tanpa mengulang unit yang
            # sudah sukses/cache-hit.
            #
            # Agar outage permanen tidak membuat Streamlit menggantung selamanya,
            # main dialihkan ke pemulihan setelah satu PUTARAN PENUH tanpa satu
            # pun keberhasilan baru. Selama masih ada kemajuan, 1->2->3 terus
            # berulang sampai seluruh target selesai.
            remaining_main = list(enumerate(translation_queue))
            main_round = 0
            main_stagnant_rounds = 0

            while remaining_main:
                main_round += 1
                round_success = 0
                round_failed = []
                current_provider = f'Google Translate {worker_count} worker'
                items = iter(remaining_main)

                with ThreadPoolExecutor(max_workers=worker_count,
                                        thread_name_prefix='translate') as pool:
                    pending = set()
                    for _ in range(min(worker_count, len(remaining_main))):
                        try:
                            pending.add(pool.submit(_translate_job, next(items)))
                        except StopIteration:
                            break

                    while pending:
                        future = next(as_completed(pending))
                        pending.remove(future)
                        try:
                            index, para, found, failed, provider = future.result()
                        except Exception:
                            # Exception worker tidak boleh menjatuhkan Engine 8.
                            # Item yang tidak diketahui hasilnya diproses ulang pada
                            # putaran utama berikutnya/pemulihan.
                            index = -1
                            para = None
                            found = []
                            failed = True
                            provider = current_provider

                        current_provider = _provider_label(provider)
                        if failed:
                            _wait_failure_slot()
                            failure_times.append(time.monotonic())
                            main_failed_attempts += 1
                            if para is not None:
                                round_failed.append((index, para))
                        else:
                            translated_count += 1
                            round_success += 1
                            italic_count += len(found)

                        # `gagal` pada translate utama adalah jumlah attempt gagal
                        # aktual. Counter dimulai dari 0, hanya naik saat suatu
                        # attempt benar-benar gagal, dan tidak pernah turun ketika
                        # retry berikutnya berhasil. Progress X/Y tetap menghitung
                        # target unik yang sudah berhasil diterjemahkan.
                        _notify(
                            progress_callback,
                            10 + int(translated_count / max(total, 1) * 80),
                            f"[{current_provider} | {translated_count}/{total} | "
                            f"berhasil={translated_count} | gagal={main_failed_attempts}]",
                        )

                        try:
                            pending.add(pool.submit(_translate_job, next(items)))
                        except StopIteration:
                            pass

                # Stabilkan urutan dokumen dan hanya retry item gagal.
                remaining_main = sorted(round_failed, key=lambda item: item[0])
                if not remaining_main:
                    break

                if round_success == 0:
                    main_stagnant_rounds += 1
                else:
                    main_stagnant_rounds = 0

                # Cooldown antar-siklus sebelum kembali ke Gelombang 1.
                # Putaran dengan kemajuan akan terus diulang.
                if main_stagnant_rounds == 0:
                    time.sleep(1.0 + random.uniform(0.10, 0.40))
                    continue

                # Tidak ada kemajuan satu putaran penuh: provider utama sedang
                # benar-benar buntu. Serahkan hanya sisa gagal ke recovery agar
                # pipeline tidak infinite-loop dan Engine 9 tetap dapat dicapai.
                break

            failed_paras = remaining_main

            # -----------------------------------------------------------------
            # PEMULIHAN BERULANG, SERIAL 1 WORKER
            # -----------------------------------------------------------------
            # Recovery hanya menerima unit yang gagal dari translate utama.
            # Setiap putaran mencoba seluruh rantai:
            # Google -> Gemini 1..15 -> Argos -> MyMemory -> DeepL.
            # Yang sukses dikeluarkan permanen; hanya yang masih gagal di-loop.
            # Recovery dibatasi maksimal 10 putaran penuh. Setelah putaran ke-10,
            # sisa yang tetap gagal ditandai merah dan pipeline diteruskan ke
            # tahap berikutnya (Engine 9); tidak ada recovery tanpa batas.
            recovery_total = len(failed_paras)
            recovery_pending = sorted(failed_paras, key=lambda item: item[0])
            recovery_success = 0
            recovery_round = 0
            MAX_RECOVERY_ROUNDS = 10
            recovery_tr = _RecoveryFallbackTranslator(
                self.source_lang, self.target_lang, self.custom_dict,
                self.italic_dict, dictionary_fingerprint
            ) if recovery_total else None

            while recovery_pending and recovery_round < MAX_RECOVERY_ROUNDS:
                recovery_round += 1
                next_pending = []
                round_success = 0

                for index, para in recovery_pending:
                    before = len(recovery_tr.failed_texts)
                    try:
                        found = _translate_para(para, recovery_tr)
                        failed = len(recovery_tr.failed_texts) > before
                    except Exception as exc:
                        found = []
                        failed = True
                        recovery_tr.last_provider = recovery_tr.last_provider or 'Google Translate'
                        recovery_tr.diagnostics.append(
                            f'Unhandled:{type(exc).__name__}:{exc}'
                        )

                    if failed:
                        next_pending.append((index, para))
                    else:
                        round_success += 1
                        recovery_success += 1
                        translated_count += 1
                        italic_count += len(found)

                    recovery_provider = recovery_tr.last_provider or 'Google Translate'
                    unresolved = len(next_pending) + max(
                        0, len(recovery_pending) - round_success - len(next_pending)
                    )
                    pct = 90 + int(recovery_success / max(recovery_total, 1) * 8)
                    _notify(
                        progress_callback, pct,
                        f"[Recovery: {recovery_provider} | "
                        f"{recovery_success}/{recovery_total} | "
                        f"berhasil={recovery_success} | gagal={unresolved}]"
                    )

                recovery_pending = sorted(next_pending, key=lambda item: item[0])
                if not recovery_pending:
                    break

                # Maksimal 10 loop recovery. Bahkan bila satu putaran tidak
                # menghasilkan kemajuan, recovery tetap mendapat kesempatan pada
                # putaran berikutnya karena provider/cooldown dapat sudah pulih.
                # Sesudah loop ke-10, sisa gagal langsung menuju tahap berikutnya.
                if recovery_round >= MAX_RECOVERY_ROUNDS:
                    break

                # Cooldown bertahap antarputaran recovery. Maksimum 30 detik;
                # memberi kesempatan circuit breaker/provider pulih tanpa spam.
                recovery_cooldown = min(30.0, 5.0 * (2 ** min(recovery_round - 1, 3)))
                time.sleep(recovery_cooldown + random.uniform(0.10, 0.50))

            still_failed = [para for _, para in recovery_pending]
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
            atomic_save_docx(doc, output_docx)
            if still_failed:
                return False, (
                    f"{len(still_failed)} bagian masih belum berhasil "
                    "diterjemahkan setelah pemulihan multilevel Google/Gemini/Argos/MyMemory/DeepL. Dokumen parsial "
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
        worker_count: int = 2,
    ) -> tuple[bool, str, str]:
        success, result = self.translate(
            input_docx=input_docx,
            output_docx=output_docx,
            progress_callback=progress_callback,
            translate_headers=translate_headers,
            worker_count=worker_count,
        )
        if success:
            return True, result, "Penerjemahan selektif selesai."
        return False, output_docx, result
