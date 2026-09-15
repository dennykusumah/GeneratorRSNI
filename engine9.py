"""
Engine9: StyleFinalizerEngine
=================================
Engine untuk menambahkan style custom "Judul" dan "Pasal" (sesuai spesifikasi
Modify Style yang dibuat manual di Word oleh pengguna) ke dalam dokumen, lalu
MENERAPKANNYA secara otomatis pada bagian-bagian dokumen yang sesuai.

Spesifikasi style (hasil dari Word "Modify Style"):

  "Judul"  — based on Heading 1, style untuk paragraf berikutnya: Body Text
             Font Arial 12pt Bold, rata tengah (center), outline level 1,
             indentasi kiri/kanan 0, spasi sebelum/sesudah 0pt,
             spasi baris single, TANPA auto-numbering.

  "Pasal"  — based on Heading 1, style untuk paragraf berikutnya: Body Text
             Font Arial 11pt Bold, rata kiri-kanan (justified),
             spasi sebelum/sesudah 0pt, spasi baris single,
             TANPA bullet/numbering Word (numId=0, dinonaktifkan eksplisit
             baik di style maupun override per-paragraf). Nomor Pasal/Subpasal
             TIDAK dibangkitkan oleh auto-numbering Word — nomor yang tampil
             murni mengikuti teks literal apa adanya dari dokumen yang
             diupload user (termasuk untuk subpasal di dalam Lampiran/Annex,
             paragraf ber-style "a2"/"a3", mis. Lampiran B -> B.1, B.2, dst.,
             angkanya tetap berasal dari teks asli, bukan dari list Word).

Penerapan otomatis:

  Style "Judul" -> heading halaman "Daftar Isi" (judul halaman itu sendiri,
      BUKAN entri di dalam daftar isi), "Prakata", "Pendahuluan" (heading-nya,
      BUKAN entri di dalam daftar isi), dan "Bibliografi".

  Style "Pasal" -> semua Pasal (paragraf ber-style "Heading 1") dan Subpasal
      (paragraf ber-style "Heading 2") yang BERBAHASA INDONESIA saja, KECUALI:
        - Sub-subpasal ke bawah (Heading 3, Heading 4, dst.) — tidak disentuh.
        - Subpasal (3.1, 3.2, ...) di bawah Pasal "Istilah dan Definisi" —
          tetap dilewati (supaya tidak membanjiri Daftar Isi/ToC). TAPI
          Pasal "Istilah dan Definisi" itu SENDIRI (mis. Pasal 3) TETAP
          diberi style "Pasal" supaya ikut muncul di Daftar Isi (ToC
          dibangun dari style "Judul" & "Pasal").
        - Seluruh bagian berbahasa Inggris — dilewati sepenuhnya.
      Style "Pasal" JUGA diterapkan pada subpasal tingkat pertama di dalam
      Lampiran/Annex berbahasa Indonesia (paragraf ber-style "a2"), dengan penomoran
      yang mengikuti huruf Lampirannya sendiri (mis. B.1, B.2, ... untuk
      Lampiran B), BUKAN nomor Heading 1/2 dari badan dokumen utama.
      Judul Lampiran itu sendiri (paragraf ber-style "ANNEX", mis.
      "Lampiran B (informatif)") TIDAK disentuh — tetap memakai style
      aslinya.

Bagian lain dokumen (tabel, isi paragraf biasa, cover, entri daftar isi,
header/footer, dst.) TIDAK diubah sama sekali.

Deteksi batas bahasa Indonesia vs Inggris:
  Dokumen hasil pipeline GeneratorRSNI selalu memiliki DUA paragraf ber-style
  "Main Title 1" (judul utama halaman pertama isi dokumen): yang pertama
  adalah judul berbahasa Indonesia (pembuka bagian ID), yang kedua adalah
  judul berbahasa Inggris (pembuka bagian EN, hasil adopsi dua-bahasa).
  Semua heading DI ANTARA kedua "Main Title 1" tsb. dianggap berbahasa
  Indonesia; semua heading SETELAH "Main Title 1" kedua dianggap berbahasa
  Inggris dan dilewati. Jika dokumen hanya punya satu bagian bahasa (tidak
  ditemukan dua "Main Title 1"), seluruh heading diperlakukan sebagai
  berbahasa Indonesia.
"""

from pipeline_utils import validate_docx, atomic_save_docx
import os
import re
import copy
import traceback
import subprocess
import shutil
import tempfile
from typing import Optional
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls

# ─────────────────────────────────────────────────────────────────────────────
# Konstanta
# ─────────────────────────────────────────────────────────────────────────────

# Teks heading Pasal yang SUBPASAL-nya (Heading 2, mis. 3.1, 3.2, ...) harus
# DILEWATI saat menerapkan style "Pasal" (supaya tidak membanjiri Daftar
# Isi/ToC). Pasal itu sendiri (Heading 1, mis. "3  Istilah dan Definisi")
# TETAP diberi style "Pasal" — lihat _collect_pasal_targets().
_SKIP_SUBSECTION_TITLES = {
    'istilah dan definisi',
    'istilah dan definisi umum',
}

# Teks heading halaman yang harus diberi style "Judul".
_JUDUL_TARGET_TEXTS = {'daftar isi', 'prakata', 'pendahuluan', 'bibliografi'}
_ENGINE5_MARKER = 'Engine5DuplicateStart'
_TOC_STYLE_NAME = 'TOC 1'
_TOC_FIELD_INSTR = ' TOC \\h \\z \\t "Judul;1;Pasal;1" '
_TOC_MARK_BOOKMARK = 'Engine9TocField'


def _norm(text: str) -> str:
    """Normalisasi teks untuk dibandingkan: lower-case, trim whitespace/tab."""
    return re.sub(r'\s+', ' ', (text or '').strip()).strip().lower()


def _style_name(paragraph) -> str:
    try:
        return paragraph.style.name if paragraph.style is not None else ''
    except Exception:
        return ''


# ─────────────────────────────────────────────────────────────────────────────
# Generator nomor Pasal/Subpasal
# ─────────────────────────────────────────────────────────────────────────────
# Dokumen sumber (hasil convert dari template ISO) menomori Pasal/Subpasal
# HANYA lewat auto-numbering Word (numId terhubung ke style Heading1/Heading2/
# a2/a3 di numbering.xml) — angka itu TIDAK pernah tersimpan sebagai teks
# literal di XML, hanya dihitung/ditampilkan oleh aplikasi Word saat dibuka.
# python-docx (dan proses translate berbasis teks di Engine9) sama sekali
# tidak melihat angka tsb. Karena numbering Word DIMATIKAN secara eksplisit
# di _apply_pasal (numId=0, supaya tidak dobel dengan render Word lain),
# nomor Pasal/Subpasal HARUS dibangkitkan & ditulis sebagai teks literal di
# sini, mengikuti urutan kemunculan heading yang sama persis dengan yang
# akan dihasilkan Word dari auto-numbering aslinya:
#   - Setiap "Heading 1" -> nomor Pasal urut (1, 2, 3, ...), reset counter
#     Subpasal ("Heading 2") setiap kali masuk Pasal baru.
#   - Setiap "Heading 2" (Subpasal, non-skip) -> "<no Pasal induk>.<urut>"
#     (mis. 4.1, 4.2, ...).
#   - Setiap heading "ANNEX" (Lampiran) -> menaikkan huruf Lampiran
#     (A, B, C, ...), reset counter subpasal Lampiran ('a2'/'a3').
#   - Setiap "a2" (subpasal Lampiran, level 1) -> "<huruf>.<urut>" (mis.
#     A.1, A.2, ...).
#   - "a3" (mis. A.1.1) tidak diberi style Pasal dan tidak masuk TOC.
_RE_LEADING_NUMBER = re.compile(r'^\s*[A-Za-z]?\d+(?:\.\d+)*\.?\s+')


def _strip_existing_leading_number(paragraph) -> None:
    """Buang prefix angka/no. Pasal yang MUNGKIN sudah ada sebagai teks
    literal di run pertama paragraf (mis. jika dokumen sumber kebetulan
    sudah punya nomor manual, atau proses ini dijalankan dua kali),
    supaya nomor baru yang dibangkitkan _insert_number_run() tidak
    dobel/duplikat (mis. "1    1    Ruang lingkup")."""
    runs = paragraph.runs
    if not runs:
        return
    first = runs[0]
    text = first.text or ''
    m = _RE_LEADING_NUMBER.match(text)
    if not m:
        return
    remainder = text[m.end():]
    if remainder:
        first.text = remainder
    else:
        # Run pertama isinya cuma nomor -> hapus run itu supaya tidak
        # menyisakan run kosong di depan.
        r_el = first._element
        parent = r_el.getparent()
        if parent is not None:
            parent.remove(r_el)


def _insert_number_run(paragraph, number_text: str) -> None:
    """Sisipkan run baru berisi '<number_text>    ' (nomor + 4 spasi,
    mengikuti format asli SNI/ISO: "1    Ruang lingkup", "4.1    Umum",
    dst.) sebagai run PALING AWAL di paragraf (persis setelah pPr &
    bookmark, sebelum teks judul Pasal)."""
    p_el = paragraph._p
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = f'{number_text}    '
    r.append(t)
    pPr = p_el.find(qn('w:pPr'))
    insert_pos = list(p_el).index(pPr) + 1 if pPr is not None else 0
    p_el.insert(insert_pos, r)


# ─────────────────────────────────────────────────────────────────────────────
# Style XML — "Judul" & "Pasal"
# ─────────────────────────────────────────────────────────────────────────────

_JUDUL_STYLE_XML = f'''<w:style {nsdecls("w")} w:type="paragraph" w:customStyle="1" w:styleId="Judul">
  <w:name w:val="Judul"/>
  <w:basedOn w:val="Heading1"/>
  <w:next w:val="BodyText"/>
  <w:qFormat/>
  <w:pPr>
    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>
    <w:outlineLvl w:val="0"/>
    <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
    <w:ind w:left="0" w:right="0" w:firstLine="0"/>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
    <w:b/>
    <w:sz w:val="24"/>
    <w:szCs w:val="24"/>
  </w:rPr>
</w:style>'''

_PASAL_STYLE_XML = f'''<w:style {nsdecls("w")} w:type="paragraph" w:customStyle="1" w:styleId="Pasal">
  <w:name w:val="Pasal"/>
  <w:basedOn w:val="Heading1"/>
  <w:next w:val="BodyText"/>
  <w:qFormat/>
  <w:pPr>
    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>
    <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
    <w:jc w:val="both"/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
    <w:b/>
    <w:sz w:val="22"/>
    <w:szCs w:val="22"/>
  </w:rPr>
</w:style>'''


def _find_style_el(doc, style_id):
    for st in doc.styles.element.findall(qn('w:style')):
        if st.get(qn('w:styleId')) == style_id:
            return st
    return None


def _ensure_style(doc, style_id, style_xml):
    """Tambahkan style ke styles.xml jika belum ada. Idempotent."""
    if _find_style_el(doc, style_id) is not None:
        return
    new_style = parse_xml(style_xml)
    doc.styles.element.append(new_style)


def _resolve_num_id(doc, style_id, _depth=0):
    """Telusuri rantai basedOn untuk menemukan w:numId yang berlaku pada style ini."""
    if _depth > 8:
        return None
    st = _find_style_el(doc, style_id)
    if st is None:
        return None
    pPr = st.find(qn('w:pPr'))
    if pPr is not None:
        numPr = pPr.find(qn('w:numPr'))
        if numPr is not None:
            numId_el = numPr.find(qn('w:numId'))
            if numId_el is not None:
                val = numId_el.get(qn('w:val'))
                if val and val != '0':
                    return val
    based = st.find(qn('w:basedOn'))
    if based is not None:
        parent_id = based.get(qn('w:val'))
        if parent_id and parent_id != style_id:
            return _resolve_num_id(doc, parent_id, _depth + 1)
    return None


def _resolve_style_ilvl(doc, style_id, _depth=0):
    """Telusuri rantai basedOn untuk menemukan w:ilvl yang dideklarasikan
    langsung pada style ini (mis. 'a2' -> ilvl 1, 'a3' -> ilvl 2, sesuai
    struktur multilevel list ANNEX di dokumen sumber). Default '0' jika
    tidak ditemukan sama sekali di sepanjang rantai."""
    if _depth > 8:
        return '0'
    st = _find_style_el(doc, style_id)
    if st is None:
        return '0'
    pPr = st.find(qn('w:pPr'))
    if pPr is not None:
        numPr = pPr.find(qn('w:numPr'))
        if numPr is not None:
            ilvl_el = numPr.find(qn('w:ilvl'))
            if ilvl_el is not None:
                val = ilvl_el.get(qn('w:val'))
                if val is not None:
                    return val
    based = st.find(qn('w:basedOn'))
    if based is not None:
        parent_id = based.get(qn('w:val'))
        if parent_id and parent_id != style_id:
            return _resolve_style_ilvl(doc, parent_id, _depth + 1)
    return '0'


def _clear_paragraph_direct_formatting(paragraph, drop_tags):
    """Hapus elemen langsung tertentu dari pPr agar formatting style baru berlaku penuh."""
    p_el = paragraph._p
    pPr = p_el.find(qn('w:pPr'))
    if pPr is None:
        pPr = parse_xml(f'<w:pPr {nsdecls("w")}/>')
        p_el.insert(0, pPr)
    for tag in drop_tags:
        el = pPr.find(qn(tag))
        if el is not None:
            pPr.remove(el)
    return pPr


def _strip_run_direct_formatting(paragraph):
    """Hapus rPr langsung pada tiap run agar teks mengikuti rPr dari style
    paragraf, TAPI pertahankan w:vertAlign (superscript/subscript) supaya
    superscript pada teks (mis. m², catatan kaki angka) tidak ikut hilang
    dan tetap berfungsi normal."""
    for r in paragraph._p.findall(qn('w:r')):
        rPr = r.find(qn('w:rPr'))
        if rPr is None:
            continue
        vert = rPr.find(qn('w:vertAlign'))
        if vert is not None:
            # PENTING: run superscript/subscript berasal dari file input.
            # Jangan menyisakan hanya w:vertAlign karena itu akan membuang
            # font, bahasa, bold/italic, ukuran, character spacing, dll.
            # Properti run asli harus tetap utuh sampai dokumen final.
            continue
        else:
            r.remove(rPr)


def _set_pstyle(paragraph, style_id):
    p_el = paragraph._p
    pPr = p_el.find(qn('w:pPr'))
    if pPr is None:
        pPr = parse_xml(f'<w:pPr {nsdecls("w")}/>')
        p_el.insert(0, pPr)
    pStyle = pPr.find(qn('w:pStyle'))
    if pStyle is None:
        pStyle = parse_xml(f'<w:pStyle {nsdecls("w")} w:val="{style_id}"/>')
        pPr.insert(0, pStyle)
    else:
        pStyle.set(qn('w:val'), style_id)


def _apply_judul(doc, paragraph, force_page_break_before=False):
    _set_pstyle(paragraph, 'Judul')
    pPr = _clear_paragraph_direct_formatting(
        paragraph, ['w:jc', 'w:spacing', 'w:ind', 'w:outlineLvl', 'w:numPr']
    )
    if force_page_break_before and pPr.find(qn('w:pageBreakBefore')) is None:
        pPr.append(parse_xml(f'<w:pageBreakBefore {nsdecls("w")}/>'))
    _strip_run_direct_formatting(paragraph)


def _apply_pasal(doc, paragraph, ilvl, num_id, number_text=None):
    _set_pstyle(paragraph, 'Pasal')
    pPr = _clear_paragraph_direct_formatting(
        paragraph, ['w:jc', 'w:spacing', 'w:ind', 'w:numPr']
    )
    # Nonaktifkan bullet/numbering Word (numId=0) secara eksplisit per-paragraf.
    # Nomor Pasal/Subpasal TIDAK lagi dibangkitkan oleh Word auto-numbering —
    # sebagai gantinya, nomor yang BENAR (dihitung oleh _collect_pasal_targets
    # berdasarkan urutan & level heading, lihat komentar di atas
    # _RE_LEADING_NUMBER) dituliskan di sini sebagai teks literal, PERSIS
    # meniru hasil auto-numbering Word yang asli (mis. "1", "4.2", "A.1").
    numPr_xml = (
        f'<w:numPr {nsdecls("w")}>'
        f'<w:ilvl w:val="0"/><w:numId w:val="0"/>'
        f'</w:numPr>'
    )
    pPr.insert(0, parse_xml(numPr_xml))

    if number_text:
        # Buang dulu nomor lama (jika ada, mis. sudah literal di sumber atau
        # sisa dari proses sebelumnya) supaya tidak dobel, baru sisipkan
        # nomor final yang benar di depan judul Pasal.
        _strip_existing_leading_number(paragraph)
        _insert_number_run(paragraph, number_text)

    _strip_run_direct_formatting(paragraph)


# ─────────────────────────────────────────────────────────────────────────────
# Engine utama
# ─────────────────────────────────────────────────────────────────────────────

def _split_run_and_italicize(run, term: str) -> bool:
    """Pecah SATU run yang teksnya mengandung `term` di tengah/tepi teks lain
    (bukan hanya sama persis) menjadi hingga 3 run: [sebelum][term][sesudah],
    dengan run [term] dipaksa italic, sedangkan run sebelum/sesudah TETAP
    memakai formatting asli run tsb (font, bold, size, dst — hanya teksnya
    yang dipotong, tidak ada properti lain yang berubah).

    Dirancang seaman mungkin: hanya menangani run dengan struktur sederhana
    (satu elemen <w:t> per run, kasus normal untuk teks paragraf biasa).
    Jika run punya struktur lain yang tidak terduga (mis. tab/break/drawing
    di dalamnya), fungsi ini TIDAK melakukan apa-apa (aman, tidak mengubah
    apa pun) — dikembalikan False supaya caller tahu belum tertangani.

    Return True jika berhasil memecah & meng-italic-kan (atau run memang
    sudah persis == term, tinggal di-italic-kan langsung tanpa split).
    """
    text = run.text
    if not text or term not in text:
        return False

    r_el = run._element
    t_els = r_el.findall(qn('w:t'))
    # Hanya tangani run dengan tepat satu <w:t> (kasus umum). Run dengan
    # elemen lain (tab, br, drawing, dst.) atau >1 <w:t> dilewati demi
    # keamanan struktur dokumen.
    other_children = [c for c in r_el if c.tag != qn('w:rPr') and c.tag != qn('w:t')]
    if len(t_els) != 1 or other_children:
        return False

    if text == term:
        run.font.italic = True
        return True

    idx = text.find(term)
    if idx == -1:
        return False
    before, after = text[:idx], text[idx + len(term):]

    parent = r_el.getparent()
    if parent is None:
        return False
    pos = list(parent).index(r_el)

    def _make_run(piece_text, force_italic):
        new_r = copy.deepcopy(r_el)
        new_t = new_r.find(qn('w:t'))
        new_t.text = piece_text
        new_t.set(qn('xml:space'), 'preserve')
        if force_italic:
            rPr = new_r.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                new_r.insert(0, rPr)
            if rPr.find(qn('w:i')) is None:
                rPr.append(OxmlElement('w:i'))
        return new_r

    new_runs = []
    if before:
        new_runs.append(_make_run(before, force_italic=False))
    new_runs.append(_make_run(term, force_italic=True))
    if after:
        new_runs.append(_make_run(after, force_italic=False))

    for offset, nr in enumerate(new_runs):
        parent.insert(pos + offset, nr)
    parent.remove(r_el)
    return True


def _enforce_italic_terms(doc, terms: list[str]) -> None:
    """Jaring pengaman terakhir: paksa teks yang PERSIS sama dengan salah
    satu `terms` (mis. "Red Green Blue") agar SELALU tampil italic, apa pun
    yang terjadi di tahap-tahap sebelumnya (terjemahan, dsb). Dijalankan
    paling akhir (sebelum doc.save) di StyleFinalizerEngine supaya jadi
    jaminan final.

    Menangani DUA kasus:
      1. Run yang teksnya PERSIS == term -> langsung di-italic-kan (cepat,
         seperti semula).
      2. Term muncul sebagai BAGIAN dari run yang lebih besar (mis. hasil
         penggabungan run oleh proses lain) -> run tsb dipecah supaya
         hanya bagian term-nya yang italic, sisanya tetap memakai
         formatting asli tanpa berubah.
    Tidak mengubah paragraf/run lain yang tidak mengandung salah satu term.
    """
    for para in doc.paragraphs:
        for run in list(para.runs):
            text = run.text
            if not text:
                continue
            for term in terms:
                if term in text:
                    _split_run_and_italicize(run, term)
                    break


def _enforce_empty_daftar_isi_page(doc) -> None:
    """Jaring pengaman terakhir: pastikan halaman "Daftar Isi" SELALU kosong
    (tanpa entri/teks apa pun di bawah judulnya), apa pun yang terjadi di
    tahap-tahap sebelumnya (mis. field TOC yang ter-update otomatis, atau
    proses lain yang tanpa sengaja menambahkan teks). Dijalankan paling
    akhir (sebelum doc.save) di StyleFinalizerEngine.

    Cakupan: dari paragraf TEPAT SETELAH judul "Daftar Isi" (kemunculan
    pertama di seluruh dokumen), berhenti pada batas TERDEKAT dari:
      1. paragraf "Prakata" atau "Pendahuluan" (heading halaman
         berikutnya — PENTING: di pipeline ini, DI + Prakata +
         Pendahuluan berbagi SATU section/sectPr yang sama, letaknya
         jauh di akhir Pendahuluan — bukan tepat setelah Daftar Isi.
         Maka deteksi sectPr SAJA tidak cukup & pernah salah menghapus
         seluruh isi Prakata/Pendahuluan; heading berikutnya harus jadi
         batas utama).
      2. paragraf ber-style Heading (badan dokumen utama sudah mulai).
      3. paragraf yang memuat <w:sectPr> (akhir section, jaga-jaga bila
         Prakata/Pendahuluan belum ada sama sekali).
      4. batas pengaman keras (_MAX_SCAN paragraf) supaya proses TIDAK
         PERNAH menyapu ke bagian dokumen yang jauh meski 1–3 gagal
         terdeteksi.
    Hanya TEKS run yang dihapus (run/paragraf itu sendiri tidak
    dihapus), supaya format/struktur dokumen tidak berubah.
    """
    _MAX_SCAN = 8   # lebih dari cukup: DI hanya berisi judul + 3 paragraf kosong
    _BOUNDARY_TITLES = {'prakata', 'pendahuluan'}

    paras = doc.paragraphs
    di_idx = None
    for i, p in enumerate(paras):
        if _norm(p.text) == 'daftar isi':
            di_idx = i
            break
    if di_idx is None:
        return

    scanned = 0
    for p in paras[di_idx + 1:]:
        if scanned >= _MAX_SCAN:
            break
        scanned += 1

        norm_text = _norm(p.text)
        if norm_text in _BOUNDARY_TITLES:
            break
        style_name = (p.style.name or '') if p.style is not None else ''
        if style_name.lower().startswith('heading'):
            break

        pPr = p._p.find(qn('w:pPr'))
        has_sectpr = pPr is not None and pPr.find(qn('w:sectPr')) is not None
        for run in p.runs:
            run.text = ''
        if has_sectpr:
            break


class StyleFinalizerEngine:
    """
    Menambahkan style "Judul" & "Pasal" ke dokumen lalu menerapkannya secara
    otomatis pada heading halaman (Judul) dan Pasal/Subpasal berbahasa
    Indonesia (Pasal), tanpa mengubah bagian dokumen lainnya.
    """

    def apply(self, input_docx: str, output_docx: str) -> tuple[bool, str]:
        try:
            doc = Document(input_docx)

            # 1) Pastikan style "Judul" & "Pasal" tersedia di styles.xml
            _ensure_style(doc, 'Judul', _JUDUL_STYLE_XML)
            _ensure_style(doc, 'Pasal', _PASAL_STYLE_XML)

            paras = doc.paragraphs
            n = len(paras)

            def style_name(p):
                return p.style.name if p.style is not None else ''

            # 2) Petakan nomor section (1-based) dan cari marker Engine5.
            # Marker adalah batas mutlak: semua pasal/subpasal setelah marker
            # merupakan duplikasi bahasa Inggris dan TIDAK boleh disentuh.
            section_by_idx = []
            section_no = 1
            marker_idx = n
            for i, paragraph in enumerate(paras):
                section_by_idx.append(section_no)
                for bm in paragraph._p.iter(qn('w:bookmarkStart')):
                    if bm.get(qn('w:name')) == _ENGINE5_MARKER:
                        marker_idx = min(marker_idx, i)
                pPr = paragraph._p.find(qn('w:pPr'))
                if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                    section_no += 1

            # 3) Kumpulkan target style "Judul" sesuai section.
            judul_targets = {}  # idx -> force_page_break_before
            annex_indices = set(
                i for i in range(n)
                if style_name(paras[i]) == 'ANNEX'
                or re.match(r'^\s*(?:lampiran|annex)\s+[A-Z0-9]', paras[i].text or '', re.I)
            )
            for i, paragraph in enumerate(paras):
                folded = _norm(paragraph.text)
                sec = section_by_idx[i]
                # Section 3: hanya heading halaman front matter.
                if sec == 3 and folded in {'daftar isi', 'prakata', 'pendahuluan'}:
                    judul_targets[i] = False
                # Content Indonesia: hanya judul Lampiran sebelum marker.
                elif sec >= 4 and i < marker_idx and re.match(r'^lampiran\s+[a-z0-9]', folded):
                    judul_targets[i] = True
                # Bibliografi adalah pengecualian yang berada sesudah salinan
                # Engine5 tetapi tetap harus masuk style Judul/TOC.
                elif sec >= 4 and folded == 'bibliografi':
                    judul_targets[i] = style_name(paragraph) == 'Biblio Title'

            for idx, force_pb in judul_targets.items():
                _apply_judul(doc, paras[idx], force_page_break_before=force_pb)

            # 4) Kumpulkan target style "Pasal": seluruh Pasal/Subpasal pada
            # Content Indonesia sebelum marker, termasuk a2 di Lampiran.
            heading1_num_id = _resolve_num_id(doc, 'Heading1')
            heading2_num_id = _resolve_num_id(doc, 'Heading2') or heading1_num_id

            # Subpasal di dalam Lampiran/Annex ('a2'/'a3'): numId & ilvl
            # style-nya SENDIRI (bukan Heading1/2) hanya dipakai sebagai
            # PENANDA bahwa paragraf tsb memang bagian dari list Lampiran
            # (untuk menentukan target). Nomor yang tampil TIDAK lagi
            # dibangkitkan dari numId ini — auto-numbering Word dinonaktifkan
            # di _apply_pasal, nomor mengikuti teks literal dokumen asli
            # (mis. B.1, B.2, ... untuk Lampiran B, apa adanya dari teks).
            annex_a2_num_id = _resolve_num_id(doc, 'a2')
            annex_a2_ilvl = _resolve_style_ilvl(doc, 'a2')

            pasal_targets = []  # (idx, ilvl, num_id, number_text)
            # Counter untuk membangkitkan nomor Pasal/Subpasal (lihat
            # penjelasan lengkap di komentar atas _RE_LEADING_NUMBER).
            h1_counter = 0          # nomor Pasal (Heading 1) berjalan
            h2_counter = 0          # nomor Subpasal (Heading 2), reset tiap Pasal baru
            annex_letter = None       # huruf Lampiran/Annex aktif, mis. A, B, C
            annex_sub_counter = 0      # nomor 'a2' (mis. B.1, B.2, ...), reset tiap Annex baru
            for i in range(n):
                sname = style_name(paras[i])
                if i in annex_indices:
                    # Ambil huruf dari judul aktual, BUKAN dari counter global.
                    # Dengan demikian:
                    #   Lampiran A -> A.1, A.2, ...
                    #   Lampiran B -> B.1, B.2, ...
                    #   Annex A    -> A.1, A.2, ...
                    #   Annex B    -> B.1, B.2, ...
                    # dan bagian Inggris tidak berubah menjadi C/D hanya karena
                    # bagian Indonesia sudah memiliki Annex A/B yang sama.
                    m_annex = re.search(
                        r'\b(?:Lampiran|Annex)\s+([A-Z])(?:\b|\s|$)',
                        paras[i].text.strip(),
                        flags=re.IGNORECASE
                    )
                    annex_letter = m_annex.group(1).upper() if m_annex else annex_letter
                    annex_sub_counter = 0
                    continue

                # Hanya section Content Indonesia dan selalu sebelum marker.
                if section_by_idx[i] < 4 or i >= marker_idx:
                    continue
                if sname in ('@Pasal', 'Pasal'):
                    # Migrasi style lama. Pertahankan nomor literal yang telah
                    # dibuat Engine7/8, termasuk A.1/A.1.1 pada Lampiran.
                    m_num = re.match(r'^\s*([A-Z]?\d+(?:\.\d+)*)\.?\s+', paras[i].text or '', re.I)
                    if m_num:
                        number = m_num.group(1).upper()
                        level = number.count('.')
                        if level <= 1:
                            pasal_targets.append((i, level, None, number))
                        else:
                            # Bersihkan hasil Engine9 lama/idempotent: Pasal
                            # tingkat 3 tidak boleh tetap memakai style Pasal,
                            # karena field TOC Word mengambil style tersebut.
                            _set_pstyle(paras[i], 'a3' if number[0].isalpha() else 'Heading3')
                    else:
                        pasal_targets.append((i, 0, None, None))
                elif sname == 'Heading 1':
                    h1_counter += 1
                    h2_counter = 0
                    pasal_targets.append((i, 0, heading1_num_id, str(h1_counter)))
                elif sname == 'Heading 2':
                    h2_counter += 1
                    pasal_targets.append((i, 1, heading2_num_id, f'{h1_counter}.{h2_counter}'))
                elif sname.lower() == 'a2' and annex_letter:
                    annex_sub_counter += 1
                    pasal_targets.append(
                        (i, annex_a2_ilvl, annex_a2_num_id,
                         f'{annex_letter}.{annex_sub_counter}')
                    )
                # a3/Heading 3+ (mis. A.1.1/1.1.1) sengaja tidak disentuh
                # Judul Lampiran sendiri (style "ANNEX") sengaja tidak disentuh

            for idx, ilvl, num_id, number_text in pasal_targets:
                _apply_pasal(doc, paras[idx], ilvl, num_id, number_text)

            # 6) Jaring pengaman terakhir: pastikan "Red Green Blue" pada
            #    Prakata SELALU tercetak italic — jaminan final sebelum
            #    dokumen disimpan.
            #    CATATAN: _enforce_empty_daftar_isi_page() SENGAJA TIDAK
            #    dipanggil lagi di sini. Engine5 (DaftarIsiEngine) kini
            #    menyisipkan field TOC asli Word di halaman Daftar Isi
            #    (lihat engine5.py) supaya daftar isi terisi otomatis —
            #    memanggil fungsi ini akan menghapus kembali isi field
            #    tersebut sehingga halaman Daftar Isi kosong lagi.
            #    Fungsi itu sendiri dibiarkan ada (tidak dihapus) di bawah
            #    supaya tidak mengubah bagian lain dari engine ini.
            _enforce_italic_terms(doc, ['Red Green Blue'])

            atomic_save_docx(doc, output_docx)
            msg = (
                f'OK: {len(judul_targets)} heading/Lampiran -> "Judul", '
                f'{len(pasal_targets)} pasal/subpasal -> "Pasal".'
            )
            return True, msg

        except Exception as e:
            import traceback
            return False, f'StyleFinalizerEngine Error: {str(e)}\n{traceback.format_exc()}'


def apply_custom_styles(input_docx: str, output_docx: str) -> tuple[bool, str]:
    """Shortcut fungsi-level untuk StyleFinalizerEngine().apply(...)."""
    return StyleFinalizerEngine().apply(input_docx, output_docx)


class TableOfContentsEngine:
    """Engine 9 final: buat style Judul/Pasal lalu sisipkan TOC Word."""

    @staticmethod
    def _ensure_toc_style(doc: Document):
        style = None
        for candidate in (_TOC_STYLE_NAME, 'toc 1'):
            try:
                style = doc.styles[candidate]
                break
            except KeyError:
                pass
        if style is None:
            style = doc.styles.add_style(_TOC_STYLE_NAME, WD_STYLE_TYPE.PARAGRAPH)
        try:
            style.base_style = doc.styles['Normal']
            style.next_paragraph_style = doc.styles['Normal']
        except KeyError:
            pass

        # Bangun ulang pPr/rPr secara deterministik agar properti lama dari
        # template (bold, indent, alignment, space-before) tidak tersisa.
        style_el = style.element
        priority = style_el.find(qn('w:uiPriority'))
        if priority is None:
            priority = OxmlElement('w:uiPriority')
            style_el.insert(3, priority)
        priority.set(qn('w:val'), '39')

        pPr = style_el.get_or_add_pPr()
        for child in list(pPr):
            pPr.remove(child)
        tabs = OxmlElement('w:tabs')
        left_tab = OxmlElement('w:tab')
        left_tab.set(qn('w:val'), 'left')
        left_tab.set(qn('w:pos'), '720')       # 1,27 cm
        tabs.append(left_tab)
        right_tab = OxmlElement('w:tab')
        right_tab.set(qn('w:val'), 'right')
        right_tab.set(qn('w:leader'), 'dot')
        right_tab.set(qn('w:pos'), '9752')     # 17,2 cm
        tabs.append(right_tab)
        pPr.append(tabs)
        pPr.append(OxmlElement('w:suppressAutoHyphens'))
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:after'), '120')      # 6 pt
        spacing.set(qn('w:line'), '240')       # single
        spacing.set(qn('w:lineRule'), 'auto')
        pPr.append(spacing)

        rPr = style_el.get_or_add_rPr()
        for child in list(rPr):
            rPr.remove(child)
        fonts = OxmlElement('w:rFonts')
        fonts.set(qn('w:ascii'), 'Arial')
        fonts.set(qn('w:hAnsi'), 'Arial')
        rPr.append(fonts)
        return style

    @staticmethod
    def _add_fld_char(run, kind: str, dirty: bool = False):
        fld = OxmlElement('w:fldChar')
        fld.set(qn('w:fldCharType'), kind)
        if dirty:
            fld.set(qn('w:dirty'), 'true')
        run._r.append(fld)

    @staticmethod
    def _remove_existing_toc(doc: Document):
        for bookmark in list(doc.element.body.iter(qn('w:bookmarkStart'))):
            if bookmark.get(qn('w:name')) != _TOC_MARK_BOOKMARK:
                continue
            node = bookmark
            while node is not None and node.tag != qn('w:p'):
                node = node.getparent()
            if node is not None and node.getparent() is not None:
                node.getparent().remove(node)
            break

    @staticmethod
    def _clear_cached_toc_entries(title: Paragraph):
        """Hapus hasil TOC lama di antara judul dan page break/Prakata.

        Saat field TOC pernah diperbarui Word, hasilnya dapat menjadi banyak
        paragraf ``TOC 1``. Menghapus paragraf field bertanda bookmark saja
        tidak cukup dan dapat mendorong TOC baru ke halaman berikutnya.
        """
        node = title._p.getnext()
        while node is not None:
            next_node = node.getnext()
            if node.tag in (qn('w:bookmarkStart'), qn('w:bookmarkEnd')):
                # Beberapa Word/LibreOffice menyimpan bookmark TOC sebagai
                # sibling paragraf, bukan di dalam paragraf field.
                bookmark_name = node.get(qn('w:name')) or ''
                if node.tag == qn('w:bookmarkEnd') or bookmark_name == _TOC_MARK_BOOKMARK:
                    node.getparent().remove(node)
                    node = next_node
                    continue
                break
            # Word lazim membungkus seluruh hasil TOC dalam w:sdt. Hapus
            # wrapper itu sebagai satu unit agar semua cache entri ikut hilang.
            if node.tag == qn('w:sdt'):
                has_toc_field = any(
                    ' TOC ' in f' {instr.text or ""} '
                    for instr in node.iter(qn('w:instrText'))
                )
                toc_styles = [
                    style.get(qn('w:val')) or ''
                    for style in node.iter(qn('w:pStyle'))
                ]
                if has_toc_field or any(
                    value.lower().replace(' ', '') == 'toc1'
                    for value in toc_styles
                ):
                    node.getparent().remove(node)
                    node = next_node
                    continue
                break
            if node.tag != qn('w:p'):
                break
            text = ''.join(t.text or '' for t in node.iter(qn('w:t'))).strip()
            pPr = node.find(qn('w:pPr'))
            has_page_break = bool(
                node.xpath('.//w:br[@w:type="page"] | ./w:pPr/w:pageBreakBefore')
            )
            has_sectpr = pPr is not None and pPr.find(qn('w:sectPr')) is not None
            if _norm(text) in {'prakata', 'pendahuluan'} or has_page_break or has_sectpr:
                break
            pstyle = pPr.find(qn('w:pStyle')) if pPr is not None else None
            style_id = pstyle.get(qn('w:val')) if pstyle is not None else ''
            is_toc = style_id.lower().replace(' ', '') == 'toc1'
            has_toc_field = any(
                ' TOC ' in f' {instr.text or ""} '
                for instr in node.iter(qn('w:instrText'))
            )
            if not text or is_toc or has_toc_field:
                node.getparent().remove(node)
            else:
                break
            node = next_node

    @staticmethod
    def _force_update_fields(doc: Document):
        settings = doc.settings.element
        node = settings.find(qn('w:updateFields'))
        if node is None:
            node = OxmlElement('w:updateFields')
            settings.append(node)
        node.set(qn('w:val'), 'true')

    @staticmethod
    def _disable_update_fields_after_finalization(docx_path: str) -> None:
        """Jangan biarkan Word menghitung ulang semua field saat file dibuka.

        Engine9 sendiri sudah membuka dokumen di Word, melakukan repaginasi,
        membangun TOC, dan menyimpan nomor halaman final. Bila
        ``w:updateFields`` tetap ``true``, Word dapat mengulang pembaruan saat
        dokumen baru dibuka (sebelum pagination tampilan stabil) dan menimpa
        hasil tersimpan menjadi i/1. PAGE field di footer tetap bekerja tanpa
        opsi global ini.
        """
        doc = Document(docx_path)
        settings = doc.settings.element
        node = settings.find(qn('w:updateFields'))
        if node is not None:
            node.set(qn('w:val'), 'false')
        atomic_save_docx(doc, docx_path)
    @staticmethod
    def _collect_entries(doc: Document, toc_title: Paragraph):
        entries = []
        for paragraph in doc.paragraphs:
            if paragraph._p is toc_title._p:
                continue
            if _style_name(paragraph) not in ('Judul', 'Pasal'):
                continue
            text = re.sub(r'\s+', ' ', paragraph.text or '').strip()
            if _style_name(paragraph) == 'Pasal':
                match = re.match(r'^([A-Z]?\d+(?:\.\d+)*)\b', text, re.I)
                # Pertahanan berlapis untuk dokumen lama: meskipun style
                # Pasal tingkat 3 masih tertinggal, jangan masukkan 1.1.1,
                # A.1.1, dan level yang lebih dalam ke TOC.
                if match and match.group(1).count('.') > 1:
                    continue
            if text:
                entries.append((text, '…'))
        return entries

    def _build_toc_paragraph(self, doc: Document):
        paragraph = doc.add_paragraph()
        paragraph.style = self._ensure_toc_style(doc)
        self._add_fld_char(paragraph.add_run(), 'begin', dirty=True)
        instr_run = paragraph.add_run()
        instr = OxmlElement('w:instrText')
        instr.set(qn('xml:space'), 'preserve')
        instr.text = _TOC_FIELD_INSTR
        instr_run._r.append(instr)
        self._add_fld_char(paragraph.add_run(), 'separate')
        self._add_fld_char(paragraph.add_run(), 'end')
        start = OxmlElement('w:bookmarkStart')
        start.set(qn('w:id'), '9001')
        start.set(qn('w:name'), _TOC_MARK_BOOKMARK)
        end = OxmlElement('w:bookmarkEnd')
        end.set(qn('w:id'), '9001')
        pPr = paragraph._p.find(qn('w:pPr'))
        paragraph._p.insert(1 if pPr is not None else 0, start)
        paragraph._p.append(end)
        return paragraph

    @staticmethod
    def _toc_anchor_after_three_blank_paragraphs(title: Paragraph):
        """Letakkan TOC sesudah tepat tiga paragraf kosong setelah judul.

        Template CNT sudah menyediakan tiga paragraf kosong. Jika salah satu
        tidak ada, tambahkan hanya yang kurang; konten lain tidak disentuh.
        """
        # Judul dan tiga blank harus tetap bersama dengan awal TOC pada
        # halaman yang sama. Blank dibuat ulang secara deterministik.
        title_pPr = title._p.get_or_add_pPr()
        for tag in ('w:pageBreakBefore',):
            old = title_pPr.find(qn(tag))
            if old is not None:
                title_pPr.remove(old)
        if title_pPr.find(qn('w:keepNext')) is None:
            title_pPr.append(OxmlElement('w:keepNext'))

        anchor = title._p
        blank_count = 0
        while blank_count < 3:
            blank = OxmlElement('w:p')
            pPr = OxmlElement('w:pPr')
            pPr.append(OxmlElement('w:keepNext'))
            spacing = OxmlElement('w:spacing')
            spacing.set(qn('w:before'), '0')
            spacing.set(qn('w:after'), '0')
            spacing.set(qn('w:line'), '240')
            spacing.set(qn('w:lineRule'), 'auto')
            pPr.append(spacing)
            blank.append(pPr)
            anchor.addnext(blank)
            anchor = blank
            blank_count += 1
        return anchor

    @staticmethod
    def _update_toc_page_numbers_with_word(docx_path: str) -> None:
        """Paksa Microsoft Word menghitung pagination lalu menjalankan
        *Update page numbers only* pada seluruh TOC sebelum file dikembalikan.

        ``w:updateFields`` hanya meminta Word memperbarui field saat dokumen
        dibuka. Karena itu Engine9 harus benar-benar membuka dokumen melalui
        Word, melakukan Repaginate(), UpdatePageNumbers(), lalu Save/Close.
        """
        if not docx_path or not os.path.isfile(docx_path):
            raise FileNotFoundError(f'File untuk update TOC tidak ditemukan: {docx_path}')
        if os.name != 'nt':
            TableOfContentsEngine._update_toc_page_numbers_on_linux(docx_path)
            return

        abs_path = os.path.abspath(docx_path)
        pywin32_error = None

        # Metode utama: Word COM via pywin32.
        try:
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            word = None
            doc = None
            try:
                word = win32com.client.DispatchEx('Word.Application')
                word.Visible = False
                word.DisplayAlerts = 0
                try:
                    word.Options.Pagination = True
                except Exception:
                    pass
                try:
                    word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
                except Exception:
                    pass

                doc = word.Documents.Open(
                    abs_path,
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Visible=False,
                    OpenAndRepair=True,
                    NoEncodingDialog=True,
                )
                try:
                    doc.Activate()
                    word.ActiveWindow.View.Type = 3  # wdPrintView
                    word.ActiveWindow.View.ShowFieldCodes = False
                except Exception:
                    pass
                doc.Repaginate()

                toc_count = int(doc.TablesOfContents.Count)
                if toc_count < 1:
                    # Field TOC mentah kadang belum masuk koleksi TablesOfContents
                    # sebelum field tersebut dimaterialisasi Word.
                    for i in range(1, int(doc.Fields.Count) + 1):
                        fld = doc.Fields.Item(i)
                        try:
                            code = str(fld.Code.Text or '').strip().upper()
                        except Exception:
                            code = ''
                        if code == 'TOC' or code.startswith('TOC '):
                            fld.Update()
                    doc.Repaginate()
                    toc_count = int(doc.TablesOfContents.Count)

                if toc_count < 1:
                    raise RuntimeError('Microsoft Word tidak menemukan objek Table of Contents.')

                # WAJIB materialisasikan ulang seluruh entry. Range.Text pada
                # field TOC baru tidak selalu kosong (dapat berisi karakter
                # field/bookmark), sehingga pengujian ``if not result_text``
                # sebelumnya salah melewati Update(). Akibatnya Word hanya
                # memperbarui cache PAGEREF lama dan semua entry menjadi i/1.
                for i in range(1, toc_count + 1):
                    toc = doc.TablesOfContents.Item(i)
                    try:
                        toc.Range.Fields.Locked = False
                    except Exception:
                        pass
                    toc.Update()

                # Pembentukan TOC dapat menggeser isi beberapa halaman. Ulangi
                # pagination + UpdatePageNumbers sampai hasil stabil.
                previous = None
                for _ in range(4):
                    doc.Repaginate()
                    for i in range(1, int(doc.TablesOfContents.Count) + 1):
                        doc.TablesOfContents.Item(i).UpdatePageNumbers()
                    current = tuple(
                        str(doc.TablesOfContents.Item(i).Range.Text or '')
                        for i in range(1, int(doc.TablesOfContents.Count) + 1)
                    )
                    if current == previous:
                        break
                    previous = current

                # Jangan pernah mengembalikan file bila Word ternyata masih
                # menyimpan nomor dummy yang sama untuk TOC panjang.
                for i in range(1, int(doc.TablesOfContents.Count) + 1):
                    lines = [x.strip() for x in str(doc.TablesOfContents.Item(i).Range.Text or '').splitlines() if x.strip()]
                    page_values = []
                    for line in lines:
                        match = re.search(r'(?:\t|\s)([ivxlcdm]+|\d+)\s*$', line, re.I)
                        if match:
                            page_values.append(match.group(1).lower())
                    if len(page_values) >= 8 and len(set(page_values)) < 2:
                        raise RuntimeError('Validasi TOC gagal: semua nomor halaman masih sama.')

                doc.Save()
                return
            finally:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=True)
                    except Exception:
                        pass
                if word is not None:
                    try:
                        word.Quit()
                    except Exception:
                        pass
                pythoncom.CoUninitialize()
        except Exception as exc:
            pywin32_error = exc

        # Fallback: PowerShell COM. Tidak membutuhkan paket pywin32, tetapi
        # tetap membutuhkan Windows dan Microsoft Word terpasang.
        escaped = abs_path.replace("'", "''")
        ps_script = rf'''
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try {{ $word.Options.Pagination = $true }} catch {{}}
    try {{ $word.AutomationSecurity = 3 }} catch {{}}

    $doc = $word.Documents.Open('{escaped}')
    try {{
        $doc.Activate()
        $word.ActiveWindow.View.Type = 3
        $word.ActiveWindow.View.ShowFieldCodes = $false
    }} catch {{}}
    $doc.Repaginate()

    $tocCount = $doc.TablesOfContents.Count
    if ($tocCount -lt 1) {{
        for ($j = 1; $j -le $doc.Fields.Count; $j++) {{
            $field = $doc.Fields.Item($j)
            $code = ''
            try {{ $code = ($field.Code.Text).Trim().ToUpperInvariant() }} catch {{}}
            if ($code -eq 'TOC' -or $code.StartsWith('TOC ')) {{
                [void]$field.Update()
            }}
        }}
        $doc.Repaginate()
        $tocCount = $doc.TablesOfContents.Count
    }}

    if ($tocCount -lt 1) {{
        throw 'Microsoft Word tidak menemukan objek Table of Contents.'
    }}

    # Selalu bangun ulang entry terlebih dahulu. Range.Text bukan indikator
    # yang andal bahwa result field TOC sudah termaterialisasi.
    for ($i = 1; $i -le $tocCount; $i++) {{
        $toc = $doc.TablesOfContents.Item($i)
        try {{ $toc.Range.Fields.Locked = $false }} catch {{}}
        [void]$toc.Update()
    }}

    $previous = $null
    for ($pass = 1; $pass -le 4; $pass++) {{
        $doc.Repaginate()
        $tocCount = $doc.TablesOfContents.Count
        for ($i = 1; $i -le $tocCount; $i++) {{
            [void]$doc.TablesOfContents.Item($i).UpdatePageNumbers()
        }}
        $current = ''
        for ($i = 1; $i -le $tocCount; $i++) {{
            $current += [string]$doc.TablesOfContents.Item($i).Range.Text
        }}
        if ($null -ne $previous -and $current -eq $previous) {{ break }}
        $previous = $current
    }}

    for ($i = 1; $i -le $tocCount; $i++) {{
        $values = New-Object System.Collections.Generic.List[string]
        $lines = ([string]$doc.TablesOfContents.Item($i).Range.Text) -split "`r?`n"
        foreach ($line in $lines) {{
            if ($line -match '(?:\t|\s)([ivxlcdm]+|\d+)\s*$') {{
                [void]$values.Add($Matches[1].ToLowerInvariant())
            }}
        }}
        if ($values.Count -ge 8 -and @($values | Select-Object -Unique).Count -lt 2) {{
            throw 'Validasi TOC gagal: semua nomor halaman masih sama.'
        }}
    }}

    $doc.Save()
}}
finally {{
    if ($doc -ne $null) {{ try {{ $doc.Close($true) }} catch {{}} }}
    if ($word -ne $null) {{ try {{ $word.Quit() }} catch {{}} }}
    if ($doc -ne $null) {{ [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }}
    if ($word -ne $null) {{ [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}}
'''
        try:
            completed = subprocess.run(
                ['powershell.exe', '-NoProfile', '-NonInteractive',
                 '-ExecutionPolicy', 'Bypass', '-Command', ps_script],
                capture_output=True, text=True, timeout=180, check=False,
            )
        except Exception as ps_exc:
            raise RuntimeError(
                'Gagal menjalankan Update page numbers only melalui Microsoft Word. '
                f'pywin32: {pywin32_error}; PowerShell: {ps_exc}'
            ) from ps_exc

        if completed.returncode != 0:
            detail = (completed.stderr or completed.stdout or '').strip()
            raise RuntimeError(
                'Microsoft Word gagal menjalankan Update page numbers only pada TOC. '
                f'pywin32: {pywin32_error}; PowerShell: {detail}'
            )

    @staticmethod
    def _roman(number: int) -> str:
        values = ((1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'),
                  (100, 'c'), (90, 'xc'), (50, 'l'), (40, 'xl'),
                  (10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i'))
        result = []
        number = max(1, int(number))
        for value, symbol in values:
            while number >= value:
                result.append(symbol)
                number -= value
        return ''.join(result)

    @staticmethod
    def _pdf_pages(docx_path: str):
        """Render DOCX dengan LibreOffice dan kembalikan teks setiap halaman."""
        soffice = shutil.which('libreoffice') or shutil.which('soffice')
        if not soffice:
            raise RuntimeError(
                'LibreOffice tidak ditemukan. Tambahkan "libreoffice-writer" '
                'ke packages.txt pada Streamlit Cloud.'
            )
        try:
            from pypdf import PdfReader
        except Exception as exc:
            raise RuntimeError('Paket pypdf diperlukan untuk menghitung halaman TOC.') from exc

        with tempfile.TemporaryDirectory(prefix='engine9_pdf_') as tmp:
            profile = os.path.join(tmp, 'profile')
            completed = subprocess.run(
                [soffice, f'-env:UserInstallation=file://{profile}', '--headless',
                 '--convert-to', 'pdf', '--outdir', tmp, os.path.abspath(docx_path)],
                capture_output=True, text=True, timeout=240, check=False,
            )
            pdf_path = os.path.join(
                tmp, os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
            )
            if completed.returncode != 0 or not os.path.isfile(pdf_path):
                detail = (completed.stderr or completed.stdout or '').strip()
                raise RuntimeError(f'LibreOffice gagal menghitung pagination: {detail}')
            reader = PdfReader(pdf_path)
            pages = [re.sub(r'\s+', ' ', page.extract_text() or '').strip().lower()
                     for page in reader.pages]
            outline = []

            def collect(items):
                for item in items or []:
                    if isinstance(item, list):
                        collect(item)
                        continue
                    try:
                        name = re.sub(r'\s+', ' ', str(item.title)).strip()
                        page_no = reader.get_destination_page_number(item)
                    except Exception:
                        continue
                    if name and page_no is not None:
                        outline.append((_norm(name), int(page_no)))

            collect(reader.outline)
            return pages, outline

    @staticmethod
    def _write_static_toc(docx_path: str, entries) -> None:
        """Ganti field TOC dengan paragraf statis ber-tab leader titik."""
        doc = Document(docx_path)
        field_p = None
        for paragraph in doc.paragraphs:
            instruction = ''.join(
                node.text or '' for node in paragraph._p.iter(qn('w:instrText'))
            )
            if ' TOC ' in f' {instruction} ':
                field_p = paragraph
                break
        if field_p is None:
            # Iterasi kedua dan selanjutnya: hapus TOC statis yang ditandai.
            marked = [p for p in doc.paragraphs
                      if p._p.get(qn('w:rsidRPr')) == 'E9000001']
            if not marked:
                raise RuntimeError('Field/hasil TOC tidak ditemukan untuk ditulis.')
            anchor = marked[0]._p
            for p in marked:
                if p._p is not anchor:
                    p._p.getparent().remove(p._p)
            field_p = marked[0]

        anchor = field_p._p
        # Kosongkan paragraf anchor lalu gunakan sebagai entry pertama.
        for child in list(anchor):
            if child.tag != qn('w:pPr'):
                anchor.remove(child)

        style = TableOfContentsEngine._ensure_toc_style(doc)
        for index, (title, page_label) in enumerate(entries):
            if index == 0:
                paragraph = field_p
            else:
                paragraph = doc.add_paragraph()
                new_p = paragraph._p
                anchor.addnext(new_p)
                anchor = new_p
            paragraph.style = style
            paragraph._p.set(qn('w:rsidRPr'), 'E9000001')
            paragraph.add_run(title)
            paragraph.add_run('\t' + page_label)
        atomic_save_docx(doc, docx_path)
    @staticmethod
    def _update_toc_page_numbers_on_linux(docx_path: str) -> None:
        """Hitung TOC secara deterministik di Streamlit Cloud/Linux.

        LibreOffice dipakai hanya sebagai layout engine. Nomor halaman dicari
        dari PDF hasil render dan ditulis ke TOC statis; jadi tidak bergantung
        pada dukungan LibreOffice terhadap field TOC khusus Microsoft Word.
        """
        doc = Document(docx_path)
        title = next((p for p in doc.paragraphs if _norm(p.text) == 'daftar isi'), None)
        if title is None:
            raise RuntimeError('Heading Daftar isi tidak ditemukan.')
        source_entries = [(re.sub(r'\s+', ' ', title.text).strip(), '…')]
        source_entries.extend(TableOfContentsEngine._collect_entries(doc, title))
        if not source_entries:
            raise RuntimeError('Tidak ada Judul/Pasal untuk membangun TOC.')

        # Materialisasikan seluruh baris terlebih dahulu agar panjang TOC sudah
        # ikut memengaruhi pagination pada render berikutnya.
        labels = ['i' if not re.match(r'^1(?:\s|\.)', text) else '1'
                  for text, _ in source_entries]
        TableOfContentsEngine._write_static_toc(
            docx_path, [(item[0], labels[i]) for i, item in enumerate(source_entries)]
        )

        previous = None
        for _ in range(4):
            pages, outline = TableOfContentsEngine._pdf_pages(docx_path)
            normalized_titles = [_norm(text) for text, _ in source_entries]
            content_index = next((i for i, text in enumerate(normalized_titles)
                                  if re.match(r'^1(?:\s|\.)', text)), None)
            if content_index is None:
                raise RuntimeError('Awal halaman TOC/konten tidak dapat dideteksi.')

            # Style Judul/Pasal memiliki outline level 1, sehingga ekspor PDF
            # LibreOffice membawa destination bookmark yang menunjuk tepat ke
            # halaman heading asli (bukan teks duplikat di halaman TOC).
            found_pages = []
            for heading in normalized_titles:
                match_page = next((page for name, page in outline if name == heading), None)
                if match_page is None:
                    # Toleransi untuk perbedaan tanda baca/line wrapping pada
                    # judul bookmark hasil ekspor.
                    match_page = next((page for name, page in outline
                                       if heading in name or name in heading), None)
                if match_page is None:
                    raise RuntimeError(f'Halaman heading tidak ditemukan: {heading[:80]}')
                found_pages.append(match_page)

            toc_page = found_pages[0]
            content_page = found_pages[content_index]
            labels = [
                (TableOfContentsEngine._roman(page - toc_page + 1)
                 if i < content_index else str(page - content_page + 1))
                for i, page in enumerate(found_pages)
            ]
            state = tuple(labels)
            TableOfContentsEngine._write_static_toc(
                docx_path, [(item[0], labels[i]) for i, item in enumerate(source_entries)]
            )
            if state == previous:
                break
            previous = state

        if len(labels) >= 8 and len(set(labels)) < 2:
            raise RuntimeError('Validasi TOC Linux gagal: nomor halaman masih seragam.')

    def insert_toc(self, input_docx: str, output_docx: str) -> str:
        if not input_docx or not os.path.isfile(input_docx):
            raise FileNotFoundError(f'File input tidak ditemukan: {input_docx}')
        validate_docx(input_docx)
        ok, message = StyleFinalizerEngine().apply(input_docx, output_docx)
        if not ok:
            raise RuntimeError(message)
        doc = Document(output_docx)
        title = next((p for p in doc.paragraphs if _norm(p.text) == 'daftar isi'), None)
        if title is None:
            raise ValueError('Heading "Daftar isi" tidak ditemukan di section 3.')
        self._remove_existing_toc(doc)
        self._clear_cached_toc_entries(title)
        toc = self._build_toc_paragraph(doc)
        anchor = self._toc_anchor_after_three_blank_paragraphs(title)
        anchor.addnext(toc._p)
        self._force_update_fields(doc)
        atomic_save_docx(doc, output_docx)
        # Jangan sediakan file untuk download sebelum Word benar-benar selesai
        # menghitung pagination dan menjalankan "Update page numbers only".
        self._update_toc_page_numbers_with_word(output_docx)
        # Hasil COM di atas adalah hasil final. Cegah Word mengulang Update All
        # saat pengguna membuka file dan merusak nomor tersimpan menjadi i/1.
        self._disable_update_fields_after_finalization(output_docx)
        return output_docx

    def process(self, input_docx: Optional[str] = None,
                output_docx: Optional[str] = None, **_kwargs):
        output_docx = output_docx or 'hasil_RSNI.docx'
        try:
            path = self.insert_toc(input_docx, output_docx)
            return True, path, (
                'Engine 9 selesai: style custom Judul/Pasal diterapkan hanya '
                'pada bagian Indonesia, TOC dibuat dari "Judul,1,Pasal,1", '
                'dan nomor halaman TOC sudah dihitung berdasarkan pagination final.'
            )
        except Exception as exc:
            return False, None, f'Engine9 Error: {exc}\n{traceback.format_exc()}'


DaftarIsiTocEngine = TableOfContentsEngine
Engine9 = TableOfContentsEngine
__all__ = [
    'StyleFinalizerEngine', 'TableOfContentsEngine', 'DaftarIsiTocEngine',
    'Engine9', 'apply_custom_styles'
]
