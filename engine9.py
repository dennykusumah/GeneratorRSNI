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
      Style "Pasal" JUGA diterapkan pada subpasal di dalam Lampiran/Annex
      berbahasa Indonesia (paragraf ber-style "a2"/"a3"), dengan penomoran
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

import os
import re
import copy
import traceback
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
_TOC_FIELD_INSTR = ' TOC \\h \\z \\t "Engine9 Judul ToC;1;Engine9 Pasal ToC;1" '
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
#     A.1, A.2, ...), reset counter 'a3'.
#   - Setiap "a3" (sub-subpasal Lampiran, level 2) -> "<huruf>.<urut a2>.
#     <urut a3>" (mis. A.1.1, A.1.2, ...).
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

_JUDUL_STYLE_XML = f'''<w:style {nsdecls("w")} w:type="paragraph" w:customStyle="1" w:styleId="Engine9JudulTOC">
  <w:name w:val="Engine9 Judul ToC"/>
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

_PASAL_STYLE_XML = f'''<w:style {nsdecls("w")} w:type="paragraph" w:customStyle="1" w:styleId="Engine9PasalTOC">
  <w:name w:val="Engine9 Pasal ToC"/>
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


# Style visual identik dengan "Pasal", tetapi sengaja TIDAK dipetakan ke
# field TOC. Dipakai untuk level ketiga dan lebih dalam (mis. 1.1.1, A.1.2)
# agar formatting dokumen tetap sama namun entri tersebut tidak masuk Daftar Isi.
_PASAL_NONTOC_STYLE_XML = f'''<w:style {nsdecls("w")} w:type="paragraph" w:customStyle="1" w:styleId="Engine9PasalNonTOC">
  <w:name w:val="Engine9 Pasal Non-TOC"/>
  <w:basedOn w:val="BodyText"/>
  <w:next w:val="BodyText"/>
  <w:pPr>
    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>
    <w:outlineLvl w:val="9"/>
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


def _set_pasal_toc_safe_style(paragraph, depth):
    """Set style Pasal secara aman untuk ToC maksimum dua tingkat.

    depth 0/1 = Pasal/Subpasal yang boleh masuk ToC.
    depth >=2 = sub-subpasal atau lebih dalam; style dibuat non-ToC DAN
    outline level paragraf dipaksa ke body-text (9) agar Word tidak dapat
    memasukkannya lewat inheritance Heading/outline level.
    """
    try:
        depth = int(depth)
    except (TypeError, ValueError):
        depth = 99
    non_toc = depth >= 2
    _set_pstyle(paragraph, 'Engine9PasalNonTOC' if non_toc else 'Engine9PasalTOC')
    pPr = paragraph._p.get_or_add_pPr()
    old_outline = pPr.find(qn('w:outlineLvl'))
    if old_outline is not None:
        pPr.remove(old_outline)
    if non_toc:
        outline = OxmlElement('w:outlineLvl')
        outline.set(qn('w:val'), '9')
        pPr.append(outline)


def _apply_judul(doc, paragraph, force_page_break_before=False):
    _set_pstyle(paragraph, 'Engine9JudulTOC')
    pPr = _clear_paragraph_direct_formatting(
        paragraph, ['w:jc', 'w:spacing', 'w:ind', 'w:outlineLvl', 'w:numPr']
    )
    if force_page_break_before and pPr.find(qn('w:pageBreakBefore')) is None:
        pPr.append(parse_xml(f'<w:pageBreakBefore {nsdecls("w")}/>'))
    _strip_run_direct_formatting(paragraph)


def _apply_pasal(doc, paragraph, ilvl, num_id, number_text=None):
    # ToC dibatasi maksimal dua tingkat. Level 0 = Pasal, level 1 = Subpasal.
    # Level >= 2 tetap diformat identik, tetapi memakai style non-TOC.
    try:
        toc_level = int(ilvl)
    except (TypeError, ValueError):
        toc_level = (number_text or '').count('.')
    _set_pasal_toc_safe_style(paragraph, toc_level)
    pPr = _clear_paragraph_direct_formatting(
        paragraph, ['w:jc', 'w:spacing', 'w:ind', 'w:numPr']
    )
    # _clear_paragraph_direct_formatting tidak boleh menghapus guard outlineLvl=9
    # pada level ketiga ke bawah. Jika sumber membawa outline heading langsung,
    # helper di atas sudah menggantinya secara deterministik.
    if toc_level >= 2:
        old_outline = pPr.find(qn('w:outlineLvl'))
        if old_outline is not None:
            pPr.remove(old_outline)
        outline = OxmlElement('w:outlineLvl')
        outline.set(qn('w:val'), '9')
        pPr.append(outline)
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
            _ensure_style(doc, 'Engine9JudulTOC', _JUDUL_STYLE_XML)
            _ensure_style(doc, 'Engine9PasalTOC', _PASAL_STYLE_XML)
            _ensure_style(doc, 'Engine9PasalNonTOC', _PASAL_NONTOC_STYLE_XML)

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
            # Content Indonesia sebelum marker, termasuk a2/a3 di Lampiran.
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
            annex_a3_num_id = _resolve_num_id(doc, 'a3') or annex_a2_num_id
            annex_a2_ilvl = _resolve_style_ilvl(doc, 'a2')
            annex_a3_ilvl = _resolve_style_ilvl(doc, 'a3')

            pasal_targets = []  # (idx, ilvl, num_id, number_text)
            # Counter untuk membangkitkan nomor Pasal/Subpasal (lihat
            # penjelasan lengkap di komentar atas _RE_LEADING_NUMBER).
            h1_counter = 0          # nomor Pasal (Heading 1) berjalan
            h2_counter = 0          # nomor Subpasal (Heading 2), reset tiap Pasal baru
            annex_letter = None       # huruf Lampiran/Annex aktif, mis. A, B, C
            annex_sub_counter = 0      # nomor 'a2' (mis. B.1, B.2, ...), reset tiap Annex baru
            annex_sub2_counter = 0     # nomor 'a3' (mis. B.1.1, ...), reset tiap 'a2' baru
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
                    annex_sub2_counter = 0
                    continue

                # Hanya section Content Indonesia dan selalu sebelum marker.
                if section_by_idx[i] < 4 or i >= marker_idx:
                    continue
                if sname in ('@Pasal', 'Pasal', 'Pasal Non-TOC', 'Engine9 Pasal ToC', 'Engine9 Pasal Non-TOC'):
                    # Migrasi style lama. Pertahankan nomor literal yang telah
                    # dibuat Engine7/8, termasuk A.1/A.1.1 pada Lampiran.
                    m_num = re.match(r'^\s*([A-Z]?\d+(?:\.\d+)*)\.?\s+', paras[i].text or '', re.I)
                    if m_num:
                        number = m_num.group(1).upper()
                        level = number.count('.')
                        pasal_targets.append((i, level, None, number))
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
                    annex_sub2_counter = 0
                    pasal_targets.append(
                        (i, annex_a2_ilvl, annex_a2_num_id,
                         f'{annex_letter}.{annex_sub_counter}')
                    )
                elif sname.lower() == 'a3' and annex_letter:
                    annex_sub2_counter += 1
                    pasal_targets.append(
                        (i, annex_a3_ilvl, annex_a3_num_id,
                         f'{annex_letter}.{annex_sub_counter}.{annex_sub2_counter}')
                    )
                # Heading 3+ (sub-subpasal) sengaja tidak disentuh
                # Judul Lampiran sendiri (style "ANNEX") sengaja tidak disentuh

            for idx, ilvl, num_id, number_text in pasal_targets:
                _apply_pasal(doc, paras[idx], ilvl, num_id, number_text)

            # 5b) Sanitasi ToC berbasis NOMOR AKTUAL, bukan hanya style sumber.
            # Dokumen ISO/IEC kompleks dapat membawa style duplikat seperti
            # styleId Pasal0 dengan display-name 'Pasal'. Word field TOC \t
            # mencocokkan display-name style, sehingga 6.4.1/7.8.1 dapat ikut
            # walaupun bukan Heading 1/2. Semua heading bernomor >= 3 tingkat
            # dipaksa ke style unik non-TOC, sedangkan 1/1.1/A.1 tetap pada
            # style unik TOC. Untuk level >=3, outlineLvl juga dipaksa 9
            # (body text) agar tidak dapat masuk ToC melalui inheritance.
            # Tampilan visual kedua style dibuat identik.
            for i in range(n):
                if section_by_idx[i] < 4 or i >= marker_idx:
                    continue
                text_now = re.sub(r'[\u200b\u200c\u200d\ufeff]', '', paras[i].text or '').strip()
                m_num = re.match(r'^([A-Z]?\d+(?:\.\d+)*)(?:\.)?(?=\s|[\u200b\u200c\u200d\ufeff]|[^0-9.]|$)', text_now, re.I)
                if not m_num:
                    continue
                number_now = m_num.group(1).upper()
                depth = number_now.count('.')
                s_now = style_name(paras[i])
                # Batasi perubahan hanya pada paragraf yang memang heading/pasal
                # atau sudah pernah disentuh engine sebelumnya.
                is_heading_like = (
                    s_now in ('@Pasal', 'Pasal', 'Pasal Non-TOC',
                              'Engine9 Pasal ToC', 'Engine9 Pasal Non-TOC')
                    or s_now.startswith('Heading ')
                    or s_now.lower() in ('a2', 'a3')
                )
                if not is_heading_like:
                    continue
                _set_pasal_toc_safe_style(paras[i], depth)

            # Invariant final: style yang dipetakan ke TOC tidak boleh pernah
            # membawa nomor tiga tingkat atau lebih, termasuk kasus tanpa spasi
            # seperti '7.8.1.1Tujuan' atau yang mengandung zero-width character.
            for i in range(n):
                if section_by_idx[i] < 4 or i >= marker_idx:
                    continue
                if style_name(paras[i]) != 'Engine9 Pasal ToC':
                    continue
                check_text = re.sub(r'[\u200b\u200c\u200d\ufeff]', '', paras[i].text or '').strip()
                m_check = re.match(
                    r'^([A-Z]?\d+(?:\.\d+)*)(?:\.)?(?=\s|[^0-9.]|$)',
                    check_text, flags=re.I
                )
                if m_check and m_check.group(1).count('.') >= 2:
                    _set_pasal_toc_safe_style(paras[i], 2)

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

            doc.save(output_docx)

            msg = (
                f'OK: {len(judul_targets)} heading/Lampiran -> "Judul", '
                f'{len(pasal_targets)} pasal/subpasal diformat; ToC maksimal dua tingkat.'
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
    def _collect_entries(doc: Document, toc_title: Paragraph):
        entries = []
        for paragraph in doc.paragraphs:
            if paragraph._p is toc_title._p:
                continue
            if _style_name(paragraph) not in ('Engine9 Judul ToC', 'Engine9 Pasal ToC'):
                continue
            text = re.sub(r'\s+', ' ', paragraph.text or '').strip()
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

    def insert_toc(self, input_docx: str, output_docx: str) -> str:
        if not input_docx or not os.path.isfile(input_docx):
            raise FileNotFoundError(f'File input tidak ditemukan: {input_docx}')
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
        doc.save(output_docx)
        return output_docx

    def process(self, input_docx: Optional[str] = None,
                output_docx: Optional[str] = None, **_kwargs):
        output_docx = output_docx or 'hasil_RSNI.docx'
        try:
            path = self.insert_toc(input_docx, output_docx)
            return True, path, (
                'Engine 9 selesai: style custom Judul/Pasal diterapkan hanya '
                'pada bagian Indonesia dan TOC dibatasi sampai Pasal/Subpasal (dua tingkat).'
            )
        except Exception as exc:
            return False, None, f'Engine9 Error: {exc}\n{traceback.format_exc()}'


DaftarIsiTocEngine = TableOfContentsEngine
Engine9 = TableOfContentsEngine
__all__ = [
    'StyleFinalizerEngine', 'TableOfContentsEngine', 'DaftarIsiTocEngine',
    'Engine9', 'apply_custom_styles'
]
