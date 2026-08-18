"""
Engine10: StyleFinalizerEngine
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

import re
import copy
from docx import Document
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


def _norm(text: str) -> str:
    """Normalisasi teks untuk dibandingkan: lower-case, trim whitespace/tab."""
    return re.sub(r'\s+', ' ', (text or '').strip()).strip().lower()


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
            # Buang semua properti langsung lain, sisakan hanya vertAlign
            for child in list(rPr):
                if child is not vert:
                    rPr.remove(child)
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

            # 2) Cari batas bagian Indonesia vs Inggris via "Main Title 1"
            main_title_idx = [i for i in range(n) if style_name(paras[i]) == 'Main Title 1']
            if len(main_title_idx) >= 2:
                id_lo, id_hi = main_title_idx[0], main_title_idx[1]
            else:
                id_lo, id_hi = 0, n

            # 3) Cari batas akhir "front matter" (sebelum heading/pasal pertama)
            first_heading_idx = n
            for i in range(n):
                if style_name(paras[i]) in ('Heading 1', 'Heading 2', 'Heading 3', 'Heading 4'):
                    first_heading_idx = i
                    break

            # 4) Kumpulkan target paragraf untuk style "Judul"
            judul_targets = {}  # idx -> force_page_break_before

            # "Daftar Isi" -> kemunculan PERTAMA di seluruh dokumen (judul halaman TOC)
            for i in range(n):
                if _norm(paras[i].text) == 'daftar isi':
                    judul_targets[i] = False
                    break

            # "Pendahuluan" -> kemunculan TERAKHIR sebelum heading/pasal pertama
            # (yaitu heading halaman Pendahuluan, bukan entri di daftar isi)
            cand = [i for i in range(first_heading_idx) if _norm(paras[i].text) == 'pendahuluan']
            if cand:
                judul_targets[cand[-1]] = False

            # "Prakata" -> semua kemunculan sebelum heading/pasal pertama
            for i in range(first_heading_idx):
                if _norm(paras[i].text) == 'prakata':
                    judul_targets[i] = False

            # "Bibliografi" -> semua kemunculan SETELAH heading/pasal pertama
            # (di badan dokumen, bukan entri di daftar isi front-matter)
            for i in range(first_heading_idx, n):
                if _norm(paras[i].text) == 'bibliografi':
                    was_biblio_title = style_name(paras[i]) == 'Biblio Title'
                    judul_targets[i] = was_biblio_title  # pertahankan page-break-before

            for idx, force_pb in judul_targets.items():
                _apply_judul(doc, paras[idx], force_page_break_before=force_pb)

            # 5) Kumpulkan target paragraf untuk style "Pasal"
            #    (hanya Heading 1 & Heading 2, hanya di rentang bahasa Indonesia).
            #    Pasal "Istilah dan Definisi" (Heading 1) SENDIRI tetap diberi
            #    style "Pasal" (supaya muncul di Daftar Isi/ToC); hanya
            #    subpasal-nya (Heading 2, mis. 3.1, 3.2, ...) yang dilewati.
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
            skip_active = False
            # Counter untuk membangkitkan nomor Pasal/Subpasal (lihat
            # penjelasan lengkap di komentar atas _RE_LEADING_NUMBER).
            h1_counter = 0          # nomor Pasal (Heading 1) berjalan
            h2_counter = 0          # nomor Subpasal (Heading 2), reset tiap Pasal baru
            annex_counter = 0       # huruf Lampiran: 0->belum ada, 1->A, 2->B, ...
            annex_sub_counter = 0   # nomor 'a2' (mis. A.1, A.2, ...), reset tiap Lampiran baru
            annex_sub2_counter = 0  # nomor 'a3' (mis. A.1.1, ...), reset tiap 'a2' baru
            for i in range(id_lo, id_hi):
                sname = style_name(paras[i])
                if sname == 'ANNEX':
                    # Judul Lampiran sendiri TIDAK disentuh (tetap style asli),
                    # tapi menandai mulainya huruf Lampiran baru untuk
                    # subpasal 'a2'/'a3' di bawahnya.
                    annex_counter += 1
                    annex_sub_counter = 0
                    annex_sub2_counter = 0
                    continue
                if sname == 'Heading 1':
                    skip_active = _norm(paras[i].text) in _SKIP_SUBSECTION_TITLES
                    h1_counter += 1
                    h2_counter = 0
                    pasal_targets.append((i, 0, heading1_num_id, str(h1_counter)))
                elif sname == 'Heading 2':
                    if skip_active:
                        continue
                    h2_counter += 1
                    pasal_targets.append((i, 1, heading2_num_id, f'{h1_counter}.{h2_counter}'))
                elif sname == 'a2' and annex_a2_num_id:
                    annex_sub_counter += 1
                    annex_sub2_counter = 0
                    letter = chr(ord('A') + max(annex_counter - 1, 0))
                    pasal_targets.append((i, annex_a2_ilvl, annex_a2_num_id, f'{letter}.{annex_sub_counter}'))
                elif sname == 'a3' and annex_a3_num_id:
                    annex_sub2_counter += 1
                    letter = chr(ord('A') + max(annex_counter - 1, 0))
                    pasal_targets.append((i, annex_a3_ilvl, annex_a3_num_id, f'{letter}.{annex_sub_counter}.{annex_sub2_counter}'))
                # Heading 3+ (sub-subpasal) sengaja tidak disentuh
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

            doc.save(output_docx)

            msg = (
                f'OK: {len(judul_targets)} heading -> "Judul", '
                f'{len(pasal_targets)} pasal/subpasal -> "Pasal".'
            )
            return True, msg

        except Exception as e:
            import traceback
            return False, f'StyleFinalizerEngine Error: {str(e)}\n{traceback.format_exc()}'


def apply_custom_styles(input_docx: str, output_docx: str) -> tuple[bool, str]:
    """Shortcut fungsi-level untuk StyleFinalizerEngine().apply(...)."""
    return StyleFinalizerEngine().apply(input_docx, output_docx)
