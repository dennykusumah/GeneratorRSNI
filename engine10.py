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


# Marker yang ditanam Engine2 pada paragraf yang berasal langsung dari
# file .docx yang di-upload pengguna. Engine10 WAJIB mempertahankan style
# paragraf tersebut; marker ini tidak memengaruhi tampilan Word.
_USER_MARK_NS = 'urn:generatorrsni:user-content'
_USER_MARK_QN = '{%s}uploaded' % _USER_MARK_NS

def _is_user_uploaded_paragraph(paragraph) -> bool:
    return paragraph._p.get(_USER_MARK_QN) == '1'


def _strip_number_for_toc(text: str) -> str:
    """Hilangkan nomor literal di awal heading agar TOC tidak dobel.
    Nomor otomatis Word tidak muncul pada paragraph.text, sehingga pada
    dokumen ISO yang masih memakai numbering ini fungsi ini tidak menghapus
    apa pun dan nomor ditambahkan oleh generator TOC dari counter."""
    return _RE_LEADING_NUMBER.sub('', (text or '').strip(), count=1)


def _bookmark_name_for_para(paragraph, index: int) -> str:
    """Buat bookmark unik untuk target TOC jika belum ada bookmark."""
    p_el = paragraph._p
    # Gunakan bookmark yang sudah ada bila paragraf sumber memang memilikinya.
    for b in p_el.findall('.//' + qn('w:bookmarkStart')):
        name = b.get(qn('w:name'))
        if name and not name.startswith('_Toc'):
            return name
    return f'_GeneratorTOC_{index}'


def _ensure_bookmark(paragraph, name: str, bookmark_id: int) -> None:
    """Pasang bookmark mengelilingi isi paragraf tanpa mengubah teks/style."""
    p_el = paragraph._p
    starts = p_el.findall('.//' + qn('w:bookmarkStart'))
    for b in starts:
        if b.get(qn('w:name')) == name:
            return

    end_id = str(bookmark_id)
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), end_id)
    start.set(qn('w:name'), name)

    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), end_id)

    # Sisipkan bookmark setelah pPr dan sebelum run pertama; end di akhir.
    pPr = p_el.find(qn('w:pPr'))
    insert_at = list(p_el).index(pPr) + 1 if pPr is not None else 0
    p_el.insert(insert_at, start)
    p_el.append(end)


def _next_bookmark_id(doc) -> int:
    ids = []
    for b in doc.element.body.iter(qn('w:bookmarkStart')):
        try:
            ids.append(int(b.get(qn('w:id'))))
        except Exception:
            pass
    return max(ids, default=0) + 1


def _make_toc_run(text: str, *, bold=False, color=None, underline=False):
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')
    rFonts.set(qn('w:cs'), 'Arial')
    rPr.append(rFonts)
    if bold:
        rPr.append(OxmlElement('w:b'))
    if color:
        c = OxmlElement('w:color'); c.set(qn('w:val'), color); rPr.append(c)
    if underline:
        u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
    r.append(rPr)
    fld = OxmlElement('w:t')
    fld.set(qn('xml:space'), 'preserve')
    fld.text = text
    r.append(fld)
    return r


def _make_pageref_field_run(bookmark_name: str):
    """Field PAGEREF yang menghasilkan nomor halaman + hyperlink."""
    rpr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Arial'); rFonts.set(qn('w:hAnsi'), 'Arial'); rFonts.set(qn('w:cs'), 'Arial')
    rpr.append(rFonts)
    color = OxmlElement('w:color'); color.set(qn('w:val'), '0070C0'); rpr.append(color)
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rpr.append(u)

    def rr(child):
        r = OxmlElement('w:r'); r.append(copy.deepcopy(rpr)); r.append(child); return r

    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = f' PAGEREF {bookmark_name} \\h '
    sep = OxmlElement('w:fldChar'); sep.set(qn('w:fldCharType'), 'separate')
    text = OxmlElement('w:t'); text.text = '1'
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    return [rr(begin), rr(instr), rr(sep), rr(text), rr(end)]


def _build_custom_toc_paragraph(doc, label: str, bookmark_name: str):
    """Buat satu entri TOC visual seperti F.docx: Arial, dot leader,
    hyperlink nomor halaman, tanpa mengubah paragraf sumber."""
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle'); pStyle.set(qn('w:val'), 'TOC1'); pPr.append(pStyle)
    p.append(pPr)

    # Label kiri.
    p.append(_make_toc_run(label))
    tab = OxmlElement('w:r'); tab.append(OxmlElement('w:tab')); p.append(tab)

    # Nomor halaman kanan; Word akan menghitung ulang PAGEREF saat field di-update.
    for r in _make_pageref_field_run(bookmark_name):
        p.append(r)
    return p


def _replace_engine5_toc_with_generated(doc, toc_entries):
    """Ganti field/placeholder TOC dari Engine5 dengan entri TOC final
    yang dibangun Engine10. Halaman, dot leader, dan hyperlink mengikuti
    tampilan F.docx, sementara isi sumber tetap mempertahankan style aslinya."""
    paras = doc.paragraphs
    di_idx = None
    for i, p in enumerate(paras):
        if _norm(p.text) == 'daftar isi':
            di_idx = i
            break
    if di_idx is None:
        return 0

    # Temukan batas TOC: dari setelah judul DI sampai sebelum Prakata.
    end_idx = len(paras)
    for j in range(di_idx + 1, len(paras)):
        if _norm(paras[j].text) == 'prakata':
            end_idx = j
            break

    candidates = []
    for p in paras[di_idx + 1:end_idx]:
        txt = p.text or ''
        if p.style is not None and p.style.name in ('TOC1', 'toc 1', 'TOC2', 'toc 2'):
            candidates.append(p)
        elif 'TOC \\h' in txt or 'Klik kanan pada teks ini' in txt:
            candidates.append(p)

    if not candidates:
        return 0

    first = candidates[0]
    parent = first._p.getparent()
    pos = list(parent).index(first._p)

    # Hapus semua paragraf TOC lama, tetapi JANGAN menghapus tiga paragraf
    # kosong yang sengaja dibuat Engine5 untuk jarak judul -> daftar isi.
    for old in candidates:
        try:
            old._p.getparent().remove(old._p)
        except Exception:
            pass

    new_nodes = [_build_custom_toc_paragraph(doc, label, bm) for label, bm in toc_entries]
    for offset, node in enumerate(new_nodes):
        parent.insert(pos + offset, node)
    return len(new_nodes)


def _collect_toc_entries(doc, id_lo, id_hi):
    """Kumpulkan target TOC tanpa mengubah style paragraf sumber.
    Semua entri dibuat level datar (TOC1), sama seperti F.docx."""
    paras = doc.paragraphs
    entries = []
    bookmark_id = _next_bookmark_id(doc)

    # Judul halaman yang dibuat generator.
    generated_titles = {'daftar isi', 'prakata', 'pendahuluan'}

    # Daftar Isi sendiri, Prakata, Pendahuluan.
    for i, p in enumerate(paras):
        norm = _norm(p.text)
        if norm in generated_titles and not _is_user_uploaded_paragraph(p):
            name = _bookmark_name_for_para(p, i)
            _ensure_bookmark(p, name, bookmark_id); bookmark_id += 1
            entries.append((p.text.strip(), name))

    # Heading/pasal sumber: tetap dengan style asli.
    h1 = 0; h2 = 0; skip_sub = False
    annex_no = 0; annex_sub = 0; annex_sub2 = 0
    for i in range(id_lo, id_hi):
        p = paras[i]
        s = p.style.name if p.style is not None else ''
        if _is_user_uploaded_paragraph(p):
            if s in ('Heading 1', 'Pasal'):
                raw = _strip_number_for_toc(p.text)
                m_num = re.match(r'^\s*(\d+)(?:\.(\d+))?\s+', p.text or '')
                if s == 'Pasal' and m_num and m_num.group(2):
                    h2 += 1
                    label = f'{m_num.group(1)}.{m_num.group(2)}    {raw}'
                else:
                    h1 += 1
                    h2 = 0
                    skip_sub = _norm(raw) in _SKIP_SUBSECTION_TITLES
                    label = f'{h1}    {raw}'
                name = _bookmark_name_for_para(p, i)
                _ensure_bookmark(p, name, bookmark_id); bookmark_id += 1
                entries.append((label, name))
            elif s == 'Heading 2':
                if skip_sub:
                    continue
                h2 += 1
                label = f'{h1}.{h2}    {_strip_number_for_toc(p.text)}'
                name = _bookmark_name_for_para(p, i)
                _ensure_bookmark(p, name, bookmark_id); bookmark_id += 1
                entries.append((label, name))
            elif s == 'Judul' and _norm(p.text) not in generated_titles:
                name = _bookmark_name_for_para(p, i)
                _ensure_bookmark(p, name, bookmark_id); bookmark_id += 1
                entries.append((p.text.strip(), name))
            elif s == 'ANNEX':
                annex_no += 1; annex_sub = 0; annex_sub2 = 0
                letter = chr(ord('A') + annex_no - 1)
                label = f'Lampiran {letter} {_strip_number_for_toc(p.text)}'
                name = _bookmark_name_for_para(p, i)
                _ensure_bookmark(p, name, bookmark_id); bookmark_id += 1
                entries.append((label, name))
            elif s == 'a2':
                annex_sub += 1; annex_sub2 = 0
                letter = chr(ord('A') + max(annex_no - 1, 0))
                label = f'{letter}.{annex_sub}    {_strip_number_for_toc(p.text)}'
                name = _bookmark_name_for_para(p, i)
                _ensure_bookmark(p, name, bookmark_id); bookmark_id += 1
                entries.append((label, name))
            elif s == 'a3':
                annex_sub2 += 1
                letter = chr(ord('A') + max(annex_no - 1, 0))
                label = f'{letter}.{annex_sub}.{annex_sub2}    {_strip_number_for_toc(p.text)}'
                name = _bookmark_name_for_para(p, i)
                _ensure_bookmark(p, name, bookmark_id); bookmark_id += 1
                entries.append((label, name))

    # Bibliografi: masukkan sekali bila berasal dari dokumen upload.
    for i, p in enumerate(paras):
        if _norm(p.text) == 'bibliografi' and _is_user_uploaded_paragraph(p):
            name = _bookmark_name_for_para(p, i)
            _ensure_bookmark(p, name, bookmark_id); bookmark_id += 1
            entries.append((p.text.strip(), name))
            break

    return entries


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
                if _norm(paras[i].text) == 'daftar isi' and not _is_user_uploaded_paragraph(paras[i]):
                    judul_targets[i] = False
                    break

            # "Pendahuluan" -> kemunculan TERAKHIR sebelum heading/pasal pertama
            # (yaitu heading halaman Pendahuluan, bukan entri di daftar isi)
            cand = [i for i in range(first_heading_idx) if _norm(paras[i].text) == 'pendahuluan' and not _is_user_uploaded_paragraph(paras[i])]
            if cand:
                judul_targets[cand[-1]] = False

            # "Prakata" -> semua kemunculan sebelum heading/pasal pertama
            for i in range(first_heading_idx):
                if _norm(paras[i].text) == 'prakata' and not _is_user_uploaded_paragraph(paras[i]):
                    judul_targets[i] = False

            # "Bibliografi" -> semua kemunculan SETELAH heading/pasal pertama
            # (di badan dokumen, bukan entri di daftar isi front-matter)
            for i in range(first_heading_idx, n):
                if _norm(paras[i].text) == 'bibliografi' and not _is_user_uploaded_paragraph(paras[i]):
                    was_biblio_title = style_name(paras[i]) == 'Biblio Title'
                    judul_targets[i] = was_biblio_title  # pertahankan page-break-before

            # "Lampiran/Annex" -> WAJIB style "Judul".
            # Dokumen sumber ISO memakai style custom "ANNEX" untuk judul
            # lampiran. Engine5 membangun TOC dengan field:
            #   TOC \h \z \t "Judul;1;Pasal;1;ANNEX;1"
            # Namun agar hasil akhir sama seperti F.docx, style ANNEX harus
            # dikonversi menjadi "Judul". Dengan begitu entri Lampiran
            # tetap muncul di TOC sebagai level 1, sekaligus tampil dengan
            # format Judul (Arial 12 pt, bold, center) seperti F.docx.
            # Teks Lampiran TIDAK diubah.
            # force_page_break_before=True -> SETIAP Lampiran/Annex (ID
            # maupun EN) WAJIB dimulai di halaman baru. Ini dijamin oleh
            # kode (properti w:pageBreakBefore pada paragraf heading-nya
            # sendiri), TIDAK bergantung pada apakah dokumen sumber sudah
            # punya page break manual sebelum Annex atau belum.
            # Simpan indeks paragraf ANNEX asli SEBELUM style-nya diubah
            # menjadi "Judul" di bawah. Step 5 (penomoran Pasal Lampiran,
            # mis. B.1, B.2, ...) butuh tahu paragraf mana yang merupakan
            # batas Lampiran baru — kalau dicek SESUDAH mutasi, style-nya
            # sudah jadi "Judul" untuk SEMUA Lampiran sehingga batas antar
            # Lampiran (mis. A -> B -> C) tidak lagi terdeteksi dan seluruh
            # subpasal akan salah dianggap masih milik Lampiran pertama.
            annex_indices = set(i for i in range(n) if style_name(paras[i]) == 'ANNEX' and not _is_user_uploaded_paragraph(paras[i]))
            for i in annex_indices:
                judul_targets[i] = True

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
                if i in annex_indices:
                    # Paragraf ini adalah judul Lampiran (style aslinya
                    # "ANNEX", sudah diubah jadi "Judul" di atas — makanya
                    # dicek lewat annex_indices, bukan sname). Menandai
                    # mulainya huruf Lampiran baru untuk subpasal 'a2'/'a3'
                    # di bawahnya.
                    annex_counter += 1
                    annex_sub_counter = 0
                    annex_sub2_counter = 0
                    continue
                if sname == 'Heading 1':
                    skip_active = _norm(paras[i].text) in _SKIP_SUBSECTION_TITLES
                    h1_counter += 1
                    h2_counter = 0
                    if not _is_user_uploaded_paragraph(paras[i]):
                        pasal_targets.append((i, 0, heading1_num_id, str(h1_counter)))
                elif sname == 'Heading 2':
                    if skip_active:
                        continue
                    h2_counter += 1
                    if not _is_user_uploaded_paragraph(paras[i]):
                        pasal_targets.append((i, 1, heading2_num_id, f'{h1_counter}.{h2_counter}'))
                elif sname == 'a2' and annex_a2_num_id:
                    annex_sub_counter += 1
                    annex_sub2_counter = 0
                    letter = chr(ord('A') + max(annex_counter - 1, 0))
                    if not _is_user_uploaded_paragraph(paras[i]):
                        pasal_targets.append((i, annex_a2_ilvl, annex_a2_num_id, f'{letter}.{annex_sub_counter}'))
                elif sname == 'a3' and annex_a3_num_id:
                    annex_sub2_counter += 1
                    letter = chr(ord('A') + max(annex_counter - 1, 0))
                    if not _is_user_uploaded_paragraph(paras[i]):
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

            # 7) TOC FINAL:
            # Engine5 membuat halaman Daftar Isi dan placeholder field. Pada
            # tahap terakhir ini Engine10 membangun entri TOC berdasarkan
            # heading final, lalu mengganti placeholder tersebut. Heading yang
            # berasal dari file upload pengguna TIDAK diubah style-nya; TOC
            # memakai bookmark + PAGEREF sehingga nomor halaman tetap dinamis.
            toc_entries = _collect_toc_entries(doc, id_lo, id_hi)
            toc_count = _replace_engine5_toc_with_generated(doc, toc_entries)

            doc.save(output_docx)

            msg = (
                f'OK: {len(judul_targets)} heading generator -> "Judul", '
                f'{len(pasal_targets)} pasal generator -> "Pasal", '
                f'{toc_count} entri TOC dibuat.'
            )
            return True, msg

        except Exception as e:
            import traceback
            return False, f'StyleFinalizerEngine Error: {str(e)}\n{traceback.format_exc()}'


def apply_custom_styles(input_docx: str, output_docx: str) -> tuple[bool, str]:
    """Shortcut fungsi-level untuk StyleFinalizerEngine().apply(...)."""
    return StyleFinalizerEngine().apply(input_docx, output_docx)
