"""
Engine5: DaftarIsiEngine (v7 - Daftar Isi = replika PERSIS field TOC F.docx)
=============================================================================
Engine untuk menyisipkan halaman "Daftar Isi" sesuai standar BSN/SNI.

v7 (field TOC & style "TOC1" dipelajari & disamakan PERSIS dengan F.docx,
menggantikan pendekatan v6 yang formatnya masih berbeda dari F.docx):
  - Halaman Daftar Isi berisi judul "Daftar Isi" (rata tengah, bold, style
    "Judul"), diikuti field TOC ASLI Word:
        { TOC \\h \\z \\t "Judul;1;Pasal;1" }
    Field ini disisipkan dalam bentuk "belum di-update" (persis seperti saat
    pengguna melakukan Insert > Table of Contents secara manual di Word) —
    satu paragraf berisi begin/instrText/separate/teks-placeholder/end.
    Karena <w:updateFields w:val="true"/> sudah diset di settings.xml
    (lihat _ensure_update_fields), Word otomatis meng-update SEMUA field
    ketika dokumen dibuka — termasuk field TOC ini — sehingga daftar isi
    langsung terisi rapi (entri + dot leader + nomor halaman + hyperlink)
    begitu file dibuka, tanpa perlu pengguna menekan F9 secara manual.
  - PEMISAH TITIK KOMA (;), bukan koma (,), pada switch \\t — dipelajari
    langsung dari field TOC di F.docx. Pemisah daftar pada switch field
    Word mengikuti Regional Settings Windows si pengguna; di locale
    Indonesia (id-ID) pemisahnya titik koma, sama seperti di F.docx. Field
    yang salah pemisah bisa gagal menghasilkan entri ("Error! No table of
    contents entries found.") — inilah salah satu sumber "error" yang
    ingin dihindari.
  - Style paragraf "TOC1" (dipakai tiap entri field di atas) SEKARANG
    disalin field-demi-field dari definisi "toc 1" asli di F.docx (lihat
    komentar lengkap di _ensure_toc_styles): TIDAK bold, TIDAK ada
    indentasi tambahan (hanging/right indent), tab kiri di 720 twips +
    tab kanan dot-leader di batas kanan area konten — persis F.docx.
    Sebelumnya (v6) style ini masih bold & pakai hanging-indent ala
    "Modify Style manual", sehingga tampilannya BEDA dari F.docx.
  - Field mengacu ke style "Judul" dan "Pasal" — SEMUA level 1
    (flat, tanpa indentasi bertingkat), sesuai style yang diterapkan
    Engine10 (StyleFinalizerEngine) pada judul halaman (Daftar Isi/Prakata/
    Pendahuluan/Bibliografi), Pasal/Subpasal berbahasa Indonesia, dan
    Lampiran (ANNEX) — persis pola yang terlihat di F.docx: "1  Ruang
    lingkup", "4.1  Umum", "Lampiran A (informatif) ..." semuanya tampil
    sebagai entri Daftar Isi.
  - Engine10 TIDAK LAGI memaksa halaman ini tetap kosong (lihat perubahan
    terkait pada engine10.py — fungsi _enforce_empty_daftar_isi_page tidak
    lagi dipanggil) supaya field TOC yang disisipkan di sini tidak dihapus
    isinya di tahap akhir pipeline.
  - Bagian lain (header/footer, page size/margin, romawi nomor halaman,
    posisi penyisipan setelah halaman Hak Cipta) TIDAK diubah.
"""

import re
import zipfile
from lxml import etree

NS_W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL_HEADER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
REL_FOOTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'

HDR_NS = ' '.join([
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"',
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
    'xmlns:o="urn:schemas-microsoft-com:office:office"',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
    'xmlns:v="urn:schemas-microsoft-com:vml"',
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
    'xmlns:w10="urn:schemas-microsoft-com:office:word"',
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"',
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"',
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"',
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
    'mc:Ignorable="w14 w15"',
])

def cm_to_twips(cm): return int(cm * 567)
def pt_to_hpts(pt):  return int(pt * 2)

def _esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def _run(text, bold=True, size_pt=12, italic=False):
    sz = pt_to_hpts(size_pt)
    b  = '<w:b/>' if bold else ''
    i  = '<w:i/>' if italic else ''
    return (
        f'<w:r><w:rPr>'
        f'<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'{b}{i}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
        f'</w:rPr><w:t xml:space="preserve">{_esc(text)}</w:t></w:r>'
    )

def _field_run(field, bold=True, size_pt=10):
    sz = pt_to_hpts(size_pt)
    b  = '<w:b/>' if bold else ''
    rpr = (
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'{b}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
        f'</w:rPr>'
    )
    return (
        f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r>{rpr}<w:instrText xml:space="preserve"> {field} </w:instrText></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>'
    )

def _build_header(title, align):
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:hdr {HDR_NS}>'
        f'<w:p><w:pPr>'
        f'<w:jc w:val="{align}"/>'
        f'<w:spacing w:before="0" w:after="0"/>'
        f'</w:pPr>{_run(title, bold=True, size_pt=12)}</w:p>'
        f'</w:hdr>'
    )

def _build_footer(copyright_text, pw, lm, rm):
    center = (pw - lm - rm) // 2
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:ftr {HDR_NS}>'
        f'<w:p><w:pPr>'
        f'<w:jc w:val="left"/>'
        f'<w:spacing w:before="0" w:after="0"/>'
        f'<w:tabs><w:tab w:val="center" w:pos="{center}"/></w:tabs>'
        f'</w:pPr>'
        f'{_run(copyright_text, bold=True, size_pt=10)}'
        f'<w:r><w:tab/></w:r>'
        f'{_field_run("PAGE", bold=True, size_pt=10)}'
        f'</w:p>'
        f'</w:ftr>'
    )


# ─────────────────────────────────────────────────────────────────────────────
# DAFTAR ISI BUILDER
# Field TOC asli Word — mengacu ke style "Judul" (judul halaman: Daftar Isi/
# Prakata/Pendahuluan/Bibliografi), "Pasal" (Pasal & Subpasal berbahasa
# Indonesia) dan "ANNEX" (judul Lampiran), SEMUA didaftarkan sebagai level 1
# supaya tampilannya flat (tanpa indentasi bertingkat) — sama seperti pada
# dokumen referensi (mis. F.docx).
#
# PENTING — pemisah SEMICOLON (;), BUKAN koma (,):
# Dipelajari langsung dari field TOC asli di F.docx (" TOC \h \z \t
# "Base_Heading;1;ANNEX;1;ANNEX Head;1;Biblio Title;1;Judul;1;Pasal;1" ") —
# daftar style pada switch \t dipisahkan titik koma. Ini BUKAN kebetulan:
# pemisah daftar (list separator) untuk switch field Word mengikuti
# pengaturan Regional Windows si pengguna, dan pada locale Indonesia
# (id-ID) — sama seperti locale Eropa lain — pemisahnya adalah TITIK KOMA,
# bukan koma. Field yang ditulis dengan koma bisa gagal di-parse / gagal
# menghasilkan entri apa pun ("Error! No table of contents entries
# found.") pada mesin Word dengan Regional Settings Indonesia. Memakai
# titik koma (sama seperti F.docx) menghindari masalah ini.
# ─────────────────────────────────────────────────────────────────────────────
_TOC_FIELD_INSTR = 'TOC \\h \\z \\t "Judul;1;Pasal;1"'
_TOC_PLACEHOLDER = (
    'Klik kanan pada teks ini lalu pilih "Update Field" '
    '(atau tekan Ctrl+A kemudian F9) untuk menampilkan Daftar Isi.'
)


def _build_di_elements(hdr_odd, hdr_even, ftr_odd, ftr_even):
    """
    Return list of raw XML strings untuk paragraf DI + inline sectPr.

    Halaman Daftar Isi berisi judul "Daftar Isi" (rata tengah, bold),
    2 paragraf kosong (spasi), lalu SATU paragraf berisi field TOC asli
    Word (belum di-update — begin/instrText/separate/placeholder/end
    semuanya dalam satu paragraf, persis seperti hasil Insert > Table of
    Contents manual di Word). Field ini otomatis ter-update & terisi penuh
    (entri + dot leader + nomor halaman + hyperlink) begitu dokumen dibuka
    di Word, karena <w:updateFields w:val="true"/> sudah diset di
    settings.xml (lihat _ensure_update_fields).
    """
    top  = cm_to_twips(3);   bottom = cm_to_twips(2)
    left = cm_to_twips(3);   right  = cm_to_twips(2)
    pw   = cm_to_twips(21);  ph     = cm_to_twips(29.7)
    # Posisi tab kanan (dot leader) = lebar area konten (page width dikurangi
    # margin kiri & kanan), persis prinsip yang dipakai F.docx (tab kanan
    # diposisikan tepat di margin kanan area teks).
    TAB  = pw - left - right

    # Style penanda "jangan diterjemahkan" — style yang SAMA persis dipakai
    # engine6 untuk melindungi halaman Prakata/Pendahuluan dari mesin
    # terjemahan (engine9). Style ini tidak perlu didefinisikan di
    # styles.xml — engine9 hanya mengecek atribut w:pStyle/@w:val secara
    # mentah, jadi cukup ditandai di sini agar halaman Daftar Isi DIJAMIN
    # tidak pernah tersentuh/terisi apapun oleh proses terjemahan.
    NT = 'BSNNoTranslate'

    def title_p():
        return (
            f'<w:p><w:pPr>'
            f'<w:pStyle w:val="{NT}"/>'
            f'<w:jc w:val="center"/>'
            f'<w:spacing w:before="0" w:after="0"/>'
            f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{TAB}"/></w:tabs>'
            f'</w:pPr>'
            f'{_run("Daftar Isi", bold=True, size_pt=12, italic=False)}'
            f'</w:p>'
        )

    def empty_p():
        return (
            f'<w:p><w:pPr>'
            f'<w:pStyle w:val="{NT}"/>'
            f'<w:spacing w:before="0" w:after="0"/>'
            f'</w:pPr></w:p>'
        )

    def toc_field_p():
        sz  = pt_to_hpts(11)
        rpr = (
            f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        )
        return (
            f'<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
            # Field TOC dibuat sebagai field Word murni. JANGAN menaruh
            # kalimat placeholder/cached result di antara separate-end,
            # karena Word dapat mempertahankannya sebagai hasil TOC lama.
            # Setelah dibuka, w:updateFields=true + w:dirty=true akan
            # memaksa Word menghitung ulang entri, dot leader, hyperlink,
            # dan nomor halaman.
            f'<w:r>{rpr}<w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
            f'<w:r>{rpr}<w:instrText xml:space="preserve"> {_TOC_FIELD_INSTR} </w:instrText></w:r>'
            f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>'
            f'</w:p>'
        )

    def sect_p():
        return (
            f'<w:p><w:pPr><w:sectPr>'
            f'<w:headerReference w:type="default" r:id="{hdr_odd}"/>'
            f'<w:headerReference w:type="even" r:id="{hdr_even}"/>'
            f'<w:footerReference w:type="default" r:id="{ftr_odd}"/>'
            f'<w:footerReference w:type="even" r:id="{ftr_even}"/>'
            f'<w:type w:val="nextPage"/>'
            f'<w:pgSz w:w="{pw}" w:h="{ph}"/>'
            f'<w:pgMar w:top="{top}" w:right="{right}" w:bottom="{bottom}" '
            f'w:left="{left}" w:header="709" w:footer="595" w:gutter="0"/>'
            f'<w:mirrorMargins/>'
            f'<w:pgNumType w:fmt="lowerRoman" w:start="1"/>'
            f'</w:sectPr></w:pPr></w:p>'
        )

    xmls = (
        [title_p(), empty_p(), empty_p(), toc_field_p()]
        + [sect_p()]
    )
    return xmls


# ─────────────────────────────────────────────────────────────────────────────
# STYLE "TOC1" — dipelajari & disalin PERSIS dari definisi style "toc 1" di
# F.docx (word/styles.xml), supaya hasil tampilan Daftar Isi identik:
#
#   <w:style w:type="paragraph" w:styleId="TOC1">
#     <w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
#     <w:uiPriority w:val="39"/>
#     <w:pPr>
#       <w:tabs>
#         <w:tab w:val="left" w:pos="720"/>
#         <w:tab w:val="right" w:leader="dot" w:pos="<lebar area konten>"/>
#       </w:tabs>
#       <w:suppressAutoHyphens/>
#       <w:spacing w:after="120" w:line="240" w:lineRule="auto"/>
#     </w:pPr>
#     <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr>
#   </w:style>
#
# Catatan penting (beda dari versi sebelumnya, sekarang diperbaiki):
#   - TIDAK bold (F.docx TIDAK menebalkan entri Daftar Isi — hanya judul
#     halaman "Daftar Isi" yang bold, itu style "Judul", bukan "TOC1").
#   - TIDAK ada w:ind (hanging/right indent) — F.docx murni memakai dua
#     tab stop (kiri 720 twips & kanan dot-leader), tanpa indentasi
#     tambahan apa pun.
#   - TIDAK ada w:jc override — mengikuti default (inherit dari Normal).
#   - Posisi tab kanan (dot leader) DIHITUNG dinamis dari lebar area
#     konten dokumen ini (page width - margin kiri - margin kanan),
#     mengikuti prinsip yang sama dengan F.docx (tab kanan pas di margin
#     kanan area teks) — bukan angka mentah hasil copy dari F.docx, karena
#     margin dokumen ini berbeda dari F.docx.
#   - w:styleId TETAP "TOC1" (dipakai field TOC & referensi lain di file
#     ini) meski w:name aslinya "toc 1" (huruf kecil, sesuai F.docx).
#
# Hanya SATU level dipakai (Judul, Pasal & ANNEX sama-sama TOC level 1),
# jadi hanya style "TOC1" yang diperlukan.
# ─────────────────────────────────────────────────────────────────────────────
def _ensure_toc_styles(files: dict) -> None:
    key = 'word/styles.xml'
    if key not in files:
        return
    styles_xml = files[key].decode('utf-8')

    pw   = cm_to_twips(21)
    left = cm_to_twips(3)
    right = cm_to_twips(2)
    TAB  = pw - left - right   # posisi tab kanan (dot leader) = lebar konten

    toc1_xml = (
        f'<w:style w:type="paragraph" w:styleId="TOC1">'
        f'<w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
        f'<w:uiPriority w:val="39"/>'
        f'<w:pPr>'
        f'<w:tabs>'
        f'<w:tab w:val="left" w:pos="720"/>'
        f'<w:tab w:val="right" w:leader="dot" w:pos="{TAB}"/>'
        f'</w:tabs>'
        f'<w:suppressAutoHyphens/>'
        f'<w:spacing w:after="120" w:line="240" w:lineRule="auto"/>'
        f'</w:pPr>'
        f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
        f'</w:style>'
    )

    # Buang definisi TOC1/TOC2 lama (jika ada dari hasil generate sebelumnya)
    # supaya selalu memakai spesifikasi terbaru, lalu sisipkan TOC1 yang baru.
    styles_xml = re.sub(
        r'<w:style\b[^>]*w:styleId="TOC[12]"[^>]*>.*?</w:style>',
        '', styles_xml, flags=re.DOTALL
    )
    styles_xml = styles_xml.replace('</w:styles>', toc1_xml + '</w:styles>')
    files[key] = styles_xml.encode('utf-8')


def _ensure_update_fields(files: dict) -> None:
    """Paksa Word melakukan refresh field saat dokumen dibuka.

    Penting: jangan hanya menambahkan <w:updateFields>. Dokumen sumber
    kadang sudah memiliki elemen tersebut dengan nilai false. Pada kondisi
    itu kode lama melakukan ``return`` sehingga TOC hasil akhir tetap
    memakai cached result lama. Di sini nilainya SELALU dipaksa true dan
    flag yang melarang update field dihapus.
    """
    key = 'word/settings.xml'
    if key not in files:
        return

    settings_xml = files[key].decode('utf-8')

    # Jangan biarkan setting ini menghalangi refresh field.
    settings_xml = re.sub(
        r'<w:doNotUpdateFields\b[^>]*/>',
        '',
        settings_xml,
        flags=re.DOTALL,
    )

    # Jika updateFields sudah ada, ganti seluruh elemennya menjadi true.
    settings_xml, n = re.subn(
        r'<w:updateFields\b[^>]*/>',
        '<w:updateFields w:val="true"/>',
        settings_xml,
        count=1,
        flags=re.DOTALL,
    )

    # Jika belum ada, sisipkan tepat setelah <w:settings>.
    if n == 0:
        m = re.search(r'<w:settings\b[^>]*>', settings_xml)
        if m:
            insert_at = m.end()
            settings_xml = (
                settings_xml[:insert_at]
                + '<w:updateFields w:val="true"/>'
                + settings_xml[insert_at:]
            )

    files[key] = settings_xml.encode('utf-8')


def _parse_elements(xml_list):
    elements = []
    for xml_str in xml_list:
        wrapped = (
            f'<root xmlns:w="{NS_W}" xmlns:r="{NS_R}" '
            f'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            f'{xml_str}</root>'
        )
        root = etree.fromstring(wrapped.encode('utf-8'))
        elements.extend(list(root))
    return elements


def _find_nth_section_paragraph_index(body, n=2):
    count = 0
    last_idx = None
    for i, child in enumerate(list(body)):
        if child.tag == f'{{{NS_W}}}p':
            pPr = child.find(f'{{{NS_W}}}pPr')
            if pPr is not None and pPr.find(f'{{{NS_W}}}sectPr') is not None:
                count += 1
                last_idx = i
                if count == n:
                    return i, count
    if last_idx is not None:
        return last_idx, count
    return 0, 0


class DaftarIsiEngine:
    """
    Engine untuk menyisipkan halaman Daftar Isi setelah halaman Hak Cipta (page 2).
    Halaman hanya berisi judul "Daftar Isi" — TIDAK ADA isi/entri apapun,
    persis seperti halaman Pendahuluan.
    Penomoran: Romawi (i, ii, iii, ...).
    """

    def insert(self, input_docx: str, output_docx: str,
               doc_title: str = 'SNI ISO XXXXX:20XX',
               copyright_text: str = '©BSN 20XX') -> tuple[bool, str]:
        try:
            # 1. Baca input
            with zipfile.ZipFile(input_docx, 'r') as z:
                files = {n: z.read(n) for n in z.namelist()}

            # 1b. Pastikan style "TOC 1"/"TOC 2" ada & field auto-update saat dibuka
            _ensure_toc_styles(files)
            _ensure_update_fields(files)

            # 3. Hitung rId dan file number berikutnya
            rels_xml = files['word/_rels/document.xml.rels'].decode('utf-8')
            max_rid  = max((int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml)), default=0)
            max_hf   = max((int(n) for n in re.findall(
                            r'Target="(?:header|footer)(\d+)\.xml"', rels_xml)), default=0)

            n = max_rid + 1
            h = max_hf + 1

            rid_ho = f'rId{n}';   rid_he = f'rId{n+1}'
            rid_fo = f'rId{n+2}'; rid_fe = f'rId{n+3}'
            f_ho = f'header{h}.xml';     f_he = f'header{h+1}.xml'
            f_fo = f'footer{h+2}.xml';   f_fe = f'footer{h+3}.xml'

            # 4. Build header/footer
            pw = cm_to_twips(21); lm = cm_to_twips(3); rm = cm_to_twips(2)
            files[f'word/{f_ho}'] = _build_header(doc_title, 'right').encode('utf-8')
            files[f'word/{f_he}'] = _build_header(doc_title, 'left').encode('utf-8')
            files[f'word/{f_fo}'] = _build_footer(copyright_text, pw, lm, rm).encode('utf-8')
            files[f'word/{f_fe}'] = _build_footer(copyright_text, pw, lm, rm).encode('utf-8')

            # 5. Update rels
            new_rels = (
                f'<Relationship Id="{rid_ho}" Type="{REL_HEADER}" Target="{f_ho}"/>\n'
                f'<Relationship Id="{rid_he}" Type="{REL_HEADER}" Target="{f_he}"/>\n'
                f'<Relationship Id="{rid_fo}" Type="{REL_FOOTER}" Target="{f_fo}"/>\n'
                f'<Relationship Id="{rid_fe}" Type="{REL_FOOTER}" Target="{f_fe}"/>\n'
            )
            files['word/_rels/document.xml.rels'] = rels_xml.replace(
                '</Relationships>', new_rels + '</Relationships>'
            ).encode('utf-8')

            # 6. Update Content Types
            ct_xml = files['[Content_Types].xml'].decode('utf-8')
            hdr_ct = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
            ftr_ct = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
            adds = ''
            for fname, ct in [(f_ho, hdr_ct), (f_he, hdr_ct), (f_fo, ftr_ct), (f_fe, ftr_ct)]:
                part = f'/word/{fname}'
                if part not in ct_xml:
                    adds += f'<Override PartName="{part}" ContentType="{ct}"/>\n'
            if adds:
                ct_xml = ct_xml.replace('</Types>', adds + '</Types>')
            files['[Content_Types].xml'] = ct_xml.encode('utf-8')

            # 7. Parse document.xml
            tree = etree.fromstring(files['word/document.xml'])
            body = tree.find(f'{{{NS_W}}}body')

            # 8. Cari posisi insert (setelah sectPr ke-2 / Hak Cipta)
            hakcip_idx, found = _find_nth_section_paragraph_index(body, n=2)
            if found < 2:
                hakcip_idx, _ = _find_nth_section_paragraph_index(body, n=1)

            insert_pos = hakcip_idx + 1

            # 9. Build & insert DI elements (field TOC native Word)
            di_xmls = _build_di_elements(rid_ho, rid_he, rid_fo, rid_fe)
            di_els  = _parse_elements(di_xmls)
            for offset, el in enumerate(di_els):
                body.insert(insert_pos + offset, el)

            # 10. Serialisasi
            files['word/document.xml'] = etree.tostring(
                tree, xml_declaration=True, encoding='UTF-8', standalone=True
            )

            # 11. Tulis output
            with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as zout:
                for prio in ['[Content_Types].xml', '_rels/.rels']:
                    if prio in files:
                        zout.writestr(prio, files[prio])
                for name, data in files.items():
                    if name not in ('[Content_Types].xml', '_rels/.rels'):
                        zout.writestr(name, data)

            return True, output_docx

        except Exception as e:
            import traceback
            return False, f'DaftarIsiEngine Error: {str(e)}\n{traceback.format_exc()}'
