"""
Engine5: DaftarIsiEngine (v4 - Native Word TOC Field)
=========================================================
Engine untuk membuat halaman Daftar Isi sesuai standar BSN/SNI.

v4 (Native TOC field):
  - Isi Daftar Isi TIDAK LAGI diketik manual / diekstrak sebagai teks statis.
  - Sebagai gantinya, disisipkan field TOC asli Word:
        { TOC \\t "Judul,1,Pasal,1" \\h \\z }
    persis seperti hasil "References > Table of Contents > Custom Table of
    Contents" di Word dengan opsi "Styles" dicentang HANYA untuk style
    "Judul" (TOC level 1) dan "Pasal" (TOC level 1) — opsi "Outline levels"
    DI-UNCHECK (sengaja TANPA switch \\u), supaya Daftar Isi HANYA berisi
    paragraf ber-style "Judul"/"Pasal" — bukan seluruh paragraf yang punya
    outline level (mis. heading asli "Heading 1"/"Heading 2" sebelum
    dikonversi, atau heading pada bagian berbahasa Inggris), yang sebelumnya
    ikut nyasar masuk Daftar Isi tanpa titik-titik/nomor halaman karena
    switch \\u membuat Word menyertakan SEMUA paragraf ber-outline-level
    (union dengan hasil \\t), bukan hanya yang ber-style Judul/Pasal.
  - Word yang menghitung isi + nomor halaman secara otomatis saat dokumen
    dibuka (bukan dihitung oleh Python), sehingga Daftar Isi selalu akurat
    mengikuti pagination final. Ditambahkan <w:updateFields w:val="true"/>
    pada settings.xml supaya field ini otomatis ter-update saat file dibuka
    di Word (tidak perlu klik kanan > Update Field secara manual).
  - Style paragraf "TOC 1" ditambahkan ke styles.xml (Arial 11 Bold, rata
    kiri, spacing after 6pt, single line spacing, outline level Body Text,
    indent left 0 / hanging 1,27 cm / right 0,88 cm — sesuai spesifikasi
    Modify Style yang dibuat manual di Word oleh pengguna) supaya tampilan
    entri Daftar Isi konsisten walau dibuka di komputer lain.
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
# DAFTAR ISI BUILDER (native Word TOC field)
# ─────────────────────────────────────────────────────────────────────────────
_TOC_FIELD_INSTR = 'TOC \\t "Judul,1,Pasal,1" \\h \\z'
_TOC_PLACEHOLDER = (
    'Klik kanan pada teks ini lalu pilih "Update Field" '
    '(atau tekan Ctrl+A kemudian F9) untuk menampilkan Daftar Isi.'
)


def _build_di_elements(hdr_odd, hdr_even, ftr_odd, ftr_even):
    """
    Return list of raw XML strings untuk paragraf DI + inline sectPr.

    Isi Daftar Isi kini berupa field TOC asli Word (bukan teks statis),
    dibangun dari style "Judul" (TOC level 1) & "Pasal" (TOC level 1) saja —
    identik dengan Table of Contents Options: Styles dicentang hanya
    untuk Judul & Pasal (keduanya level 1), style lain di-uncheck,
    Show levels: 1.
    """
    TAB  = 9061
    top  = cm_to_twips(3);   bottom = cm_to_twips(2)
    left = cm_to_twips(3);   right  = cm_to_twips(2)
    pw   = cm_to_twips(21);  ph     = cm_to_twips(29.7)

    def title_p():
        return (
            f'<w:p><w:pPr>'
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
            f'<w:r>{rpr}<w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
            f'<w:r>{rpr}<w:instrText xml:space="preserve"> {_TOC_FIELD_INSTR} </w:instrText></w:r>'
            f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r>{rpr}<w:t xml:space="preserve">{_esc(_TOC_PLACEHOLDER)}</w:t></w:r>'
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
        [title_p(), empty_p(), empty_p(), empty_p()]
        + [toc_field_p()]
        + [sect_p()]
    )
    return xmls


# ─────────────────────────────────────────────────────────────────────────────
# STYLE "TOC 1" — sesuai Modify Style yang dibuat manual di Word:
#   Based on: Normal, Next: Normal, Font Arial 11 Bold, rata kiri (Left),
#   Outline level: Body Text, Spacing before 0pt / after 6pt / Single,
#   Indentation: Left 0 cm, Hanging 1,27 cm, Right 0,88 cm.
# Hanya SATU level dipakai (Judul & Pasal sama-sama TOC level 1), jadi
# hanya style "TOC 1" yang diperlukan.
# ─────────────────────────────────────────────────────────────────────────────
def _ensure_toc_styles(files: dict) -> None:
    key = 'word/styles.xml'
    if key not in files:
        return
    styles_xml = files[key].decode('utf-8')

    TAB     = 9061
    hang    = cm_to_twips(1.27)   # 720
    rindent = cm_to_twips(0.88)   # 498

    toc1_xml = (
        f'<w:style w:type="paragraph" w:customStyle="1" w:styleId="TOC1">'
        f'<w:name w:val="TOC 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
        f'<w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/>'
        f'<w:pPr>'
        f'<w:jc w:val="left"/>'
        f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{TAB}"/></w:tabs>'
        f'<w:spacing w:before="0" w:after="120" w:line="240" w:lineRule="auto"/>'
        f'<w:ind w:left="0" w:right="{rindent}" w:hanging="{hang}"/>'
        f'</w:pPr>'
        f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'<w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
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
    """Set <w:updateFields w:val="true"/> supaya field TOC otomatis
    ter-update (isi + nomor halaman) begitu dokumen dibuka di Word."""
    key = 'word/settings.xml'
    if key not in files:
        return
    settings_xml = files[key].decode('utf-8')
    if '<w:updateFields' in settings_xml:
        return
    m = re.search(r'<w:settings[^>]*>', settings_xml)
    if not m:
        return
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
    Isi Daftar Isi diambil OTOMATIS dari heading dokumen (sama dengan logika engine2).
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
