"""
Engine7: InfoPendukungEngine
=============================
Menyisipkan halaman "Informasi pendukung terkait perumus standar"
sebagai SECTION BARU setelah halaman Bibliografi (= akhir dokumen isi).

Aturan:
- Section baru (sectPr tersendiri) — MENGGANTIKAN sectPr terakhir dokumen.
- Margin : Top 3 cm | Left 3 cm | Bottom 2 cm | Right 2 cm.
- Header/Footer : BERSIH (blank paragraph, tanpa teks/logo/nomor halaman).
- Penomoran halaman : TIDAK ADA (tidak ada pgNumType).
- Isi : persis seperti dokumen referensi BSN.
  Daftar bernomor [1]–[5] disimulasikan secara visual (hanging indent).
  Tabel "Susunan keanggotaan" : 3 kolom, tanpa border, persis dokumen.
"""

import re
import zipfile
from lxml import etree

NS_W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_R  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
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


# ─────────────────────────────────────────────────────────────────────────────
# UNIT HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def cm_to_twips(cm): return int(cm * 567)
def pt_to_hpts(pt):  return int(pt * 2)

def _esc(t: str) -> str:
    return (t.replace('&','&amp;').replace('<','&lt;')
             .replace('>','&gt;').replace('"','&quot;'))


# ─────────────────────────────────────────────────────────────────────────────
# BLANK HEADER / FOOTER
# ─────────────────────────────────────────────────────────────────────────────
def _blank_header() -> bytes:
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:hdr {HDR_NS}>'
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>'
        f'</w:hdr>'
    ).encode('utf-8')

def _blank_footer() -> bytes:
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:ftr {HDR_NS}>'
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>'
        f'</w:ftr>'
    ).encode('utf-8')


# ─────────────────────────────────────────────────────────────────────────────
# XML PARAGRAPH BUILDERS
# ─────────────────────────────────────────────────────────────────────────────
FONT = '<w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>'

def _rpr(bold=False, size_pt=11, color=None) -> str:
    b  = '<w:b/>' if bold else ''
    sz = pt_to_hpts(size_pt)
    cl = f'<w:color w:val="{color}"/>' if color else ''
    return f'<w:rPr>{FONT}{b}{cl}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/><w:noproof/></w:rPr>'

def _run(text: str, bold=False, size_pt=11, color=None) -> str:
    return f'<w:r>{_rpr(bold,size_pt,color)}<w:t xml:space="preserve">{_esc(text)}</w:t></w:r>'

def _para(runs_xml: str, align='left', space_after=0,
          ind_left=0, ind_hanging=0, ind_first=0) -> str:
    jc  = f'<w:jc w:val="{align}"/>' if align != 'left' else ''
    sa  = f'<w:spacing w:before="0" w:after="{space_after}"/>'
    ind = ''
    if ind_left or ind_hanging or ind_first:
        parts = []
        if ind_left:    parts.append(f'w:left="{ind_left}"')
        if ind_hanging: parts.append(f'w:hanging="{ind_hanging}"')
        if ind_first:   parts.append(f'w:firstLine="{ind_first}"')
        ind = f'<w:ind {" ".join(parts)}/>'
    return (
        f'<w:p><w:pPr>{sa}{jc}{ind}'
        f'<w:rPr>{FONT}<w:noproof/></w:rPr>'
        f'</w:pPr>{runs_xml}</w:p>'
    )

def _empty() -> str:
    return _para('')


# ─────────────────────────────────────────────────────────────────────────────
# TABLE BUILDER  (persis dokumen referensi)
# Kolom: Label | Separator | Nilai
# Width : 1829  |    717    |  3261  (dxa)
# ─────────────────────────────────────────────────────────────────────────────
def _table_row(label: str, value_paragraphs: list, ind_left_col3: int = 0) -> str:
    """
    Buat satu baris tabel.
    value_paragraphs: list of strings (teks untuk tiap paragraf di kolom nilai).
    Kolom label & separator masing-masing 1 paragraf; kolom nilai bisa multi-paragraf.
    """
    def tc_single(w_dxa: int, text: str, ind_left: int) -> str:
        ind = f'<w:ind w:left="{ind_left}"/>'
        return (
            f'<w:tc>'
            f'<w:tcPr><w:tcW w:w="{w_dxa}" w:type="dxa"/>'
            f'<w:tcMar>'
            f'<w:top w:w="0" w:type="dxa"/><w:left w:w="115" w:type="dxa"/>'
            f'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="115" w:type="dxa"/>'
            f'</w:tcMar></w:tcPr>'
            f'<w:p><w:pPr><w:spacing w:after="0"/>{ind}'
            f'<w:rPr>{FONT}<w:noproof/></w:rPr></w:pPr>'
            f'<w:r><w:rPr>{FONT}<w:noproof/></w:rPr>'
            f'<w:t xml:space="preserve">{_esc(text)}</w:t></w:r></w:p>'
            f'</w:tc>'
        )

    def tc_multi(w_dxa: int, paragraphs: list) -> str:
        paras_xml = ''
        for text in paragraphs:
            paras_xml += (
                f'<w:p><w:pPr><w:spacing w:after="0"/>'
                f'<w:rPr>{FONT}<w:noproof/></w:rPr></w:pPr>'
                f'<w:r><w:rPr>{FONT}<w:noproof/></w:rPr>'
                f'<w:t xml:space="preserve">{_esc(text)}</w:t></w:r></w:p>'
            )
        return (
            f'<w:tc>'
            f'<w:tcPr><w:tcW w:w="{w_dxa}" w:type="dxa"/>'
            f'<w:tcMar>'
            f'<w:top w:w="0" w:type="dxa"/><w:left w:w="115" w:type="dxa"/>'
            f'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="115" w:type="dxa"/>'
            f'</w:tcMar></w:tcPr>'
            f'{paras_xml}'
            f'</w:tc>'
        )

    return (
        f'<w:tr>'
        + tc_single(1829, label, 425)
        + tc_single(717,  ':',   425)
        + tc_multi(3261, value_paragraphs)
        + f'</w:tr>'
    )

def _build_table() -> str:
    """Tabel Susunan keanggotaan — 4 baris utama, Anggota dengan 8 anggota bernomor."""
    # Baris Anggota: nilai berupa 8 baris bernomor 1–8
    anggota_values = [f'{i}   xxxxxxxxx' for i in range(1, 9)]

    rows = [
        ('Ketua',       ['xxxxxxx']),
        ('Wakil Ketua', ['xxxxxxx']),
        ('Sekretaris',  ['xxxxxxx']),
        ('Anggota',     anggota_values),
    ]
    tbl_pr = (
        '<w:tblPr>'
        '<w:tblW w:w="5807" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblLook w:val="0400" w:firstRow="0" w:lastRow="0" '
        'w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        '</w:tblPr>'
        '<w:tblGrid>'
        '<w:gridCol w:w="1829"/><w:gridCol w:w="717"/><w:gridCol w:w="3261"/>'
        '</w:tblGrid>'
    )
    return '<w:tbl>' + tbl_pr + ''.join(_table_row(l, v) for l, v in rows) + '</w:tbl>'


# ─────────────────────────────────────────────────────────────────────────────
# FULL CONTENT BUILDER
# ─────────────────────────────────────────────────────────────────────────────
def _build_content() -> list:
    """
    Kembalikan list string XML (paragraf + tabel) untuk halaman info pendukung.
    Penomoran [1]–[5] disimulasikan dengan teks inline + hanging indent.
    """
    xmls = []

    # ── Judul ────────────────────────────────────────────────────────────
    xmls.append(_para(
        _run('Informasi pendukung terkait perumus standar', bold=True, size_pt=12),
        align='center'
    ))
    xmls.append(_empty())
    xmls.append(_empty())
    xmls.append(_empty())

    # Pending label width: left=567, hanging=567
    # Label "[n]" kemudian teks (visual identical dengan numPr fmt=[%1])
    def list_item(n: int, text: str, bold=True) -> str:
        label = f'[{n}]'
        # label dalam satu run, lalu tab, lalu teks
        runs = (
            _run(label, bold=bold, size_pt=11)
            + f'<w:r>{_rpr(bold,11)}<w:tab/></w:r>'
            + _run(text,  bold=bold, size_pt=11)
        )
        return _para(runs, ind_left=567, ind_hanging=567)

    def sub_item(text: str, ind_left=567) -> str:
        return _para(_run(text, size_pt=11), ind_left=ind_left)

    # ── [1] Komtek perumus SNI ────────────────────────────────────────────
    xmls.append(list_item(1, 'Komtek perumus SNI'))
    xmls.append(_para(
        _run('Komite Teknis xx-yy zzzzzz', size_pt=11),
        ind_first=567
    ))
    xmls.append(_empty())

    # ── [2] Susunan keanggotaan ───────────────────────────────────────────
    xmls.append(list_item(2, 'Susunan keanggotaan Komite perumus SNI'))
    xmls.append(_build_table())   # tabel langsung (bukan paragraf)
    xmls.append(_empty())

    # ── [3] Konseptor ─────────────────────────────────────────────────────
    xmls.append(list_item(3, 'Konseptor terjemahan rancangan SNI'))
    xmls.append(sub_item('xxxxxxxxxxxxxxxxxxxxx'))
    xmls.append(sub_item('yyyyyyyyyyyyyyyyyyyyy'))
    xmls.append(_empty())

    # ── [4] Editor ───────────────────────────────────────────────────────
    xmls.append(list_item(4, 'Editor rancangan SNI'))
    xmls.append(sub_item('xxxxxxxxxxxxxxxxxxxxx'))
    xmls.append(sub_item('yyyyyyyyyyyyyyyyyyyyy'))
    xmls.append(_empty())

    # ── [5] Sekretariat ──────────────────────────────────────────────────
    xmls.append(list_item(5, 'Sekretariat pengelola Komtek perumus SNI'))
    xmls.append(sub_item('xxxxxxxxxxxxxxxxxxxxx'))
    xmls.append(sub_item('yyyyyyyyyyyyyyyyyyyyy'))

    return xmls


# ─────────────────────────────────────────────────────────────────────────────
# PARSE HELPER
# ─────────────────────────────────────────────────────────────────────────────
def _parse_elements(xml_list: list) -> list:
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


# ─────────────────────────────────────────────────────────────────────────────
# ENGINE CLASS
# ─────────────────────────────────────────────────────────────────────────────
class InfoPendukungEngine:
    """
    Menambahkan halaman "Informasi pendukung terkait perumus standar"
    sebagai section baru setelah Bibliografi (akhir dokumen isi).

    Strategi:
    - Ambil sectPr terakhir (body sectPr), jadikan sectPr section sebelumnya
      dengan cara menyisipkannya sebagai inline sectPr pada paragraf
      tepat sebelum konten baru.
    - Konten baru + sectPr baru (blank header/footer, margin custom) ditambahkan
      ke akhir body sebelum body's final sectPr (yang kini sudah dipindah).
    """

    def append(
        self,
        input_docx:  str,
        output_docx: str,
    ) -> tuple[bool, str]:
        try:
            # 1. Baca input
            with zipfile.ZipFile(input_docx, 'r') as z:
                files = {n: z.read(n) for n in z.namelist()}

            # 2. Hitung rId & file number baru
            rels_xml = files['word/_rels/document.xml.rels'].decode('utf-8')
            max_rid = max((int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml)), default=0)
            max_hf  = max((int(n) for n in re.findall(
                r'Target="(?:header|footer)(\d+)\.xml"', rels_xml)), default=0)

            n = max_rid + 1
            h = max_hf  + 1
            rid_hd = f'rId{n}'
            rid_ft = f'rId{n+1}'
            f_hd   = f'header{h}.xml'
            f_ft   = f'footer{h+1}.xml'

            # 3. Simpan blank header/footer
            files[f'word/{f_hd}'] = _blank_header()
            files[f'word/{f_ft}'] = _blank_footer()

            # 4. Update rels
            new_rels = (
                f'<Relationship Id="{rid_hd}" Type="{REL_HEADER}" Target="{f_hd}"/>\n'
                f'<Relationship Id="{rid_ft}" Type="{REL_FOOTER}" Target="{f_ft}"/>\n'
            )
            files['word/_rels/document.xml.rels'] = rels_xml.replace(
                '</Relationships>', new_rels + '</Relationships>'
            ).encode('utf-8')

            # 5. Update Content Types
            ct_xml  = files['[Content_Types].xml'].decode('utf-8')
            hdr_ct  = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
            ftr_ct  = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
            adds = ''
            for fname, ct in [(f_hd, hdr_ct), (f_ft, ftr_ct)]:
                part = f'/word/{fname}'
                if part not in ct_xml:
                    adds += f'<Override PartName="{part}" ContentType="{ct}"/>\n'
            if adds:
                ct_xml = ct_xml.replace('</Types>', adds + '</Types>')
            files['[Content_Types].xml'] = ct_xml.encode('utf-8')

            # 6. Parse document.xml
            tree = etree.fromstring(files['word/document.xml'])
            body = tree.find(f'{{{NS_W}}}body')

            # 7. Ambil body's final sectPr (anak langsung terakhir body)
            #    Ini adalah sectPr section terakhir (Isi/Bibliografi).
            #    Kita pertahankan sebagai-is — section baru kita tambahkan
            #    sebagai inline sectPr pada paragraf kosong, lalu diikuti
            #    konten info pendukung, lalu body's sectPr (blank).
            #
            #    Namun cara paling bersih: GANTIKAN body's final sectPr
            #    menjadi sectPr section info pendukung (blank hdr/ftr, margin khusus),
            #    dan sisipkan inline sectPr untuk section sebelumnya TEPAT
            #    sebelum konten baru.
            #
            #    Langkah:
            #    a. Cari body's final sectPr element
            #    b. Clone-nya sebagai inline sectPr → sisipkan sebagai
            #       paragraf kosong di akhir body (sebelum final sectPr)
            #    c. Ganti body's final sectPr dengan sectPr info pendukung
            #    d. Sisipkan konten info pendukung sebelum body's (modified) sectPr

            body_children = list(body)

            # Temukan body's final sectPr (langsung child body, bukan inline)
            final_sectPr = body.find(f'{{{NS_W}}}sectPr')
            if final_sectPr is None:
                # Coba ambil dari paragraf terakhir
                for child in reversed(body_children):
                    if child.tag == f'{{{NS_W}}}p':
                        pPr = child.find(f'{{{NS_W}}}pPr')
                        if pPr is not None:
                            sp = pPr.find(f'{{{NS_W}}}sectPr')
                            if sp is not None:
                                final_sectPr = sp
                                break

            # Clone sectPr isi sebagai inline sectPr (untuk paragraf pemisah)
            import copy
            if final_sectPr is not None:
                inline_sectPr_clone = copy.deepcopy(final_sectPr)
            else:
                # Fallback: buat sectPr minimal
                inline_sectPr_clone = etree.fromstring(
                    f'<w:sectPr xmlns:w="{NS_W}"><w:type w:val="nextPage"/></w:sectPr>'
                )

            # Paragraf pemisah: paragraf kosong dengan inline sectPr (clone isi)
            sep_para_xml = (
                f'<w:p xmlns:w="{NS_W}"><w:pPr>'
                f'<w:spacing w:before="0" w:after="0"/>'
                f'</w:pPr></w:p>'
            )
            sep_para = etree.fromstring(sep_para_xml)
            sep_pPr  = sep_para.find(f'{{{NS_W}}}pPr')
            sep_pPr.append(inline_sectPr_clone)

            # 8. Margin & sectPr untuk section info pendukung
            top    = cm_to_twips(3)
            left   = cm_to_twips(3)
            bottom = cm_to_twips(2)
            right  = cm_to_twips(2)
            pw     = cm_to_twips(21)
            ph     = cm_to_twips(29.7)

            new_sectPr_xml = (
                f'<w:sectPr xmlns:w="{NS_W}" '
                f'xmlns:r="{NS_R}">'
                f'<w:headerReference w:type="default" r:id="{rid_hd}"/>'
                f'<w:headerReference w:type="even"    r:id="{rid_hd}"/>'
                f'<w:headerReference w:type="first"   r:id="{rid_hd}"/>'
                f'<w:footerReference w:type="default" r:id="{rid_ft}"/>'
                f'<w:footerReference w:type="even"    r:id="{rid_ft}"/>'
                f'<w:footerReference w:type="first"   r:id="{rid_ft}"/>'
                f'<w:type w:val="nextPage"/>'
                f'<w:pgSz w:w="{pw}" w:h="{ph}"/>'
                f'<w:pgMar w:top="{top}" w:right="{right}" w:bottom="{bottom}" '
                f'w:left="{left}" w:header="709" w:footer="595" w:gutter="0"/>'
                f'</w:sectPr>'
            )
            new_sectPr_el = etree.fromstring(new_sectPr_xml)

            # 9. Sisipkan: paragraf pemisah + konten + ganti/tambah sectPr

            # Cari index body's final sectPr atau akhir body
            final_sectPr_body = body.find(f'{{{NS_W}}}sectPr')  # langsung child body
            if final_sectPr_body is not None:
                idx = list(body).index(final_sectPr_body)
                # Sisipkan sep_para sebelum final_sectPr
                body.insert(idx, sep_para)
                # Ganti final_sectPr dengan new_sectPr
                body.remove(final_sectPr_body)
                # Parse & insert konten sebelum final sectPr position
                content_els = _parse_elements(_build_content())
                for i, el in enumerate(content_els):
                    body.insert(idx + 1 + i, el)
                # Tambahkan new_sectPr sebagai child body terakhir
                body.append(new_sectPr_el)
            else:
                # Tidak ada body sectPr langsung — append saja
                body.append(sep_para)
                for el in _parse_elements(_build_content()):
                    body.append(el)
                # Tambah paragraf dengan inline sectPr baru
                p_last_xml = f'<w:p xmlns:w="{NS_W}" xmlns:r="{NS_R}"><w:pPr></w:pPr></w:p>'
                p_last = etree.fromstring(p_last_xml)
                p_last.find(f'{{{NS_W}}}pPr').append(new_sectPr_el)
                body.append(p_last)

            # 10. Serialisasi & tulis output
            files['word/document.xml'] = etree.tostring(
                tree, xml_declaration=True, encoding='UTF-8', standalone=True
            )

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
            return False, f'InfoPendukungEngine Error: {str(e)}\n{traceback.format_exc()}'
    