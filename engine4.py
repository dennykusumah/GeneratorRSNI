"""
Engine4: CoverPageEngine - SNI ISO Cover/Sampul Generator
===========================================================
Engine untuk membuat halaman sampul/cover sesuai standar BSN (SNI ISO).

Ketentuan Cover:
- Header: Logo SNI (kiri atas, 2.29cm x 3.23cm), "Standar Nasional Indonesia" (Arial 12 bold, rata kiri),
          Nomor SNI (kanan atas, Arial 14 bold, rata kanan), "(Ditetapkan oleh BSN Tahun XXXX)" (Arial 12 bold, rata kanan)
          Garis bawah header: 2.25 pt
- Footer: ICS dan nomor ICS (Arial 11 bold, kiri), Logo BSN (1.5cm x 7cm, kanan)
          Garis atas footer: 1.5 pt
- Isi: Judul bahasa Indonesia (Arial 18 bold, center)
        Judul bahasa Inggris (Arial 16 bold italic, center)
        Nomor standar acuan dalam kurung (Arial 12 bold, center)
- Margin: sama dengan dokumen isi (Top 3cm, Inside 3cm, Bottom 2cm, Outside 2cm)
- Cover adalah section TERSENDIRI tanpa nomor halaman.
"""

import os
import re
import zipfile


# ──────────────────────────────────────────────────────────
# UNIT CONVERSIONS
# ──────────────────────────────────────────────────────────
def cm_to_emu(cm):
    return int(cm * 360000)

def cm_to_twips(cm):
    return int(cm * 567)

def pt_to_half_pts(pt):
    return int(pt * 2)


# ──────────────────────────────────────────────────────────
# XML HELPER FUNCTIONS
# All namespaces declared at ROOT level - never inline.
# ──────────────────────────────────────────────────────────

# Full namespace declarations for header/footer root element
HDR_NAMESPACES = ' '.join([
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


def xml_escape(text):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def make_inline_image(rId, cx_emu, cy_emu, doc_id, name):
    """Inline image - NO namespace redeclarations (declared at root)."""
    return (
        f'<w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx_emu}" cy="{cy_emu}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{doc_id}" name="{xml_escape(name)}"/>'
        f'<wp:cNvGraphicFramePr/>'
        f'<a:graphic>'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic>'
        f'<pic:nvPicPr>'
        f'<pic:cNvPr id="0" name="{xml_escape(name)}"/>'
        f'<pic:cNvPicPr preferRelativeResize="0"/>'
        f'</pic:nvPicPr>'
        f'<pic:blipFill>'
        f'<a:blip r:embed="{rId}"/>'
        f'<a:srcRect/><a:stretch><a:fillRect/></a:stretch>'
        f'</pic:blipFill>'
        f'<pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx_emu}" cy="{cy_emu}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</pic:spPr>'
        f'</pic:pic>'
        f'</a:graphicData>'
        f'</a:graphic>'
        f'</wp:inline>'
        f'</w:drawing>'
    )


def make_run(text, font="Arial", size_pt=11, bold=False, italic=False):
    rpr = (
        f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" w:eastAsia="{font}"/>'
    )
    if bold:
        rpr += '<w:b/>'
    if italic:
        rpr += '<w:i/>'
    rpr += f'<w:sz w:val="{pt_to_half_pts(size_pt)}"/>'
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r>'


def make_tab_run():
    return '<w:r><w:tab/></w:r>'


def make_spacing(before=0, after=0, line=240):
    return f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'


def make_right_tabs(pos=9072):
    return f'<w:tabs><w:tab w:val="right" w:pos="{pos}"/></w:tabs>'


def make_border_bottom(sz_eighth_pts=18):
    return f'<w:pBdr><w:bottom w:val="single" w:sz="{sz_eighth_pts}" w:space="1" w:color="000000"/></w:pBdr>'


def make_border_top(sz_eighth_pts=12):
    return f'<w:pBdr><w:top w:val="single" w:sz="{sz_eighth_pts}" w:space="1" w:color="000000"/></w:pBdr>'


# ──────────────────────────────────────────────────────────
# HEADER / FOOTER XML BUILDERS
# ──────────────────────────────────────────────────────────

def make_linebreak_run():
    """A run containing only a line break (w:br) — stays within same paragraph."""
    return '<w:r><w:br/></w:r>'


def build_header_xml(sni_rId, sni_cx, sni_cy, sni_number, bsn_year):
    """
    Layout sesuai standar BSN menggunakan tabel 2 kolom:

    | Kolom kiri (logo + "Standar Nasional Indonesia") | Kolom kanan (nomor SNI + Ditetapkan) |
    |--------------------------------------------------|--------------------------------------|
    | [Logo SNI]                                       | SNI ISO XXXXX:20XX   (Arial 14 bold) |
    |                                                  | (Ditetapkan oleh BSN Tahun XXXX)     |
    | Standar Nasional Indonesia  (Arial 12 bold)      |                                      |
    -------------------------------------------------------------------------------------------
    [garis bawah 2.25pt di paragraf setelah tabel]
    """
    year_text = f'(Ditetapkan oleh BSN Tahun {bsn_year})'
    logo = make_inline_image(sni_rId, sni_cx, sni_cy, 1001, "sni_logo")

    # Usable width = A4(21cm) - left margin(3cm) - right margin(2cm) = 16cm = 9072 twips
    # Left col: harus cukup untuk "Standar Nasional Indonesia" Arial 12 bold dalam 1 baris (~7cm)
    # Right col: sisanya untuk teks SNI number
    left_col_w  = cm_to_twips(7.0)   # ~3969 twips - cukup untuk teks 1 baris
    right_col_w = cm_to_twips(9.0)   # ~5103 twips → total 16cm

    def tbl_cell(content_xml, w_twips, valign='top', jc='left'):
        return (
            f'<w:tc>'
            f'<w:tcPr>'
            f'<w:tcW w:w="{w_twips}" w:type="dxa"/>'
            f'<w:tcBorders>'
            f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'</w:tcBorders>'
            f'<w:tcMar>'
            f'<w:top w:w="0" w:type="dxa"/>'
            f'<w:left w:w="0" w:type="dxa"/>'
            f'<w:bottom w:w="0" w:type="dxa"/>'
            f'<w:right w:w="0" w:type="dxa"/>'
            f'</w:tcMar>'
            f'<w:vAlign w:val="{valign}"/>'
            f'</w:tcPr>'
            f'{content_xml}'
            f'</w:tc>'
        )

    def tbl_para(runs_xml, jc='left', spacing_before=0, spacing_after=0):
        return (
            f'<w:p>'
            f'<w:pPr>'
            f'<w:jc w:val="{jc}"/>'
            f'<w:spacing w:before="{spacing_before}" w:after="{spacing_after}" w:line="240" w:lineRule="auto"/>'
            f'<w:contextualSpacing w:val="0"/>'
            f'</w:pPr>'
            f'{runs_xml}'
            f'</w:p>'
        )

    # Left cell: logo para + "Standar Nasional Indonesia" para
    left_logo_para   = tbl_para(f'<w:r>{logo}</w:r>')
    left_text_para   = tbl_para(make_run("Standar Nasional Indonesia", size_pt=12, bold=True))
    left_cell = tbl_cell(left_logo_para + left_text_para, left_col_w, valign='top')

    # Right cell: SNI number + (Ditetapkan...) stacked, right-aligned
    right_sni_para   = tbl_para(make_run(sni_number, size_pt=14, bold=True), jc='right')
    right_year_para  = tbl_para(make_run(year_text, size_pt=12, bold=True), jc='right')
    right_cell = tbl_cell(right_sni_para + right_year_para, right_col_w, valign='top', jc='right')

    # Table row - no fixed height, let content determine height
    tbl_row = (
        f'<w:tr>'
        f'{left_cell}{right_cell}'
        f'</w:tr>'
    )

    # Table
    total_w = left_col_w + right_col_w
    table = (
        f'<w:tbl>'
        f'<w:tblPr>'
        f'<w:tblStyle w:val="TableGrid"/>'
        f'<w:tblW w:w="{total_w}" w:type="dxa"/>'
        f'<w:tblInd w:w="0" w:type="dxa"/>'
        f'<w:tblBorders>'
        f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'</w:tblBorders>'
        f'<w:tblCellMar>'
        f'<w:top w:w="0" w:type="dxa"/>'
        f'<w:left w:w="0" w:type="dxa"/>'
        f'<w:bottom w:w="0" w:type="dxa"/>'
        f'<w:right w:w="0" w:type="dxa"/>'
        f'</w:tblCellMar>'
        f'<w:tblLook w:val="0000"/>'
        f'</w:tblPr>'
        f'<w:tblGrid>'
        f'<w:gridCol w:w="{left_col_w}"/>'
        f'<w:gridCol w:w="{right_col_w}"/>'
        f'</w:tblGrid>'
        f'{tbl_row}'
        f'</w:tbl>'
    )

    # Border paragraph below the table (garis bawah 2.25pt)
    # Dirapatkan ke "Standar Nasional Indonesia": before=0, line minimal
    border_para = (
        f'<w:p>'
        f'<w:pPr>'
        f'<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
        f'<w:contextualSpacing w:val="0"/>'
        f'{make_border_bottom(18)}'
        f'</w:pPr>'
        f'</w:p>'
    )

    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:hdr {HDR_NAMESPACES}>'
        f'{table}'
        f'{border_para}'
        f'</w:hdr>'
    )


def make_anchor_image(rId, cx_emu, cy_emu, doc_id, name, pos_x_emu=0, pos_y_emu=0):
    """Floating (anchored) image — right-aligned relative to right margin, vertically offset from bottom margin."""
    return (
        f'<w:drawing>'
        f'<wp:anchor distT="0" distB="0" distL="0" distR="0" '
        f'simplePos="0" relativeHeight="251659264" behindDoc="0" '
        f'locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="rightMargin">'
        f'<wp:align>right</wp:align>'
        f'</wp:positionH>'
        f'<wp:positionV relativeFrom="bottomMargin">'
        f'<wp:posOffset>{pos_y_emu}</wp:posOffset>'
        f'</wp:positionV>'
        f'<wp:extent cx="{cx_emu}" cy="{cy_emu}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:wrapNone/>'
        f'<wp:docPr id="{doc_id}" name="{xml_escape(name)}"/>'
        f'<wp:cNvGraphicFramePr/>'
        f'<a:graphic>'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic>'
        f'<pic:nvPicPr>'
        f'<pic:cNvPr id="0" name="{xml_escape(name)}"/>'
        f'<pic:cNvPicPr preferRelativeResize="0"/>'
        f'</pic:nvPicPr>'
        f'<pic:blipFill>'
        f'<a:blip r:embed="{rId}"/>'
        f'<a:srcRect/><a:stretch><a:fillRect/></a:stretch>'
        f'</pic:blipFill>'
        f'<pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx_emu}" cy="{cy_emu}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</pic:spPr>'
        f'</pic:pic>'
        f'</a:graphicData>'
        f'</a:graphic>'
        f'</wp:anchor>'
        f'</w:drawing>'
    )


def make_anchor_text(text, font="Arial", size_pt=11, bold=False,
                     doc_id=1003, name="ics_text",
                     cx_emu=0, cy_emu=0, pos_y_emu=0,
                     pos_h_relative='page', pos_h_offset=0):
    """Floating text box — horizontally positioned by offset from page/margin, vertically from bottomMargin."""
    sz = int(size_pt * 2)
    text_esc = xml_escape(text)
    return (
        f'<w:drawing>'
        f'<wp:anchor distT="0" distB="0" distL="0" distR="0" '
        f'simplePos="0" relativeHeight="251659263" behindDoc="0" '
        f'locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="{pos_h_relative}">'
        f'<wp:posOffset>{pos_h_offset}</wp:posOffset>'
        f'</wp:positionH>'
        f'<wp:positionV relativeFrom="bottomMargin">'
        f'<wp:posOffset>{pos_y_emu}</wp:posOffset>'
        f'</wp:positionV>'
        f'<wp:extent cx="{cx_emu}" cy="{cy_emu}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:wrapNone/>'
        f'<wp:docPr id="{doc_id}" name="{xml_escape(name)}"/>'
        f'<wp:cNvGraphicFramePr/>'
        f'<a:graphic>'
        f'<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wps:wsp>'
        f'<wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>'
        f'<wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx_emu}" cy="{cy_emu}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</wps:spPr>'
        f'<wps:txbx>'
        f'<w:txbxContent>'
        f'<w:p>'
        f'<w:pPr><w:jc w:val="left"/><w:spacing w:before="0" w:after="0"/></w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
        f'{"<w:b/>" if bold else ""}'
        f'<w:sz w:val="{sz}"/>'
        f'</w:rPr>'
        f'<w:t>{text_esc}</w:t>'
        f'</w:r>'
        f'</w:p>'
        f'</w:txbxContent>'
        f'</wps:txbx>'
        f'<wps:bodyPr insFocus="0" rot="0" spcFirstLastPara="0" '
        f'vertOverflow="overflow" horzOverflow="overflow" '
        f'vert="horz" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" '
        f'numCol="1" spcCol="0" rtlCol="0" fromWordArt="0" '
        f'anchor="t" anchorCtr="0" forceAA="0" compatLnSpc="1"/>'
        f'</wps:wsp>'
        f'</a:graphicData>'
        f'</a:graphic>'
        f'</wp:anchor>'
        f'</w:drawing>'
    )


def build_footer_xml(bsn_rId, bsn_cx, bsn_cy, ics_number):
    """
    Layout footer sesuai referensi BSN — tabel 2 kolom (mirip header):

    | ICS XX.XXX.XX  (kiri, Arial 11 bold) | [Logo BSN, inline, kanan] |
    ───────────────────────────────────────────────────────────────────── ← garis 1.5pt BAWAH (rapat)
    """
    bsn_logo = make_inline_image(bsn_rId, bsn_cx, bsn_cy, 1002, "bsn_logo")

    # Lebar kolom sama dengan header agar garis rata: total 16cm
    left_col_w  = cm_to_twips(7.0)
    right_col_w = cm_to_twips(9.0)

    def ftr_cell(content_xml, w_twips, valign='center', jc='left'):
        return (
            f'<w:tc>'
            f'<w:tcPr>'
            f'<w:tcW w:w="{w_twips}" w:type="dxa"/>'
            f'<w:tcBorders>'
            f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'</w:tcBorders>'
            f'<w:tcMar>'
            f'<w:top w:w="0" w:type="dxa"/>'
            f'<w:left w:w="0" w:type="dxa"/>'
            f'<w:bottom w:w="0" w:type="dxa"/>'
            f'<w:right w:w="0" w:type="dxa"/>'
            f'</w:tcMar>'
            f'<w:vAlign w:val="{valign}"/>'
            f'</w:tcPr>'
            f'{content_xml}'
            f'</w:tc>'
        )

    def ftr_para(runs_xml, jc='left'):
        return (
            f'<w:p>'
            f'<w:pPr>'
            f'<w:jc w:val="{jc}"/>'
            f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:contextualSpacing w:val="0"/>'
            f'</w:pPr>'
            f'{runs_xml}'
            f'</w:p>'
        )

    ics_cell = ftr_cell(
        ftr_para(make_run(f'ICS {ics_number}', size_pt=11, bold=True)),
        left_col_w, valign='bottom', jc='left'
    )
    bsn_cell = ftr_cell(
        ftr_para(f'<w:r>{bsn_logo}</w:r>', jc='right'),
        right_col_w, valign='bottom', jc='right'
    )

    total_w = left_col_w + right_col_w
    ftr_table = (
        f'<w:tbl>'
        f'<w:tblPr>'
        f'<w:tblStyle w:val="TableGrid"/>'
        f'<w:tblW w:w="{total_w}" w:type="dxa"/>'
        f'<w:tblInd w:w="0" w:type="dxa"/>'
        f'<w:tblBorders>'
        f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'</w:tblBorders>'
        f'<w:tblCellMar>'
        f'<w:top w:w="0" w:type="dxa"/>'
        f'<w:left w:w="0" w:type="dxa"/>'
        f'<w:bottom w:w="0" w:type="dxa"/>'
        f'<w:right w:w="0" w:type="dxa"/>'
        f'</w:tblCellMar>'
        f'<w:tblLook w:val="0000"/>'
        f'</w:tblPr>'
        f'<w:tblGrid>'
        f'<w:gridCol w:w="{left_col_w}"/>'
        f'<w:gridCol w:w="{right_col_w}"/>'
        f'</w:tblGrid>'
        f'<w:tr>{ics_cell}{bsn_cell}</w:tr>'
        f'</w:tbl>'
    )

    # Garis bawah 1.5pt — mepet ke tabel (sama teknik dengan garis header)
    border_para = (
        f'<w:p>'
        f'<w:pPr>'
        f'<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
        f'<w:contextualSpacing w:val="0"/>'
        f'{make_border_bottom(12)}'
        f'</w:pPr>'
        f'</w:p>'
    )

    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:ftr {HDR_NAMESPACES}>'
        f'{ftr_table}'
        f'{border_para}'
        f'</w:ftr>'
    )


# ──────────────────────────────────────────────────────────
# COVER BODY PARAGRAPH BUILDERS
# ──────────────────────────────────────────────────────────

def build_cover_body_xml(title_id, title_en, ref_standard, n_spacer=14):
    """Build XML fragment for cover body content (without <w:body> wrapper)."""
    spacer = f'<w:p><w:pPr>{make_spacing()}</w:pPr></w:p>'
    parts = []

    for _ in range(n_spacer):
        parts.append(spacer)

    if title_id:
        parts.append(
            f'<w:p>'
            f'<w:pPr>{make_spacing()}<w:jc w:val="center"/></w:pPr>'
            f'{make_run(title_id, size_pt=18, bold=True)}'
            f'</w:p>'
        )
        parts.append(spacer)

    if title_en:
        parts.append(
            f'<w:p>'
            f'<w:pPr>{make_spacing()}<w:jc w:val="center"/></w:pPr>'
            f'{make_run(title_en, size_pt=16, bold=True, italic=True)}'
            f'</w:p>'
        )
        parts.append(spacer)

    if ref_standard:
        ref_text = ref_standard if ref_standard.startswith('(') else f'({ref_standard})'
        parts.append(
            f'<w:p>'
            f'<w:pPr>{make_spacing()}<w:jc w:val="center"/></w:pPr>'
            f'{make_run(ref_text, size_pt=12, bold=True)}'
            f'</w:p>'
        )

    return ''.join(parts)


def build_copyright_body_xml(bsn_year, iso_year, city="Jakarta"):
    """
    Halaman hak cipta — layout PERSIS sesuai dokumen referensi BSN.

    Spesifikasi (dari XML dokumen asli):
    - Font: Arial 9pt (sz=18), bold, semua teks
    - Tabel: width=9072 dxa, tblInd=-8, border single sz=6 color=000000
    - Cell margin: top=100, left=115, bottom=100, right=115 dxa
    - Para indent: left=167, right=176 khusus paragraf panjang
    - Spacing: after=0, line=240, lineRule=auto
    - 39 spacer paragraf mendorong tabel ke bawah
    """
    def cp(text="", ind_right=0):
        rpr = (
            '<w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
            '<w:b/>'
            '<w:sz w:val="18"/>'
            '<w:szCs w:val="18"/>'
        )
        run = (
            f'<w:r><w:rPr>{rpr}</w:rPr>'
            f'<w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r>'
        ) if text else ''
        right_ind = f' w:right="{ind_right}"' if ind_right else ''
        return (
            f'<w:p>'
            f'<w:pPr>'
            f'<w:jc w:val="both"/>'
            f'<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:ind w:left="167"{right_ind}/>'
            f'<w:rPr>{rpr}</w:rPr>'
            f'</w:pPr>'
            f'{run}'
            f'</w:p>'
        )

    def cp_url(text):
        rpr_p = (
            '<w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
            '<w:b/><w:sz w:val="18"/><w:szCs w:val="18"/>'
        )
        rpr_r = (
            '<w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
            '<w:b/>'
            '<w:color w:val="000000" w:themeColor="text1"/>'
            '<w:sz w:val="18"/><w:szCs w:val="18"/>'
            '<w:u w:val="none"/>'
        )
        return (
            f'<w:p>'
            f'<w:pPr><w:jc w:val="both"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:ind w:left="167"/><w:rPr>{rpr_p}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr_r}</w:rPr><w:t>{xml_escape(text)}</w:t></w:r>'
            f'</w:p>'
        )

    box_content = (
        cp(f'\u00a9 ISO {iso_year} \u2013 All rights reserved')
        + cp(f'\u00a9 BSN {bsn_year} untuk kepentingan adopsi standar \u00a9 ISO menjadi SNI \u2013 Semua hak dilindungi')
        + cp()
        + cp(
            'Hak cipta dilindungi undang-undang. Dilarang mengumumkan dan memperbanyak sebagian atau '
            'seluruh isi dokumen ini dengan cara dan dalam bentuk apapun serta dilarang mendistribusikan '
            'dokumen ini baik secara elektronik maupun tercetak tanpa izin tertulis BSN',
            ind_right=176
        )
        + cp()
        + cp()
        + cp('BSN')
        + cp('Email: dokinfo@bsn.go.id')
        + cp_url('www.bsn.go.id')
        + cp()
        + cp()
        + cp()
        + cp(f'Diterbitkan di {city}')
    )

    box_w  = 9072
    bdr_sz = 6

    bordered_table = (
        f'<w:tbl>'
        f'<w:tblPr>'
        f'<w:tblStyle w:val="TableGrid"/>'
        f'<w:tblW w:w="{box_w}" w:type="dxa"/>'
        f'<w:tblInd w:w="-9" w:type="dxa"/>'
        f'<w:tblBorders>'
        f'<w:top    w:val="single" w:sz="{bdr_sz}" w:space="0" w:color="000000"/>'
        f'<w:left   w:val="single" w:sz="{bdr_sz}" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="{bdr_sz}" w:space="0" w:color="000000"/>'
        f'<w:right  w:val="single" w:sz="{bdr_sz}" w:space="0" w:color="000000"/>'
        f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'</w:tblBorders>'
        f'<w:tblCellMar>'
        f'<w:top    w:w="100" w:type="dxa"/>'
        f'<w:left   w:w="115" w:type="dxa"/>'
        f'<w:bottom w:w="100" w:type="dxa"/>'
        f'<w:right  w:w="115" w:type="dxa"/>'
        f'</w:tblCellMar>'
        f'<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/>'
        f'</w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{box_w}"/></w:tblGrid>'
        f'<w:tr>'
        f'<w:tc>'
        f'<w:tcPr>'
        f'<w:tcW w:w="{box_w}" w:type="dxa"/>'
        f'<w:tcBorders>'
        f'<w:top    w:val="single" w:sz="{bdr_sz}" w:space="0" w:color="000000"/>'
        f'<w:left   w:val="single" w:sz="{bdr_sz}" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="{bdr_sz}" w:space="0" w:color="000000"/>'
        f'<w:right  w:val="single" w:sz="{bdr_sz}" w:space="0" w:color="000000"/>'
        f'</w:tcBorders>'
        f'</w:tcPr>'
        f'{box_content}'
        f'</w:tc>'
        f'</w:tr>'
        f'</w:tbl>'
    )

    push_spacers = ''.join(
        '<w:p><w:pPr>'
        '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        '</w:pPr></w:p>'
        for _ in range(39)
    )

    return push_spacers + bordered_table
def make_rels_with_image(media_filename):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        f'Target="media/{media_filename}"/>'
        '</Relationships>'
    )

def make_empty_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    )


# ──────────────────────────────────────────────────────────
# MAIN ENGINE CLASS
# ──────────────────────────────────────────────────────────

class CoverPageEngine:
    """
    Engine untuk membuat halaman sampul/cover sesuai standar BSN (SNI ISO).

    Cara pakai:
        engine4 = CoverPageEngine()
        success, msg = engine4.prepend_cover(
            input_docx="dokumen_isi.docx",
            output_docx="dengan_cover.docx",
            sni_number="SNI ISO 19659-2:2020",
            bsn_year="2020",
            title_id="Judul Bahasa Indonesia",
            title_en="English Title",
            ics_number="45.060.01"
        )

    Pastikan sni_logo.jpeg dan bsn_logo.jpg ada di folder yang sama dengan engine4.py.
    """

    DEFAULT_SNI_LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sni_logo.jpg")
    DEFAULT_BSN_LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bsn_logo.jpg")

    def __init__(self, sni_logo_path=None, bsn_logo_path=None):
        self.sni_logo_path = sni_logo_path or self.DEFAULT_SNI_LOGO
        self.bsn_logo_path = bsn_logo_path or self.DEFAULT_BSN_LOGO

    def prepend_cover(
        self,
        input_docx: str,
        output_docx: str,
        sni_number: str = "SNI ISO XXXXX:20XX",
        bsn_year: str = "20XX",
        iso_year: str = "",
        title_id: str = "",
        title_en: str = "",
        ref_standard: str = "",
        ics_number: str = "XX.XXX.XX",
        copyright_city: str = "Jakarta",
    ):
        """
        Prepend halaman cover SNI ke dokumen yang sudah ada.

        Returns:
            (True, output_path) atau (False, error_message)
        """
        try:
            if not ref_standard:
                ref_standard = sni_number.replace("SNI ", "", 1) + ", IDT"

            # Load logos
            sni_bytes = self._load_file(self.sni_logo_path)
            bsn_bytes = self._load_file(self.bsn_logo_path)
            sni_ext = self._get_ext(self.sni_logo_path, 'jpeg')
            bsn_ext = self._get_ext(self.bsn_logo_path, 'jpeg')

            # EMU sizes
            sni_cx, sni_cy = cm_to_emu(3.23), cm_to_emu(2.29)   # width=2.29cm, height=3.23cm (portrait)
            bsn_cx, bsn_cy = cm_to_emu(7.0),  cm_to_emu(1.5)

            # Margin in twips
            top_tw    = cm_to_twips(3)
            bottom_tw = cm_to_twips(2)
            left_tw   = cm_to_twips(3)
            right_tw  = cm_to_twips(2)
            hdr_tw    = cm_to_twips(1.5)
            ftr_tw    = cm_to_twips(1.3)

            # Load body ZIP
            with zipfile.ZipFile(input_docx, 'r') as z:
                body = {n: z.read(n) for n in z.namelist()}

            # ── Find next available file index
            nums = [int(m.group(1)) for n in body
                    for m in [re.search(r'(?:header|footer)(\d+)\.xml$', n)] if m]
            next_n = (max(nums) + 1) if nums else 10

            hdr_file    = f"header{next_n}.xml"
            ftr_file    = f"footer{next_n + 1}.xml"
            hdr_rels    = f"_rels/header{next_n}.xml.rels"
            ftr_rels    = f"_rels/footer{next_n + 1}.xml.rels"
            # Copyright section — blank header/footer (section tersendiri)
            cpr_hdr_file = f"header{next_n + 2}.xml"
            cpr_ftr_file = f"footer{next_n + 3}.xml"
            sni_media   = f"cover_sni.{sni_ext}"
            bsn_media   = f"cover_bsn.{bsn_ext}"

            # ── Find next available rId
            rels_str = body.get('word/_rels/document.xml.rels', b'').decode('utf-8', errors='replace')
            rid_nums = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels_str)]
            next_rid = (max(rid_nums) + 1) if rid_nums else 50
            hdr_rid = f"rId{next_rid}"
            ftr_rid = f"rId{next_rid + 1}"
            cpr_hdr_rid = f"rId{next_rid + 2}"
            cpr_ftr_rid = f"rId{next_rid + 3}"

            # ── Build header / footer XML
            hdr_xml = build_header_xml("rId1", sni_cx, sni_cy, sni_number, bsn_year)
            ftr_xml = build_footer_xml("rId1", bsn_cx, bsn_cy, ics_number)

            # ── Build cover body content
            cover_content = build_cover_body_xml(title_id, title_en, ref_standard)

            # ── Build halaman hak cipta (page 2 dari section cover)
            copyright_content = build_copyright_body_xml(
                bsn_year=bsn_year,
                iso_year=iso_year if iso_year else bsn_year,
                city=copyright_city
            )

            # ── sectPr section 1: COVER — referensi header/footer berlogo
            cover_sectPr = (
                f'<w:sectPr>'
                f'<w:headerReference w:type="default" r:id="{hdr_rid}"/>'
                f'<w:footerReference w:type="default" r:id="{ftr_rid}"/>'
                f'<w:pgSz w:w="11906" w:h="16838"/>'
                f'<w:pgMar w:top="{top_tw}" w:right="{right_tw}" '
                f'w:bottom="{bottom_tw}" w:left="{left_tw}" '
                f'w:header="{hdr_tw}" w:footer="{ftr_tw}" w:gutter="0"/>'
                f'<w:pgNumType w:fmt="none"/>'
                f'<w:type w:val="nextPage"/>'
                f'</w:sectPr>'
            )
            cover_sep = f'<w:p><w:pPr>{cover_sectPr}</w:pPr></w:p>'

            # ── sectPr section 2: HAK CIPTA — margin khusus (berbeda dari cover)
            # top=2cm, inside(left)=2cm, bottom=3cm, outside(right)=3cm
            cpr_top_tw     = cm_to_twips(2)
            cpr_bottom_tw  = cm_to_twips(3)
            cpr_inside_tw  = cm_to_twips(2)
            cpr_outside_tw = cm_to_twips(3)
            cpr_sectPr = (
                f'<w:sectPr>'
                f'<w:headerReference w:type="default" r:id="{cpr_hdr_rid}"/>'
                f'<w:footerReference w:type="default" r:id="{cpr_ftr_rid}"/>'
                f'<w:pgSz w:w="11906" w:h="16838"/>'
                f'<w:pgMar w:top="{cpr_top_tw}" w:right="{cpr_outside_tw}" '
                f'w:bottom="{cpr_bottom_tw}" w:left="{cpr_inside_tw}" '
                f'w:header="{hdr_tw}" w:footer="{ftr_tw}" w:gutter="0"/>'
                f'<w:pgNumType w:fmt="none"/>'
                f'<w:type w:val="nextPage"/>'
                f'</w:sectPr>'
            )
            cpr_sep = f'<w:p><w:pPr>{cpr_sectPr}</w:pPr></w:p>'

            # ── Inject cover content into document.xml immediately after <w:body>
            doc_xml = body['word/document.xml'].decode('utf-8', errors='replace')
            m = re.search(r'<w:body>', doc_xml)
            if not m:
                raise ValueError("No <w:body> found in document.xml")
            ins = m.end()
            new_doc_xml = doc_xml[:ins] + cover_content + cover_sep + copyright_content + cpr_sep + doc_xml[ins:]

            # ── Check if body's final sectPr already has header/footer references
            # If NOT, we must add empty header/footer to BREAK inheritance from cover section.
            # Otherwise body pages will display the cover header/footer.
            body_end_pos = new_doc_xml.rfind('</w:body>')
            last_sect_pos = new_doc_xml.rfind('<w:sectPr', 0, body_end_pos)
            last_sect_end = new_doc_xml.find('</w:sectPr>', last_sect_pos) + len('</w:sectPr>')
            last_sect_xml = new_doc_xml[last_sect_pos:last_sect_end]
            body_has_hdr = 'headerReference' in last_sect_xml

            # rIds for body blank header/footer (only used if body has no own header/footer)
            body_hdr_rid = f"rId{next_rid + 4}"
            body_ftr_rid = f"rId{next_rid + 5}"
            body_hdr_file = f"header{next_n + 4}.xml"
            body_ftr_file = f"footer{next_n + 5}.xml"

            if not body_has_hdr:
                # Inject headerReference + footerReference into the final sectPr
                # Insert right after the opening <w:sectPr...> tag
                sect_open_end = new_doc_xml.find('>', last_sect_pos) + 1
                body_refs = (
                    f'<w:headerReference w:type="default" r:id="{body_hdr_rid}"/>'
                    f'<w:footerReference w:type="default" r:id="{body_ftr_rid}"/>'
                )
                new_doc_xml = (
                    new_doc_xml[:sect_open_end]
                    + body_refs
                    + new_doc_xml[sect_open_end:]
                )

            # ── Add relationships to document.xml.rels
            hdr_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
            ftr_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
            new_rels = (
                f'<Relationship Id="{hdr_rid}" Type="{hdr_type}" Target="{hdr_file}"/>'
                f'<Relationship Id="{ftr_rid}" Type="{ftr_type}" Target="{ftr_file}"/>'
                f'<Relationship Id="{cpr_hdr_rid}" Type="{hdr_type}" Target="{cpr_hdr_file}"/>'
                f'<Relationship Id="{cpr_ftr_rid}" Type="{ftr_type}" Target="{cpr_ftr_file}"/>'
            )
            if not body_has_hdr:
                new_rels += (
                    f'<Relationship Id="{body_hdr_rid}" Type="{hdr_type}" Target="{body_hdr_file}"/>'
                    f'<Relationship Id="{body_ftr_rid}" Type="{ftr_type}" Target="{body_ftr_file}"/>'
                )
            new_rels_str = rels_str.replace('</Relationships>', new_rels + '</Relationships>')

            # ── Update [Content_Types].xml
            ct_xml = body.get('[Content_Types].xml', b'').decode('utf-8', errors='replace')
            hdr_ct = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
            ftr_ct = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'

            for part_name, ct in [
                (f'/word/{hdr_file}', hdr_ct),
                (f'/word/{ftr_file}', ftr_ct),
                (f'/word/{cpr_hdr_file}', hdr_ct),
                (f'/word/{cpr_ftr_file}', ftr_ct),
            ]:
                if part_name not in ct_xml:
                    ct_xml = ct_xml.replace(
                        '</Types>',
                        f'<Override PartName="{part_name}" ContentType="{ct}"/></Types>'
                    )

            if not body_has_hdr:
                for part_name, ct in [
                    (f'/word/{body_hdr_file}', hdr_ct),
                    (f'/word/{body_ftr_file}', ftr_ct),
                ]:
                    if part_name not in ct_xml:
                        ct_xml = ct_xml.replace(
                            '</Types>',
                            f'<Override PartName="{part_name}" ContentType="{ct}"/></Types>'
                        )

            for ext, ct in [('jpeg', 'image/jpeg'), ('jpg', 'image/jpeg'), ('png', 'image/png')]:
                if f'Extension="{ext}"' not in ct_xml:
                    ct_xml = ct_xml.replace(
                        '</Types>',
                        f'<Default Extension="{ext}" ContentType="{ct}"/></Types>'
                    )

            # Blank header/footer XML (minimal valid OOXML - just one empty paragraph)
            blank_hdr_xml = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:hdr {HDR_NAMESPACES}>'
                f'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>'
                f'</w:hdr>'
            )
            blank_ftr_xml = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:ftr {HDR_NAMESPACES}>'
                f'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>'
                f'</w:ftr>'
            )
            blank_rels_xml = make_empty_rels()

            # ── Assemble output files dict
            out = dict(body)
            out['word/document.xml']                  = new_doc_xml.encode('utf-8')
            out['word/_rels/document.xml.rels']       = new_rels_str.encode('utf-8')
            out['[Content_Types].xml']                = ct_xml.encode('utf-8')
            out[f'word/{hdr_file}']                   = hdr_xml.encode('utf-8')
            out[f'word/{ftr_file}']                   = ftr_xml.encode('utf-8')
            out[f'word/{hdr_rels}']                   = (make_rels_with_image(sni_media) if sni_bytes else make_empty_rels()).encode('utf-8')
            out[f'word/{ftr_rels}']                   = (make_rels_with_image(bsn_media) if bsn_bytes else make_empty_rels()).encode('utf-8')
            # Copyright section — blank header/footer (bersih, tanpa logo)
            out[f'word/{cpr_hdr_file}']               = blank_hdr_xml.encode('utf-8')
            out[f'word/{cpr_ftr_file}']               = blank_ftr_xml.encode('utf-8')
            out[f'word/_rels/{cpr_hdr_file}.rels']    = blank_rels_xml.encode('utf-8')
            out[f'word/_rels/{cpr_ftr_file}.rels']    = blank_rels_xml.encode('utf-8')
            if sni_bytes:
                out[f'word/media/{sni_media}']        = sni_bytes
            if bsn_bytes:
                out[f'word/media/{bsn_media}']        = bsn_bytes

            # Blank header/footer for body section (only if body had no own header/footer)
            if not body_has_hdr:
                out[f'word/{body_hdr_file}']                    = blank_hdr_xml.encode('utf-8')
                out[f'word/{body_ftr_file}']                    = blank_ftr_xml.encode('utf-8')
                out[f'word/_rels/{body_hdr_file}.rels']         = blank_rels_xml.encode('utf-8')
                out[f'word/_rels/{body_ftr_file}.rels']         = blank_rels_xml.encode('utf-8')

            # ── Write output ZIP (Content_Types and _rels/.rels first for Word compatibility)
            priority_first = ['[Content_Types].xml', '_rels/.rels']
            with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as zout:
                for name in priority_first:
                    if name in out:
                        zout.writestr(name, out[name])
                for name, data in out.items():
                    if name not in priority_first:
                        zout.writestr(name, data)

            return True, output_docx

        except Exception as e:
            import traceback
            traceback.print_exc()
            return False, f"CoverPageEngine Error: {str(e)}"

    # ──────────────────────────────────────────────────────
    # HELPERS
    # ──────────────────────────────────────────────────────

    def _load_file(self, path):
        if path and os.path.exists(path):
            with open(path, 'rb') as f:
                return f.read()
        return None

    def _get_ext(self, path, default='jpeg'):
        if not path:
            return default
        ext = os.path.splitext(path)[1].lstrip('.').lower()
        if ext == 'jpg':
            ext = 'jpeg'
        return ext or default
