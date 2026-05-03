"""
Engine5: DaftarIsiEngine (v3 - Auto Heading Extraction)
=========================================================
Engine untuk membuat halaman Daftar Isi sesuai standar BSN/SNI.

v3 (Auto):
  - Isi Daftar Isi diambil OTOMATIS dari heading dokumen (logika identik engine2).
  - Heading numbered (1, 2, 3...) → level 0 (no indent)
  - Sub-heading (1.1, 1.2...) → level 1 (indent 360 twips)
  - Sub-sub-heading (1.1.1...) → level 2 (indent 720 twips)
  - Special heading (Bibliografi, Lampiran, Annex, dll) → level 0
  - Fixed header selalu ada di atas: Kata Pendahuluan, Daftar Isi, Pendahuluan
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
# HEADING EXTRACTOR (logika identik engine2)
# Returns list of (display_text, indent_level)
#   indent_level 0 = no indent (main chapter / special)
#   indent_level 1 = sub-chapter (1.1, 1.2, ...)
#   indent_level 2 = sub-sub-chapter (1.1.1, ...)
# ─────────────────────────────────────────────────────────────────────────────
def extract_headings_from_docx(docx_path: str) -> list:
    """
    Membaca dokumen dan mengekstrak heading bernomor untuk Daftar Isi,
    PERSIS seperti Word auto-TOC:

    Strategi (prioritas):
      1. Jika dokumen punya Heading style (Heading 1/2/3...) → gunakan style,
         auto-nomori berdasarkan urutan (1, 1.1, 1.1.1, dst.)
      2. Jika tidak ada Heading style → cari paragraf bold dengan teks
         yang sudah bernomor (hasil engine2): "1    Ruang Lingkup", dst.

    Returns:
        list of (display_text: str, level: int)
    """
    try:
        from docx import Document
        doc = Document(docx_path)
    except Exception:
        return []

    paragraphs = list(doc.paragraphs)
    re_numbered = re.compile(r'^(\d[\d\.]*)\.?\s+(\S.*)')

    # ── Cek apakah dokumen menggunakan Heading style ──────────────────
    has_heading_styles = any(
        p.style and p.style.name.startswith('Heading ')
        for p in paragraphs
    )

    headings = []

    if has_heading_styles:
        # ── STRATEGI 1: Dari Heading style (seperti Word auto-TOC) ─────
        # Auto-number sesuai urutan per level
        counters = {}

        for p in paragraphs:
            txt = p.text.strip()
            if not txt:
                continue
            style = p.style.name if p.style else ''
            if not style.startswith('Heading '):
                continue

            try:
                h_level = int(style.split()[-1]) - 1  # Heading 1→0, Heading 2→1, dst.
            except (ValueError, IndexError):
                continue

            level_num = h_level + 1  # 1-based untuk counter
            counters[level_num] = counters.get(level_num, 0) + 1
            # Reset semua sub-level di bawahnya
            for k in list(counters.keys()):
                if k > level_num:
                    counters[k] = 0

            # Bangun string nomor: "1", "1.1", "1.1.1", dll.
            num_parts = [str(counters.get(i, 1)) for i in range(1, level_num + 1)]
            num_str = '.'.join(num_parts)

            display = f'{num_str}    {txt}'
            headings.append((display, h_level))

    else:
        # ── STRATEGI 2: Dari teks bernomor bold (output engine2) ───────
        re_list_item = re.compile(
            r'^(?:[a-z]\)|[a-z]\.|[A-Z]\)|[A-Z]\.\|'
            r'\([a-z]\)|\([0-9]+\)|[ivxlcdm]+\.|[IVXLCDM]+\.)\s+',
            re.IGNORECASE
        )
        re_note = re.compile(r'^(note|catatan)\b', re.IGNORECASE)

        title_processed = False

        for p in paragraphs:
            txt = p.text.strip()
            if not txt:
                continue

            is_bold = any(r.bold for r in p.runs)
            if not is_bold and p.style:
                try:
                    if p.style.font and p.style.font.bold:
                        is_bold = True
                except Exception:
                    pass

            match_number = re_numbered.match(txt)

            if not title_processed:
                if not match_number:
                    title_processed = True
                continue

            title_processed = True

            if not match_number:
                continue
            if bool(re_list_item.match(txt)):
                continue
            if bool(re_note.match(txt)):
                continue
            if any(r.font.size and r.font.size.pt == 10 for r in p.runs):
                continue

            num_part   = match_number.group(1).rstrip('.')
            title_part = match_number.group(2).strip()
            dot_count  = num_part.count('.')

            display = f'{num_part}    {title_part}'
            headings.append((display, dot_count))

    return headings

# ─────────────────────────────────────────────────────────────────────────────
# DAFTAR ISI BUILDER
# ─────────────────────────────────────────────────────────────────────────────
# Indent twips per level
LEVEL_INDENT = {0: 0, 1: 360, 2: 720}

def _build_di_elements(hdr_odd, hdr_even, ftr_odd, ftr_even, heading_entries=None):
    """
    Return list of raw XML strings untuk paragraf DI + inline sectPr.

    heading_entries: list of (text, level) dari extract_headings_from_docx().
                     Jika None → gunakan fallback statis.
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

    def entry_p(text, level=0):
        indent_val = LEVEL_INDENT.get(level, 0)
        ind = f'<w:ind w:left="{indent_val}"/>' if indent_val else ''
        sz  = pt_to_hpts(11)
        run = (
            f'<w:r><w:rPr>'
            f'<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
            f'<w:b/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
            f'</w:rPr><w:t xml:space="preserve">{_esc(text)}</w:t></w:r>'
        ) if text else ''
        return (
            f'<w:p><w:pPr>'
            f'<w:jc w:val="left"/>'
            f'<w:spacing w:before="0" w:after="120"/>'
            f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{TAB}"/></w:tabs>'
            f'{ind}'
            f'</w:pPr>'
            f'{run}<w:r><w:tab/></w:r>'
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

    # ── Fallback statis jika tidak ada heading yang diekstrak ──
    if not heading_entries:
        heading_entries = [
            ('1    Ruang Lingkup', 0),
            ('2    Acuan Normatif', 0),
            ('3    Istilah dan Definisi', 0),
            ('4    ...', 0),
            ('5    ...', 0),
            ('Lampiran A (informatif)', 0),
            ('Bibliografi', 0),
        ]

    # ── Header tetap ──
    fixed_top = [
        ('Kata Pendahuluan', 0),
        ('Daftar Isi', 0),
        ('Pendahuluan', 0),
    ]

    all_entries = fixed_top + heading_entries

    xmls = (
        [title_p(), empty_p(), empty_p(), empty_p()]
        + [entry_p(text, level) for text, level in all_entries]
        + [sect_p()]
    )
    return xmls


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
            # 1. Ekstrak heading dari input docx
            heading_entries = extract_headings_from_docx(input_docx)

            # 2. Baca input
            with zipfile.ZipFile(input_docx, 'r') as z:
                files = {n: z.read(n) for n in z.namelist()}

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

            # 9. Build & insert DI elements (dengan heading otomatis)
            di_xmls = _build_di_elements(rid_ho, rid_he, rid_fo, rid_fe, heading_entries)
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