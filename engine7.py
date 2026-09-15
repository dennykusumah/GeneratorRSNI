"""
Engine7: DocxOptimizerEngine - Updated Version
===============================================
Engine untuk merapikan dokumen Word sesuai standar ISO/SNI
Digunakan oleh app.py untuk menu "2. Rapikan (Word -> ISO Std)"
"""

from pipeline_utils import validate_docx, atomic_save_docx
import os
import re
import zipfile
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.text import WD_TAB_ALIGNMENT


# Namespace
WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def _normalize_bibliography_text(text):
    value = (text or '').replace('\xa0', ' ')
    value = re.sub(r'[\u200b\u200c\u200d\u2060\ufeff\u00ad]', '', value)
    return re.sub(r'\s+', ' ', value).strip()


def _is_bibliography_paragraph(paragraph):
    """Deteksi heading Bibliography, termasuk running header IEC/CISPR."""
    text = _normalize_bibliography_text(paragraph.text)
    key = re.sub(r'[\s:;,.\-–—]+$', '', text).casefold()
    key = re.sub(r'^(?:[a-z]\.?|\d+(?:\.\d+)*)\s+', '', key)
    if key in {
        'bibliography', 'bibliographical references', 'daftar pustaka'
    }:
        return True
    ppr = paragraph._p.find(qn('w:pPr'))
    style_el = ppr.find(qn('w:pStyle')) if ppr is not None else None
    style = style_el.get(qn('w:val'), '').casefold() if style_el is not None else ''
    heading_style = any(
        token in style for token in ('heading', 'title', 'judul', 'biblio')
    )
    if key == 'references' and heading_style:
        return True
    has_word = bool(re.search(r'(?<![a-z])bibliography(?![a-z])', key))
    running_header = bool(re.search(
        r'\b(?:ISO(?:\s*/\s*IEC)?|IEC|CISPR)\b'
        r'\s+(?:19|20)\d{2}\b'
        r'[^A-Za-z0-9]{0,15}\bbibliography\b',
        key,
        flags=re.IGNORECASE,
    ))
    return has_word and (heading_style or running_header)

def remove_all_hyperlinks(doc):
    """
    Hapus semua hyperlink dalam dokumen dan jadikan teks biasa.
    Mempertahankan formatting run (bold, italic, font size, dll).
    """
    from lxml import etree
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    
    # Cari semua elemen hyperlink di seluruh document body
    body = doc.element.body
    hyperlinks = body.findall(f'.//{{{W}}}hyperlink')
    
    for hyperlink in hyperlinks:
        parent = hyperlink.getparent()
        if parent is None:
            continue
        
        # Dapatkan posisi hyperlink di parent
        idx = list(parent).index(hyperlink)
        
        # Pindahkan semua child <w:r> dari hyperlink ke parent (menggantikan hyperlink)
        children = list(hyperlink)
        for i, child in enumerate(children):
            parent.insert(idx + i, child)
        
        # Hapus elemen hyperlink (sudah kosong)
        parent.remove(hyperlink)


def flatten_broken_ref_fields(doc, font_name='Arial', font_size=11):
    """Ubah field REF internal menjadi teks statis agar bookmark hilang tidak
    menghasilkan ``Error! Reference source not found.`` saat Word menekan F9.

    Nilai hasil field yang tersimpan (misalnya 4.2/4.3/5) dipertahankan. Jika
    cache field sudah rusak, nilai diturunkan dari nama bookmark.
    """
    flattened = 0

    # REF/PAGEREF dapat berada di body maupun part header/footer bawaan ISO.
    # Proses setiap part tepat sekali; header yang linked sering berbagi part.
    roots = [doc.element.body]
    seen_parts = set()
    for section in doc.sections:
        for container in (
            section.header, section.first_page_header, section.even_page_header,
            section.footer, section.first_page_footer, section.even_page_footer,
        ):
            part_key = str(container.part.partname)
            if part_key in seen_parts:
                continue
            seen_parts.add(part_key)
            roots.append(container._element)

    def field_match(instruction):
        return re.search(
            r'\b(?:REF|PAGEREF)\s+([^\s\\]+)',
            instruction or '',
            re.IGNORECASE,
        )

    def static_value(match, cached_text=''):
        value = (cached_text or '').strip()
        if not value or 'reference source not found' in value.casefold():
            bookmark = match.group(1)
            value = re.sub(
                r'^(?:Section|Clause|Annex)_sec_', '', bookmark,
                flags=re.IGNORECASE,
            )
        return value

    def static_run(value):
        replacement = OxmlElement('w:r')
        r_pr = OxmlElement('w:rPr')
        r_fonts = OxmlElement('w:rFonts')
        for attr in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
            r_fonts.set(qn(f'w:{attr}'), font_name)
        r_pr.append(r_fonts)
        color = OxmlElement('w:color')
        color.set(qn('w:val'), '000000')
        r_pr.append(color)
        underline = OxmlElement('w:u')
        underline.set(qn('w:val'), 'none')
        r_pr.append(underline)
        for tag in ('w:sz', 'w:szCs'):
            size = OxmlElement(tag)
            size.set(qn('w:val'), str(int(font_size * 2)))
            r_pr.append(size)
        replacement.append(r_pr)
        text_node = OxmlElement('w:t')
        text_node.text = value
        replacement.append(text_node)
        return replacement

    # Bentuk sederhana: <w:fldSimple w:instr="REF ...">hasil</w:fldSimple>.
    for field in [node for root in roots for node in root.xpath('.//w:fldSimple')]:
        match = field_match(field.get(qn('w:instr')) or '')
        if not match:
            continue
        cached = ''.join(
            node.text or '' for node in field.iter(qn('w:t'))
        )
        value = static_value(match, cached)
        parent = field.getparent()
        parent.insert(parent.index(field), static_run(value))
        parent.remove(field)
        flattened += 1

    # Bentuk kompleks yang dibungkus hyperlink internal.
    hyperlinks = [
        node for root in roots
        for node in root.xpath('.//w:hyperlink[.//w:instrText]')
    ]
    for hyperlink in hyperlinks:
        instructions = ' '.join(
            node.text or '' for node in hyperlink.iter(qn('w:instrText'))
        )
        match = field_match(instructions)
        if not match:
            continue

        result_parts = []
        in_result = False
        for run in hyperlink.iter(qn('w:r')):
            field_chars = run.findall(qn('w:fldChar'))
            field_type = (
                field_chars[0].get(qn('w:fldCharType')) if field_chars else None
            )
            if field_type == 'separate':
                in_result = True
                continue
            if field_type == 'end':
                break
            if in_result:
                result_parts.extend(
                    node.text or '' for node in run.findall(qn('w:t'))
                )

        visible_text = static_value(match, ''.join(result_parts))
        if not visible_text:
            continue
        parent = hyperlink.getparent()
        parent.insert(parent.index(hyperlink), static_run(visible_text))
        parent.remove(hyperlink)
        flattened += 1

    # Bentuk kompleks biasa: run begin/instr/separate/result/end langsung di
    # paragraf, tanpa pembungkus hyperlink.
    for paragraph in [
        node for root in roots for node in root.xpath('.//w:p[.//w:instrText]')
    ]:
        changed = True
        while changed:
            changed = False
            children = list(paragraph)
            for begin_index, child in enumerate(children):
                begin_chars = [
                    node for node in child.findall(qn('w:fldChar'))
                    if node.get(qn('w:fldCharType')) == 'begin'
                ]
                if not begin_chars:
                    continue
                depth = 0
                separate_index = None
                end_index = None
                for index in range(begin_index, len(children)):
                    node = children[index]
                    for field_char in node.findall(qn('w:fldChar')):
                        field_type = field_char.get(qn('w:fldCharType'))
                        if field_type == 'begin':
                            depth += 1
                        elif field_type == 'separate' and depth == 1:
                            separate_index = index
                        elif field_type == 'end':
                            depth -= 1
                            if depth == 0:
                                end_index = index
                                break
                    if end_index is not None:
                        break
                if separate_index is None or end_index is None:
                    continue
                instruction = ' '.join(
                    node.text or ''
                    for element in children[begin_index:separate_index + 1]
                    for node in element.iter(qn('w:instrText'))
                )
                match = field_match(instruction)
                if not match:
                    continue
                cached = ''.join(
                    node.text or ''
                    for element in children[separate_index + 1:end_index]
                    for node in element.iter(qn('w:t'))
                )
                value = static_value(match, cached)
                for element in children[begin_index:end_index + 1]:
                    paragraph.remove(element)
                paragraph.insert(begin_index, static_run(value))
                flattened += 1
                changed = True
                break

    # Jaring pengaman untuk error yang sudah berubah menjadi teks literal dan
    # tidak lagi mempunyai kode field.
    error_pattern = re.compile(
        r'Error!\s*Reference source not found\.?', re.IGNORECASE
    )
    for text_node in [node for root in roots for node in root.xpath('.//w:t')]:
        if text_node.text and error_pattern.search(text_node.text):
            text_node.text = error_pattern.sub('', text_node.text)
    return flattened


def _has_image(paragraph):
    """Cek apakah paragraf mengandung gambar"""
    p_el = paragraph._element
    drawing_tag = f'{{{WNS}}}drawing'
    pict_tag    = f'{{{WNS}}}pict'
    for descendant in p_el.iter():
        if descendant.tag in (drawing_tag, pict_tag):
            return True
    return False

def _is_truly_empty(paragraph):
    """Paragraf kosong = tidak ada teks DAN tidak ada gambar"""
    if _has_image(paragraph):
        return False
    return not paragraph.text.strip()

def _has_page_break(paragraph):
    """Paragraf tanpa teks dapat tetap bermakna bila membawa page break."""
    return any(
        br.get(qn('w:type')) == 'page'
        for br in paragraph._p.xpath('.//w:br')
    ) or bool(paragraph._p.xpath('./w:pPr/w:pageBreakBefore'))

def set_document_margins(doc, top_cm=3, inside_cm=3, bottom_cm=2, outside_cm=2,
                         section_indexes=None):
    """Set margin hanya pada section yang masuk cakupan Engine 7."""
    selected = set(section_indexes) if section_indexes is not None else None
    for index, section in enumerate(doc.sections):
        if selected is not None and index not in selected:
            continue
        section.top_margin = Cm(top_cm)
        section.bottom_margin = Cm(bottom_cm)
        section.left_margin = Cm(inside_cm)
        section.right_margin = Cm(outside_cm)

#headerfooter
def setup_headers_footers(doc, doc_title="SNI ISO XXXXX:2025", copyright_text="©BSN 2025"):

    # =====================================================
    # HELPER
    # =====================================================
    def clear_container(container):
        container.is_linked_to_previous = False
        for tbl in list(container.tables):
            tbl._element.getparent().remove(tbl._element)
        for p in list(container.paragraphs):
            p._element.getparent().remove(p._element)

    def add_field(run, field_name):
        fld_begin = OxmlElement('w:fldChar')
        fld_begin.set(qn('w:fldCharType'), 'begin')

        instr = OxmlElement('w:instrText')
        instr.set(qn('xml:space'), 'preserve')
        instr.text = field_name

        fld_sep = OxmlElement('w:fldChar')
        fld_sep.set(qn('w:fldCharType'), 'separate')

        fld_end = OxmlElement('w:fldChar')
        fld_end.set(qn('w:fldCharType'), 'end')

        run._r.append(fld_begin)
        run._r.append(instr)
        run._r.append(fld_sep)
        run._r.append(fld_end)

    # =====================================================
    # NORMALISASI SECTION (A4 + MIRROR)
    # =====================================================
    for section in doc.sections:

        # ---------- FORCE A4 ----------
        top = section.top_margin
        bottom = section.bottom_margin
        left = section.left_margin
        right = section.right_margin
        header_dist = section.header_distance
        footer_dist = section.footer_distance

        section.page_width = Cm(21)
        section.page_height = Cm(29.7)


        section.top_margin = top
        section.bottom_margin = bottom
        section.left_margin = left
        section.right_margin = right
        section.header_distance = header_dist
        section.footer_distance = footer_dist
        section.footer_distance = Pt(35)   # ±1.2 cm (ideal ISO look)


        # ---------- FORCE MIRROR ----------
        sectPr = section._sectPr
        for el in sectPr.findall(qn('w:mirrorMargins')):
            sectPr.remove(el)

        mirror = OxmlElement('w:mirrorMargins')
        sectPr.append(mirror)

        # =====================================================
        # AKTIFKAN ODD/EVEN
        # =====================================================
        section.different_first_page_header_footer = False
        section.odd_and_even_pages_header_footer = True

        # =====================================================
        # RESET PAGE NUMBERING
        # =====================================================
        for el in sectPr.findall(qn('w:pgNumType')):
            sectPr.remove(el)

        pgNumType = OxmlElement('w:pgNumType')
        pgNumType.set(qn('w:start'), '1')
        sectPr.append(pgNumType)

        # =====================================================
        # HAPUS HEADER/FOOTER LAMA
        # =====================================================
        clear_container(section.header)
        clear_container(section.footer)
        clear_container(section.first_page_header)
        clear_container(section.first_page_footer)
        clear_container(section.even_page_header)
        clear_container(section.even_page_footer)

        # =====================================================
        # HEADER (TIDAK DIUBAH – SESUAI KODE ANDA)
        # =====================================================
        header = section.header
        p_header = header.add_paragraph()
        p_header.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        run_header = p_header.add_run(doc_title)
        run_header.font.name = "Arial"
        run_header.font.size = Pt(12)
        run_header.bold = True

        even_header = section.even_page_header
        p_even_header = even_header.add_paragraph()
        p_even_header.alignment = WD_ALIGN_PARAGRAPH.LEFT

        run_even = p_even_header.add_run(doc_title)
        run_even.font.name = "Arial"
        run_even.font.size = Pt(12)
        run_even.bold = True

        # =====================================================
        # HITUNG TENGAH PRESISI
        # =====================================================
        usable_width = int(section.page_width - section.left_margin - section.right_margin)
        center_pos = usable_width // 2

        # =====================================================
        # BUILDER FOOTER
        # =====================================================
        def build_footer(container):

            p = container.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1

            tabs = p.paragraph_format.tab_stops
            tabs.clear_all()
            tabs.add_tab_stop(center_pos, WD_TAB_ALIGNMENT.CENTER)

            # COPYRIGHT
            run_left = p.add_run(copyright_text)
            run_left.font.name = "Arial"
            run_left.font.size = Pt(10)
            run_left.bold = True

            p.add_run("\t")

            # PAGE
            run_page = p.add_run()
            run_page.font.name = "Arial"
            run_page.font.size = Pt(10)
            run_page.bold = True
            add_field(run_page, "PAGE")

            # TEXT
            run_mid = p.add_run(" dari ")
            run_mid.font.name = "Arial"
            run_mid.font.size = Pt(10)
            run_mid.bold = True

            # TOTAL GLOBAL
            run_total = p.add_run()
            run_total.font.name = "Arial"
            run_total.font.size = Pt(10)
            run_total.bold = True
            add_field(run_total, "NUMPAGES")

        # FOOTER GANJIL
        build_footer(section.footer)

        # FOOTER GENAP
        build_footer(section.even_page_footer)


# Implementasi khusus Engine 7. Definisi ini sengaja menggantikan helper lama
# di atas agar hanya section 3 dan 4 yang memperoleh header/footer SNI.
def setup_headers_footers(doc, doc_title="SNI ISO XXXXX-X:XXXX",
                          copyright_text=None):
    """Pasang header/footer pada seluruh section isi secara dinamis.

    Section 1–2 adalah Cover/Copyright dan section terakhir adalah informasi
    perumus. Section di antaranya dapat berjumlah lebih dari dua karena tabel
    landscape pada dokumen ISO membuat section layout tambahan.
    """
    if len(doc.sections) < 5:
        raise ValueError('Output Engine 6 harus mempunyai sedikitnya lima section.')

    current_year = datetime.now().year
    copyright_text = copyright_text or f"© BSN {current_year}"
    doc_title = re.sub(r'\s+', ' ', doc_title or '').strip()
    if not doc_title:
        doc_title = "SNI ISO XXXXX-X:XXXX"

    # Minta Microsoft Word memperbarui PAGE/SECTIONPAGES saat file dibuka.
    settings = doc.settings.element
    update_fields = settings.find(qn('w:updateFields'))
    if update_fields is None:
        update_fields = OxmlElement('w:updateFields')
        settings.append(update_fields)
    update_fields.set(qn('w:val'), 'true')

    # Header genap/ganjil adalah pengaturan tingkat dokumen di OOXML.
    # Mengaktifkannya hanya sebagai atribut section tidak cukup bagi Word.
    doc.settings.odd_and_even_pages_header_footer = True

    def clear_container(container):
        container.is_linked_to_previous = False
        for table in list(container.tables):
            table._element.getparent().remove(table._element)
        for paragraph in list(container.paragraphs):
            paragraph._element.getparent().remove(paragraph._element)

    def add_field(run, field_name):
        begin = OxmlElement('w:fldChar')
        begin.set(qn('w:fldCharType'), 'begin')
        begin.set(qn('w:dirty'), 'true')
        instruction = OxmlElement('w:instrText')
        instruction.set(qn('xml:space'), 'preserve')
        instruction.text = f' {field_name} '
        separate = OxmlElement('w:fldChar')
        separate.set(qn('w:fldCharType'), 'separate')
        result = OxmlElement('w:t')
        result.text = '1'
        end = OxmlElement('w:fldChar')
        end.set(qn('w:fldCharType'), 'end')
        run._r.extend((begin, instruction, separate, result, end))

    def style_run(run, size=10):
        run.font.name = 'Arial'
        run.font.size = Pt(size)
        run.font.bold = True
        run.font.color.rgb = RGBColor(127, 127, 127)
        r_pr = run._element.get_or_add_rPr()
        r_fonts = r_pr.get_or_add_rFonts()
        for attr in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
            r_fonts.set(qn(f'w:{attr}'), 'Arial')

    def build_header(container, alignment):
        clear_container(container)
        paragraph = container.add_paragraph()
        paragraph.alignment = alignment
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0
        run = paragraph.add_run(doc_title)
        style_run(run, size=11)

    def build_footer(container, section, include_section_total=True):
        clear_container(container)
        paragraph = container.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0

        usable_width = int(
            section.page_width - section.left_margin - section.right_margin
        )
        tabs = paragraph.paragraph_format.tab_stops
        tabs.clear_all()
        tabs.add_tab_stop(usable_width // 2, WD_TAB_ALIGNMENT.CENTER)

        copyright_run = paragraph.add_run(copyright_text)
        style_run(copyright_run, size=10)
        paragraph.add_run('\t')

        page_run = paragraph.add_run()
        style_run(page_run, size=10)
        add_field(page_run, 'PAGE')
        if include_section_total:
            separator_run = paragraph.add_run(' dari ')
            style_run(separator_run, size=10)
            total_run = paragraph.add_run()
            style_run(total_run, size=10)
            add_field(total_run, 'SECTIONPAGES')

    # index 2 = bagian awal (Romawi). Index 3 sampai sebelum section terakhir
    # adalah Content/Bibliography (angka), termasuk layout landscape tambahan.
    formatted_section_indexes = range(2, len(doc.sections) - 1)
    for section_index in formatted_section_indexes:
        number_format = 'lowerRoman' if section_index == 2 else 'decimal'
        section = doc.sections[section_index]
        section.different_first_page_header_footer = False
        section.header_distance = Cm(1.25)
        section.footer_distance = Cm(1.25)

        pg_num = section._sectPr.find(qn('w:pgNumType'))
        if pg_num is None:
            pg_num = OxmlElement('w:pgNumType')
            section._sectPr.append(pg_num)
        # Nomor hanya dimulai ulang pada section awal dan Content pertama.
        # Section layout Content berikutnya harus melanjutkan nomor halaman.
        if section_index in (2, 3):
            pg_num.set(qn('w:start'), '1')
        else:
            pg_num.attrib.pop(qn('w:start'), None)
        pg_num.set(qn('w:fmt'), number_format)

        # Header default dipakai untuk halaman ganjil, sedangkan header even
        # dipakai untuk halaman genap ketika evenAndOddHeaders aktif.
        build_header(section.header, WD_ALIGN_PARAGRAPH.RIGHT)
        build_header(section.even_page_header, WD_ALIGN_PARAGRAPH.LEFT)
        build_header(section.first_page_header, WD_ALIGN_PARAGRAPH.RIGHT)
        # Section Romawi hanya menampilkan nomor halaman. Semua section Content
        # tetap memakai format "halaman dari total halaman section".
        include_section_total = section_index >= 3
        build_footer(section.footer, section, include_section_total)
        build_footer(section.even_page_footer, section, include_section_total)
        build_footer(section.first_page_footer, section, include_section_total)




class DocxOptimizerEngine:
    """
    Engine untuk optimasi dokumen Word sesuai standar ISO/SNI
    Compatible dengan app.py interface
    """
    
    def process(self, input_path, output_path, font_name="Arial", font_size=11,
                enable_headers=False, doc_title="", copyright_text="©BSN 2025"):
        """
        Process dokumen Word untuk formatting ISO/SNI
        
        Args:
            input_path: Path input .docx
            output_path: Path output .docx
            font_name: Font name (default: Arial)
            font_size: Font size (default: 11)
            enable_headers: Enable header/footer setup (default: False)
            doc_title: Document title for header
            copyright_text: Copyright text for footer
            
        Returns:
            (success: bool, message: str)
        """
        try:
            if not input_path or not os.path.isfile(input_path):
                raise FileNotFoundError(f"File input tidak ditemukan: {input_path}")
            doc = Document(input_path)

            # Jangan menolak dokumen hanya berdasarkan jumlah section.
            # ISO/IEC dapat membawa section break tambahan yang sah; validasi
            # struktural paket DOCX lebih stabil daripada hitungan absolut.
            validate_docx(input_path)

            # Field REF dari dokumen ISO sering kehilangan bookmark setelah
            # trim/duplikasi section. Bekukan hasil yang sudah terlihat agar
            # pembaruan field Word tidak menghasilkan pesan error referensi.
            flatten_broken_ref_fields(doc, font_name, font_size)

            all_paragraphs = list(doc.paragraphs)
            toc_index = next(
                (i for i, p in enumerate(all_paragraphs)
                 if re.sub(r'\s+', ' ', p.text or '').strip().casefold() == 'daftar isi'),
                None,
            )
            bibliography_index = next(
                (i for i, p in enumerate(all_paragraphs)
                 if _is_bibliography_paragraph(p)),
                None,
            )
            perumus_index = next(
                (i for i, p in enumerate(all_paragraphs)
                 if re.sub(r'\s+', ' ', p.text or '').strip().casefold()
                 == 'informasi pendukung terkait perumus standar'),
                None,
            )
            if toc_index is None:
                raise ValueError('Heading Daftar isi tidak ditemukan.')
            if perumus_index is None:
                raise ValueError('Heading informasi perumus SNI tidak ditemukan.')
            if not (toc_index < perumus_index):
                raise ValueError(
                    'Urutan dokumen harus Daftar isi → informasi perumus SNI.'
                )
            if bibliography_index is not None and not (
                toc_index < bibliography_index < perumus_index
            ):
                raise ValueError(
                    'Urutan dokumen harus Daftar isi → Bibliography → informasi perumus SNI.'
                )

            # Hanya paragraf dalam rentang ini yang boleh diubah. Paragraf
            # pemisah section 5 berada sebelum heading perumus dan tetap
            # dipertahankan oleh guard sectPr pada tahap penghapusan blank.
            target_paragraph_elements = {
                p._p for p in all_paragraphs[toc_index:perumus_index]
            }
            body_children = list(doc.element.body)
            toc_body_index = body_children.index(all_paragraphs[toc_index]._p)
            perumus_body_index = body_children.index(all_paragraphs[perumus_index]._p)
            target_body_elements = set(body_children[toc_body_index:perumus_body_index])

            # Hyperlink sudah dibersihkan oleh Engine 1. Engine 7 sengaja
            # tidak menyentuh tautan di Cover, Copyright, atau section 5.

            # Fix ukuran font autonumbering "Annex %1" → 12pt (24 half-points)
            # Label "Annex A" dirender dari numbering lvl rPr, bukan dari run paragraf
            _WNS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            
            # Fix heading style sizes AND fonts so auto-generated numbers match
            _sz_val = str(font_size * 2)  # half-points (11pt = 22)
            _styles_el = doc.styles.element
            for _style in _styles_el.iter(f'{{{_WNS_W}}}style'):
                _styleId = _style.get(f'{{{_WNS_W}}}styleId', '')
                if _styleId.startswith('Heading') and 'Char' not in _styleId:
                    _rPr = _style.find(f'{{{_WNS_W}}}rPr')
                    if _rPr is None:
                        from docx.oxml import OxmlElement as _OxmlElement
                        _rPr = _OxmlElement('w:rPr')
                        _style.append(_rPr)
                    # Fix font size
                    for _sz in _rPr.findall(f'{{{_WNS_W}}}sz'):
                        _sz.set(f'{{{_WNS_W}}}val', _sz_val)
                    for _sz in _rPr.findall(f'{{{_WNS_W}}}szCs'):
                        _sz.set(f'{{{_WNS_W}}}val', _sz_val)
                    # Fix font to Arial - remove old rFonts and replace
                    for _rf in _rPr.findall(f'{{{_WNS_W}}}rFonts'):
                        _rPr.remove(_rf)
                    from lxml import etree as _etree
                    _rFonts_new = _etree.SubElement(_rPr, f'{{{_WNS_W}}}rFonts')
                    _rFonts_new.set(f'{{{_WNS_W}}}ascii', font_name)
                    _rFonts_new.set(f'{{{_WNS_W}}}hAnsi', font_name)
                    _rFonts_new.set(f'{{{_WNS_W}}}cs', font_name)
                    _rPr.insert(0, _rFonts_new)

            # Fix numbering.xml rFonts AND Size
            # LOGIKA BARU: Cek Level Text, jika mengandung 'Annex' -> 12pt, lainnya 11pt
            from lxml import etree as _etree
            _num_part = doc.part.numbering_part
            if _num_part is not None:
                _num_root = _num_part._element
                for _lvl in _num_root.iter(f'{{{_WNS_W}}}lvl'):
                    _rPr = _lvl.find(f'{{{_WNS_W}}}rPr')
                    if _rPr is None:
                        _rPr = _etree.SubElement(_lvl, f'{{{_WNS_W}}}rPr')
                    
                    # 1. Replace rFonts
                    for _rf in _rPr.findall(f'{{{_WNS_W}}}rFonts'):
                        _rPr.remove(_rf)
                    _rFonts_new = _etree.Element(f'{{{_WNS_W}}}rFonts')
                    _rFonts_new.set(f'{{{_WNS_W}}}ascii', font_name)
                    _rFonts_new.set(f'{{{_WNS_W}}}hAnsi', font_name)
                    _rFonts_new.set(f'{{{_WNS_W}}}cs', font_name)
                    _rPr.insert(0, _rFonts_new)

                    # 2. Determine Size based on Level Text (Annex vs Others)
                    _sz_val_current = _sz_val # Default 11pt (22 half-points)
                    _lt = _lvl.find(f'{{{_WNS_W}}}lvlText')
                    if _lt is not None:
                        _lt_val = _lt.get(f'{{{_WNS_W}}}val', '') or ''
                        # Jika lvlText mengandung kata 'annex', paksa ukuran 12pt (24 half-points)
                        if 'annex' in _lt_val.lower():
                            _sz_val_current = '24'
                    
                    # Remove old size
                    for _sz in _rPr.findall(f'{{{_WNS_W}}}sz'):
                        _rPr.remove(_sz)
                    for _szCs in _rPr.findall(f'{{{_WNS_W}}}szCs'):
                        _rPr.remove(_szCs)
                    
                    # Add new size
                    _sz_new = _etree.SubElement(_rPr, f'{{{_WNS_W}}}sz')
                    _sz_new.set(f'{{{_WNS_W}}}val', _sz_val_current)
                    
                    _szCs_new = _etree.SubElement(_rPr, f'{{{_WNS_W}}}szCs')
                    _szCs_new.set(f'{{{_WNS_W}}}val', _sz_val_current)

            def clean_format(paragraph, is_heading=False):
                pf = paragraph.paragraph_format
                pf.space_before = Pt(0)
                pf.space_after  = Pt(0)
                pf.line_spacing = 1.0
                for run in paragraph.runs:
                    run.font.name = font_name
                    run.font.size = Pt(font_size)  # selalu 11pt, termasuk heading/pasal

            def normalize_explicit_heading_spacing(paragraph, number_pattern):
                """Ubah pemisah nomor--judul menjadi empat spasi tanpa
                membangun ulang paragraf.

                ``paragraph.text = ...`` menghapus seluruh run lama dan ikut
                menghilangkan ``w:vertAlign`` (superscript/subscript), italic,
                serta format karakter lain. Helper ini hanya mengganti rentang
                whitespace pada run yang sudah ada sehingga rPr asal tetap utuh.
                """
                full_text = ''.join(run.text or '' for run in paragraph.runs)
                match = re.match(
                    rf'^({number_pattern})(\s+)(?=\S)', full_text,
                    re.IGNORECASE,
                )
                if not match:
                    return False

                whitespace_start, whitespace_end = match.span(2)
                offset = 0
                replacement_written = False
                for run in paragraph.runs:
                    original = run.text or ''
                    run_start = offset
                    run_end = offset + len(original)
                    offset = run_end
                    overlap_start = max(run_start, whitespace_start)
                    overlap_end = min(run_end, whitespace_end)
                    if overlap_start >= overlap_end:
                        continue
                    local_start = overlap_start - run_start
                    local_end = overlap_end - run_start
                    replacement = '    ' if not replacement_written else ''
                    run.text = original[:local_start] + replacement + original[local_end:]
                    replacement_written = True
                return replacement_written

            def format_note_preserving_runs(paragraph, full_text):
                """Bold-kan label NOTE/CATATAN tanpa meratakan run isi.

                Persamaan, satuan, dan simbol pada isi note sering disimpan
                sebagai run superscript/subscript. Karena itu run tidak boleh
                dihapus dan paragraf tidak boleh dibuat ulang.
                """
                label_match = re.match(
                    r'^(?:NOTE|CATATAN)(?:\s+\d+)?'
                    r'(?:\s+to\s+entry)?:?',
                    full_text,
                    re.IGNORECASE,
                )
                label_end = label_match.end() if label_match else 0
                offset = 0
                for run in paragraph.runs:
                    run_text = run.text or ''
                    run_start = offset
                    run_end = offset + len(run_text)
                    offset = run_end
                    run.font.size = Pt(10)
                    run.font.name = font_name
                    # Format vertikal/italic dan XML rPr lainnya dipertahankan.
                    run.bold = bool(label_end and run_start < label_end)

            # Style baku untuk heading front matter yang diminta pengguna.
            # Style dibuat bila belum tersedia, kemudian dinormalisasi agar
            # tidak bergantung pada format bawaan dokumen sumber.
            try:
                judul_style = doc.styles['@Judul']
            except KeyError:
                judul_style = doc.styles.add_style('@Judul', WD_STYLE_TYPE.PARAGRAPH)
            judul_style.font.name = font_name
            judul_style.font.size = Pt(12)
            judul_style.font.bold = True
            judul_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            judul_style.paragraph_format.space_before = Pt(0)
            judul_style.paragraph_format.space_after = Pt(0)
            judul_style.paragraph_format.line_spacing = 1.0

            def format_front_matter_heading(paragraph):
                """Terapkan @Judul dan tepat tiga paragraf kosong sesudahnya."""
                paragraph.style = judul_style
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pf = paragraph.paragraph_format
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                pf.line_spacing = 1.0
                for run in paragraph.runs:
                    run.font.name = font_name
                    run.font.size = Pt(12)
                    run.bold = True
                    run.font.color.rgb = RGBColor(0, 0, 0)

                # Hapus blank lama yang langsung mengikuti heading agar proses
                # berulang tetap idempoten, lalu buat tepat tiga blank single.
                next_el = paragraph._element.getnext()
                while next_el is not None and next_el.tag == f'{{{WNS}}}p':
                    text = ''.join(
                        node.text or '' for node in next_el.iter(f'{{{WNS}}}t')
                    ).strip()
                    has_break = bool(next_el.xpath('.//w:br | ./w:pPr/w:sectPr'))
                    if text or has_break:
                        break
                    following = next_el.getnext()
                    next_el.getparent().remove(next_el)
                    next_el = following

                for _ in range(3):
                    blank_after = doc.add_paragraph('')
                    clean_format(blank_after)
                    blank_after.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    paragraph._element.addnext(blank_after._element)

            def next_nonempty_body_text(paragraph):
                """Ambil teks paragraf isi berikutnya tanpa mengubah dokumen."""
                next_el = next_nonempty_body_paragraph_element(paragraph)
                if next_el is None:
                    return ''
                text = ''.join(
                    node.text or '' for node in next_el.iter(f'{{{WNS}}}t')
                ).strip()
                return re.sub(r'\s+', ' ', text)

            def next_nonempty_body_paragraph_element(paragraph):
                """Ambil elemen paragraf isi berikutnya yang tidak kosong."""
                next_el = paragraph._element.getnext()
                while next_el is not None:
                    if next_el.tag == f'{{{WNS}}}p':
                        text = ''.join(
                            node.text or '' for node in next_el.iter(f'{{{WNS}}}t')
                        ).strip()
                        if text:
                            return next_el
                    next_el = next_el.getnext()
                return None

            def format_adoption_bullet(paragraph):
                """Rapikan bullet adopsi: tanda pisah menggantung, teks sejajar."""
                raw = re.sub(r'\s+', ' ', paragraph.text or '').strip()
                content = re.sub(r'^[—–-]\s*', '', raw)
                paragraph.text = f'—\t{content}'
                try:
                    paragraph.style = doc.styles['Normal']
                except KeyError:
                    pass
                p_pr = paragraph._element.get_or_add_pPr()
                num_pr = p_pr.find(qn('w:numPr'))
                if num_pr is not None:
                    p_pr.remove(num_pr)
                pf = paragraph.paragraph_format
                indent = Cm(0.75)
                pf.left_indent = indent
                pf.first_line_indent = -indent
                pf.tab_stops.clear_all()
                pf.tab_stops.add_tab_stop(indent, WD_TAB_ALIGNMENT.LEFT)
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                pf.line_spacing = 1.0
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                for run in paragraph.runs:
                    run.font.name = font_name
                    run.font.size = Pt(font_size)
                    run.bold = False

            def strip_stray_midtext_tabs(paragraph):
                """Ganti tab liar di tengah kalimat (mis. dari konversi PDF
                sumber ISO) menjadi satu spasi, tanpa mengubah tab yang
                memang disengaja untuk hanging indent/bullet/nomor pasal
                (tab semacam itu selalu didahului tanda seperti '—', ')',
                atau '.', bukan huruf/angka biasa).

                Panjang teks per run tidak berubah (tab dan spasi sama-sama
                satu karakter), jadi penggantian dilakukan langsung per run
                tanpa perlu membangun ulang paragraf atau memetakan offset.
                """
                full_text = ''.join(run.text or '' for run in paragraph.runs)
                if '\t' not in full_text:
                    return False
                any_changed = False
                offset = 0
                for run in paragraph.runs:
                    original = run.text or ''
                    if '\t' in original:
                        chars = list(original)
                        run_changed = False
                        for i, ch in enumerate(chars):
                            if ch != '\t':
                                continue
                            global_idx = offset + i
                            before = full_text[global_idx - 1] if global_idx > 0 else ''
                            after = full_text[global_idx + 1] if global_idx + 1 < len(full_text) else ''
                            if re.match(r'\w', before) and re.match(r'\w', after):
                                chars[i] = ' '
                                run_changed = True
                        if run_changed:
                            run.text = ''.join(chars)
                            any_changed = True
                    offset += len(original)
                return any_changed

            def is_subpasal_3(text):
                return bool(re.match(r'^3\.\d+(\s|$)', text))

            # Patterns
            re_split_number = re.compile(r'^(\d[\d\.]*\.?)\s+(.*)')
            re_annex_sub = re.compile(r'^([A-Z]\.\d+[\d\.]*\.?)\s+(.*)')
            re_list_item = re.compile(
                r'^(?:[a-z]\)|[a-z]\.|[A-Z]\)|[A-Z]\.|\([a-z]\)|\([0-9]+\)|[ivxlcdm]+\.|[IVXLCDM]+\.)\s+',
                re.IGNORECASE
            )
            re_copyright = re.compile(r'©\s*ISO.*All\s*rights\s*reserved.*', re.IGNORECASE)
            re_bab = re.compile(r'^(BAB|PASAL|CHAPTER|ARTICLE|SECTION)\s+([IVXLCDM]+|\d+)', re.IGNORECASE)

            # Preprocessing: hapus blank paragraphs
            # Kecuali di bagian tail setelah entri bibliography pertama (preserve spacing asli)
            re_bib_entry = re.compile(r'^\[\d+\]')
            past_bibliography = False
            for p in list(doc.paragraphs):
                if p._p not in target_paragraph_elements:
                    continue
                if not past_bibliography and re_bib_entry.match(p.text.strip()):
                    past_bibliography = True
                if past_bibliography:
                    continue  # Jaga blank asli di section tail
                if _is_truly_empty(p):
                    # Jangan pernah menghapus paragraf pembawa section break.
                    if p._p.xpath('./w:pPr/w:sectPr') or _has_page_break(p):
                        continue
                    try:
                        p._element.getparent().remove(p._element)
                    except:
                        pass

            paragraphs = [
                p for p in doc.paragraphs if p._p in target_paragraph_elements
            ]
            tables = [
                table for table in doc.tables if table._element in target_body_elements
            ]

            # Tambah enter setelah tabel
            for table in tables:
                try:
                    table_element = table._element
                    parent = table_element.getparent()
                    siblings = list(parent)
                    table_index = siblings.index(table_element)
                    
                    need_blank = True
                    if table_index + 1 < len(siblings):
                        next_el = siblings[table_index + 1]
                        p_tag = f'{{{WNS}}}p'
                        if next_el.tag == p_tag:
                            next_text = ''.join(t.text or '' for t in next_el.iter(f'{{{WNS}}}t'))
                            if not next_text.strip():
                                need_blank = False
                    
                    if need_blank:
                        blank_p = doc.add_paragraph("")
                        clean_format(blank_p)
                        table_element.addnext(blank_p._element)
                except:
                    pass

            # Tambah enter setelah gambar
            for p in list(doc.paragraphs):
                if p._p not in target_paragraph_elements:
                    continue
                if not _has_image(p):
                    continue
                try:
                    p_element = p._element
                    parent = p_element.getparent()
                    siblings = list(parent)
                    p_index = siblings.index(p_element)
                    
                    need_blank = True
                    if p_index + 1 < len(siblings):
                        next_el = siblings[p_index + 1]
                        p_tag = f'{{{WNS}}}p'
                        if next_el.tag == p_tag:
                            next_text = ''.join(t.text or '' for t in next_el.iter(f'{{{WNS}}}t'))
                            if not next_text.strip():
                                need_blank = False
                    
                    if need_blank:
                        blank_p = doc.add_paragraph("")
                        clean_format(blank_p)
                        p_element.addnext(blank_p._element)
                except:
                    pass

            paragraphs = [
                p for p in doc.paragraphs if p._p in target_paragraph_elements
            ]

            # State
            title_processed = False
            in_pasal_3_area = False
            in_bibliography_area = False
            in_annex_area = False
            annex_title_zone = 0  # countdown: paragraf setelah "Annex X" yang harus rapat (informative + judul)
            scope_after_content_title_elements = set()

            # Iterasi paragraf
            for p in paragraphs:
                txt = p.text.strip()

                if _has_image(p):
                    clean_format(p)
                    continue

                if not txt:
                    clean_format(p)
                    continue

                if re_copyright.search(txt):
                    try:
                        p._element.getparent().remove(p._element)
                    except:
                        pass
                    continue

                # Prakata dan Introduction selalu memakai @Judul, terlepas
                # dari style/format langsung yang dibawa dokumen sebelumnya.
                if re.fullmatch(r'(prakata|introduction)', txt, re.IGNORECASE):
                    format_front_matter_heading(p)
                    title_processed = True
                    continue

                # Judul Content adalah paragraf @Judul yang langsung diikuti
                # pasal 1 Scope/Ruang lingkup. Terapkan tepat tiga enter
                # sebelum pasal tanpa bergantung pada format dokumen sumber.
                if (
                    p.style and p.style.name == '@Judul'
                    and re.match(
                        r'^1(?:\s{1,4}|\t)+(?:scope|ruang\s+lingkup)\b',
                        next_nonempty_body_text(p),
                        re.IGNORECASE,
                    )
                ):
                    scope_el = next_nonempty_body_paragraph_element(p)
                    if scope_el is not None:
                        scope_after_content_title_elements.add(scope_el)
                    format_front_matter_heading(p)
                    title_processed = True
                    continue

                # Deteksi
                match_number = re_split_number.match(txt)
                match_annex_sub = re_annex_sub.match(txt)
                match_bab = re_bab.match(txt)
                has_heading_style = p.style and "Heading" in p.style.name
                has_numbering = (p._element.pPr is not None and p._element.pPr.numPr is not None)
                is_bold_para = any(run.bold for run in p.runs)

                if not is_bold_para and p.style:
                    try:
                        if p.style.font and p.style.font.bold:
                            is_bold_para = True
                    except:
                        pass

                is_list_item = bool(re_list_item.match(txt))
                is_list_item_exception = is_list_item or (is_bold_para and re.match(r'^[A-Z]\s{2,}', txt))
                is_note_exception = txt.lower().startswith('note') or txt.lower().startswith('catatan')
                is_example_exception = bool(re.match(
                    r'^(?:[—–-]\s*)?(?:EXAMPLE|CONTOH)\b',
                    txt,
                    re.IGNORECASE,
                ))
                is_small_font_exception = any(run.font.size and run.font.size.pt == 10 for run in p.runs)
                is_term_definition_exception = p.style and any(
                    term_type in p.style.name for term_type in ['Term', 'Definition']
                )

                is_real_heading = bool(
                    (not is_list_item) and (not is_example_exception) and (
                        match_number or match_annex_sub or has_heading_style or has_numbering or match_bab or
                        (p.style and p.style.name.upper() in ('ANNEX', 'ANNEX HEADING')) or
                        (is_bold_para and title_processed and not is_list_item_exception 
                         and not is_note_exception and not is_small_font_exception 
                         and not is_term_definition_exception)
                    )
                )


                # ---- GUARD: ANNEX style selalu diproses sebagai special heading,
                #      tidak boleh jatuh ke blok "Judul Utama" meski title_processed=False ----
                _early_annex = p.style and p.style.name.upper() in ('ANNEX', 'ANNEX HEADING')
                if _early_annex and not title_processed:
                    # Tandai title sudah diproses agar paragraf berikutnya tidak dikira judul
                    title_processed = True

                if not title_processed and not is_real_heading:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    clean_format(p)
                    for run in p.runs:
                        run.bold = True
                        run.font.size = Pt(12)
                        run.font.name = font_name
                        # Paksa warna hitam via XML
                        rPr = run._r.get_or_add_rPr()
                        for color_el in rPr.findall(
                            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color'
                        ):
                            rPr.remove(color_el)
                        color_new = OxmlElement('w:color')
                        color_new.set(qn('w:val'), '000000')
                        rPr.insert(0, color_new)
                    for _ in range(2):
                        blank_after = doc.add_paragraph("")
                        clean_format(blank_after)
                        p._element.addnext(blank_after._element)
                    title_processed = True
                    continue

                # B. HEADING
                if is_real_heading:
                    # Figure/Table
                    _is_fig = bool(re.match(r'^(figure|fig\.?|gambar)\b', txt, re.IGNORECASE))
                    _is_tbl = bool(re.match(r'^(table|tabel)\b', txt, re.IGNORECASE))
                    if _is_fig or _is_tbl:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        clean_format(p)
                        # Caption mengikuti bold/non-bold dokumen input. Jangan
                        # memaksa judul Table/Figure menjadi bold.
                        blank_after = doc.add_paragraph("")
                        clean_format(blank_after)
                        p._element.addnext(blank_after._element)
                        continue

                    # Special heading
                    _is_biblio_title = p.style and p.style.name == 'Biblio Title'
                    _is_annex_style = p.style and p.style.name.upper() in ('ANNEX', 'ANNEX HEADING')
                    # Salinan bahasa Inggris dari Engine 5 dapat membawa style
                    # lain (mis. @Judul), tetapi teksnya tetap diawali "Annex".
                    # Perlakukan keduanya sebagai heading Annex yang sama.
                    _is_annex_heading = bool(_is_annex_style) or bool(re.match(
                        r'^(annex|lampiran)\b', txt, re.IGNORECASE
                    ))
                    _is_special = _is_biblio_title or _is_annex_heading or bool(re.match(
                        r'^(bibliography|bibliografi|annex|lampiran|foreword|kata\s+pengantar|index|indeks)',
                        txt, re.IGNORECASE
                    ))
                    if _is_special:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        pf = p.paragraph_format
                        pf.space_before = Pt(0)
                        pf.space_after = Pt(0)

                        if _is_annex_heading:
                            # Heading Annex: Arial 12 bold, center, single,
                            # before/after 0 pt, termasuk salinan Engine 5.
                            pf.line_spacing = 1.0
                            # ANNEX: 12pt bold centered
                            _annex_font_size = 12
                            if not p.runs:
                                r = p.add_run(p.text)
                                r.bold = True
                                r.font.size = Pt(_annex_font_size)
                                r.font.name = font_name
                            else:
                                for run in p.runs:
                                    run.bold = True
                                    run.font.size = Pt(_annex_font_size)
                                    run.font.name = font_name

                            # Hapus <w:br/> ganda sebelum teks judul agar mepet dengan (informative)
                            # Pola: run berisi HANYA <w:br/> tanpa teks → hapus jika run berikutnya punya teks
                            W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                            runs_el = p._element.findall(f'{{{W}}}r')
                            for ri, r_el in enumerate(runs_el):
                                r_texts = r_el.findall(f'{{{W}}}t')
                                r_brs = r_el.findall(f'{{{W}}}br')
                                # Run hanya berisi br tanpa teks → kandidat untuk dihapus
                                if r_brs and not r_texts:
                                    # Cek apakah run berikutnya punya teks judul (bukan br saja)
                                    if ri + 1 < len(runs_el):
                                        next_el = runs_el[ri + 1]
                                        next_brs = next_el.findall(f'{{{W}}}br')
                                        next_texts = next_el.findall(f'{{{W}}}t')
                                        # Jika run berikutnya JUGA dimulai dengan br sebelum teks → hapus br ini
                                        if next_brs and next_texts:
                                            # Hapus br dari run berikutnya (br ekstra sebelum judul)
                                            for br in next_brs:
                                                next_el.remove(br)
                        else:
                            # Non-ANNEX special headings (Bibliography, Foreword, dll): 11pt
                            pf.line_spacing = 1.0
                            if not p.runs:
                                r = p.add_run(p.text)
                                r.bold = True
                                r.font.size = Pt(font_size)
                                r.font.name = font_name
                            else:
                                for run in p.runs:
                                    run.bold = True
                                    run.font.size = Pt(font_size)
                                    run.font.name = font_name

                        if re.match(r'^(bibliography|bibliografi)', txt, re.IGNORECASE):
                            in_bibliography_area = True

                        # ANNEX style: satu paragraf berisi (informative) + judul
                        # Tambah 3 blank setelah → total 3 sebelum sub-heading pertama
                        if _is_annex_heading:
                            in_annex_area = True
                            in_pasal_3_area = False  # Reset agar should_add_enter benar di annex
                            annex_title_zone = 0  # reset, tidak dipakai untuk ANNEX style
                            for _ in range(3):
                                blank_after = doc.add_paragraph("")
                                clean_format(blank_after)
                                p._element.addnext(blank_after._element)
                        else:
                            for _ in range(3):
                                blank_after = doc.add_paragraph("")
                                clean_format(blank_after)
                                p._element.addnext(blank_after._element)
                        continue

                    # Regular heading
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    clean_format(p)
                    for run in p.runs:
                        run.bold = True

                    # Track pasal 3
                    txt_lower = txt.lower()
                    if 'terms and definition' in txt_lower:
                        in_pasal_3_area = True
                    elif match_number:
                        num_str = match_number.group(1).strip('.').strip('(').strip(')')
                        first_char = num_str[0] if num_str else '0'
                        if first_char >= '4':
                            in_pasal_3_area = False
                        elif first_char == '3':
                            in_pasal_3_area = True

                    # Tipe heading
                    is_main_chapter = False
                    if match_bab:
                        is_main_chapter = True
                    elif match_number:
                        num_str = match_number.group(1).strip('.').strip('(').strip(')')
                        is_main_chapter = (num_str.count('.') == 0)
                    elif has_heading_style:
                        is_main_chapter = ("Heading 1" in p.style.name or p.style.name == "Heading")
                    elif has_numbering:
                        try:
                            num_pr = p._element.pPr.numPr
                            if num_pr.ilvl is not None:
                                is_main_chapter = (num_pr.ilvl.val == 0)
                            else:
                                is_main_chapter = True
                        except:
                            is_main_chapter = True

                    is_sub3 = is_subpasal_3(txt)
                    # Annex sub-heading: pola huruf (C.1, A.2, dll) ATAU style 'a2' di dalam annex area
                    is_annex_sub = in_annex_area and (
                        bool(match_annex_sub) or
                        (p.style and p.style.name.lower() in ('a2', 'a3', 'annex sub', 'annex subheading'))
                    )

                    # Spacing sebelum: a2 di annex TIDAK pakai insert_before.
                    # Blank sudah diatur oleh paragraf sebelumnya via should_add_enter + extra_for_annex_sub.
                    # Di area bibliography/tail → jangan tambah blank (spacing sudah preserved dari asli)
                    if not is_annex_sub and not in_bibliography_area:
                        num_enters_before = 1 if (is_main_chapter or is_sub3) else 0
                        if p._element in scope_after_content_title_elements:
                            num_enters_before = 0
                        for _ in range(num_enters_before):
                            blank = p.insert_paragraph_before("")
                            clean_format(blank)

                    # Format ulang
                    if is_annex_sub:
                        # Format sub-heading annex: left align, bold
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        if match_annex_sub:
                            normalize_explicit_heading_spacing(
                                p, r'[A-Z]\.\d+(?:\.\d+)*\.?'
                            )
                        for run in p.runs:
                            run.bold = True
                            run.font.name = font_name
                            run.font.size = Pt(font_size)
                    elif match_annex_sub:
                        normalize_explicit_heading_spacing(
                            p, r'[A-Z]\.\d+(?:\.\d+)*\.?'
                        )
                        for run in p.runs:
                            run.bold = True
                            run.font.name = font_name
                            run.font.size = Pt(font_size)
                    elif match_number and not match_bab:
                        normalize_explicit_heading_spacing(
                            p, r'\d[\d.]*\.?'
                        )
                        for run in p.runs:
                            run.bold = True
                            run.font.name = font_name
                            run.font.size = Pt(font_size) # 11pt

                    # Spacing setelah
                    # Di area bibliography/tail → jangan tambah blank
                    if not is_sub3 and not in_bibliography_area:
                        blank_after = doc.add_paragraph("")
                        clean_format(blank_after)
                        p._element.addnext(blank_after._element)

                # C. PARAGRAF BIASA
                else:
                    # Tab liar di tengah kalimat (contoh: "specifies<TAB>fire"
                    # dari dokumen ISO sumber) → jadi satu spasi biasa.
                    strip_stray_midtext_tabs(p)

                    # Bibliography - MATCH SCREENSHOT FORMAT  
                    is_biblio_entry = p.style and 'Biblio' in p.style.name and 'Title' not in p.style.name
                    bib_match = re.match(r'^\[(\d+)\]\s*(.*)', txt, re.DOTALL)
                    if is_biblio_entry or (in_bibliography_area and bib_match):
                        in_bibliography_area = True
                        
                        # CRITICAL: Remove the style first to avoid inherited formatting
                        try:
                            p.style = 'Normal'
                        except:
                            pass
                        
                        # Apply formatting directly to existing runs
                        first_comma_found = False
                        
                        for run in p.runs:
                            run.font.name = font_name
                            run.font.size = Pt(font_size)
                            run.bold = False
                            
                            run_text = run.text or ''
                            
                            # Number runs and tabs are never italic
                            if run_text.strip() in ['[', ']', '\t'] or run_text.strip().isdigit():
                                run.italic = False
                            elif not first_comma_found:
                                if ',' in run_text:
                                    first_comma_found = True
                                    run.italic = False
                                else:
                                    run.italic = False
                            else:
                                run.italic = True
                        
                        # Set paragraph format with hanging indent
                        pf = p.paragraph_format
                        indent_size = Cm(1.25)
                        pf.left_indent = indent_size
                        pf.first_line_indent = -indent_size
                        pf.space_before = Pt(0)
                        pf.space_after = Pt(0)
                        pf.line_spacing = 1.0
                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        
                        # Add tab stop
                        pf.tab_stops.clear_all()
                        pf.tab_stops.add_tab_stop(indent_size, WD_TAB_ALIGNMENT.LEFT)
                        
                        continue

                    is_table_title = bool(re.match(r'^(table|tabel)\b', txt, re.IGNORECASE))
                    is_figure_title = bool(re.match(r'^(figure|fig\.?|gambar)\b', txt, re.IGNORECASE))

                    if is_table_title or is_figure_title:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        # Pertahankan format run input. Sebagian standar ISO
                        # menggunakan caption Table/Figure non-bold.
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                    clean_format(p)

                    # Dua butir contoh adopsi pada Prakata harus memakai
                    # hanging indent dan tab stop, sehingga baris lanjutan
                    # tepat sejajar dengan awal teks setelah tanda pisah.
                    if (
                        re.match(r'^[—–-]\s*ISO/IEC\b', txt, re.IGNORECASE)
                        and re.search(r'telah\s+diadopsi\s+dengan\s+tingkat\s+keselarasan',
                                      txt, re.IGNORECASE)
                    ):
                        format_adoption_bullet(p)

                    # List item
                    if is_list_item:
                        for run in p.runs:
                            run.bold = False

                    is_bold_non_heading = is_bold_para and not is_real_heading and not is_list_item

                    # Example / Contoh → font size 10 dan non-bold.
                    if re.match(r'^(?:[—–-]\s*)?(EXAMPLE|CONTOH)\b', txt, re.IGNORECASE):
                        for run in p.runs:
                            run.font.size = Pt(10)
                            run.font.name = font_name
                            # Butir contoh seperti "— EXAMPLE 2 ..." pada ISO
                            # bersifat non-bold; jangan diwarisi sebagai heading.
                            run.bold = False

                    # Note
                    if txt.lower().startswith('note') or txt.lower().startswith('catatan'):
                        full_text = p.text
                        format_note_preserving_runs(p, full_text)

                    # Spacing setelah
                    should_add_enter = True
                    if in_pasal_3_area:
                        should_add_enter = not is_bold_non_heading
                    elif in_bibliography_area:
                        should_add_enter = False  # Jaga spacing asli di bibliography/tail
                    else:
                        current_idx = paragraphs.index(p)
                        if current_idx + 1 < len(paragraphs):
                            next_p = paragraphs[current_idx + 1]
                            next_txt = next_p.text.strip()
                            if is_subpasal_3(next_txt):
                                should_add_enter = True

                    # Jangan tambah blank di paragraf terakhir
                    current_idx = paragraphs.index(p)
                    is_last_para = (current_idx == len(paragraphs) - 1)

                    if should_add_enter and not is_last_para:
                        blank_after = doc.add_paragraph("")
                        clean_format(blank_after)
                        p._element.addnext(blank_after._element)

                    # Di area annex: tambah 1 blank ekstra jika paragraf berikutnya a2
                    # Rule dari referensi:
                    #   Body Text  → 1 blank → a2
                    #   Note       → 2 blank → a2
                    #   List       → 2 blank → a2  (blank ekstra pakai style List Continue 1)
                    if in_annex_area and should_add_enter and not is_last_para:
                        current_idx = paragraphs.index(p)
                        if current_idx + 1 < len(paragraphs):
                            next_p = paragraphs[current_idx + 1]
                            next_s = next_p.style.name if next_p.style else ''
                            next_is_a2 = next_s.lower() in ('a2', 'a3')
                            cur_s = p.style.name if p.style else ''
                            is_body_text = cur_s in ('Body Text', 'Normal')
                            if next_is_a2 and not is_body_text:
                                extra = doc.add_paragraph("")
                                clean_format(extra)
                                p._element.addnext(extra._element)

            def _element_text(element):
                return ''.join(
                    node.text or '' for node in element.iter(f'{{{WNS}}}t')
                ).strip()

            def _is_removable_blank_element(element):
                if element is None or element.tag != f'{{{WNS}}}p':
                    return False
                if _element_text(element):
                    return False
                return not bool(element.xpath('.//w:br | ./w:pPr/w:sectPr'))

            def _set_exact_blank_before(paragraph, blank_count):
                previous = paragraph._element.getprevious()
                while _is_removable_blank_element(previous):
                    before_previous = previous.getprevious()
                    previous.getparent().remove(previous)
                    previous = before_previous
                for _ in range(blank_count):
                    blank = paragraph.insert_paragraph_before('')
                    clean_format(blank)

            # Annex: pasal pertama (A.1/B.1/dst.) berjarak tepat tiga enter
            # dari heading Annex; pasal Annex berikutnya tetap dua enter.
            for annex_heading in list(doc.paragraphs):
                annex_text = re.sub(r'\s+', ' ', annex_heading.text or '').strip()
                if re.match(
                    r'^[A-Z]\.\d+(?:\.\d+)*(?:\.?)(?:\s|$)',
                    annex_text,
                    re.IGNORECASE,
                ):
                    previous = annex_heading._element.getprevious()
                    while _is_removable_blank_element(previous):
                        previous = previous.getprevious()
                    previous_text = _element_text(previous) if previous is not None else ''
                    follows_annex_title = bool(re.match(
                        r'^(?:annex|lampiran)\s+[A-Z0-9]\b',
                        previous_text,
                        re.IGNORECASE,
                    ))
                    _set_exact_blank_before(annex_heading, 3 if follows_annex_title else 2)
                    annex_heading.paragraph_format.space_before = Pt(0)
                    annex_heading.paragraph_format.space_after = Pt(0)
                    annex_heading.paragraph_format.line_spacing = 1.0

            # Jaminan final untuk heading Annex, termasuk hasil duplikasi
            # Engine 5 yang style-nya bukan lagi ANNEX.
            for annex_title in list(doc.paragraphs):
                annex_title_text = re.sub(
                    r'\s+', ' ', annex_title.text or ''
                ).strip()
                if not re.match(
                    r'^(?:annex|lampiran)\s+[A-Z0-9]\b',
                    annex_title_text,
                    re.IGNORECASE,
                ):
                    continue
                annex_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                annex_title.paragraph_format.space_before = Pt(0)
                annex_title.paragraph_format.space_after = Pt(0)
                annex_title.paragraph_format.line_spacing = 1.0
                for run in annex_title.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(12)
                    run.bold = True

            # Bibliography: semua entri spacing after 0 pt, single spacing,
            # dan tepat satu paragraf kosong di antara dua entri.
            current_paragraphs = list(doc.paragraphs)
            bibliography_start = next(
                (i for i, paragraph in enumerate(current_paragraphs)
                 if _is_bibliography_paragraph(paragraph)),
                None,
            )
            bibliography_end = next(
                i for i, paragraph in enumerate(current_paragraphs)
                if (paragraph.text or '').strip().casefold()
                == 'informasi pendukung terkait perumus standar'
            )
            bibliography_entries = [
                paragraph
                for paragraph in (
                    current_paragraphs[bibliography_start + 1:bibliography_end]
                    if bibliography_start is not None else []
                )
                if re.match(r'^\[\d+\]', (paragraph.text or '').strip())
            ]
            for entry in bibliography_entries:
                entry.paragraph_format.space_before = Pt(0)
                entry.paragraph_format.space_after = Pt(0)
                entry.paragraph_format.line_spacing = 1.0

            for current_entry, next_entry in zip(
                bibliography_entries, bibliography_entries[1:]
            ):
                cursor = current_entry._element.getnext()
                between = []
                while cursor is not None and cursor is not next_entry._element:
                    between.append(cursor)
                    cursor = cursor.getnext()
                if cursor is not next_entry._element:
                    continue
                if not all(_is_removable_blank_element(el) for el in between):
                    continue
                for element in between:
                    element.getparent().remove(element)
                blank = doc.add_paragraph('')
                clean_format(blank)
                current_entry._element.addnext(blank._element)

            # Set margins
            # Section 3 (Daftar isi/Prakata/Introduction) sampai section tepat
            # sebelum informasi perumus diformat. Rentang ini juga mencakup
            # section tambahan untuk tabel/halaman landscape.
            set_document_margins(
                doc, top_cm=3, inside_cm=3, bottom_cm=2, outside_cm=2,
                section_indexes=range(2, len(doc.sections) - 1),
            )

            # Setup headers/footers (jika diminta)
            if enable_headers and doc_title:
                setup_headers_footers(doc, doc_title, copyright_text)


            atomic_save_docx(doc, output_path)
            return True, output_path

        except Exception as e:
            import traceback
            return False, f"Gagal: {str(e)}\n{traceback.format_exc()}"


class DaftarIsiBibliographyFormatterEngine:
    """Adapter Engine 7 untuk pipeline dashboard Engine 1–6."""

    def process(self, input_docx: str, output_docx: str, **kwargs):
        sni_number = re.sub(
            r'\s+', ' ', kwargs.pop('sni_number', '') or ''
        ).strip() or 'SNI ISO XXXXX-X:XXXX'
        current_year = datetime.now().year
        ok, result = DocxOptimizerEngine().process(
            input_path=input_docx,
            output_path=output_docx,
            enable_headers=True,
            doc_title=sni_number,
            copyright_text=f'© BSN {current_year}',
            **kwargs,
        )
        if not ok:
            return False, None, result

        # Validasi paket akhir, bukan hanya object Document di memori. Engine 7
        # tidak boleh meloloskan REF/PAGEREF atau teks error ke file unduhan.
        broken_reference_parts = []
        field_pattern = re.compile(
            r'(?:<w:instrText[^>]*>\s*|w:instr="[^"]*)'
            r'(?:REF|PAGEREF)\b',
            re.IGNORECASE,
        )
        # Periksa hanya document.xml serta header/footer yang benar-benar
        # direferensikan section akhir. Paket ISO kadang masih menyimpan part
        # header/footer yatim (tidak dipakai) dengan cache REF rusak; part itu
        # tidak tampil di Word dan tidak boleh menggagalkan hasil yang valid.
        validated_doc = Document(result)
        active_parts = {'word/document.xml'}
        for section in validated_doc.sections:
            for container in (
                section.header, section.first_page_header,
                section.even_page_header, section.footer,
                section.first_page_footer, section.even_page_footer,
            ):
                active_parts.add(str(container.part.partname).lstrip('/'))
        with zipfile.ZipFile(result, 'r') as archive:
            for part_name in sorted(active_parts):
                if part_name not in archive.namelist():
                    continue
                xml_text = archive.read(part_name).decode('utf-8', errors='ignore')
                if (
                    'reference source not found' in xml_text.casefold()
                    or field_pattern.search(xml_text)
                ):
                    broken_reference_parts.append(part_name)
        if broken_reference_parts:
            return (
                False,
                None,
                'Engine 7 masih menemukan referensi Word yang rusak pada: '
                + ', '.join(broken_reference_parts),
            )

        check = Document(result)
        if not check.settings.odd_and_even_pages_header_footer:
            return False, None, 'Header genap dan ganjil belum diaktifkan.'
        for section_index in range(2, len(check.sections) - 1):
            expected_format = 'lowerRoman' if section_index == 2 else 'decimal'
            pg_num = check.sections[section_index]._sectPr.find(qn('w:pgNumType'))
            if pg_num is None or pg_num.get(qn('w:fmt')) != expected_format:
                return (
                    False,
                    None,
                    f'Format nomor halaman section {section_index + 1} tidak sesuai.',
                )
            header_text = ' '.join(
                p.text for p in check.sections[section_index].header.paragraphs
            ).strip()
            even_header_text = ' '.join(
                p.text
                for p in check.sections[section_index].even_page_header.paragraphs
            ).strip()
            footer_text = ' '.join(
                p.text for p in check.sections[section_index].footer.paragraphs
            ).strip()
            odd_paragraphs = check.sections[section_index].header.paragraphs
            even_paragraphs = check.sections[section_index].even_page_header.paragraphs
            header_alignment_ok = (
                bool(odd_paragraphs)
                and odd_paragraphs[-1].alignment == WD_ALIGN_PARAGRAPH.RIGHT
                and bool(even_paragraphs)
                and even_paragraphs[-1].alignment == WD_ALIGN_PARAGRAPH.LEFT
            )
            if (
                sni_number not in header_text
                or sni_number not in even_header_text
                or not header_alignment_ok
                or f'© BSN {current_year}' not in footer_text
            ):
                return False, None, 'Header atau footer Engine 7 tidak berhasil dibuat.'
        return (
            True,
            result,
            'Formatting Engine 7 berhasil diterapkan dari Daftar isi sampai '
            f'Bibliography. Header ganjil rata kanan dan header genap rata kiri '
            f'memakai {sni_number}; footer memakai © BSN '
            f'{current_year}. Section 3 tetap Romawi dan seluruh section Content '
            'tetap memakai angka desimal.',
        )


# API publik Engine 7. Ekspor eksplisit ini mencegah ketidakcocokan nama kelas
# antara app.py baru dan file Engine 7 dari revisi sebelumnya.
ENGINE7_API_VERSION = '7.2'
Engine7 = DaftarIsiBibliographyFormatterEngine
Engine7Formatter = DaftarIsiBibliographyFormatterEngine


def get_engine7_class():
    """Kembalikan kelas adapter resmi yang digunakan dashboard."""
    return DaftarIsiBibliographyFormatterEngine


__all__ = [
    'DaftarIsiBibliographyFormatterEngine',
    'DocxOptimizerEngine',
    'Engine7',
    'Engine7Formatter',
    'ENGINE7_API_VERSION',
    'get_engine7_class',
]
