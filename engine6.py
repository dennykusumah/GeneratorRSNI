"""
Engine6: PrakataPendahuluanEngine (v5 - Fixed)
==============================================
Engine untuk menyisipkan halaman Prakata dan Pendahuluan setelah Daftar Isi.

Perbaikan v5:
- Menggunakan font "Arial" untuk Em Dash bullet (sebelumnya "Symbol" menyebabkan error tampilan).
- Indent mepet margin (Flush Left).
- Deteksi duplikat definisi numbering.

Aturan:
- SATU SECTION dengan Daftar Isi (konten disisipkan SEBELUM sectPr penutup DI).
- Margin, penomoran Romawi, header/footer mewarisi dari sectPr DI.
- Seluruh run menyertakan <w:noproof/>.
"""

import re
import zipfile
from lxml import etree

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


# ─────────────────────────────────────────────────────────────────────────────
# XML HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _esc(t: str) -> str:
    return (t.replace('&', '&amp;')
             .replace('<', '&lt;')
             .replace('>', '&gt;')
             .replace('"', '&quot;'))


def _rpr(bold=False, italic=False, size_pt=11, color=None) -> str:
    """Return <w:rPr> — selalu menyertakan <w:noproof/>."""
    b  = '<w:b/>'  if bold   else ''
    i  = '<w:i/>'  if italic else ''
    sz = int(size_pt * 2)
    cl = f'<w:color w:val="{color}"/>' if color else ''
    return (
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'{b}{i}{cl}'
        f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
        f'<w:noproof/>'
        f'</w:rPr>'
    )


def _run(text: str, bold=False, italic=False, size_pt=11, color=None) -> str:
    return f'<w:r>{_rpr(bold, italic, size_pt, color)}<w:t xml:space="preserve">{_esc(text)}</w:t></w:r>'


def _para(runs_xml: str, align='both', space_after=0, indent_left=0) -> str:
    jc  = f'<w:jc w:val="{align}"/>'
    sa  = f'<w:spacing w:before="0" w:after="{space_after}"/>'
    ind = f'<w:ind w:left="{indent_left}"/>' if indent_left else ''
    return (
        f'<w:p><w:pPr>{sa}{jc}{ind}'
        f'<w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:noproof/></w:rPr>'
        f'</w:pPr>{runs_xml}</w:p>'
    )


def _empty(align='both') -> str:
    return _para('', align=align)


# ─────────────────────────────────────────────────────────────────────────────
# NUMBERING (Em-Dash Bullet List)
# ─────────────────────────────────────────────────────────────────────────────

def _patch_numbering(numbering_bytes: bytes | None) -> tuple[bytes, int]:
    """
    Sisipkan abstractNum + num em-dash ke numbering.xml.
    PERBAIKAN: Gunakan font "Arial" agar karakter Em Dash (—) tampil benar.
    Indent: left=360, hanging=360 (Mepet Margin/Flush Left).
    """
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    
    # Cek apakah sudah ada definisi untuk em-dash (string literal)
    em_dash_marker = '—'
    if numbering_bytes is not None:
        try:
            tree_check = etree.fromstring(numbering_bytes)
            # Cari abstractNum yang lvlText-nya em dash
            for ab in tree_check.findall(f'{{{NS}}}abstractNum'):
                for lvl in ab.findall(f'{{{NS}}}lvl'):
                    txt_el = lvl.find(f'{{{NS}}}lvlText')
                    if txt_el is not None and txt_el.get(f'{{{NS}}}val') == em_dash_marker:
                        # Sudah ada! Ambil numId yang ada
                        for num in tree_check.findall(f'{{{NS}}}num'):
                            # Kita cari yang paling terakhir atau buat baru? Lebih aman buat baru tapi refer ke abstrak ini
                            # Atau gunakan yang sudah ada?
                            # Agar konsisten, kita buat numId baru yang merujuk ke abstractNum yang sudah ada.
                            ab_id = ab.get(f'{{{NS}}}abstractNumId')
                            
                            # Cari max numId
                            existing_nums = tree_check.findall(f'{{{NS}}}num')
                            max_id = max((int(n.get(f'{{{NS}}}numId', 0)) for n in existing_nums), default=0)
                            new_num_id = str(max_id + 1)
                            
                            num_xml = f'<w:num w:numId="{new_num_id}" xmlns:w="{NS}"><w:abstractNumId w:val="{ab_id}"/></w:num>'
                            tree_check.append(etree.fromstring(num_xml))
                            
                            return (
                                etree.tostring(tree_check, xml_declaration=True, encoding='UTF-8', standalone=True),
                                new_num_id
                            )
        except Exception:
            pass # Jika gagal parsing, buat baru

    if numbering_bytes is None:
        abstract_id = 1
        num_id      = 1
        xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:numbering xmlns:w="{NS}">'
            f'<w:abstractNum w:abstractNumId="{abstract_id}">'
            f'<w:multiLevelType w:val="singleLevel"/>'
            f'<w:lvl w:ilvl="0">'
            f'<w:start w:val="1"/><w:numFmt w:val="bullet"/>'
            f'<w:lvlText w:val="\u2014"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>'
            f'<w:rPr>'
            # PERBAIKAN: Font Arial, bukan Symbol
            f'<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
            f'<w:sz w:val="22"/><w:szCs w:val="22"/>'
            f'</w:rPr>'
            f'</w:lvl></w:abstractNum>'
            f'<w:num w:numId="{num_id}">'
            f'<w:abstractNumId w:val="{abstract_id}"/>'
            f'</w:num>'
            f'</w:numbering>'
        )
        return xml.encode('utf-8'), num_id

    # ── Parse numbering.xml yang sudah ada ──────────────────────────────────
    tree = etree.fromstring(numbering_bytes)

    # ── REPAIR URUTAN ──────────────────────────────────────────────────────
    children_list = list(tree)
    first_num_idx = next(
        (i for i, el in enumerate(children_list) if el.tag == f'{{{NS}}}num'),
        None
    )
    if first_num_idx is not None:
        out_of_order = [
            (i, el) for i, el in enumerate(children_list)
            if el.tag == f'{{{NS}}}abstractNum' and i > first_num_idx
        ]
        for _, el in reversed(out_of_order):
            tree.remove(el)
            first_num = tree.find(f'{{{NS}}}num')
            first_num.addprevious(el)

    existing_abstract_ids = [
        int(el.get(f'{{{NS}}}abstractNumId', 0))
        for el in tree.findall(f'{{{NS}}}abstractNum')
    ]
    existing_num_ids = [
        int(el.get(f'{{{NS}}}numId', 0))
        for el in tree.findall(f'{{{NS}}}num')
    ]

    abstract_id = (max(existing_abstract_ids) + 1) if existing_abstract_ids else 1
    num_id      = (max(existing_num_ids)      + 1) if existing_num_ids      else 1

    # ── Buat elemen abstractNum baru (Arial Font) ──────────────────────────
    abs_xml = (
        f'<w:abstractNum w:abstractNumId="{abstract_id}" xmlns:w="{NS}">'
        f'<w:multiLevelType w:val="singleLevel"/>'
        f'<w:lvl w:ilvl="0">'
        f'<w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'<w:lvlText w:val="\u2014"/><w:lvlJc w:val="left"/>'
        f'<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>'
        f'<w:rPr>'
        # PERBAIKAN: Font Arial
        f'<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'<w:sz w:val="22"/><w:szCs w:val="22"/>'
        f'</w:rPr>'
        f'</w:lvl>'
        f'</w:abstractNum>'
    )
    abs_el = etree.fromstring(abs_xml.encode('utf-8'))

    # ── Buat elemen num baru ─────────────────────────────────────────────────
    num_xml = (
        f'<w:num w:numId="{num_id}" xmlns:w="{NS}">'
        f'<w:abstractNumId w:val="{abstract_id}"/>'
        f'</w:num>'
    )
    num_el = etree.fromstring(num_xml.encode('utf-8'))

    # ── Sisipkan dengan urutan BENAR ─────────────────────────────────────────
    first_num = tree.find(f'{{{NS}}}num')
    if first_num is not None:
        first_num.addprevious(abs_el)
    else:
        tree.append(abs_el)

    tree.append(num_el)

    return (
        etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True),
        num_id,
    )


def _bullet_para(runs_xml: str, num_id: int) -> str:
    """
    Paragraf bullet resmi Word dengan simbol em-dash.
    """
    return (
        f'<w:p>'
        f'<w:pPr>'
        f'<w:spacing w:before="0" w:after="0"/>'
        f'<w:jc w:val="both"/>'
        f'<w:numPr>'
        f'<w:ilvl w:val="0"/>'
        f'<w:numId w:val="{num_id}"/>'
        f'</w:numPr>'
        f'<w:ind w:left="360" w:hanging="360"/>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'<w:noproof/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'{runs_xml}'
        f'</w:p>'
    )


def _page_break_para() -> str:
    return (
        '<w:p><w:pPr>'
        '<w:spacing w:before="0" w:after="0"/>'
        '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:noproof/></w:rPr>'
        '</w:pPr>'
        '<w:r><w:rPr><w:noproof/></w:rPr><w:br w:type="page"/></w:r>'
        '</w:p>'
    )


# ─────────────────────────────────────────────────────────────────────────────
# CONTENT BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

def _build_prakata(sni_number, title_id, title_en, ref_standard, bsn_year, num_id: int = 1):
    xmls = []
    xmls.append(_page_break_para())
    xmls.append(_para(_run('Prakata', bold=True, size_pt=12), align='center'))
    xmls.append(_empty(align='center'))
    xmls.append(_empty(align='center'))
    xmls.append(_empty())

    p1_runs = (
        _run(f'{sni_number}, ')
        + _run(title_id, italic=True)
        + _run(', merupakan standar yang disusun dengan jalur adopsi tingkat keselarasan identik dari ')
        + _run(ref_standard)
        + _run(', ')
        + _run(title_en, italic=True)
        + _run(f', dengan metode adopsi terjemahan dua bahasa dan ditetapkan oleh BSN Tahun {bsn_year}.')
    )
    xmls.append(_para(p1_runs, align='both'))
    xmls.append(_empty())

    p2 = (
        f'Dalam Standar ini istilah \u201cthis International Standard\u201d pada standar '
        f'{ref_standard} yang diadopsi diganti dengan \u201cthis Standard\u201d dan '
        f'diterjemahkan menjadi \u201cStandar ini\u201d.'
    )
    xmls.append(_para(_run(p2), align='both'))
    xmls.append(_empty())

    xmls.append(_para(_run(
        'Terdapat standar yang dijadikan sebagai acuan normatif dalam Standar ini '
        'telah diadopsi menjadi SNI, yaitu:'
    ), align='both'))
    xmls.append(_empty())
    for li in [
        'ISO/IEC XXXX-X:YYYY, ZZZZ, telah diadopsi dengan tingkat keselarasan identik menjadi SNI ISO/IEC XXXX-X:YYYY, ZZZZ'
    ]:
        xmls.append(_bullet_para(_run(li), num_id))
    xmls.append(_empty())

    for li in [
        'ISO/IEC XXXX-X:YYYY, ZZZZ, telah diadopsi dengan tingkat keselarasan identik menjadi SNI ISO/IEC XXXX-X:YYYY, ZZZZ'
    ]:
        xmls.append(_bullet_para(_run(li), num_id))
    xmls.append(_empty())

    p_runs = (
        _run('Standar ini disusun oleh Komite Teknis XX-YY, ZZZZ. ')
        + _run('Standar ini telah dibahas melalui rapat teknis dan disepakati dalam rapat konsensus ')
        + _run('pada tanggal XXXX di YYYY, yang dihadiri oleh para pemangku kepentingan (')
        + _run('stakeholders', italic=True)
        + _run(') terkait yaitu perwakilan dari pemerintah, pelaku usaha, konsumen, dan pakar. ')
        + _run('Standar ini telah melalui tahap jajak pendapat pada tanggal XXXX sampai dengan YYYY ')
        + _run('dengan hasil akhir disetujui menjadi SNI.')
        )

    xmls.append(_para(p_runs, align='both'))

    xmls.append(_para(_run(
        'Untuk menghindari kesalahan dalam penggunaan Standar ini, disarankan bagi pengguna '
        'standar menggunakan dokumen SNI yang dicetak dengan tinta berwarna (dapat mencantumkan '
        'kode tingkat warna Red Green Blue (RGB) jika diperlukan untuk cetak gambar dengan '
        'warna yang lebih akurat).'
    ), align='both'))
    xmls.append(_empty())

    xmls.append(_para(_run(
        f'Apabila pengguna menemukan keraguan dalam Standar ini, maka disarankan untuk melihat '
        f'standar aslinya, yaitu {ref_standard}, dan/atau dokumen terkait lain yang menyertainya.'
    ), align='both'))
    xmls.append(_empty())

    xmls.append(_para(_run(
        'Perlu diperhatikan bahwa kemungkinan beberapa unsur dari Standar ini dapat berupa '
        'kekayaan intelektual. Namun selama proses perumusan SNI, Badan Standardisasi Nasional '
        'telah memperhatikan penyelesaian terhadap kemungkinan adanya kekayan intelektual terkait '
        'substansi SNI. Apabila setelah penetapan SNI masih terdapat permasalahan terkait kekayaan '
        'intelektual, Badan Standardisasi Nasional tidak bertanggung jawab mengenai bukti, validitas, '
        'dan ruang lingkup dari kekayaan intelektual tersebut. Badan Standardisasi Nasional tidak '
        'bertanggung jawab mengenai bukti, validitas, dan ruang lingkup dari kekayaan intelektual tersebut.'
    ), align='both'))

    for _ in range(6):
        xmls.append(_empty(align='center'))

    return xmls


def _build_pendahuluan():
    xmls = []
    xmls.append(_page_break_para())
    xmls.append(_para(_run('Pendahuluan', bold=True, size_pt=12), align='center'))
    xmls.append(_empty(align='center'))
    xmls.append(_empty(align='center'))
    xmls.append(_empty(align='center'))

    lorem = (
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit. '
        'Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. '
        'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip '
        'ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit '
        'esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non '
        'proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'
    )
    xmls.append(_para(_run(lorem, bold=False, size_pt=11, color='FF0000'), align='both'))
    xmls.append(_empty(align='both'))
    lorem = (
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit. '
        'Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. '
        'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip '
        'ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit '
        'esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non '
        'proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'
    )
    xmls.append(_para(_run(lorem, bold=False, size_pt=11, color='FF0000'), align='both'))
    return xmls


# ─────────────────────────────────────────────────────────────────────────────
# PARSE HELPER
# ─────────────────────────────────────────────────────────────────────────────

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


# ─────────────────────────────────────────────────────────────────────────────
# SETTINGS PATCHER
# ─────────────────────────────────────────────────────────────────────────────

def _patch_settings(settings_bytes: bytes) -> bytes:
    xml = settings_bytes.decode('utf-8')
    inject = ''
    if '<w:hideSpellingErrors' not in xml:
        inject += '<w:hideSpellingErrors/>\n'
    if '<w:hideGrammaticalErrors' not in xml:
        inject += '<w:hideGrammaticalErrors/>\n'

    if not inject:
        return settings_bytes

    m = re.search(r'<w:settings\b[^>]*>', xml)
    if m:
        pos = m.end()
        xml = xml[:pos] + '\n' + inject + xml[pos:]
    else:
        xml = xml.replace('</w:settings>', inject + '</w:settings>')

    return xml.encode('utf-8')


# ─────────────────────────────────────────────────────────────────────────────
# ENGINE CLASS
# ─────────────────────────────────────────────────────────────────────────────

class PrakataPendahuluanEngine:

    def insert(
        self,
        input_docx:   str,
        output_docx:  str,
        sni_number:   str = 'SNI ISO/IEC XXXX:20XX',
        title_id:     str = 'Judul Bahasa Indonesia',
        title_en:     str = 'Title in English',
        ref_standard: str = 'ISO XXXX:20XX',
        bsn_year:     str = '20XX',
    ) -> tuple[bool, str]:
        try:
            with zipfile.ZipFile(input_docx, 'r') as z:
                files = {n: z.read(n) for n in z.namelist()}

            if 'word/settings.xml' in files:
                files['word/settings.xml'] = _patch_settings(files['word/settings.xml'])

            numbering_existed = 'word/numbering.xml' in files
            files['word/numbering.xml'], emdash_num_id = _patch_numbering(
                files.get('word/numbering.xml')
            )

            if not numbering_existed:
                rels_key = 'word/_rels/document.xml.rels'
                if rels_key in files:
                    rels_xml = files[rels_key].decode('utf-8')
                    num_rel = (
                        '<Relationship Id="rIdNumbering" '
                        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" '
                        'Target="numbering.xml"/>'
                    )
                    if 'numbering.xml' not in rels_xml:
                        rels_xml = rels_xml.replace('</Relationships>', num_rel + '</Relationships>')
                        files[rels_key] = rels_xml.encode('utf-8')
                ct_key = '[Content_Types].xml'
                if ct_key in files:
                    ct_xml = files[ct_key].decode('utf-8')
                    num_ct = (
                        '<Override PartName="/word/numbering.xml" '
                        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
                    )
                    if 'numbering.xml' not in ct_xml:
                        ct_xml = ct_xml.replace('</Types>', num_ct + '</Types>')
                        files[ct_key] = ct_xml.encode('utf-8')

            tree = etree.fromstring(files['word/document.xml'])
            body = tree.find(f'{{{NS_W}}}body')
            children = list(body)

            di_sect_idx = None
            for i, child in enumerate(children):
                if child.tag != f'{{{NS_W}}}p':
                    continue
                pPr = child.find(f'{{{NS_W}}}pPr')
                if pPr is None:
                    continue
                sectPr = pPr.find(f'{{{NS_W}}}sectPr')
                if sectPr is None:
                    continue
                pgNum = sectPr.find(f'{{{NS_W}}}pgNumType')
                if pgNum is not None and pgNum.get(f'{{{NS_W}}}fmt') == 'lowerRoman':
                    di_sect_idx = i
                    break

            if di_sect_idx is None:
                count = 0
                for i, child in enumerate(children):
                    if child.tag != f'{{{NS_W}}}p':
                        continue
                    pPr = child.find(f'{{{NS_W}}}pPr')
                    if pPr is None:
                        continue
                    if pPr.find(f'{{{NS_W}}}sectPr') is not None:
                        count += 1
                        if count == 3:
                            di_sect_idx = i
                            break
                if di_sect_idx is None:
                    return False, 'PrakataPendahuluanEngine: sectPr Daftar Isi tidak ditemukan.'

            all_xmls = (
                _build_prakata(sni_number, title_id, title_en, ref_standard, bsn_year,
                               num_id=emdash_num_id)
                + _build_pendahuluan()
            )
            elements = _parse_elements(all_xmls)
            for offset, el in enumerate(elements):
                body.insert(di_sect_idx + offset, el)

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
            return False, f'PrakataPendahuluanEngine Error: {str(e)}\n{traceback.format_exc()}'
