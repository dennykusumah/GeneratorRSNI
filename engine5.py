"""Engine 5 — duplikasi Introduction dan Content Inggris sebelum Bibliography.

Input yang diharapkan adalah keluaran Engine 4 dengan empat section. Salinan
Introduction dan seluruh blok Content (termasuk Annex) ditempatkan di section 4
tepat sebelum Bibliography. Tidak ada section baru yang dibuat.
"""

from __future__ import annotations

import os
import re
from copy import deepcopy

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class IntroductionContentDuplicatorEngine:
    MARKER_NAME = "Engine5DuplicateStart"

    @staticmethod
    def _normalize(text: str) -> str:
        return re.sub(r"\s+", " ", text or "").strip()

    @classmethod
    def _element_text(cls, element) -> str:
        return cls._normalize("".join(element.xpath(".//w:t/text()")))

    @staticmethod
    def _has_section_break(element) -> bool:
        if element.tag != qn("w:p"):
            return False
        ppr = element.find(qn("w:pPr"))
        return ppr is not None and ppr.find(qn("w:sectPr")) is not None

    @staticmethod
    def _page_break_paragraph():
        paragraph = OxmlElement("w:p")
        run = OxmlElement("w:r")
        page_break = OxmlElement("w:br")
        page_break.set(qn("w:type"), "page")
        run.append(page_break)
        paragraph.append(run)
        return paragraph

    @staticmethod
    def _bookmark_ids(doc: Document) -> set[int]:
        ids = set()
        for item in doc.element.xpath(".//w:bookmarkStart"):
            value = item.get(qn("w:id"))
            if value and value.isdigit():
                ids.add(int(value))
        return ids

    def _marker_paragraph(self, bookmark_id: int):
        paragraph = self._page_break_paragraph()
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(bookmark_id))
        start.set(qn("w:name"), self.MARKER_NAME)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(bookmark_id))
        paragraph.insert(0, start)
        paragraph.insert(1, end)
        return paragraph

    def _already_processed(self, doc: Document) -> bool:
        return any(
            item.get(qn("w:name")) == self.MARKER_NAME
            for item in doc.element.xpath(".//w:bookmarkStart")
        )

    @staticmethod
    def _force_content_decimal(doc: Document):
        """Nomor angka berlanjut pada semua layout section Content.

        Dokumen dengan tabel landscape dapat mempunyai section 4, 5, dst.
        Hanya section Content pertama yang memulai dari 1; section berikutnya
        meneruskan PAGE tanpa restart.
        """
        for index, section in enumerate(doc.sections[3:], start=3):
            section_properties = section._sectPr
            page_number_type = section_properties.find(qn("w:pgNumType"))
            if page_number_type is None:
                page_number_type = OxmlElement("w:pgNumType")
                section_properties.append(page_number_type)
            page_number_type.set(qn("w:fmt"), "decimal")
            if index == 3:
                page_number_type.set(qn("w:start"), "1")
            else:
                page_number_type.attrib.pop(qn("w:start"), None)

    def duplicate_before_bibliography(self, input_docx: str, output_docx: str):
        try:
            if not input_docx or not os.path.isfile(input_docx):
                raise FileNotFoundError(f"File input tidak ditemukan: {input_docx}")

            doc = Document(input_docx)
            input_section_count = len(doc.sections)
            if input_section_count < 4:
                raise ValueError(
                    f"Output Engine 4 harus memiliki minimal 4 section; "
                    f"ditemukan {input_section_count}."
                )

            if self._already_processed(doc):
                self._force_content_decimal(doc)
                doc.save(output_docx)
                return True, output_docx

            body = doc.element.body
            elements = list(body)
            break_indices = [
                index for index, element in enumerate(elements)
                if self._has_section_break(element)
            ]
            if len(break_indices) < 3:
                raise ValueError(
                    "Batas section tidak sesuai: dokumen harus mempunyai "
                    "sedikitnya tiga pemisah section."
                )
            section4_start = break_indices[2] + 1

            intro_index = next(
                (
                    index for index, element in enumerate(elements[:break_indices[2]])
                    if element.tag == qn("w:p")
                    and self._element_text(element).casefold() == "introduction"
                ),
                None,
            )
            bibliography_index = next(
                (
                    index for index, element in enumerate(elements[section4_start:], section4_start)
                    if element.tag == qn("w:p")
                    and self._element_text(element).casefold() == "bibliography"
                ),
                None,
            )
            if bibliography_index is None:
                raise ValueError("Heading Bibliography tidak ditemukan di section 4.")

            introduction = (
                elements[intro_index:break_indices[2]]
                if intro_index is not None else []
            )
            content = elements[section4_start:bibliography_index]
            if not content:
                raise ValueError("Blok Content kosong.")

            bibliography = elements[bibliography_index]
            used_ids = self._bookmark_ids(doc)
            marker_id = max(used_ids, default=0) + 1

            # Setiap addprevious mempertahankan urutan bila dilakukan berurutan.
            bibliography.addprevious(self._marker_paragraph(marker_id))
            for element in introduction:
                bibliography.addprevious(deepcopy(element))
            # Marker sudah mengandung page break. Page break kedua hanya
            # diperlukan untuk memisahkan Introduction dari Content. Pada
            # standar tanpa Introduction, dua break berurutan membuat satu
            # halaman kosong dengan header/footer.
            if introduction:
                bibliography.addprevious(self._page_break_paragraph())
            for element in content:
                bibliography.addprevious(deepcopy(element))

            copied_section_breaks = sum(
                1 for element in introduction + content
                if self._has_section_break(element)
            )
            expected_section_count = input_section_count + copied_section_breaks
            self._force_content_decimal(doc)
            out_dir = os.path.dirname(os.path.abspath(output_docx))
            os.makedirs(out_dir, exist_ok=True)
            doc.save(output_docx)

            check = Document(output_docx)
            if len(check.sections) != expected_section_count:
                raise RuntimeError("Engine 5 mengubah jumlah section dokumen.")
            for section_index, section in enumerate(check.sections[3:], start=3):
                fmt = section._sectPr.find(qn("w:pgNumType"))
                if fmt is None or fmt.get(qn("w:fmt")) != "decimal":
                    raise RuntimeError(
                        f"Penomoran halaman content section {section_index + 1} "
                        "bukan angka desimal."
                    )
            return True, output_docx
        except Exception as exc:
            return False, f"IntroductionContentDuplicatorEngine Error: {exc}"

    def process(self, input_docx: str, output_docx: str, **_kwargs):
        ok, result = self.duplicate_before_bibliography(input_docx, output_docx)
        if not ok:
            return False, None, result
        return (
            True,
            result,
            "Salinan Content berbahasa Inggris beserta Introduction jika "
            "tersedia berhasil disisipkan tepat sebelum Bibliography. "
            "Penomoran seluruh layout section Content menggunakan angka desimal.",
        )


# Alias ringkas untuk pemakaian langsung bila diperlukan.
Engine5 = IntroductionContentDuplicatorEngine
