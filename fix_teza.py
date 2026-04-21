#!/usr/bin/env python3
"""
Repara stilurile Heading 1/2/3 in teza_anca_aliciuc_v1_compressed.docx
si insereaza un camp TOC automat pe care Word il va popula la deschidere.
"""

import re
import copy
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

INPUT  = "teza-master/teza_anca_aliciuc_v1_compressed.docx"
OUTPUT = "teza_anca_aliciuc_v1_fixed.docx"

# ─── Heading patterns ────────────────────────────────────────────────────────
H1_PATTERNS = re.compile(
    r'^(CUPRINS|INTRODUCERE|CAPITOLUL\s+[IVX]+\.|CONCLUZII|BIBLIOGRAFIE|ANEXA?\s+\d+)',
    re.IGNORECASE
)
H2_PATTERN = re.compile(r'^\d+\.\d+[\.\s]')
H3_PATTERN = re.compile(
    r'^(ZIUA\s+\d+\s*[-–]|Ziua\s+\d+\s*[-–]|Activitatea\s+\d+:)',
    re.IGNORECASE
)

def get_run_size_pt(para):
    for r in para.runs:
        if r.font.size:
            return round(r.font.size / 12700)
    return None

def is_bold(para):
    return any(r.bold for r in para.runs if r.text.strip())

def classify(para):
    text = para.text.strip()
    if not text:
        return None
    size = get_run_size_pt(para)
    bold = is_bold(para)

    if bold and size == 14 and H1_PATTERNS.match(text):
        return "H1"
    if bold and size == 14:
        return "H1"
    if bold and size == 13:
        return "H2"
    if H2_PATTERN.match(text) and bold:
        return "H2"
    if H3_PATTERN.match(text) and bold:
        return "H3"
    return None

# ─── Setup heading styles in document ────────────────────────────────────────
def setup_heading_styles(doc):
    h1 = doc.styles['Heading 1']
    h1.font.name = 'Times New Roman'
    h1.font.size = Pt(14)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(0, 0, 0)
    h1.paragraph_format.space_before = Pt(0)
    h1.paragraph_format.space_after  = Pt(12)
    h1.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.first_line_indent = Cm(0)

    h2 = doc.styles['Heading 2']
    h2.font.name = 'Times New Roman'
    h2.font.size = Pt(13)
    h2.font.bold = True
    h2.font.color.rgb = RGBColor(0, 0, 0)
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after  = Pt(6)
    h2.paragraph_format.first_line_indent = Cm(0)

    h3 = doc.styles['Heading 3']
    h3.font.name = 'Times New Roman'
    h3.font.size = Pt(12)
    h3.font.bold = True
    h3.font.italic = True
    h3.font.color.rgb = RGBColor(0, 0, 0)
    h3.paragraph_format.space_before = Pt(6)
    h3.paragraph_format.space_after  = Pt(3)
    h3.paragraph_format.first_line_indent = Cm(0)

# ─── Build TOC field XML ─────────────────────────────────────────────────────
def make_toc_xml():
    """Returneaza XML-ul pentru un camp TOC automat (Word il va popula)."""
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    toc_para = OxmlElement('w:p')

    # fldChar begin
    r1 = OxmlElement('w:r')
    fc1 = OxmlElement('w:fldChar')
    fc1.set(qn('w:fldCharType'), 'begin')
    fc1.set(qn('w:dirty'), 'true')
    r1.append(fc1)
    toc_para.append(r1)

    # instrText
    r2 = OxmlElement('w:r')
    it = OxmlElement('w:instrText')
    it.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    it.text = ' TOC \\o "1-3" \\h \\z \\u '
    r2.append(it)
    toc_para.append(r2)

    # fldChar end
    r3 = OxmlElement('w:r')
    fc3 = OxmlElement('w:fldChar')
    fc3.set(qn('w:fldCharType'), 'end')
    r3.append(fc3)
    toc_para.append(r3)

    return toc_para

# ─── Main ─────────────────────────────────────────────────────────────────────
def main():
    doc = Document(INPUT)
    setup_heading_styles(doc)

    paragraphs = doc.paragraphs
    body = doc.element.body

    # ── Pasul 1: aplica stilurile Heading ────────────────────────────────────
    cuprins_idx   = None
    introducere_idx = None
    applied = {"H1": 0, "H2": 0, "H3": 0}

    for i, para in enumerate(paragraphs):
        level = classify(para)
        if level is None:
            continue

        text = para.text.strip()
        if text == "CUPRINS":
            cuprins_idx = i
        if text == "INTRODUCERE" and introducere_idx is None:
            introducere_idx = i

        old_style = para.style.name
        if level == "H1":
            para.style = doc.styles['Heading 1']
        elif level == "H2":
            para.style = doc.styles['Heading 2']
        elif level == "H3":
            para.style = doc.styles['Heading 3']

        # Pastreaza fontul TNR si dimensiunea pe runs
        for run in para.runs:
            run.font.name = 'Times New Roman'
            if level == "H1":
                run.font.size = Pt(14)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
            elif level == "H2":
                run.font.size = Pt(13)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
            elif level == "H3":
                run.font.size = Pt(12)
                run.font.bold = True
                run.font.italic = True
                run.font.color.rgb = RGBColor(0, 0, 0)

        applied[level] += 1

    print(f"Stiluri aplicate: H1={applied['H1']}, H2={applied['H2']}, H3={applied['H3']}")
    print(f"CUPRINS la paragraful {cuprins_idx}, INTRODUCERE la {introducere_idx}")

    # ── Pasul 2: sterge continutul manual al cuprinsului (doturi) ────────────
    if cuprins_idx is not None and introducere_idx is not None:
        # Paragrafele dintre CUPRINS si INTRODUCERE = TOC manual de sters
        toc_paras = paragraphs[cuprins_idx + 1 : introducere_idx]
        print(f"Sterg {len(toc_paras)} paragrafe din cuprinsul manual...")

        for p in toc_paras:
            p_elem = p._element
            p_elem.getparent().remove(p_elem)

        # ── Pasul 3: insereaza campul TOC dupa paragraful CUPRINS ────────────
        cuprins_elem = paragraphs[cuprins_idx]._element
        toc_xml = make_toc_xml()

        # Insereaza imediat dupa CUPRINS
        cuprins_elem.addnext(toc_xml)
        print("Camp TOC inserat.")

    # ── Salveaza ─────────────────────────────────────────────────────────────
    doc.save(OUTPUT)
    print(f"\nSalvat: {OUTPUT}")
    print("\nIMPORTANT: Deschide fisierul in Word si:")
    print("  1. Apasa Ctrl+A (selecteaza tot)")
    print("  2. Apasa F9 (actualizeaza toate campurile)")
    print("  3. Alege 'Update entire table'")
    print("  4. Paginile din cuprins vor aparea automat!")

if __name__ == '__main__':
    main()
