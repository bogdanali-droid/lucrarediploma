#!/usr/bin/env python3
"""Generează teza.docx din fișierele markdown, format UPSC Chișinău."""

import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = "/home/user/lucrarediploma/teza-master"

CHAPTERS = [
    (BASE + "/introducere.md",  False),
    (BASE + "/capitolul_1.md",  True),
    (BASE + "/capitolul_2.md",  True),
    (BASE + "/capitolul_3.md",  True),
    (BASE + "/concluzie.md",    False),
    (BASE + "/bibliografie.md", False),
]

def set_margins(doc, top_cm=1.5, bottom_cm=2.5, left_cm=2.5, right_cm=1.5):
    for section in doc.sections:
        section.top_margin    = Cm(top_cm)
        section.bottom_margin = Cm(bottom_cm)
        section.left_margin   = Cm(left_cm)
        section.right_margin  = Cm(right_cm)

def add_page_numbers(doc):
    """Adaugă numere de pagină jos-centru."""
    for section in doc.sections:
        footer = section.footer
        para = footer.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.text = 'PAGE'
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)

def fmt_run(run, bold=False, italic=False):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run.bold = bold
    run.italic = italic

def set_para_fmt(para, first_indent=True, space_before=0, space_after=0):
    pf = para.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after  = Pt(space_after)
    pf.line_spacing = Pt(18)  # 1.5 × 12pt
    if first_indent:
        pf.first_line_indent = Cm(1.25)
    else:
        pf.first_line_indent = Cm(0)

def add_inline(para, text):
    """Adaugă text cu bold/italic inline (suportă **bold**, *italic*, ***bold-italic***)."""
    # split on bold-italic, bold, italic markers
    pattern = re.compile(r'(\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*.*?\*)')
    parts = pattern.split(text)
    for part in parts:
        if part.startswith('***') and part.endswith('***'):
            r = para.add_run(part[3:-3])
            fmt_run(r, bold=True, italic=True)
        elif part.startswith('**') and part.endswith('**'):
            r = para.add_run(part[2:-2])
            fmt_run(r, bold=True)
        elif part.startswith('*') and part.endswith('*') and len(part) > 2:
            r = para.add_run(part[1:-1])
            fmt_run(r, italic=True)
        else:
            r = para.add_run(part)
            fmt_run(r)

def process_file(doc, filepath, page_break_before=False):
    with open(filepath, encoding='utf-8') as f:
        lines = f.readlines()

    if page_break_before:
        doc.add_page_break()

    skip_meta = False
    i = 0
    while i < len(lines):
        line = lines[i].rstrip('\n')

        # Skip metadata/comment blocks
        if line.startswith('> '):
            i += 1
            continue
        if line.strip() == '---':
            i += 1
            continue

        # H1
        if line.startswith('# ') and not line.startswith('## '):
            title = line[2:].strip()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after  = Pt(12)
            pf.first_line_indent = Cm(0)
            r = p.add_run(title.upper())
            r.font.name = 'Times New Roman'
            r.font.size = Pt(14)
            r.bold = True
            i += 1
            continue

        # H2
        if line.startswith('## ') and not line.startswith('### '):
            title = line[3:].strip()
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(12)
            pf.space_after  = Pt(6)
            pf.first_line_indent = Cm(0)
            r = p.add_run(title)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(13)
            r.bold = True
            i += 1
            continue

        # H3
        if line.startswith('### '):
            title = line[4:].strip()
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(6)
            pf.space_after  = Pt(3)
            pf.first_line_indent = Cm(0)
            r = p.add_run(title)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(12)
            r.bold = True
            r.italic = True
            i += 1
            continue

        # H4 (#### or *Activitatea*)
        if line.startswith('#### '):
            title = line[5:].strip()
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(4)
            pf.space_after  = Pt(2)
            pf.first_line_indent = Cm(0)
            r = p.add_run(title)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(12)
            r.bold = True
            i += 1
            continue

        # FOTO marker
        if line.strip().startswith('[FOTO'):
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(6)
            pf.space_after  = Pt(6)
            pf.first_line_indent = Cm(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(line.strip())
            r.font.name = 'Times New Roman'
            r.font.size = Pt(10)
            r.italic = True
            r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
            i += 1
            continue

        # Numbered list (bibliography: "1. Bruner...")
        m = re.match(r'^(\d+)\.\s+(.*)', line)
        if m:
            num = m.group(1)
            content = m.group(2)
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after  = Pt(3)
            pf.left_indent  = Cm(1.0)
            pf.first_line_indent = Cm(-1.0)
            pf.line_spacing = Pt(18)
            r = p.add_run(num + '. ')
            fmt_run(r)
            add_inline(p, content)
            i += 1
            continue

        # Bullet list
        if line.startswith('- ') or line.startswith('* '):
            content = line[2:].strip()
            p = doc.add_paragraph(style='List Bullet')
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after  = Pt(2)
            pf.line_spacing = Pt(18)
            pf.left_indent  = Cm(1.0)
            pf.first_line_indent = Cm(0)
            add_inline(p, content)
            i += 1
            continue

        # Numbered list (objectives: "1. ...")
        m2 = re.match(r'^(\d+)\.\s+(.*)', line)
        if m2:
            p = doc.add_paragraph(style='List Number')
            set_para_fmt(p, first_indent=False)
            add_inline(p, m2.group(2))
            i += 1
            continue

        # Horizontal rule becomes nothing
        if line.strip() == '---':
            i += 1
            continue

        # Quote/note line
        if line.startswith('>'):
            i += 1
            continue

        # Empty line
        if line.strip() == '':
            i += 1
            continue

        # Normal paragraph
        p = doc.add_paragraph()
        set_para_fmt(p, first_indent=True)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        add_inline(p, line.strip())
        i += 1

def main():
    doc = Document()

    # Default style
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    set_margins(doc, top_cm=1.5, bottom_cm=2.5, left_cm=2.5, right_cm=1.5)
    add_page_numbers(doc)

    first = True
    for filepath, pb in CHAPTERS:
        process_file(doc, filepath, page_break_before=(not first))
        first = False

    out = "/home/user/lucrarediploma/teza_anca_aliciuc.docx"
    doc.save(out)
    print(f"Salvat: {out}")

if __name__ == '__main__':
    main()
