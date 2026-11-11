#!/usr/bin/env python3
"""Structured HTML -> editable legal-style DOCX example.

Usage:
    python converter.py examples/legal-agreement.html output.docx

Requirements: beautifulsoup4, python-docx

Demonstrated features: real multilevel Word numbering for headings and legal
paragraphs; real Word heading styles; nested (a)/(i) lists; updateable TOC
field; simple editable tables; PAGE field; footer, page layout.

This intentionally handles a known, controlled HTML schema. Production
conversion requires adapting to the client's HTML and Word template.
"""
from __future__ import annotations
import sys
import re
from pathlib import Path
from bs4 import BeautifulSoup, Tag, NavigableString
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

BLACK = RGBColor(16, 16, 16)
FONT = 'Times New Roman'


def add_num(doc, levels, *, alias=None):
    """Create a real Word multilevel numbering definition, returning its numId."""
    root = doc.part.numbering_part.element
    aids = [int(e.get(qn('w:abstractNumId'))) for e in root.findall(qn('w:abstractNum'))]
    nids = [int(e.get(qn('w:numId'))) for e in root.findall(qn('w:num'))]
    aid = max(aids, default=-1) + 1
    nid = max(nids, default=0) + 1
    abstract = OxmlElement('w:abstractNum')
    abstract.set(qn('w:abstractNumId'), str(aid))
    multi = OxmlElement('w:multiLevelType')
    multi.set(qn('w:val'), 'multilevel' if len(levels) > 1 else 'singleLevel')
    abstract.append(multi)
    if alias:
        name = OxmlElement('w:name')
        name.set(qn('w:val'), alias)
        abstract.append(name)
    for i, cfg in enumerate(levels):
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(i))
        for tag, val in [('start', '1'), ('numFmt', cfg['format']),
                         ('lvlText', cfg['text']), ('suff', 'space')]:
            el = OxmlElement(f'w:{tag}')
            el.set(qn('w:val'), val)
            lvl.append(el)
        if i:
            restart = OxmlElement('w:lvlRestart')
            restart.set(qn('w:val'), '1')
            lvl.append(restart)
        ppr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), str(cfg.get('left', 720 + i * 420)))
        ind.set(qn('w:hanging'), str(cfg.get('hanging', 360)))
        ppr.append(ind)
        lvl.append(ppr)
        rpr = OxmlElement('w:rPr')
        rf = OxmlElement('w:rFonts')
        rf.set(qn('w:ascii'), FONT)
        rf.set(qn('w:hAnsi'), FONT)
        rpr.append(rf)
        lvl.append(rpr)
        abstract.append(lvl)
    root.append(abstract)
    num = OxmlElement('w:num')
    num.set(qn('w:numId'), str(nid))
    aid_el = OxmlElement('w:abstractNumId')
    aid_el.set(qn('w:val'), str(aid))
    num.append(aid_el)
    root.append(num)
    return nid


def apply_num(p, nid, level):
    ppr = p._p.get_or_add_pPr()
    numpr = ppr.find(qn('w:numPr'))
    if numpr is None:
        numpr = OxmlElement('w:numPr')
        ppr.append(numpr)
    for tag, val in [('ilvl', level), ('numId', nid)]:
        node = numpr.find(qn(f'w:{tag}'))
        if node is None:
            node = OxmlElement(f'w:{tag}')
            numpr.append(node)
        node.set(qn('w:val'), str(val))


def inline(p, el, bold=False, italic=False):
    for node in el.children:
        if isinstance(node, NavigableString):
            val = str(node).replace('\xa0', ' ')
            if val.strip() or (p.runs and val):
                r = p.add_run(val)
                r.bold = bold
                r.italic = italic
        elif isinstance(node, Tag):
            if node.name in ['strong', 'b']:
                inline(p, node, True, italic)
            elif node.name in ['em', 'i']:
                inline(p, node, bold, True)
            elif node.name == 'br':
                p.add_run().add_break()
            elif node.name not in ['ol', 'ul']:
                inline(p, node, bold, italic)


def add_field(p, instruction, placeholder='1'):
    r = p.add_run()
    for kind in ['begin', 'instruction', 'separate', 'display', 'end']:
        if kind in ['begin', 'separate', 'end']:
            el = OxmlElement('w:fldChar')
            el.set(qn('w:fldCharType'), kind)
        elif kind == 'instruction':
            el = OxmlElement('w:instrText')
            el.set(qn('xml:space'), 'preserve')
            el.text = instruction
        else:
            el = OxmlElement('w:t')
            el.text = placeholder
        r._r.append(el)


def styles(doc):
    sec = doc.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    sec.top_margin = Inches(0.9)
    sec.bottom_margin = Inches(0.82)
    sec.left_margin = Inches(1.05)
    sec.right_margin = Inches(1.0)
    sec.header_distance = Inches(0.34)
    sec.footer_distance = Inches(0.35)
    s = doc.styles
    normal = s['Normal']
    normal.font.name = FONT
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = BLACK
    normal.paragraph_format.line_spacing = 1.16
    normal.paragraph_format.space_after = Pt(7)
    if 'Legal Title' not in s:
        s.add_style('Legal Title', WD_STYLE_TYPE.PARAGRAPH)
    for name, size, before, after in [
        ('Legal Title', 16, 6, 10),
        ('Heading 1', 12, 18, 7),
        ('Heading 2', 10.5, 11, 4),
    ]:
        st = s[name]
        st.font.name = FONT
        rf = st._element.get_or_add_rPr().rFonts
        if rf is not None:
            for attr in ('asciiTheme', 'hAnsiTheme', 'csTheme', 'eastAsiaTheme'):
                if rf.get(qn('w:' + attr)) is not None:
                    del rf.attrib[qn('w:' + attr)]
            for attr in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
                rf.set(qn('w:' + attr), FONT)
        st.font.size = Pt(size)
        st.font.color.rgb = BLACK
        st.font.bold = True
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.line_spacing = 1.05
    for name in ['Sample Kicker', 'Sample Note', 'Legal List',
                 'Legal Small', 'Sample TOC Title']:
        if name not in s:
            s.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        s[name].font.name = FONT
        s[name].font.color.rgb = BLACK
    kicker = s['Sample Kicker']
    kicker.font.size = Pt(8)
    kicker.font.bold = True
    kicker.paragraph_format.space_after = Pt(2)
    note = s['Sample Note']
    note.font.size = Pt(8.7)
    note.paragraph_format.space_after = Pt(7)
    legal = s['Legal List']
    legal.font.size = Pt(10.5)
    legal.paragraph_format.space_after = Pt(5)
    legal.paragraph_format.line_spacing = 1.14
    small = s['Legal Small']
    small.font.size = Pt(9)
    small.paragraph_format.space_after = Pt(6)
    toc_title = s['Sample TOC Title']
    toc_title.font.size = Pt(11)
    toc_title.font.bold = True
    toc_title.paragraph_format.space_before = Pt(12)
    toc_title.paragraph_format.space_after = Pt(9)
    for name in ['TOC 1', 'TOC 2']:
        if name not in s:
            s.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st = s[name]
        st.font.name = FONT
        st.font.size = Pt(10)
        st.font.color.rgb = BLACK
        st.paragraph_format.space_after = Pt(4)
    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    rr = hp.add_run('ILLUSTRATIVE DOCUMENT  |  HTML-TO-WORD DEMO')
    rr.font.name = FONT
    rr.font.size = Pt(7.3)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run('FICTIONAL SAMPLE — NOT FOR LEGAL USE     |     Page ')
    r.font.size = Pt(8)
    add_field(fp, ' PAGE ', '1')


def inline_as_nested(p, node, bold=False, italic=False):
    for c in node.children:
        if isinstance(c, NavigableString):
            r = p.add_run(re.sub(r'\s+', ' ', str(c)))
            r.bold = bold
            r.italic = italic
        elif isinstance(c, Tag):
            inline_as_nested(p, c, bold or c.name in ['b', 'strong'],
                             italic or c.name in ['em', 'i'])


def do_list(doc, ol, nid=None, level=0):
    if nid is None:
        nid = add_num(doc, [
            {'format': 'lowerLetter', 'text': '(%1)', 'left': 780, 'hanging': 340},
            {'format': 'lowerRoman', 'text': '(%2)', 'left': 1140, 'hanging': 340},
        ], alias='Legal clause list')
    for li in ol.find_all('li', recursive=False):
        p = doc.add_paragraph(style='Legal List')
        apply_num(p, nid, level)
        for c in li.contents:
            if isinstance(c, NavigableString):
                t = re.sub(r'\s+', ' ', str(c))
                if t.strip():
                    p.add_run(t)
            elif isinstance(c, Tag) and c.name not in ('ol', 'ul'):
                inline_as_nested(p, c, bold=c.name in ['b', 'strong'],
                                 italic=c.name in ['em', 'i'])
        for nested in li.find_all('ol', recursive=False):
            do_list(doc, nested, nid, level + 1)


def add_table(doc, el):
    rows = el.select('tr')
    ncols = len(rows[0].find_all(['th', 'td'], recursive=False))
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    sec = doc.sections[0]
    width = sec.page_width - sec.left_margin - sec.right_margin
    widths = [width * 0.29, width * 0.28, width * 0.43]
    for r_idx, html_row in enumerate(rows):
        cells = html_row.find_all(['th', 'td'], recursive=False)
        for c_idx, cell_node in enumerate(cells):
            cell = table.cell(r_idx, c_idx)
            cell.width = widths[c_idx]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.style = 'Normal'
            p.paragraph_format.space_after = Pt(2)
            inline(p, cell_node)
            for run in p.runs:
                run.font.size = Pt(9)
                if r_idx == 0:
                    run.bold = True
            pr = cell._tc.get_or_add_tcPr()
            b = OxmlElement('w:tcBorders')
            for edge in ['top', 'left', 'bottom', 'right']:
                e = OxmlElement(f'w:{edge}')
                e.set(qn('w:val'), 'single')
                e.set(qn('w:sz'), '5')
                e.set(qn('w:color'), '444444')
                b.append(e)
            pr.append(b)
    rp = table.rows[0]._tr.get_or_add_trPr()
    rep = OxmlElement('w:tblHeader')
    rep.set(qn('w:val'), 'true')
    rp.append(rep)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def convert(src, dst):
    soup = BeautifulSoup(Path(src).read_text(encoding='utf-8'), 'html.parser')
    doc = Document()
    styles(doc)
    hdr = soup.select_one('body > header')
    for el in hdr.find_all(recursive=False):
        if el.name == 'h1':
            p = doc.add_paragraph(style='Legal Title')
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif 'kicker' in el.get('class', []):
            p = doc.add_paragraph(style='Sample Kicker')
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif 'subtitle' in el.get('class', []):
            p = doc.add_paragraph(style='Sample Note')
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif 'disclaimer' in el.get('class', []):
            p = doc.add_paragraph(style='Sample Note')
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p = doc.add_paragraph(style='Normal')
        inline(p, el)
    p = doc.add_paragraph('TABLE OF CONTENTS', style='Sample TOC Title')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fieldp = doc.add_paragraph()
    add_field(fieldp, ' TOC \\o "1-2" \\h \\z \\u ',
              'Right-click and update field (or Ctrl+A, F9)')
    doc.add_page_break()
    clause_nid = add_num(doc, [
        {'format': 'decimal', 'text': '%1.', 'left': 440, 'hanging': 320},
        {'format': 'decimal', 'text': '%1.%2', 'left': 600, 'hanging': 420},
    ], alias='Agreement article and section numbers')
    for sec in soup.select('main > section'):
        for el in sec.find_all(recursive=False):
            if el.name == 'h2':
                p = doc.add_paragraph(style='Heading 1')
                apply_num(p, clause_nid, 0)
                inline(p, el)
            elif el.name == 'h3':
                p = doc.add_paragraph(style='Heading 2')
                apply_num(p, clause_nid, 1)
                inline(p, el)
            elif el.name == 'p':
                p = doc.add_paragraph(
                    style='Legal Small' if 'small' in el.get('class', []) else 'Normal'
                )
                inline(p, el)
            elif el.name == 'ol':
                do_list(doc, el)
            elif el.name == 'table':
                add_table(doc, el)
    settings = doc.settings._element
    upd = settings.find(qn('w:updateFields'))
    if upd is None:
        upd = OxmlElement('w:updateFields')
        settings.append(upd)
    upd.set(qn('w:val'), 'true')
    core = doc.core_properties
    core.title = 'Illustrative Professional Services Agreement — HTML to Word Sample'
    core.subject = 'Automatic section and paragraph numbering, editable DOCX, TOC'
    core.author = 'Novel Hari Keswono | Portfolio Demonstration'
    core.keywords = 'HTML, DOCX, legal-style formatting, multilevel numbering, TOC'
    doc.save(dst)
    print(f'Created {dst}')


if __name__ == '__main__':
    if len(sys.argv) != 3:
        raise SystemExit('Usage: python converter.py source.html output.docx')
    convert(sys.argv[1], sys.argv[2])
