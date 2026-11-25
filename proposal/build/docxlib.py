"""Small helpers on top of python-docx that apply the DEB SV UGM thesis format.

Format rules (Panduan Tugas Akhir 2025): A4, Times New Roman 12, 1.5 line spacing
(1.0 for tables, captions, references, lists of contents), margins top 4 / left 4 /
right 3 / bottom 3 cm, justified text, first paragraph under a heading flush left and
later paragraphs indented 1.25 cm, chapter titles centred bold capitals, roman page
numbers centred at the bottom for front matter and arabic numbers bottom-right for the
main body, table captions above (bold, centred), figure captions below.
"""
import re

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT = "Times New Roman"
TOKEN = re.compile(r"(\*\*.+?\*\*|\*.+?\*|_\{.+?\}|\^\{.+?\})")


def add_runs(par, text, size=None, bold=False, italic=False):
    """Add text with inline markup: **bold**, *italic*, _{sub}, ^{sup}."""
    for part in TOKEN.split(text):
        if not part:
            continue
        b, i, sub, sup = bold, italic, False, False
        if part.startswith("**"):
            part, b = part[2:-2], True
        elif part.startswith("*"):
            part, i = part[1:-1], not italic
        elif part.startswith("_{"):
            part, sub = part[2:-1], True
        elif part.startswith("^{"):
            part, sup = part[2:-1], True
        r = par.add_run(part)
        r.bold, r.italic = b, i
        r.font.subscript, r.font.superscript = sub, sup
        r.font.name = FONT
        if size:
            r.font.size = Pt(size)
    return par


def set_cell_shading(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def no_split(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def add_field(par, instr):
    r = par.add_run()
    f1 = OxmlElement("w:fldChar")
    f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instr
    f2 = OxmlElement("w:fldChar")
    f2.set(qn("w:fldCharType"), "end")
    r._r.append(f1)
    r._r.append(it)
    r._r.append(f2)
    r.font.name = FONT
    r.font.size = Pt(12)


def set_page_numbering(section, fmt, start=1):
    sectPr = section._sectPr
    pg = sectPr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType")
        sectPr.append(pg)
    pg.set(qn("w:fmt"), fmt)
    pg.set(qn("w:start"), str(start))


class Thesis:
    def __init__(self):
        self.doc = Document()
        self._setup_styles()
        s = self.doc.sections[0]
        self._margins(s)
        self.after_heading = True
        self.tables = []   # (number, title)
        self.figures = []  # (number, title)
        self.headings = []  # (level, text)

    # ------------------------------------------------------------------ setup
    @staticmethod
    def _margins(s):
        s.page_width, s.page_height = Cm(21.0), Cm(29.7)
        s.top_margin, s.left_margin = Cm(4), Cm(4)
        s.right_margin, s.bottom_margin = Cm(3), Cm(3)
        s.footer_distance = Cm(1.5)

    def _setup_styles(self):
        st = self.doc.styles
        normal = st["Normal"]
        normal.font.name = FONT
        normal.font.size = Pt(12)
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        pf = normal.paragraph_format
        pf.line_spacing = 1.5
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        for name, size, align in [("Heading 1", 12, WD_ALIGN_PARAGRAPH.CENTER),
                                  ("Heading 2", 12, WD_ALIGN_PARAGRAPH.LEFT),
                                  ("Heading 3", 12, WD_ALIGN_PARAGRAPH.LEFT)]:
            h = st[name]
            h.font.name = FONT
            h.font.size = Pt(size)
            h.font.bold = True
            h.font.italic = False
            h.font.color.rgb = RGBColor(0, 0, 0)
            rpr = h.element.get_or_add_rPr()
            rf = rpr.find(qn("w:rFonts"))
            if rf is None:
                rf = OxmlElement("w:rFonts")
                rpr.append(rf)
            for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                rf.set(qn(k), FONT)
            for k in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
                if rf.get(qn(k)) is not None:
                    del rf.attrib[qn(k)]
            p = h.paragraph_format
            p.alignment = align
            p.line_spacing = 1.5
            p.space_before = Pt(12) if name != "Heading 1" else Pt(0)
            p.space_after = Pt(6) if name != "Heading 1" else Pt(18)
            p.keep_with_next = True
            p.first_line_indent = Cm(0)
        st["Heading 3"].font.bold = True
        st["Heading 3"].font.italic = False

    # --------------------------------------------------------------- sections
    def new_section(self, page_fmt=None, start=1, footer_align=None):
        s = self.doc.add_section(WD_SECTION.NEW_PAGE)
        self._margins(s)
        s.footer.is_linked_to_previous = False
        for p in s.footer.paragraphs:
            p.text = ""
        if page_fmt:
            set_page_numbering(s, page_fmt, start)
            p = s.footer.paragraphs[0]
            p.alignment = footer_align
            add_field(p, "PAGE")
        return s

    def page_break(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ----------------------------------------------------------------- blocks
    def chapter(self, number_label, title):
        h = self.doc.add_paragraph(style="Heading 1")
        add_runs(h, number_label, bold=True)
        h.add_run().add_break()
        add_runs(h, title, bold=True)
        self.headings.append((1, f"{number_label} {title}"))
        self.after_heading = True

    def front_title(self, title, toc=True):
        h = self.doc.add_paragraph(style="Heading 1")
        add_runs(h, title, bold=True)
        if toc:
            self.headings.append((0, title))
        self.after_heading = True

    def h2(self, text):
        add_runs(self.doc.add_paragraph(style="Heading 2"), text, bold=True)
        self.headings.append((2, text))
        self.after_heading = True

    def h3(self, text):
        p = self.doc.add_paragraph(style="Heading 3")
        add_runs(p, text, bold=True)
        self.headings.append((3, text))
        self.after_heading = True

    def p(self, text, indent=None, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=0,
          keep_next=False):
        par = self.doc.add_paragraph()
        pf = par.paragraph_format
        pf.alignment = align
        if indent is None:
            pf.first_line_indent = Cm(0) if self.after_heading else Cm(1.25)
        else:
            pf.first_line_indent = Cm(indent)
        pf.space_after = Pt(space_after)
        pf.keep_with_next = keep_next
        add_runs(par, text)
        self.after_heading = False
        return par

    def items(self, entries, style="num", left=1.0, hang=0.6):
        """Numbered (1., 2.), lettered (a., b.) or bracketed ((1), (2)) list."""
        for i, text in enumerate(entries):
            if style == "num":
                label = f"{i + 1}."
            elif style == "alpha":
                label = f"{'abcdefghijklmnopqrstuvwxyz'[i]}."
            elif style == "paren":
                label = f"({i + 1})"
            else:
                label = "•"
            par = self.doc.add_paragraph()
            pf = par.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            pf.left_indent = Cm(left)
            pf.first_line_indent = Cm(-hang)
            pf.tab_stops.add_tab_stop(Cm(left))
            add_runs(par, label + "\t" + text)
        self.after_heading = False

    def equation(self, text, number):
        par = self.doc.add_paragraph()
        pf = par.paragraph_format
        pf.tab_stops.add_tab_stop(Cm(7.0), WD_TAB_ALIGNMENT.CENTER)
        pf.tab_stops.add_tab_stop(Cm(14.0), WD_TAB_ALIGNMENT.RIGHT)
        pf.space_before = Pt(4)
        pf.space_after = Pt(4)
        pf.first_line_indent = Cm(0)
        add_runs(par, "\t" + text + f"\t({number})")
        self.after_heading = False

    def caption_table(self, number, title):
        par = self.doc.add_paragraph()
        pf = par.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.line_spacing = 1.0
        pf.space_before = Pt(10)
        pf.space_after = Pt(6)
        pf.keep_with_next = True
        add_runs(par, f"Table {number}. {title}", bold=True)
        self.tables.append((number, title))

    def source(self, text):
        par = self.doc.add_paragraph()
        pf = par.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.line_spacing = 1.0
        pf.space_before = Pt(3)
        pf.space_after = Pt(10)
        add_runs(par, text, size=10, italic=True)
        self.after_heading = True

    def table(self, header, rows, widths, size=10, bold_first_col=False, align_center_cols=()):
        t = self.doc.add_table(rows=1, cols=len(header))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        hdr = t.rows[0]
        repeat_header(hdr)
        for j, text in enumerate(header):
            c = hdr.cells[j]
            set_cell_shading(c, "D9E2F3")
            par = c.paragraphs[0]
            par.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            par.paragraph_format.line_spacing = 1.0
            add_runs(par, text, size=size, bold=True)
        for row in rows:
            r = t.add_row()
            no_split(r)
            for j, text in enumerate(row):
                c = r.cells[j]
                lines = text.split("\n")
                for k, line in enumerate(lines):
                    par = c.paragraphs[0] if k == 0 else c.add_paragraph()
                    par.paragraph_format.line_spacing = 1.0
                    par.paragraph_format.space_after = Pt(1)
                    par.paragraph_format.first_line_indent = Cm(0)
                    par.paragraph_format.alignment = (
                        WD_ALIGN_PARAGRAPH.CENTER if j in align_center_cols
                        else WD_ALIGN_PARAGRAPH.LEFT)
                    add_runs(par, line, size=size, bold=(bold_first_col and j == 0))
        for row in t.rows:
            for j, w in enumerate(widths):
                row.cells[j].width = Cm(w)
        tblPr = t._tbl.tblPr
        for tag in ("w:tblW", "w:tblLayout"):
            old = tblPr.find(qn(tag))
            if old is not None:
                tblPr.remove(old)
        tw = OxmlElement("w:tblW")
        tw.set(qn("w:w"), str(int(sum(widths) * 567)))
        tw.set(qn("w:type"), "dxa")
        tblPr.append(tw)
        lay = OxmlElement("w:tblLayout")
        lay.set(qn("w:type"), "fixed")
        tblPr.append(lay)
        grid = t._tbl.tblGrid
        for j, gc in enumerate(grid.findall(qn("w:gridCol"))):
            gc.set(qn("w:w"), str(int(widths[j] * 567)))
        if len(rows) <= 10:  # keep short tables together with their source line
            for row in t.rows:
                for c in row.cells:
                    for par in c.paragraphs:
                        par.paragraph_format.keep_with_next = True
        self.after_heading = True
        return t

    def figure(self, path, number, title, width_cm=14.0, source=None):
        par = self.doc.add_paragraph()
        par.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.first_line_indent = Cm(0)
        par.paragraph_format.space_before = Pt(8)
        par.paragraph_format.keep_with_next = True
        par.paragraph_format.line_spacing = 1.0
        par.add_run().add_picture(str(path), width=Cm(width_cm))
        cap = self.doc.add_paragraph()
        cp = cap.paragraph_format
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.line_spacing = 1.0
        cp.space_before = Pt(4)
        cp.space_after = Pt(2 if source else 10)
        add_runs(cap, f"Figure {number}. {title}", bold=True)
        self.figures.append((number, title))
        if source:
            s = self.doc.add_paragraph()
            s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            s.paragraph_format.line_spacing = 1.0
            s.paragraph_format.space_after = Pt(10)
            add_runs(s, source, size=10, italic=True)
        self.after_heading = True

    def toc_line(self, text, page, level=0, size=12):
        par = self.doc.add_paragraph()
        pf = par.paragraph_format
        pf.line_spacing = 1.0
        pf.space_after = Pt(4)
        pf.left_indent = Cm(0.6 * level)
        pf.first_line_indent = Cm(0)
        pf.tab_stops.add_tab_stop(Cm(14.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        add_runs(par, f"{text}\t{page}", size=size)

    def reference(self, text):
        par = self.doc.add_paragraph()
        pf = par.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.line_spacing = 1.0
        pf.space_after = Pt(6)
        pf.left_indent = Cm(1.0)
        pf.first_line_indent = Cm(-1.0)
        add_runs(par, text)

    def centered(self, text, size=12, bold=False, italic=False, space_after=0, space_before=0):
        par = self.doc.add_paragraph()
        pf = par.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.first_line_indent = Cm(0)
        pf.line_spacing = 1.15
        pf.space_after = Pt(space_after)
        pf.space_before = Pt(space_before)
        add_runs(par, text, size=size, bold=bold, italic=italic)
        return par

    def save(self, path):
        self.doc.save(path)
