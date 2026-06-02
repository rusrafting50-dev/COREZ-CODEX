# docx_utils.py — HTML → DOCX export via python-docx + htmldocx (no external binaries)
import re
from io import BytesIO
from bs4 import BeautifulSoup, NavigableString
from docx import Document
from docx.shared import Pt, Cm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from htmldocx import HtmlToDocx
from flask import send_file

DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

# Usable widths (page − margins)
_PAGE_WIDTH_LANDSCAPE_CM = 27.2   # 29.7 − 1.5 − 1.0
_PAGE_WIDTH_PORTRAIT_CM  = 18.5   # 21.0 − 1.5 − 1.0

_BLOCK = frozenset({
    'div','p','h1','h2','h3','h4','h5','h6',
    'ul','ol','li','table','thead','tbody','tfoot','tr','th','td',
    'blockquote','section','article','header','footer','nav','figure',
})


# ─── Document setup ──────────────────────────────────────────────────────────

def _make_base_doc(portrait=False):
    doc = Document()
    sec = doc.sections[0]
    if portrait:
        sec.page_width    = Cm(21.0)
        sec.page_height   = Cm(29.7)
        sec.left_margin   = Cm(1.5)
        sec.right_margin  = Cm(1.0)
        sec.top_margin    = Cm(1.2)
        sec.bottom_margin = Cm(1.5)
    else:
        sec.page_width    = Cm(29.7)
        sec.page_height   = Cm(21.0)
        sec.left_margin   = Cm(1.5)
        sec.right_margin  = Cm(1.0)
        sec.top_margin    = Cm(1.5)
        sec.bottom_margin = Cm(1.5)
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(10)
    return doc


# ─── CSS helpers ─────────────────────────────────────────────────────────────

def _css(tag, prop):
    """Return value of CSS property from style attr, or ''."""
    m = re.search(r'(?:^|;)\s*' + re.escape(prop) + r'\s*:\s*([^;]+)', tag.get('style', ''))
    return m.group(1).strip() if m else ''


def _px(tag):
    """Return explicit width in px from style, or None."""
    for prop in ('width', 'min-width'):
        m = re.search(r'(?:^|;)\s*' + re.escape(prop) + r'\s*:\s*(\d+(?:\.\d+)?)\s*px', tag.get('style', ''))
        if m:
            return float(m.group(1))
    return None


# ─── Cell XML helpers ────────────────────────────────────────────────────────

def _set_cell_valign(cell, align):
    """Set vertical alignment of a table cell: 'top', 'center', 'bottom'."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = tcPr.find(qn('w:vAlign'))
    if vAlign is None:
        vAlign = OxmlElement('w:vAlign')
        tcPr.append(vAlign)
    vAlign.set(qn('w:val'), align)


def _set_cell_bottom_border(cell, size=6, color='000000'):
    """Draw a single bottom border on a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    existing = tcBorders.find(qn('w:bottom'))
    if existing is not None:
        tcBorders.remove(existing)
    border = OxmlElement('w:bottom')
    border.set(qn('w:val'), 'single')
    border.set(qn('w:sz'), str(size))
    border.set(qn('w:color'), color)
    tcBorders.append(border)


def _set_cell_width(cell, twips_val):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcW')):
        tcPr.remove(old)
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(twips_val)))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)


# ─── Pre-processing: signature blocks ────────────────────────────────────────

def _process_sig_blocks(soup):
    """
    Convert each .sig-block into a borderless 2-column table so that
    the FIO column starts at the same horizontal position on every row.
    No underline between position and FIO.
    """
    for sig_block in soup.find_all(class_='sig-block'):
        sig_rows = sig_block.find_all(class_='sig-row')
        rows_html = ''
        for sig_row in sig_rows:
            pos_el = sig_row.find(class_='sig-pos')
            fio_el = sig_row.find(class_='sig-fio')
            if fio_el is None:
                kids = [c for c in sig_row.children
                        if getattr(c, 'name', None) and c is not pos_el]
                fio_el = kids[0] if kids else None
            pos_text = pos_el.get_text(strip=True) if pos_el else ''
            fio_text = fio_el.get_text(strip=True) if fio_el else ''
            rows_html += (
                '<tr>'
                f'<td style="border:none;width:260px"><b>{pos_text}</b></td>'
                f'<td style="border:none;width:440px">{fio_text}</td>'
                '</tr>'
            )
        if not rows_html:
            sig_block.decompose()
            continue
        tbl_html = (
            '<table style="width:100%;border:none;margin-top:12pt">'
            + rows_html +
            '</table>'
        )
        tbl_tag = BeautifulSoup(tbl_html, 'html.parser').find('table')
        sig_block.replace_with(tbl_tag)

    # Fallback: orphaned sig-rows not inside a sig-block
    for sig_row in soup.find_all(class_='sig-row'):
        pos_el = sig_row.find(class_='sig-pos')
        fio_el = sig_row.find(class_='sig-fio')
        if fio_el is None:
            kids = [c for c in sig_row.children
                    if getattr(c, 'name', None) and c is not pos_el]
            fio_el = kids[0] if kids else None
        pos_text = pos_el.get_text(strip=True) if pos_el else ''
        fio_text = fio_el.get_text(strip=True) if fio_el else ''
        p = soup.new_tag('p')
        if pos_text:
            b = soup.new_tag('b')
            b.string = pos_text
            p.append(b)
        p.append(NavigableString('  ' + fio_text))
        sig_row.replace_with(p)


# ─── Pre-processing: flex space-between → date left / venue right ─────────────

def _process_flex_between(soup):
    """
    Convert d-flex justify-content-between containers (date/venue row) to a
    3-column, 2-row borderless table:
      Row 1: [date value, bottom-border, centered] | [empty] | [place value, bottom-border, centered]
      Row 2: [«Дата проведения», top, centered]     | [empty] | [«Место проведения», top, centered]
    """
    def _is_flex_between(tag):
        if tag.name != 'div':
            return False
        classes = tag.get('class') or []
        style   = tag.get('style', '')
        return (('justify-content-between' in classes) or
                ('justify-content:space-between' in style))

    def _extract_value_label(child_div):
        inners = [c for c in child_div.children if getattr(c, 'name', None)]
        if len(inners) >= 2:
            return inners[0].get_text(' ', strip=True), inners[1].get_text(' ', strip=True)
        return child_div.get_text(' ', strip=True), ''

    for div in soup.find_all(_is_flex_between):
        block_children = [c for c in div.children if getattr(c, 'name', None)]
        if len(block_children) < 2:
            div.unwrap()
            continue

        left_value, left_label   = _extract_value_label(block_children[0])
        right_value, right_label = _extract_value_label(block_children[-1])

        if left_label or right_label:
            # Structured 2-row, 3-column table with value row and label row
            # widths in px: 280+140+280 = 700 (proportional to portrait page)
            tbl_html = (
                '<table class="date-place-table" style="width:100%;border:none">'
                '<tr>'
                f'<td style="border:none;text-align:center;width:280px">{left_value}</td>'
                '<td style="border:none;width:140px"></td>'
                f'<td style="border:none;text-align:center;width:280px">{right_value}</td>'
                '</tr>'
                '<tr>'
                f'<td style="border:none;text-align:center;font-size:8pt">{left_label}</td>'
                '<td style="border:none"></td>'
                f'<td style="border:none;text-align:center;font-size:8pt">{right_label}</td>'
                '</tr>'
                '</table>'
            )
        else:
            tbl_html = (
                '<table style="width:100%;border:none">'
                '<tr>'
                f'<td style="border:none;text-align:left">{left_value}</td>'
                f'<td style="border:none;text-align:right">{right_value}</td>'
                '</tr>'
                '</table>'
            )

        tbl_tag = BeautifulSoup(tbl_html, 'html.parser').find('table')
        div.replace_with(tbl_tag)


# ─── Pre-processing: div → p ─────────────────────────────────────────────────

def _apply_bold(tag, soup):
    fw = _css(tag, 'font-weight')
    if fw in ('bold', '700', '800', '900'):
        b = soup.new_tag('b')
        for child in list(tag.children):
            child.extract()
            b.append(child)
        tag.append(b)


def _pre_line_to_br(tag, soup):
    if _css(tag, 'white-space') == 'pre-line':
        for tn in list(tag.find_all(string=True)):
            if '\n' not in str(tn):
                continue
            parts = str(tn).split('\n')
            parent = tn.parent
            pos = list(parent.children).index(tn)
            tn.extract()
            fragment = []
            for i, part in enumerate(parts):
                fragment.append(NavigableString(part))
                if i < len(parts) - 1:
                    fragment.append(soup.new_tag('br'))
            for j, node in enumerate(fragment):
                parent.insert(pos + j, node)


def _divs_to_paragraphs(soup):
    for div in reversed(soup.find_all('div')):
        # Promote text-center class to inline style so HtmlToDocx picks it up
        if 'text-center' in (div.get('class') or []):
            style = div.get('style', '').rstrip(';')
            if 'text-align' not in style:
                div['style'] = (style + ';text-align:center').lstrip(';')
        _apply_bold(div, soup)
        _pre_line_to_br(div, soup)
        has_block = any(getattr(c, 'name', None) in _BLOCK for c in div.children)
        if has_block:
            div.unwrap()
        else:
            div.name = 'p'
    for p in soup.find_all('p'):
        _apply_bold(p, soup)


# ─── Pre-processing: normalize tables (expand colspan, strip rowspan) ─────────

def _normalize_tables(soup):
    for table in soup.find_all('table'):
        for tr in table.find_all('tr'):
            # Preserve colspan for section-row rows — HtmlToDocx merges them natively
            if 'section-row' in (tr.get('class') or []):
                continue
            for cell in list(tr.find_all(['td', 'th'])):
                colspan = int(cell.get('colspan', 1))
                if colspan > 1:
                    for _ in range(colspan - 1):
                        extra = soup.new_tag(cell.name)
                        cell.insert_after(extra)
                cell.attrs.pop('colspan', None)
                cell.attrs.pop('rowspan', None)
    return soup


# ─── Width extraction (after normalize) ──────────────────────────────────────

def _extract_table_widths(soup):
    """For each table return list of (px_width | None) per column from header row."""
    result = []
    for table in soup.find_all('table'):
        first_row = table.find('tr')
        if not first_row:
            result.append([])
            continue
        result.append([_px(c) for c in first_row.find_all(['th', 'td'])])
    return result


# ─── Post-processing: apply column widths to DOCX ────────────────────────────

def _apply_table_widths(doc, table_widths, page_width_cm=None):
    """Scale extracted px widths to the page and apply them to DOCX tables."""
    if page_width_cm is None:
        page_width_cm = _PAGE_WIDTH_LANDSCAPE_CM
    CM_TO_TWIPS = 566.929
    total_twips = int(page_width_cm * CM_TO_TWIPS)

    for t_idx, table in enumerate(doc.tables):
        if t_idx >= len(table_widths):
            break
        widths_px = table_widths[t_idx]
        if not widths_px:
            continue

        known = [(i, w) for i, w in enumerate(widths_px) if w is not None]
        if not known:
            continue

        total_known_px = sum(w for _, w in known)
        n_unknown      = len(widths_px) - len(known)
        reserved_for_unknown = (total_twips * 0.10) if n_unknown else 0
        available = total_twips - reserved_for_unknown
        scale = available / total_known_px

        default_unknown = (reserved_for_unknown / n_unknown) if n_unknown else 0

        col_twips = [
            int(w * scale) if w is not None else int(default_unknown)
            for w in widths_px
        ]

        for row in table.rows:
            for c_idx, cell in enumerate(row.cells):
                if c_idx < len(col_twips):
                    _set_cell_width(cell, col_twips[c_idx])


# ─── Post-processing: date/place table borders and alignment ─────────────────

def _post_process_date_place_tables(doc):
    """
    Find date/place tables (3 columns, 2 rows) and apply:
      Row 0: value cells → bottom-aligned, bottom border, centered
      Row 1: label cells → top-aligned, centered
    """
    for table in doc.tables:
        try:
            n_rows = len(table.rows)
            n_cols = len(table.columns)
        except Exception:
            continue
        if n_rows != 2 or n_cols != 3:
            continue
        # Value row
        for col_idx in [0, 2]:
            try:
                cell = table.rows[0].cells[col_idx]
                _set_cell_valign(cell, 'bottom')
                _set_cell_bottom_border(cell)
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                pass
        # Label row
        for col_idx in [0, 2]:
            try:
                cell = table.rows[1].cells[col_idx]
                _set_cell_valign(cell, 'top')
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                pass


def _add_page_breaks_after_sig_tables(doc):
    """Insert a page-break paragraph after each table that contains 'Главный'."""
    for table in doc.tables:
        has_sig = any(
            'Главный секретарь' in cell.text
            for row in table.rows
            for cell in row.cells
        )
        if not has_sig:
            continue
        p = OxmlElement('w:p')
        r = OxmlElement('w:r')
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        r.append(br)
        p.append(r)
        table._tbl.addnext(p)


# ─── Main entry point ─────────────────────────────────────────────────────────

def send_docx(html_string, filename):
    # Detect page orientation from CSS
    portrait = ('size: A4 portrait' in html_string or 'size:A4 portrait' in html_string)
    page_width_cm = _PAGE_WIDTH_PORTRAIT_CM if portrait else _PAGE_WIDTH_LANDSCAPE_CM

    soup = BeautifulSoup(html_string, 'html.parser')

    # 1. Strip navigation and CDN
    for el in soup.select('.no-print, link[rel="stylesheet"], script'):
        el.decompose()

    # 2. Structure preprocessing (order matters)
    _process_sig_blocks(soup)
    _process_flex_between(soup)
    _divs_to_paragraphs(soup)
    _normalize_tables(soup)

    # 3. Extract widths after normalization (colspan already expanded)
    table_widths = _extract_table_widths(soup)

    # 4. HTML → DOCX
    doc = _make_base_doc(portrait=portrait)
    HtmlToDocx().add_html_to_document(str(soup), doc)

    # 5. Apply column widths
    _apply_table_widths(doc, table_widths, page_width_cm)

    # 6. Post-process: date/place table borders and alignment
    _post_process_date_place_tables(doc)

    # 7. Insert page break after each signature table
    _add_page_breaks_after_sig_tables(doc)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_старт_спринт_docx(сор, секции, filename):
    """Generate Sprint start protocol DOCX matching the official sample layout."""
    from docx.shared import Mm, Pt

    TNR   = 'Times New Roman'
    C     = WD_ALIGN_PARAGRAPH.CENTER
    # Column grid widths (dxa) from sample — 8 columns
    GRID  = [675, 1843, 808, 1055, 2531, 1370, 1387, 928]
    NC    = len(GRID)
    TW    = sum(GRID)   # 10597 dxa ≈ 185 mm

    # ── micro-helpers ─────────────────────────────────────────────────────────

    def _sp(p):
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def _run(p, text, pt=8, bold=False):
        r = p.add_run(text)
        r.font.name = TNR
        r.font.size = Pt(pt)
        if bold:
            r.font.bold = True
        return r

    def _para(doc, text='', align=None, pt=10, bold=False):
        p = doc.add_paragraph()
        _sp(p)
        if align:
            p.alignment = align
        if text:
            _run(p, text, pt, bold)
        return p

    def _line_br(p):
        p.add_run()._r.append(OxmlElement('w:br'))

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr')
            tc.insert(0, pr)
        return pr

    def _borders(cell, sz=6):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')):
            pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')
            b.append(el)
        pr.append(b)

    def _no_border_inside(cell):
        """Remove insideH/insideV (used after merge to avoid phantom inner lines)."""
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is not None:
            for side in ('insideH', 'insideV'):
                el = b.find(qn(f'w:{side}'))
                if el is not None:
                    b.remove(el)

    def _cw(cell, w):
        pr = _tcpr(cell)
        el = pr.find(qn('w:tcW'))
        if el is None:
            el = OxmlElement('w:tcW')
            pr.append(el)
        el.set(qn('w:w'), str(w))
        el.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        v = pr.find(qn('w:vAlign'))
        if v is None:
            v = OxmlElement('w:vAlign')
            pr.append(v)
        v.set(qn('w:val'), val)

    def _fill(cell, text, align=None, pt=8, bold=False):
        p = cell.paragraphs[0]
        p.clear()
        _sp(p)
        if align:
            p.alignment = align
        if text:
            _run(p, text, pt, bold)
        _valign(cell)

    def _set_grid(tbl_obj):
        tbl = tbl_obj._tbl
        for old in tbl.findall(qn('w:tblGrid')):
            tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in GRID:
            col = OxmlElement('w:gridCol')
            col.set(qn('w:w'), str(w))
            grid.append(col)
        tblPr = tbl.find(qn('w:tblPr'))
        tblPr.addnext(grid) if tblPr is not None else tbl.insert(0, grid)

    def _tbl_w(tbl_obj):
        tbl = tbl_obj._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        el = tblPr.find(qn('w:tblW'))
        if el is None:
            el = OxmlElement('w:tblW')
            tblPr.append(el)
        el.set(qn('w:w'), str(TW))
        el.set(qn('w:type'), 'dxa')

    def add_row(tbl):
        return tbl.add_row()

    def border_row(row, sz=6):
        for cell in row.cells:
            _borders(cell, sz)

    # ── document ──────────────────────────────────────────────────────────────

    doc  = Document()
    sect = doc.sections[0]
    sect.page_height    = Mm(297)
    sect.page_width     = Mm(210)
    sect.top_margin     = Mm(12)
    sect.bottom_margin  = Mm(15)
    sect.left_margin    = Mm(15)
    sect.right_margin   = Mm(10)

    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    # ── Organisations header ───────────────────────────────────────────────────
    if сор.организаторы_блок:
        lines = [l.strip() for l in сор.организаторы_блок.split('\n')]
        p = _para(doc, align=C)
        for i, line in enumerate(lines):
            if i > 0:
                _line_br(p)
            qi = next((line.index(q) for q in ('"', '«') if q in line), None)
            if qi is None:
                _run(p, line, pt=10, bold=True)
            else:
                before  = line[:qi].strip()
                quoted  = line[qi:]
                combined = (before + ' ' + quoted) if before else quoted
                if len(combined) <= 80:
                    _run(p, combined, pt=10, bold=True)
                else:
                    if before:
                        _run(p, before, pt=10, bold=True)
                        _line_br(p)
                    _run(p, quoted, pt=10, bold=True)

    _para(doc, сор.название or '',                      align=C, pt=12, bold=True)
    _para(doc, 'Стартовый протокол',                    align=C, pt=12, bold=True)
    _para(doc, 'Группа спортивных дисциплин СПРИНТ',    align=C, pt=10, bold=True)

    # ── Main table ─────────────────────────────────────────────────────────────
    tbl = doc.add_table(rows=0, cols=NC)
    _tbl_w(tbl)
    _set_grid(tbl)

    # Row 0 — date / venue values
    r = add_row(tbl)
    cs = list(r.cells)
    cs[0].merge(cs[1])          # date: cols 0+1
    cs[2].merge(cs[4])          # spacer: cols 2+3+4
    cs[5].merge(cs[7])          # venue: cols 5+6+7
    for c in (cs[0], cs[2], cs[5]):
        _borders(c)
    _fill(cs[0], сор.даты_проведения   or '', align=C, pt=10, bold=True)
    _fill(cs[2], '')
    _fill(cs[5], сор.место_проведения  or '', align=C, pt=10, bold=True)

    # Row 1 — date / venue labels
    r = add_row(tbl)
    cs = list(r.cells)
    cs[0].merge(cs[1])
    cs[2].merge(cs[4])
    cs[5].merge(cs[7])
    for c in (cs[0], cs[2], cs[5]):
        _borders(c)
    _fill(cs[0], 'Дата проведения',   align=C, pt=8, bold=True)
    _fill(cs[2], '')
    _fill(cs[5], 'Место проведения',  align=C, pt=8, bold=True)

    # Row 2 — column headers
    r = add_row(tbl)
    cs = list(r.cells)
    cs[3].merge(cs[4])          # ФИО header: cols 3+4
    border_row(r)
    # №\nп/п — two runs separated by a line-break
    p = cs[0].paragraphs[0]
    p.clear(); _sp(p); p.alignment = C; _valign(cs[0])
    _run(p, '№', pt=8, bold=True)
    _line_br(p)
    _run(p, 'п/п', pt=8, bold=True)
    _fill(cs[1], 'Команда',                                 align=C, pt=8, bold=True)
    _fill(cs[2], 'Номер',                                   align=C, pt=8, bold=True)
    _fill(cs[3], 'Фамилия, имя, отчество участников',       align=C, pt=8, bold=True)
    _fill(cs[5], 'Территория',                              align=C, pt=8, bold=True)
    _fill(cs[6], 'Порядковый номер согласно жеребьёвке',    align=C, pt=8, bold=True)
    _fill(cs[7], 'Время старта',                            align=C, pt=8, bold=True)

    # ── Sections ───────────────────────────────────────────────────────────────
    for метка, команды in секции:
        if not команды:
            continue

        # Section header — all 8 cols merged
        r = add_row(tbl)
        cs = list(r.cells)
        cs[0].merge(cs[NC - 1])
        _borders(cs[0])
        _fill(cs[0], метка, align=C, pt=9, bold=True)

        # Data rows
        for пп, к in enumerate(команды, 1):
            r = add_row(tbl)
            cs = list(r.cells)
            cs[3].merge(cs[4])
            border_row(r)

            спортсмены = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
            фио  = ', '.join(
                с.фамилия_имя_отчество for с in спортсмены if с.фамилия_имя_отчество
            )
            время = (к.время_старта_спринт_план or '').split('.')[0]

            _fill(cs[0], str(пп),                                               align=C)
            _fill(cs[1], к.название or '—')
            _fill(cs[2], str(к.нагрудный_номер or ''),                          align=C)
            _fill(cs[3], фио)
            _fill(cs[5], к.субъект.название if к.субъект else '—')
            _fill(cs[6], str(к.стартовый_номер_спринт or ''),                   align=C)
            _fill(cs[7], время,                                                  align=C)

    # ── Signatures ─────────────────────────────────────────────────────────────
    for pos, fio in (('Главный судья', сор.главный_судья or ''),
                     ('Главный секретарь', сор.главный_секретарь or '')):
        r = add_row(tbl)
        cs = list(r.cells)
        cs[0].merge(cs[3])
        cs[4].merge(cs[7])
        _borders(cs[0]); _borders(cs[4])
        _fill(cs[0], pos, pt=10)
        _fill(cs[4], fio, pt=10)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)
