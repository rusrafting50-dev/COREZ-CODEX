# docx_utils.py — HTML → DOCX export via python-docx + htmldocx (no external binaries)
import re
from io import BytesIO
from bs4 import BeautifulSoup, NavigableString
from docx import Document
from docx.shared import Pt, Cm, Twips
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from htmldocx import HtmlToDocx
from flask import send_file

DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

# A4 landscape usable width (29.7 cm − 1.5 cm left − 1.0 cm right)
_PAGE_WIDTH_CM = 27.2

_BLOCK = frozenset({
    'div','p','h1','h2','h3','h4','h5','h6',
    'ul','ol','li','table','thead','tbody','tfoot','tr','th','td',
    'blockquote','section','article','header','footer','nav','figure',
})


# ─── Document setup ──────────────────────────────────────────────────────────

def _make_base_doc():
    doc = Document()
    sec = doc.sections[0]
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


# ─── Pre-processing: signature blocks ────────────────────────────────────────

def _process_sig_blocks(soup):
    """
    Convert each .sig-row into a single paragraph:
      «Должность   ____________________   Фамилия И.О.»
    """
    for sig_row in soup.find_all(class_='sig-row'):
        pos_el = sig_row.find(class_='sig-pos')
        fio_el = sig_row.find(class_='sig-fio')
        # Fallback: second non-pos child
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
        gap = '  ' + '_' * 30 + '  '
        p.append(NavigableString(gap + fio_text))
        sig_row.replace_with(p)


# ─── Pre-processing: flex space-between → date left / venue right ─────────────

def _process_flex_between(soup):
    """
    Convert d-flex justify-content-between containers (date/venue row) to a
    borderless 2-column table so they sit on opposite sides.
    """
    def _is_flex_between(tag):
        if tag.name != 'div':
            return False
        classes = tag.get('class') or []
        style   = tag.get('style', '')
        return (('justify-content-between' in classes) or
                ('justify-content:space-between' in style))

    for div in soup.find_all(_is_flex_between):
        block_children = [c for c in div.children if getattr(c, 'name', None)]
        if len(block_children) < 2:
            div.unwrap()
            continue

        left_text  = block_children[0].get_text(' ', strip=True)
        right_text = block_children[-1].get_text(' ', strip=True)

        tbl_html = (
            '<table style="width:100%;border:none">'
            '<tr>'
            f'<td style="border:none;text-align:left">{left_text}</td>'
            f'<td style="border:none;text-align:right">{right_text}</td>'
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

def _set_cell_width(cell, twips_val):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcW')):
        tcPr.remove(old)
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(twips_val)))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)


def _apply_table_widths(doc, table_widths):
    """Scale extracted px widths to the page and apply them to DOCX tables."""
    CM_TO_TWIPS = 566.929
    total_twips = int(_PAGE_WIDTH_CM * CM_TO_TWIPS)

    for t_idx, table in enumerate(doc.tables):
        if t_idx >= len(table_widths):
            break
        widths_px = table_widths[t_idx]
        if not widths_px:
            continue

        # Separate known and unknown widths
        known = [(i, w) for i, w in enumerate(widths_px) if w is not None]
        if not known:
            continue

        total_known_px = sum(w for _, w in known)
        n_unknown      = len(widths_px) - len(known)
        # Reserve 10% of total width for unknown columns
        reserved_for_unknown = (total_twips * 0.10) if n_unknown else 0
        available = total_twips - reserved_for_unknown
        scale = available / total_known_px  # twips per px

        default_unknown = (reserved_for_unknown / n_unknown) if n_unknown else 0

        col_twips = [
            int(w * scale) if w is not None else int(default_unknown)
            for w in widths_px
        ]

        # Apply to every cell in every row
        for row in table.rows:
            for c_idx, cell in enumerate(row.cells):
                if c_idx < len(col_twips):
                    _set_cell_width(cell, col_twips[c_idx])


# ─── Main entry point ─────────────────────────────────────────────────────────

def send_docx(html_string, filename):
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
    doc = _make_base_doc()
    HtmlToDocx().add_html_to_document(str(soup), doc)

    # 5. Apply column widths
    _apply_table_widths(doc, table_widths)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)
