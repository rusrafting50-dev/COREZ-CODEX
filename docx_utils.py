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
    # Explicitly set w:orient so Word doesn't have to infer from dimensions
    pgSz = sec._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if portrait:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')
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
                f'<td style="border:none;width:350px"><b>{pos_text}</b></td>'
                f'<td style="border:none;width:350px">{fio_text}</td>'
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

        if _css(div, 'font-weight') in ('bold', '700', '800', '900'):
            left_value  = f'<b>{left_value}</b>'
            right_value = f'<b>{right_value}</b>'

        if left_label or right_label:
            # Structured 2-row, 3-column table with value row and label row.
            # Date/venue cells get near-zero width (1px) so _apply_table_widths
            # maps them to ~0 twips in ALL rows; combined with autofit+noWrap in
            # _post_process_date_place_tables this gives true minimum-content width.
            tbl_html = (
                '<table class="date-place-table" style="width:100%;border:none">'
                '<tr>'
                f'<td style="border:none;text-align:center;width:1px">{left_value}</td>'
                '<td style="border:none;width:1000px"></td>'
                f'<td style="border:none;text-align:center;width:1px">{right_value}</td>'
                '</tr>'
                '<tr>'
                f'<td style="border:none;text-align:center;font-size:8pt;width:1px"><b>{left_label}</b></td>'
                '<td style="border:none;width:1000px"></td>'
                f'<td style="border:none;text-align:center;font-size:8pt;width:1px"><b>{right_label}</b></td>'
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
            # Tag with font-size for post-conversion processing (htmldocx ignores CSS font-size).
            fs = _css(div, 'font-size')
            if fs:
                div['data-docx-fs'] = fs
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
    """For each table return list of (px_width | None) per column.

    Prefers <colgroup><col> widths when present — handles colspan/rowspan
    headers where first-row cell count doesn't match actual column count.
    Falls back to first header-row cell widths.
    """
    result = []
    for table in soup.find_all('table'):
        cols = table.find_all('col')
        if cols:
            result.append([_px(c) for c in cols])
            continue
        first_row = table.find('tr')
        if not first_row:
            result.append([])
            continue
        result.append([_px(c) for c in first_row.find_all(['th', 'td'])])
    return result


# ─── Post-processing: apply column widths to DOCX ────────────────────────────

def _apply_table_widths(doc, table_widths, page_width_cm=None, scale_factor=1.0):
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

        if n_unknown == 0:
            # All columns known: distribute proportionally with round(), last column
            # absorbs rounding remainder so columns sum exactly to target width.
            target = round(total_twips * scale_factor)
            col_twips = [round(w / total_known_px * total_twips * scale_factor)
                         for w in widths_px]
            col_twips[-1] = target - sum(col_twips[:-1])
        else:
            reserved_for_unknown = total_twips * 0.10
            available = total_twips - reserved_for_unknown
            scale = available / total_known_px
            default_unknown = reserved_for_unknown / n_unknown
            col_twips = [
                int(w * scale * scale_factor) if w is not None
                else int(default_unknown * scale_factor)
                for w in widths_px
            ]

        # Set table-level preferred width to match column sum
        table_target = int(sum(col_twips))
        tblPr = table._tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            table._tbl.insert(0, tblPr)
        old_tblW = tblPr.find(qn('w:tblW'))
        if old_tblW is not None:
            tblPr.remove(old_tblW)
        new_tblW = OxmlElement('w:tblW')
        new_tblW.set(qn('w:w'), str(table_target))
        new_tblW.set(qn('w:type'), 'dxa')
        tblPr.append(new_tblW)

        for row in table.rows:
            # Collect cells and their _tc refs first to hold strong references.
            # lxml proxy objects get new ids after GC, so id(cell._tc) is only
            # stable while a strong reference to the proxy is held.
            row_cells = list(row.cells)
            row_tcs   = [c._tc for c in row_cells]
            tc_widths: dict[int, int] = {}
            for c_idx, tc in enumerate(row_tcs):
                tc_id = id(tc)
                if c_idx < len(col_twips):
                    tc_widths[tc_id] = tc_widths.get(tc_id, 0) + col_twips[c_idx]
            applied: set[int] = set()
            for cell, tc in zip(row_cells, row_tcs):
                tc_id = id(tc)
                if tc_id in applied:
                    continue
                applied.add(tc_id)
                if tc_id in tc_widths:
                    _set_cell_width(cell, tc_widths[tc_id])


# ─── Post-processing: date/place table borders and alignment ─────────────────

def _post_process_date_place_tables(doc, border_size=6, bold=True):
    """
    Find date/place header tables by label text ('Дата проведения' / 'Место проведения')
    instead of column count, which is unreliable after direct XML edits.

    Row 0 (value cells, first and last column):
      - vertical align bottom, bottom border, horizontal center, noWrap
    Row 1 (label cells): vertical align top, horizontal center
    """
    for table in doc.tables:
        tbl = table._tbl
        trs = tbl.findall(qn('w:tr'))
        if len(trs) != 2:
            continue
        all_text = ''.join((t.text or '') for t in tbl.iter(qn('w:t')))
        if 'Дата проведения' not in all_text and 'Место проведения' not in all_text:
            continue

        # Switch to autofit so noWrap can expand cells beyond fixed widths
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        old_layout = tblPr.find(qn('w:tblLayout'))
        if old_layout is not None:
            tblPr.remove(old_layout)
        layout = OxmlElement('w:tblLayout')
        layout.set(qn('w:type'), 'autofit')
        tblPr.append(layout)

        # Value row — first and last TC
        value_tcs = trs[0].findall(qn('w:tc'))
        if len(value_tcs) < 2:
            continue
        for tc in [value_tcs[0], value_tcs[-1]]:
            # noWrap + bottom align via XML (always reliable)
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = OxmlElement('w:tcPr')
                tc.insert(0, tcPr)
            # Replace any fixed tcW with type="auto" so Word sizes column to
            # content — combined with noWrap this gives minimum width, no wrap
            old_tcW = tcPr.find(qn('w:tcW'))
            if old_tcW is not None:
                tcPr.remove(old_tcW)
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:w'), '0')
            tcW.set(qn('w:type'), 'auto')
            tcPr.append(tcW)
            if tcPr.find(qn('w:noWrap')) is None:
                tcPr.append(OxmlElement('w:noWrap'))
            vAlign = tcPr.find(qn('w:vAlign'))
            if vAlign is None:
                vAlign = OxmlElement('w:vAlign')
                tcPr.append(vAlign)
            vAlign.set(qn('w:val'), 'bottom')
            # Bottom border and horizontal center via python-docx
            try:
                col_idx = list(trs[0].findall(qn('w:tc'))).index(tc)
                cell = table.rows[0].cells[col_idx]
                _set_cell_bottom_border(cell, size=border_size)
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                pass

        # Label row — first and last TC: top align, center, bold (or explicit remove bold)
        label_tcs = trs[1].findall(qn('w:tc'))
        for tc in ([label_tcs[0], label_tcs[-1]] if len(label_tcs) >= 2 else label_tcs):
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = OxmlElement('w:tcPr')
                tc.insert(0, tcPr)
            vAlign = tcPr.find(qn('w:vAlign'))
            if vAlign is None:
                vAlign = OxmlElement('w:vAlign')
                tcPr.append(vAlign)
            vAlign.set(qn('w:val'), 'top')
            try:
                col_idx = list(trs[1].findall(qn('w:tc'))).index(tc)
                cell = table.rows[1].cells[col_idx]
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                pass
            for w_r in tc.iter(qn('w:r')):
                rPr = w_r.find(qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr')
                    w_r.insert(0, rPr)
                if bold:
                    if rPr.find(qn('w:b')) is None:
                        rPr.append(OxmlElement('w:b'))
                    if rPr.find(qn('w:bCs')) is None:
                        rPr.append(OxmlElement('w:bCs'))
                else:
                    for tag in ('w:b', 'w:bCs'):
                        el = rPr.find(qn(tag))
                        if el is not None:
                            rPr.remove(el)

        # When bold=False also strip bold from value cells (may have been set by htmldocx)
        if not bold:
            for tc in [value_tcs[0], value_tcs[-1]]:
                for w_r in tc.iter(qn('w:r')):
                    rPr = w_r.find(qn('w:rPr'))
                    if rPr is None:
                        continue
                    for tag in ('w:b', 'w:bCs'):
                        el = rPr.find(qn(tag))
                        if el is not None:
                            rPr.remove(el)


def _add_page_breaks_after_sig_tables(doc):
    """Insert a page-break paragraph after each sig table except the last one."""
    sig_tables = [
        table for table in doc.tables
        if any(
            'Главный секретарь' in cell.text
            for row in table.rows
            for cell in row.cells
        )
    ]
    for table in sig_tables[:-1]:  # skip the last — no page break needed after final sig
        p = OxmlElement('w:p')
        r = OxmlElement('w:r')
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        r.append(br)
        p.append(r)
        table._tbl.addnext(p)


# ─── Post-processing: scale table fonts ──────────────────────────────────────

def _scale_table_fonts(doc, scale_factor):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        current = run.font.size
                        if current is None:
                            current = Pt(10)
                        run.font.size = int(current * scale_factor)


# ─── Post-processing: vertical text (writing-mode:vertical-*) ────────────────

def _apply_vertical_text(doc, soup):
    """Find th/td cells with writing-mode:vertical-* and apply btLr text direction."""
    vertical_texts = set()
    for cell_tag in soup.find_all(['th', 'td']):
        has_vertical = False
        for node in [cell_tag] + list(cell_tag.descendants):
            if not hasattr(node, 'get'):
                continue
            if re.search(r'writing-mode\s*:\s*vertical', node.get('style', '')):
                has_vertical = True
                break
        if has_vertical:
            text = cell_tag.get_text(strip=True)
            if text:
                vertical_texts.add(text)

    if not vertical_texts:
        return

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip() not in vertical_texts:
                    continue
                tcPr = cell._tc.get_or_add_tcPr()
                existing = tcPr.find(qn('w:textDirection'))
                if existing is not None:
                    tcPr.remove(existing)
                td_el = OxmlElement('w:textDirection')
                td_el.set(qn('w:val'), 'btLr')
                tcPr.append(td_el)
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER


# ─── Post-processing: uniform single borders on data tables ──────────────────

def _set_all_table_borders(doc, data_table_indices, size=3, color='333333'):
    """Overwrite cell borders on all data tables with uniform single lines."""
    for t_idx, table in enumerate(doc.tables):
        if t_idx not in data_table_indices:
            continue
        # Clear table-level borders (set to none) so they don't overlap cell borders
        tblPr = table._tbl.find(qn('w:tblPr'))
        if tblPr is not None:
            old = tblPr.find(qn('w:tblBorders'))
            if old is not None:
                tblPr.remove(old)
        tblBorders = OxmlElement('w:tblBorders')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            b = OxmlElement(f'w:{side}')
            b.set(qn('w:val'), 'none')
            tblBorders.append(b)
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            table._tbl.insert(0, tblPr)
        tblPr.append(tblBorders)
        for row in table.rows:
            for cell in row.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcBorders = tcPr.find(qn('w:tcBorders'))
                if tcBorders is None:
                    tcBorders = OxmlElement('w:tcBorders')
                    tcPr.append(tcBorders)
                for side in ('top', 'bottom', 'left', 'right'):
                    existing = tcBorders.find(qn(f'w:{side}'))
                    if existing is not None:
                        tcBorders.remove(existing)
                    border = OxmlElement(f'w:{side}')
                    border.set(qn('w:val'), 'single')
                    border.set(qn('w:sz'), str(size))
                    border.set(qn('w:color'), color)
                    tcBorders.append(border)


# ─── Post-processing: center all cells in all tables ─────────────────────────

_TNR = 'Times New Roman'


def _center_all_cells(doc):
    """
    For every table cell (direct XML, avoids python-docx stale-cache after merges):
      - vertical align: center
      - horizontal align (jc): center
      - spacing before/after: 0  (equivalent to Word's «Без интервала» style)
      - font: Times New Roman on runs and as paragraph-level run-property default
    """
    for table in doc.tables:
        for tr in table._tbl.findall(qn('w:tr')):
            for tc in tr.findall(qn('w:tc')):
                # ── vertical center ──────────────────────────────────────────
                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is None:
                    tcPr = OxmlElement('w:tcPr')
                    tc.insert(0, tcPr)
                vAlign = tcPr.find(qn('w:vAlign'))
                if vAlign is None:
                    vAlign = OxmlElement('w:vAlign')
                    tcPr.append(vAlign)
                vAlign.set(qn('w:val'), 'center')

                for w_p in tc.findall('.//' + qn('w:p')):
                    pPr = w_p.find(qn('w:pPr'))
                    if pPr is None:
                        pPr = OxmlElement('w:pPr')
                        w_p.insert(0, pPr)

                    # ── horizontal center ─────────────────────────────────────
                    jc = pPr.find(qn('w:jc'))
                    if jc is None:
                        jc = OxmlElement('w:jc')
                        pPr.append(jc)
                    jc.set(qn('w:val'), 'center')

                    # ── «Без интервала»: no space before/after ────────────────
                    spacing = pPr.find(qn('w:spacing'))
                    if spacing is None:
                        spacing = OxmlElement('w:spacing')
                        pPr.append(spacing)
                    spacing.set(qn('w:before'), '0')
                    spacing.set(qn('w:after'), '0')

                    # ── paragraph-level TNR default ───────────────────────────
                    p_rPr = pPr.find(qn('w:rPr'))
                    if p_rPr is None:
                        p_rPr = OxmlElement('w:rPr')
                        pPr.append(p_rPr)
                    _set_tnr(p_rPr)

                    # ── TNR on every run ──────────────────────────────────────
                    for w_r in w_p.findall('.//' + qn('w:r')):
                        rPr = w_r.find(qn('w:rPr'))
                        if rPr is None:
                            rPr = OxmlElement('w:rPr')
                            w_r.insert(0, rPr)
                        _set_tnr(rPr)


def _set_tnr(rPr_el):
    """Set Times New Roman on an rPr or paragraph-rPr element."""
    rFonts = rPr_el.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr_el.insert(0, rFonts)
    rFonts.set(qn('w:ascii'),  _TNR)
    rFonts.set(qn('w:hAnsi'),  _TNR)
    rFonts.set(qn('w:cs'),     _TNR)
    rFonts.set(qn('w:eastAsia'), _TNR)


# ─── Post-processing: noWrap on specific header cells ────────────────────────

def _nowrap_cells_by_text(doc, texts):
    """Set w:noWrap on every table cell whose text matches one of the given strings."""
    for table in doc.tables:
        seen: set[int] = set()
        for row in table.rows:
            for cell in row.cells:
                tc_id = id(cell._tc)
                if tc_id in seen:
                    continue
                seen.add(tc_id)
                if cell.text.strip() in texts:
                    tcPr = cell._tc.get_or_add_tcPr()
                    if tcPr.find(qn('w:noWrap')) is None:
                        tcPr.append(OxmlElement('w:noWrap'))


# ─── Post-processing: merge section-row cells HtmlToDocx expands ─────────────

def _merge_section_rows(doc, section_row_texts, font_half_pts=None):
    """
    HtmlToDocx expands colspan into N separate empty cells instead of a merged cell.
    For each row whose first cell text matches a known section-row label and
    whose remaining cells are empty, collapse them into one cell spanning all columns.
    """
    if not section_row_texts:
        return
    for table in doc.tables:
        tbl = table._tbl
        tblGrid = tbl.find(qn('w:tblGrid'))
        if tblGrid is None:
            continue
        n_grid_cols = len(tblGrid.findall(qn('w:gridCol')))
        if n_grid_cols < 2:
            continue
        for tr in tbl.findall(qn('w:tr')):
            tcs = tr.findall(qn('w:tc'))
            if len(tcs) < 2:
                continue
            first_text = ''.join(
                (t.text or '') for t in tcs[0].iter() if t.tag == qn('w:t')
            ).strip()
            if first_text not in section_row_texts:
                continue
            rest_empty = all(
                ''.join((t.text or '') for t in tc.iter() if t.tag == qn('w:t')).strip() == ''
                for tc in tcs[1:]
            )
            if not rest_empty:
                continue
            # Apply gridSpan to first TC and remove the rest
            first_tcPr = tcs[0].find(qn('w:tcPr'))
            if first_tcPr is None:
                first_tcPr = OxmlElement('w:tcPr')
                tcs[0].insert(0, first_tcPr)
            old_gs = first_tcPr.find(qn('w:gridSpan'))
            if old_gs is not None:
                first_tcPr.remove(old_gs)
            gs = OxmlElement('w:gridSpan')
            gs.set(qn('w:val'), str(n_grid_cols))
            first_tcPr.append(gs)
            for tc in tcs[1:]:
                tr.remove(tc)
            # Apply bold and optional font size to merged cell runs
            for w_r in tcs[0].iter(qn('w:r')):
                rPr = w_r.find(qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr')
                    w_r.insert(0, rPr)
                if rPr.find(qn('w:b')) is None:
                    rPr.append(OxmlElement('w:b'))
                if rPr.find(qn('w:bCs')) is None:
                    rPr.append(OxmlElement('w:bCs'))
                if font_half_pts:
                    for tag_name in ('w:sz', 'w:szCs'):
                        el = rPr.find(qn(tag_name))
                        if el is None:
                            el = OxmlElement(tag_name)
                            rPr.append(el)
                        el.set(qn('w:val'), str(font_half_pts))


# ─── Post-processing: left-align specific data-table columns ─────────────────

def _left_align_data_columns(doc, header_texts):
    """
    For each table whose first row contains a header matching one of header_texts,
    left-align (jc=left) every cell in that column across all rows.
    Works at XML level to avoid python-docx cache issues after merges.
    """
    for table in doc.tables:
        tbl = table._tbl
        trs = tbl.findall(qn('w:tr'))
        if not trs:
            continue
        header_tcs = trs[0].findall(qn('w:tc'))
        left_cols: set[int] = set()
        for col_idx, tc in enumerate(header_tcs):
            cell_text = ''.join((t.text or '') for t in tc.iter(qn('w:t'))).strip()
            if cell_text in header_texts:
                left_cols.add(col_idx)
        if not left_cols:
            continue
        for tr in trs[1:]:  # skip header row — it stays centered
            for col_idx, tc in enumerate(tr.findall(qn('w:tc'))):
                if col_idx not in left_cols:
                    continue
                for w_p in tc.findall('.//' + qn('w:p')):
                    pPr = w_p.find(qn('w:pPr'))
                    if pPr is None:
                        pPr = OxmlElement('w:pPr')
                        w_p.insert(0, pPr)
                    jc = pPr.find(qn('w:jc'))
                    if jc is None:
                        jc = OxmlElement('w:jc')
                        pPr.append(jc)
                    jc.set(qn('w:val'), 'left')


# ─── Post-processing: signature table rows — left-align + minimum height ──────

_SIG_KEYWORDS = frozenset({'Главный секретарь', 'Главный судья', 'Начальник дистанции'})


def _post_process_sig_tables(doc, row_height_twips=500):
    """
    Find signature tables (those containing any of _SIG_KEYWORDS) and:
      - left-align every cell
      - set minimum row height (~2× normal line height)
      - noWrap on the FIO column (col index 1) so names stay on one line
    Covers both sprint start protocol (Главный секретарь) and act (Начальник дистанции).
    """
    for table in doc.tables:
        all_text = ' '.join(cell.text for row in table.rows for cell in row.cells)
        if not any(kw in all_text for kw in _SIG_KEYWORDS):
            continue
        for tr in table._tbl.findall(qn('w:tr')):
            # Minimum row height
            trPr = tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = OxmlElement('w:trPr')
                tr.insert(0, trPr)
            old = trPr.find(qn('w:trHeight'))
            if old is not None:
                trPr.remove(old)
            trH = OxmlElement('w:trHeight')
            trH.set(qn('w:val'), str(row_height_twips))
            trH.set(qn('w:hRule'), 'atLeast')
            trPr.append(trH)
            tcs = tr.findall(qn('w:tc'))
            for col_idx, tc in enumerate(tcs):
                # Left-align every cell
                for w_p in tc.findall('.//' + qn('w:p')):
                    pPr = w_p.find(qn('w:pPr'))
                    if pPr is None:
                        pPr = OxmlElement('w:pPr')
                        w_p.insert(0, pPr)
                    jc = pPr.find(qn('w:jc'))
                    if jc is None:
                        jc = OxmlElement('w:jc')
                        pPr.append(jc)
                    jc.set(qn('w:val'), 'left')
                # noWrap on FIO column (col 1) so names don't wrap
                if col_idx == 1:
                    tcPr = tc.find(qn('w:tcPr'))
                    if tcPr is None:
                        tcPr = OxmlElement('w:tcPr')
                        tc.insert(0, tcPr)
                    if tcPr.find(qn('w:noWrap')) is None:
                        tcPr.append(OxmlElement('w:noWrap'))


# ─── Post-processing: left-align ALL cells in specified tables ───────────────

def _center_table_header_rows(doc, table_indices):
    """Re-center the first (header) row — h+v — after all-left override."""
    for t_idx, table in enumerate(doc.tables):
        if t_idx not in table_indices:
            continue
        rows = table._tbl.findall(qn('w:tr'))
        if not rows:
            continue
        for tc in rows[0].findall(qn('w:tc')):
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = OxmlElement('w:tcPr')
                tc.insert(0, tcPr)
            vAlign = tcPr.find(qn('w:vAlign'))
            if vAlign is None:
                vAlign = OxmlElement('w:vAlign')
                tcPr.append(vAlign)
            vAlign.set(qn('w:val'), 'center')
            for w_p in tc.findall('.//' + qn('w:p')):
                pPr = w_p.find(qn('w:pPr'))
                if pPr is None:
                    pPr = OxmlElement('w:pPr')
                    w_p.insert(0, pPr)
                jc = pPr.find(qn('w:jc'))
                if jc is None:
                    jc = OxmlElement('w:jc')
                    pPr.append(jc)
                jc.set(qn('w:val'), 'center')


def _left_align_all_in_tables(doc, table_indices):
    """Override horizontal alignment to left for every cell in the given tables."""
    for t_idx, table in enumerate(doc.tables):
        if t_idx not in table_indices:
            continue
        for tr in table._tbl.findall(qn('w:tr')):
            for tc in tr.findall(qn('w:tc')):
                for w_p in tc.findall('.//' + qn('w:p')):
                    pPr = w_p.find(qn('w:pPr'))
                    if pPr is None:
                        pPr = OxmlElement('w:pPr')
                        w_p.insert(0, pPr)
                    jc = pPr.find(qn('w:jc'))
                    if jc is None:
                        jc = OxmlElement('w:jc')
                        pPr.append(jc)
                    jc.set(qn('w:val'), 'left')


# ─── Post-processing: apply font sizes to paragraphs (htmldocx ignores CSS font-size) ──

def _apply_para_font_sizes(doc, size_map):
    """Set w:sz/w:szCs on runs in paragraphs whose normalised text matches size_map."""
    def _norm(text):
        return ' '.join(text.split())[:80]

    for para in doc.paragraphs:
        key = _norm(para.text)
        if not key or key not in size_map:
            continue
        half_pts = size_map[key]
        for run in para.runs:
            rPr = run._r.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                run._r.insert(0, rPr)
            for tag_name in ('w:sz', 'w:szCs'):
                el = rPr.find(qn(tag_name))
                if el is None:
                    el = OxmlElement(tag_name)
                    rPr.append(el)
                el.set(qn('w:val'), str(half_pts))


def _apply_section_row_font_size(doc, section_row_texts, font_half_pts):
    """Final-pass: force font size on merged section-row cells.

    Detection uses structure, not text: a section-row has exactly ONE tc whose
    gridSpan equals the table's total column count.  This is immune to any
    Unicode / whitespace mismatch between the HTML and the converted DOCX.
    Runs last so nothing else can overwrite the size.
    """
    if not font_half_pts:
        return
    for table in doc.tables:
        tbl = table._tbl
        tbl_grid = tbl.find(qn('w:tblGrid'))
        if tbl_grid is None:
            continue
        n_cols = len(tbl_grid.findall(qn('w:gridCol')))
        if n_cols < 2:
            continue
        for tr in tbl.findall(qn('w:tr')):
            tcs = tr.findall(qn('w:tc'))
            if len(tcs) != 1:
                continue
            tcPr = tcs[0].find(qn('w:tcPr'))
            if tcPr is None:
                continue
            gs = tcPr.find(qn('w:gridSpan'))
            if gs is None:
                continue
            if int(gs.get(qn('w:val'), '1')) != n_cols:
                continue
            # This row spans all columns — it is a section-row merged cell
            for w_r in tcs[0].iter(qn('w:r')):
                rPr = w_r.find(qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr')
                    w_r.insert(0, rPr)
                for tag_name in ('w:sz', 'w:szCs'):
                    el = rPr.find(qn(tag_name))
                    if el is None:
                        el = OxmlElement(tag_name)
                        rPr.append(el)
                    el.set(qn('w:val'), str(font_half_pts))


def _apply_table_cell_font_sizes(doc, col_sizes, skip_first_row=False, only_first_row=False):
    """Apply font sizes to specific table cells by (table_idx, col_idx) → half_pts."""
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            if skip_first_row and r_idx == 0:
                continue
            if only_first_row and r_idx > 0:
                break
            for c_idx, cell in enumerate(row.cells):
                half_pts = col_sizes.get((t_idx, c_idx))
                if half_pts is None:
                    continue
                for para in cell.paragraphs:
                    for run in para.runs:
                        rPr = run._r.find(qn('w:rPr'))
                        if rPr is None:
                            rPr = OxmlElement('w:rPr')
                            run._r.insert(0, rPr)
                        for tag_name in ('w:sz', 'w:szCs'):
                            el = rPr.find(qn(tag_name))
                            if el is None:
                                el = OxmlElement(tag_name)
                                rPr.append(el)
                            el.set(qn('w:val'), str(half_pts))


# ─── Post-processing: remove empty paragraphs directly after tables ──────────

def _apply_font_after(doc, after_text, font_half_pts):
    """Apply font_half_pts to all body children (paragraphs + tables) that come
    AFTER the first body paragraph containing after_text.
    Also removes space-before/after on those body paragraphs (Без интервала).
    """
    body = doc.element.body
    found = False
    for child in list(body):
        if child.tag == qn('w:sectPr'):
            continue
        if not found:
            if child.tag == qn('w:p'):
                text = ''.join((t.text or '') for t in child.iter(qn('w:t')))
                if after_text in text:
                    found = True
            continue  # skip everything up to and including the marker paragraph
        # ── font size on all runs ────────────────────────────────────────────
        for w_r in child.iter(qn('w:r')):
            rPr = w_r.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                w_r.insert(0, rPr)
            for tag_name in ('w:sz', 'w:szCs'):
                el = rPr.find(qn(tag_name))
                if el is None:
                    el = OxmlElement(tag_name)
                    rPr.append(el)
                el.set(qn('w:val'), str(font_half_pts))
        # ── no-spacing on direct body paragraphs ────────────────────────────
        if child.tag == qn('w:p'):
            pPr = child.find(qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                child.insert(0, pPr)
            spacing = pPr.find(qn('w:spacing'))
            if spacing is None:
                spacing = OxmlElement('w:spacing')
                pPr.append(spacing)
            spacing.set(qn('w:before'), '0')
            spacing.set(qn('w:after'), '0')


def _remove_empty_paragraphs_after_tables(doc):
    """Remove empty paragraphs that appear immediately after a table (e.g., between
    the date/venue layout table and the main data table)."""
    body = doc.element.body
    children = list(body)
    to_remove = []
    for i, child in enumerate(children):
        if child.tag != qn('w:p'):
            continue
        if i == 0 or children[i - 1].tag != qn('w:tbl'):
            continue
        has_text = any((t.text or '').strip() for t in child.iter(qn('w:t')))
        has_break = child.find('.//' + qn('w:br')) is not None
        if not has_text and not has_break:
            to_remove.append(child)
    for p in to_remove:
        body.remove(p)


# ─── Post-processing: remove trailing page-break paragraph ──────────────────

def _remove_trailing_page_break(doc):
    """Remove the last w:p that consists only of a page break (skip trailing w:sectPr)."""
    body = doc.element.body
    # body always ends with w:sectPr; find the last w:p before it
    last_p = None
    for child in reversed(list(body)):
        if child.tag == qn('w:p'):
            last_p = child
            break
        if child.tag == qn('w:sectPr'):
            continue
        break  # hit a table or other block — stop
    if last_p is None:
        return
    has_text = any((t.text or '').strip() for t in last_p.iter(qn('w:t')))
    page_breaks = [b for b in last_p.iter(qn('w:br'))
                   if b.get(qn('w:type')) == 'page']
    if page_breaks and not has_text:
        body.remove(last_p)


# ─── Post-processing: remove leading empty paragraphs from document body ──────

def _remove_leading_empty_paragraphs(doc, max_remove=5):
    """Remove empty paragraphs that HtmlToDocx leaves at the start of the document."""
    body = doc.element.body
    removed = 0
    while removed < max_remove:
        children = list(body)
        if not children or children[0].tag != qn('w:p'):
            break
        has_text = any((t.text or '').strip() for t in children[0].iter(qn('w:t')))
        if has_text:
            break
        body.remove(children[0])
        removed += 1


# ─── Main entry point ─────────────────────────────────────────────────────────

def send_docx(html_string, filename, portrait=None, date_place_border_size=6, date_place_bold=True, font_after_pt=None, font_after_marker=None):
    # Detect page orientation from CSS (can be overridden by caller)
    if portrait is None:
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

    # 2b. Collect font-size hints from paragraphs tagged by _divs_to_paragraphs.
    # HtmlToDocx ignores CSS font-size, so we apply it via OOXML after conversion.
    def _norm_text(text):
        return ' '.join(text.split())[:80]

    font_size_map = {}  # normalised_text -> half_points
    for p in soup.find_all(attrs={'data-docx-fs': True}):
        fs_str = p.get('data-docx-fs', '')
        try:
            pt = float(fs_str.replace('pt', '').replace('px', '').strip())
            key = _norm_text(p.get_text())
            if key:
                font_size_map[key] = int(pt * 2)
        except ValueError:
            pass
        del p['data-docx-fs']

    # 2c. Collect per-column font sizes from <th>/<td> style before normalization.
    # Only scans data tables (those with <thead>). <th> → header_cell_sizes (row 0),
    # <td> → body_cell_sizes (rows 1+). Keeps them separate to avoid collisions.
    header_cell_sizes = {}  # (table_idx, col_idx) -> half_pts  (from <th>)
    body_cell_sizes   = {}  # (table_idx, col_idx) -> half_pts  (from <td>)
    for t_idx, table_tag in enumerate(soup.find_all('table')):
        if not table_tag.find('thead'):
            continue
        for row in table_tag.find_all('tr'):
            if 'section-row' in (row.get('class') or []):
                continue
            for c_idx, cell in enumerate(row.find_all('th')):
                fs = _css(cell, 'font-size')
                if fs and (t_idx, c_idx) not in header_cell_sizes:
                    try:
                        header_cell_sizes[(t_idx, c_idx)] = int(float(
                            fs.replace('pt', '').replace('px', '').strip()) * 2)
                    except ValueError:
                        pass
            for c_idx, cell in enumerate(row.find_all('td')):
                fs = _css(cell, 'font-size')
                if fs and (t_idx, c_idx) not in body_cell_sizes:
                    try:
                        body_cell_sizes[(t_idx, c_idx)] = int(float(
                            fs.replace('pt', '').replace('px', '').strip()) * 2)
                    except ValueError:
                        pass

    # Collect section-row font size (first section-row td inline style)
    section_row_font_half_pts = None
    first_sr = soup.find(class_='section-row')
    if first_sr:
        sr_td = first_sr.find('td')
        if sr_td:
            fs = _css(sr_td, 'font-size')
            if fs:
                try:
                    section_row_font_half_pts = int(float(
                        fs.replace('pt', '').replace('px', '').strip()) * 2)
                except ValueError:
                    pass

    # Collect section-row texts before normalization so we can re-merge later
    section_row_texts = {
        tr.get_text(strip=True)
        for tr in soup.find_all(class_='section-row')
        if tr.get_text(strip=True)
    }
    _normalize_tables(soup)

    # 3. Extract widths after normalization (colspan already expanded)
    table_widths = _extract_table_widths(soup)

    # 3b. Mark data tables (those with <thead>) — layout/signature tables are excluded
    data_table_indices = {
        i for i, t in enumerate(soup.find_all('table'))
        if t.find('thead')
    }

    # 3c. Mark tables that want all cells left-aligned (class="all-left")
    all_left_table_indices = {
        i for i, t in enumerate(soup.find_all('table'))
        if 'all-left' in (t.get('class') or [])
    }

    # 4. HTML → DOCX
    doc = _make_base_doc(portrait=portrait)
    HtmlToDocx().add_html_to_document(str(soup), doc)

    # 5. Apply column widths (page_width_cm already reflects orientation)
    _apply_table_widths(doc, table_widths, page_width_cm)

    # 5b. Restore vertical text lost by HtmlToDocx (writing-mode:vertical-*)
    _apply_vertical_text(doc, soup)

    # 5c. Re-merge section-row cells that HtmlToDocx expands into N empty cells
    _merge_section_rows(doc, section_row_texts, font_half_pts=section_row_font_half_pts)

    # 5d. Center all cells + No Spacing + Times New Roman (overridden where needed below)
    _center_all_cells(doc)

    # 5d2. Override to left-align ALL cells in tables marked with class="all-left",
    # then restore center alignment in header row (row 0) of those tables.
    if all_left_table_indices:
        _left_align_all_in_tables(doc, all_left_table_indices)
        _center_table_header_rows(doc, all_left_table_indices & data_table_indices)

    # 5e. Left-align columns that should not be centered
    _left_align_data_columns(doc, {
        'Команда',
        'Фамилия, имя, отчество участников',
        'Территория',
    })

    # 5f. noWrap on specific header cells that must stay on one line
    _nowrap_cells_by_text(doc, {'Фамилия, имя, отчество участников', '№ п/п'})

    # 6. Post-process: date/place table — bottom-align value cells, border, noWrap
    _post_process_date_place_tables(doc, border_size=date_place_border_size, bold=date_place_bold)

    # 6b. Signature tables: left-align + 2× row height
    _post_process_sig_tables(doc)

    # 7. Insert page break after each signature table
    _add_page_breaks_after_sig_tables(doc)

    # 8. Overwrite borders on data tables with uniform single 0.375pt lines
    _set_all_table_borders(doc, data_table_indices)

    # 9. Remove empty paragraphs HtmlToDocx leaves at the top of the document
    _remove_leading_empty_paragraphs(doc)

    # 10. Apply font sizes collected before conversion (htmldocx ignores CSS font-size)
    if font_size_map:
        _apply_para_font_sizes(doc, font_size_map)
    if body_cell_sizes:
        _apply_table_cell_font_sizes(doc, body_cell_sizes, skip_first_row=True)
    if header_cell_sizes:
        _apply_table_cell_font_sizes(doc, header_cell_sizes, only_first_row=True)

    # 10.5. After-marker font: apply given pt size to all body children after the marker paragraph
    if font_after_pt and font_after_marker:
        _apply_font_after(doc, font_after_marker, int(font_after_pt * 2))

    # 11. Remove empty paragraphs directly after tables (e.g., between date/venue and data table)
    _remove_empty_paragraphs_after_tables(doc)

    # 12. Remove trailing page-break paragraph (artifact when sig table is last element)
    _remove_trailing_page_break(doc)

    # 13. Final-pass font size on section-row merged cells (structure-based, not text-based)
    _apply_section_row_font_size(doc, None, section_row_font_half_pts)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_акт_дистанции_docx(сор, дистанции, тип, filename):
    """Generate акт дистанции DOCX directly via python-docx, matching the official sample layout."""
    from docx.shared import Mm, Pt, Inches

    TNR = 'Times New Roman'
    C = WD_ALIGN_PARAGRAPH.CENTER
    J = WD_ALIGN_PARAGRAPH.JUSTIFY
    HALF_IN = Inches(0.5)   # first-line indent (457 200 EMU) — matches sample
    TW = 10488              # text width in dxa (185 mm)

    # ── helpers ──────────────────────────────────────────────────────────────

    def _sp(p, before=0, after=0):
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)

    def _run(p, text, pt=12, bold=False):
        r = p.add_run(text)
        r.font.name = TNR
        r.font.size = Pt(pt)
        if bold:
            r.font.bold = True
        return r

    def _p(doc, text='', align=None, pt=12, bold=False, fi=None, li=None):
        para = doc.add_paragraph()
        _sp(para)
        if align is not None:
            para.alignment = align
        if fi is not None:
            para.paragraph_format.first_line_indent = fi
        if li is not None:
            para.paragraph_format.left_indent = li
        if text:
            _run(para, text, pt, bold)
        return para

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr')
            tc.insert(0, pr)
        return pr

    def _cell_borders(cell, bottom=None, top=None):
        pr = _tcpr(cell)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{side}')
            sz = bottom if side == 'bottom' and bottom else (top if side == 'top' and top else None)
            if sz:
                el.set(qn('w:val'), 'single')
                el.set(qn('w:sz'), str(sz))
                el.set(qn('w:space'), '0')
                el.set(qn('w:color'), 'auto')
            else:
                el.set(qn('w:val'), 'nil')
            b.append(el)
        pr.append(b)

    def _cell_w(cell, w, wtype='dxa'):
        pr = _tcpr(cell)
        el = pr.find(qn('w:tcW'))
        if el is None:
            el = OxmlElement('w:tcW')
            pr.append(el)
        el.set(qn('w:w'), str(w))
        el.set(qn('w:type'), wtype)

    def _no_wrap(cell):
        pr = _tcpr(cell)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))

    def _cell_text(cell, text, align=None, pt=12, bold=False):
        para = cell.paragraphs[0]
        para.clear()
        _sp(para)
        if align:
            para.alignment = align
        if text:
            _run(para, text, pt, bold)

    def _tbl_no_borders(tbl_obj):
        tblPr = tbl_obj._tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl_obj._tbl.insert(0, tblPr)
        for old in tblPr.findall(qn('w:tblBorders')):
            tblPr.remove(old)
        b = OxmlElement('w:tblBorders')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')
            b.append(el)
        tblPr.append(b)

    def _tbl_width(tbl_obj, w):
        tblPr = tbl_obj._tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl_obj._tbl.insert(0, tblPr)
        el = tblPr.find(qn('w:tblW'))
        if el is None:
            el = OxmlElement('w:tblW')
            tblPr.append(el)
        el.set(qn('w:w'), str(w))
        el.set(qn('w:type'), 'dxa')

    def _tbl_grid(tbl_obj, widths):
        tbl = tbl_obj._tbl
        for old in tbl.findall(qn('w:tblGrid')):
            tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in widths:
            col = OxmlElement('w:gridCol')
            col.set(qn('w:w'), str(w))
            grid.append(col)
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is not None:
            tblPr.addnext(grid)
        else:
            tbl.insert(0, grid)

    def _page_break(doc):
        para = doc.add_paragraph()
        _sp(para)
        r = para.add_run()
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        r._r.append(br)

    # ── document setup ────────────────────────────────────────────────────────

    doc = Document()
    sect = doc.sections[0]
    sect.page_height = Mm(297)
    sect.page_width = Mm(210)
    sect.top_margin = Mm(12)
    sect.bottom_margin = Mm(15)
    sect.left_margin = Mm(15)
    sect.right_margin = Mm(10)
    sect.header_distance = Mm(10)
    sect.footer_distance = Mm(10)

    # Remove the default empty paragraph that python-docx creates
    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    # ── pages ─────────────────────────────────────────────────────────────────

    for idx, д in enumerate(дистанции):
        if idx > 0:
            _page_break(doc)

        # ── Organisations ──
        if сор.организаторы_блок:
            lines = [l.strip() for l in сор.организаторы_блок.split('\n')]
            para = _p(doc, align=C)

            def _line_break(p):
                p.add_run()._r.append(OxmlElement('w:br'))

            for i, line in enumerate(lines):
                if i > 0:
                    _line_break(para)
                # Split at opening quote char if present
                quote_idx = next(
                    (line.index(q) for q in ('"', '«') if q in line), None
                )
                if quote_idx is None:
                    _run(para, line, pt=12)
                else:
                    before  = line[:quote_idx].strip()
                    quoted  = line[quote_idx:]
                    combined = (before + ' ' + quoted) if before else quoted
                    if len(combined) <= 80:
                        # Fits on one line — output together
                        _run(para, combined, pt=12)
                    else:
                        # Doesn't fit — before-quote on one line, quoted part on next
                        if before:
                            _run(para, before, pt=12)
                            _line_break(para)
                        _run(para, quoted, pt=12)

        _p(doc)  # empty line

        # ── Competition name ──
        _p(doc, сор.название or '', align=C, pt=14, bold=True)

        _p(doc)  # empty line between name and date/venue table

        # ── Date / Venue table ──
        dv = doc.add_table(rows=2, cols=3)
        _tbl_width(dv, TW)
        _tbl_grid(dv, [2061, 5015, 3412])
        _tbl_no_borders(dv)

        c_d0, c_s0, c_v0 = dv.rows[0].cells
        for c in (c_d0, c_s0, c_v0):
            _cell_w(c, 0, 'auto')
        _no_wrap(c_d0)
        _no_wrap(c_v0)
        _cell_borders(c_d0, bottom=3)
        _cell_borders(c_v0, bottom=3)
        _cell_borders(c_s0)
        _cell_text(c_d0, сор.даты_проведения or '', align=C)
        _cell_text(c_v0, сор.место_проведения or '', align=WD_ALIGN_PARAGRAPH.RIGHT)
        _cell_text(c_s0, '')

        c_d1, c_s1, c_v1 = dv.rows[1].cells
        _cell_w(c_d1, 10)
        _cell_w(c_s1, 10467)
        _cell_w(c_v1, 11)
        _no_wrap(c_d1)
        _no_wrap(c_v1)
        _cell_borders(c_d1, top=3)
        _cell_borders(c_v1, top=3)
        _cell_borders(c_s1)
        _cell_text(c_d1, 'Дата проведения', align=C, pt=8)
        _cell_text(c_v1, 'Место проведения', align=C, pt=8)
        _cell_text(c_s1, '')

        _p(doc)  # empty line

        # ── АКТ ──
        _p(doc, 'АКТ', align=C, bold=True)
        _p(doc, 'утверждения спортивной дистанции', align=C)
        _p(doc, f'для спортивных дисциплин: «{д.название or "—"}»', align=C)

        _p(doc, align=J, fi=HALF_IN)  # empty with fi

        # ── Preambula ──
        судья = сор.главный_судья or ''
        нач_д = сор.начальник_дистанции or ''
        нач_сп = сор.начальник_спасательной_службы or ''
        preambula = (
            f'Мы, нижеподписавшиеся, главный судья соревнований {судья}, '
            f'начальник дистанции {нач_д}, '
            f'начальник спасательной службы {нач_сп} '
            'составили настоящий акт о том, что проведён оперативно-технический осмотр '
            'спортивной дистанции и определены основные её характеристики.'
        )
        _p(doc, preambula, align=J, fi=HALF_IN)
        _p(doc, 'В результате оперативно-технического осмотра установлено:', align=J, fi=HALF_IN)

        # ── Parameters ──
        def прот(n, text):
            _p(doc, text, align=J, fi=HALF_IN)

        протяж   = д.протяженность or ''
        катег    = д.категория or ''
        ср_вр    = д.среднее_время or ''
        лимит    = д.лимит_времени or ''
        старт    = д.вариант_старта or ''
        финиш    = д.вариант_финиша or ''
        страхов  = д.страховка or ''

        прот(1, f'1. Общая протяжённость дистанции: {протяж + " м." if протяж else ""}')
        прот(2, f'2. Категория определяющего препятствия: {катег + " к. с." if катег else ""}')
        прот(3, f'3. Среднее время прохождения: {ср_вр + " мин." if ср_вр else ""}')
        прот(4, f'4. Лимит времени для прохождения дистанции: {лимит + " мин." if лимит else ""}')
        прот(5, f'5. Вариант организации старта: {старт}')

        if тип == 'слалом':
            обр = f' ({д.кол_обратных})' if д.кол_обратных else ''
            прот(6, f'6. Количество ворот (в т.ч. обратных): {д.кол_ворот or ""}{обр}')
            прот(7, f'7. Конструкция и параметры ворот: {д.конструкция_ворот or ""}')
            прот(8, f'8. Вариант организации финиша: {финиш}')
            n_стр = 9
        elif тип == 'пс':
            прот(6, f'6. Конструкция и параметры буев: {д.буи or ""}')
            прот(7, f'7. Вариант организации финиша: {финиш}')
            n_стр = 8
        elif тип == 'дг':
            прот(6, f'6. Количество, конструкция и параметры буев: {д.буи or ""}')
            прот(7, f'7. Вариант организации финиша: {финиш}')
            n_стр = 8
        else:  # спринт
            прот(6, f'6. Вариант организации финиша: {финиш}')
            n_стр = 7

        _p(doc, f'{n_стр}. Описание видов страховки:', align=J, li=HALF_IN)
        if страхов:
            for line in страхов.split('\n'):
                line = line.strip()
                if line:
                    _p(doc, line, align=J, li=HALF_IN)

        _p(doc, align=J)
        _p(doc, 'Приложение: схема дистанции.', align=J)
        _p(doc, align=J)

        # ── Signature table ──
        sig = doc.add_table(rows=3, cols=2)
        _tbl_no_borders(sig)
        COL1 = 5920
        COL2 = TW - COL1

        sig_rows = [
            ('Главный судья соревнований',    сор.главный_судья or ''),
            ('Начальник дистанции',            сор.начальник_дистанции or ''),
            ('Начальник спасательной службы',  сор.начальник_спасательной_службы or ''),
        ]
        ROW_H = 680  # 1.2 cm in twips (1.2 × 1440 / 2.54 ≈ 680)
        for ri, (pos, fio) in enumerate(sig_rows):
            row = sig.rows[ri]
            # Set row height = 1.2 cm (exact)
            trPr = row._tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = OxmlElement('w:trPr')
                row._tr.insert(0, trPr)
            trH = trPr.find(qn('w:trHeight'))
            if trH is None:
                trH = OxmlElement('w:trHeight')
                trPr.append(trH)
            trH.set(qn('w:val'), str(ROW_H))
            trH.set(qn('w:hRule'), 'exact')

            c0, c1 = row.cells
            _cell_w(c0, COL1)
            _cell_w(c1, COL2)
            _cell_borders(c0)
            _cell_borders(c1)
            _cell_text(c0, pos, pt=12)
            _cell_text(c1, fio, pt=12)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_старт_спринт_docx(сор, группы, filename):
    """Generate Sprint start protocol DOCX.

    группы — list of sections groups; each group produces one complete protocol
    block (header + table). Typically two groups: R4 and R6.
    Each group is a list of (метка, команды) pairs.
    """
    from docx.shared import Mm, Pt

    TNR   = 'Times New Roman'
    C     = WD_ALIGN_PARAGRAPH.CENTER
    GRID  = [675, 1843, 808, 1055, 2531, 1370, 1387, 928]
    NC    = len(GRID)
    TW    = sum(GRID)

    W01   = GRID[0] + GRID[1]
    W234  = GRID[2] + GRID[3] + GRID[4]
    W567  = GRID[5] + GRID[6] + GRID[7]
    W34   = GRID[3] + GRID[4]
    W0123 = GRID[0] + GRID[1] + GRID[2] + GRID[3]
    W4567 = GRID[4] + GRID[5] + GRID[6] + GRID[7]

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

    def _page_br(doc):
        p = doc.add_paragraph()
        _sp(p)
        run = p.add_run()
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        run._r.append(br)

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

    def _bottom_border(cell, sz=6):
        """Only bottom underline, no other borders."""
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')):
            pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')
            b.append(el)
        bot = OxmlElement('w:bottom')
        bot.set(qn('w:val'), 'single')
        bot.set(qn('w:sz'), str(sz))
        bot.set(qn('w:space'), '0')
        bot.set(qn('w:color'), 'auto')
        b.append(bot)
        pr.append(b)

    def _no_border(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')):
            pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')
            b.append(el)
        pr.append(b)

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

    def _row_h(row, cm):
        """Set exact row height in cm."""
        val = round(cm * 1440 / 2.54)
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            row._tr.insert(0, trPr)
        trH = trPr.find(qn('w:trHeight'))
        if trH is None:
            trH = OxmlElement('w:trHeight')
            trPr.append(trH)
        trH.set(qn('w:val'), str(val))
        trH.set(qn('w:hRule'), 'exact')

    def add_row(tbl):
        return tbl.add_row()

    def border_row(row, sz=6):
        for cell in row.cells:
            _borders(cell, sz)

    # ── builders ──────────────────────────────────────────────────────────────

    def _add_header(doc):
        """Add org block + competition name + protocol titles."""
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
            _para(doc)  # blank line after org block

        _para(doc, сор.название or '', align=C, pt=12, bold=True)
        _para(doc)  # blank line
        _para(doc, 'Стартовый протокол', align=C, pt=12, bold=True)
        _para(doc)  # blank line
        _para(doc, 'Группа спортивных дисциплин СПРИНТ', align=C, pt=10, bold=True)

    def _add_table(doc, секции):
        """Add the start protocol table for one group of sections."""
        tbl = doc.add_table(rows=0, cols=NC)
        _tbl_w(tbl)
        _set_grid(tbl)

        # Row 0 — date / venue values: bottom border only, no box, h=0.6cm
        r = add_row(tbl)
        _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[1]);  _cw(cs[0], W01)
        cs[2].merge(cs[4]);  _cw(cs[2], W234)
        cs[5].merge(cs[7]);  _cw(cs[5], W567)
        _bottom_border(cs[0])
        _no_border(cs[2])
        _bottom_border(cs[5])
        _fill(cs[0], сор.даты_проведения  or '', align=C, pt=10, bold=True)
        _fill(cs[2], '')
        _fill(cs[5], сор.место_проведения or '', align=C, pt=10, bold=True)

        # Row 1 — date / venue labels: no borders, h=0.6cm
        r = add_row(tbl)
        _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[1]);  _cw(cs[0], W01)
        cs[2].merge(cs[4]);  _cw(cs[2], W234)
        cs[5].merge(cs[7]);  _cw(cs[5], W567)
        _no_border(cs[0]); _no_border(cs[2]); _no_border(cs[5])
        _fill(cs[0], 'Дата проведения',  align=C, pt=8, bold=True)
        _fill(cs[2], '')
        _fill(cs[5], 'Место проведения', align=C, pt=8, bold=True)

        # Row 2 — column headers: full borders
        r = add_row(tbl)
        cs = list(r.cells)
        cs[3].merge(cs[4]);  _cw(cs[3], W34)
        for i, w in enumerate(GRID):
            if i != 4:
                _cw(cs[i], w)
        border_row(r)
        p = cs[0].paragraphs[0]
        p.clear(); _sp(p); p.alignment = C; _valign(cs[0])
        _run(p, '№', pt=8, bold=True)
        _line_br(p)
        _run(p, 'п/п', pt=8, bold=True)
        _fill(cs[1], 'Команда',                              align=C, pt=8, bold=True)
        _fill(cs[2], 'Номер',                                align=C, pt=8, bold=True)
        _fill(cs[3], 'Фамилия, имя, отчество участников',   align=C, pt=8, bold=True)
        _fill(cs[5], 'Территория',                           align=C, pt=8, bold=True)
        _fill(cs[6], 'Порядковый номер согласно жеребьёвке', align=C, pt=8, bold=True)
        _fill(cs[7], 'Время старта',                         align=C, pt=8, bold=True)

        # Sections
        for метка, команды in секции:
            if not команды:
                continue

            r = add_row(tbl)
            _row_h(r, 0.6)
            cs = list(r.cells)
            cs[0].merge(cs[NC - 1]);  _cw(cs[0], TW)
            _borders(cs[0])
            _fill(cs[0], метка, align=C, pt=9, bold=True)

            for пп, к in enumerate(команды, 1):
                r = add_row(tbl)
                cs = list(r.cells)
                cs[3].merge(cs[4]);  _cw(cs[3], W34)
                for i, w in enumerate(GRID):
                    if i != 4:
                        _cw(cs[i], w)
                border_row(r)

                спортсмены = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
                фио  = ', '.join(
                    с.фамилия_имя_отчество for с in спортсмены if с.фамилия_имя_отчество
                )
                время = (к.время_старта_спринт_план or '').split('.')[0]

                _fill(cs[0], str(пп),                                            align=C)
                _fill(cs[1], к.название or '—')
                _fill(cs[2], str(к.нагрудный_номер or ''),                       align=C)
                _fill(cs[3], фио)
                _fill(cs[5], к.субъект.название if к.субъект else '—')
                _fill(cs[6], str(к.стартовый_номер_спринт or ''),                align=C)
                _fill(cs[7], время,                                               align=C)

        # Signatures: no borders, h=1.2cm
        for pos, fio in (('Главный судья', сор.главный_судья or ''),
                         ('Главный секретарь', сор.главный_секретарь or '')):
            r = add_row(tbl)
            _row_h(r, 1.2)
            cs = list(r.cells)
            cs[0].merge(cs[3]);  _cw(cs[0], W0123)
            cs[4].merge(cs[7]);  _cw(cs[4], W4567)
            _no_border(cs[0]); _no_border(cs[4])
            _fill(cs[0], pos, pt=10)
            _fill(cs[4], fio, pt=10)

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

    for i, секции in enumerate(группы):
        if i > 0:
            _page_br(doc)
        _add_header(doc)
        _add_table(doc, секции)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_старт_слалом_docx(сор, секции, filename):
    """Generate Slalom start protocol DOCX matching the official sample layout.

    секции: list of (метка, команды, мб_колонки) — one page per section.
      мб_колонки: list of {'заг': label, 'поле': field_name} (1 or 2 entries).
    """
    from docx.shared import Mm, Pt

    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER

    # 12-column grid (dxa) — cols 0-11: №, Команда, Номер(2+3), ФИО, Территория(5+6), Оч1, Оч2, Сумма, Врем1, Врем2
    # Visible ratios: 4/14/6/33/12/15/8/8 %
    GRID = [448, 1567, 336, 336, 3693, 672, 672, 560, 560, 560, 896, 896]
    NC   = len(GRID)
    TW   = sum(GRID)   # 11196 dxa ≈ 197.5 mm

    # Pre-computed merged widths
    W_DATE  = GRID[0]+GRID[1]+GRID[2]                               # 2351 cols 0-2
    W_MID   = GRID[3]+GRID[4]+GRID[5]                               # 4701 cols 3-5
    W_VENUE = GRID[6]+GRID[7]+GRID[8]+GRID[9]+GRID[10]+GRID[11]    # 4144 cols 6-11
    W_НОМ   = GRID[2]+GRID[3]                                        # 672  cols 2-3
    W_ТЕРР  = GRID[5]+GRID[6]                                        # 1344 cols 5-6
    W_МБ    = GRID[7]+GRID[8]+GRID[9]                                # 1680 cols 7-9
    W_SIG1  = 3905                                                    # signature left
    W_SIG2  = TW - W_SIG1                                            # signature right

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

    def _page_br(doc):
        p = doc.add_paragraph()
        _sp(p)
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        p.add_run()._r.append(br)

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr')
            tc.insert(0, pr)
        return pr

    def _borders(cell, sz=4):
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

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')):
            pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')
            b.append(el)
        bot = OxmlElement('w:bottom')
        bot.set(qn('w:val'), 'single')
        bot.set(qn('w:sz'), str(sz))
        bot.set(qn('w:space'), '0')
        bot.set(qn('w:color'), 'auto')
        b.append(bot)
        pr.append(b)

    def _no_border(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')):
            pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')
            b.append(el)
        pr.append(b)

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
        p.style = 'No Spacing'
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
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None:
            lay = OxmlElement('w:tblLayout')
            tblLook = tblPr.find(qn('w:tblLook'))
            if tblLook is not None:
                tblLook.addprevious(lay)
            else:
                tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')

    def _row_h(row, cm):
        val = round(cm * 1440 / 2.54)
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            row._tr.insert(0, trPr)
        trH = trPr.find(qn('w:trHeight'))
        if trH is None:
            trH = OxmlElement('w:trHeight')
            trPr.append(trH)
        trH.set(qn('w:val'), str(val))
        trH.set(qn('w:hRule'), 'exact')

    def _row_h_min(row, cm):
        """Minimum row height (hRule=atLeast) — allows row to grow for btLr text."""
        val = round(cm * 1440 / 2.54)
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            row._tr.insert(0, trPr)
        trH = trPr.find(qn('w:trHeight'))
        if trH is None:
            trH = OxmlElement('w:trHeight')
            trPr.append(trH)
        trH.set(qn('w:val'), str(val))
        trH.set(qn('w:hRule'), 'atLeast')

    def add_row(tbl):
        return tbl.add_row()

    def border_row(row, sz=4):
        for cell in row.cells:
            _borders(cell, sz)

    def _vmerge_restart(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')):
            pr.remove(old)
        vm = OxmlElement('w:vMerge')
        vm.set(qn('w:val'), 'restart')
        pr.append(vm)

    def _vmerge_cont(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')):
            pr.remove(old)
        vm = OxmlElement('w:vMerge')
        pr.append(vm)
        for para in cell.paragraphs:
            para.clear()

    # ── score helpers ─────────────────────────────────────────────────────────

    def _fmt_score(v):
        if v is None or v == 0:
            return ''
        return str(int(v)) if float(v) == int(v) else f'{v:.1f}'

    def _score(к, поле):
        if поле == 'очки_спринт':
            р = к.результат_спринт
            return _fmt_score(р.очки if р else None)
        elif поле == 'очки_пс':
            р = next((r for r in к.результат_пс if r.итоговое_место is not None), None)
            return _fmt_score(р.очки if р else None)
        elif поле == 'очки_дг':
            р = к.результат_дг
            return _fmt_score(р.очки if р else None)
        return ''

    def _sum_scores(к, колонки):
        total = 0.0
        has = False
        for кол in колонки:
            s = _score(к, кол['поле'])
            if s:
                try:
                    total += float(s); has = True
                except ValueError:
                    pass
        return _fmt_score(total) if has else ''

    def _time(к, попытка):
        рез = next((р for р in к.результат_слалом if р.номер_попытки == попытка), None)
        if рез and рез.результат_попытки is not None:
            s = float(рез.результат_попытки)
            m = int(s // 60); sec = s % 60
            return f'{m:02d}:{sec:05.2f}'
        return ''

    def _vert_cell(cell):
        """Vertical text btLr (bottom-to-top rotation), valign bottom, para centered."""
        pr = _tcpr(cell)
        # Remove any existing textDirection
        for old in pr.findall(qn('w:textDirection')):
            pr.remove(old)
        td = OxmlElement('w:textDirection')
        td.set(qn('w:val'), 'btLr')
        # Must appear before w:vAlign in tcPr per OOXML spec
        va = pr.find(qn('w:vAlign'))
        if va is not None:
            va.addprevious(td)
        else:
            pr.append(td)
        _valign(cell, 'bottom')
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ── builders ──────────────────────────────────────────────────────────────

    def _add_header(doc):
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
            _para(doc)

        def _spacing_15(p):
            pPr = p._p.get_or_add_pPr()
            sp = pPr.find(qn('w:spacing'))
            if sp is None:
                sp = OxmlElement('w:spacing')
                pPr.append(sp)
            sp.set(qn('w:line'), '360')
            sp.set(qn('w:lineRule'), 'auto')

        p4 = doc.add_paragraph(); p4.style = 'No Spacing'; _sp(p4); p4.add_run().font.size = Pt(4)
        p = _para(doc, сор.название or '', align=C, pt=12, bold=True); _spacing_15(p)
        p = _para(doc, 'Стартовый протокол', align=C, pt=12, bold=True); _spacing_15(p)
        p = _para(doc, 'Группа спортивных дисциплин СЛАЛОМ', align=C, pt=10, bold=True); _spacing_15(p)

    def _add_table(doc, метка, команды, мб_колонки):
        # Pad/truncate мб_колонки to exactly 2 entries
        кол1_заг  = мб_колонки[0]['заг']  if len(мб_колонки) > 0 else ''
        кол1_поле = мб_колонки[0]['поле'] if len(мб_колонки) > 0 else None
        кол2_заг  = мб_колонки[1]['заг']  if len(мб_колонки) > 1 else ''
        кол2_поле = мб_колонки[1]['поле'] if len(мб_колонки) > 1 else None

        def _set_fixed(tbl_obj, cols_w):
            tbl = tbl_obj._tbl
            tblPr = tbl.find(qn('w:tblPr'))
            if tblPr is None:
                tblPr = OxmlElement('w:tblPr')
                tbl.insert(0, tblPr)
            el = tblPr.find(qn('w:tblW'))
            if el is None:
                el = OxmlElement('w:tblW')
                tblPr.append(el)
            el.set(qn('w:w'), str(sum(cols_w)))
            el.set(qn('w:type'), 'dxa')
            lay = tblPr.find(qn('w:tblLayout'))
            if lay is None:
                lay = OxmlElement('w:tblLayout')
                tblLook = tblPr.find(qn('w:tblLook'))
                if tblLook is not None:
                    tblLook.addprevious(lay)
                else:
                    tblPr.append(lay)
            lay.set(qn('w:type'), 'fixed')
            for old in tbl.findall(qn('w:tblGrid')):
                tbl.remove(old)
            grid = OxmlElement('w:tblGrid')
            for w in cols_w:
                gc = OxmlElement('w:gridCol')
                gc.set(qn('w:w'), str(w))
                grid.append(gc)
            tblPr.addnext(grid)

        def _sep():
            p = doc.add_paragraph()
            p.style = 'No Spacing'
            pPr = p._p.get_or_add_pPr()
            sp = pPr.find(qn('w:spacing'))
            if sp is None:
                sp = OxmlElement('w:spacing')
                pPr.append(sp)
            sp.set(qn('w:before'), '0')
            sp.set(qn('w:after'), '0')
            sp.set(qn('w:line'), '20')
            sp.set(qn('w:lineRule'), 'exact')
            p.add_run().font.size = Pt(1)

        def _hdr_cell_margins(row, dxa=28):
            for tc in row._tr.findall(qn('w:tc')):
                pr = tc.find(qn('w:tcPr'))
                if pr is None:
                    pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
                for old in pr.findall(qn('w:tcMar')):
                    pr.remove(old)
                mar = OxmlElement('w:tcMar')
                for side in ('left', 'right'):
                    el = OxmlElement(f'w:{side}')
                    el.set(qn('w:w'), str(dxa))
                    el.set(qn('w:type'), 'dxa')
                    mar.append(el)
                pr.append(mar)

        # ── Table A: date / venue (autofit columns to content) ───────────────
        tA = doc.add_table(rows=0, cols=3)
        # Set total width = TW; let autofit size columns to text content
        _tA = tA._tbl; _tApr = _tA.find(qn('w:tblPr'))
        _el = _tApr.find(qn('w:tblW'))
        if _el is None:
            _el = OxmlElement('w:tblW'); _tApr.append(_el)
        _el.set(qn('w:w'), str(TW)); _el.set(qn('w:type'), 'dxa')

        # Row 0 — date/venue values: bottom border only
        r = add_row(tA)
        cs = list(r.cells)
        _bottom_border(cs[0]); _no_border(cs[1]); _bottom_border(cs[2])
        _fill(cs[0], сор.даты_проведения  or '', align=C, pt=10, bold=True)
        _fill(cs[1], '')
        _fill(cs[2], сор.место_проведения or '', align=C, pt=10, bold=True)

        # Row 1 — date/venue labels: no borders
        r = add_row(tA)
        cs = list(r.cells)
        _no_border(cs[0]); _no_border(cs[1]); _no_border(cs[2])
        _fill(cs[0], 'Дата проведения',  align=C, pt=8, bold=True)
        _fill(cs[1], '')
        _fill(cs[2], 'Место проведения', align=C, pt=8, bold=True)

        _sep()

        # ── Table B: headers + data (fixed layout) ────────────────────────────
        тbl = doc.add_table(rows=0, cols=NC)
        _tbl_w(тbl)
        _set_grid(тbl)

        # Row 2 — header row 1 (Очки Многоборье merged over cols 7-9)
        r = add_row(тbl)
        cs = list(r.cells)
        cs[2].merge(cs[3]);  _cw(cs[2], W_НОМ)
        cs[5].merge(cs[6]);  _cw(cs[5], W_ТЕРР)
        cs[7].merge(cs[9]);  _cw(cs[7], W_МБ)
        for i, w in enumerate(GRID):
            if i not in (2, 3, 5, 6, 7, 8, 9):
                _cw(cs[i], w)
        border_row(r); _hdr_cell_margins(r)
        # №\nп/п
        p = cs[0].paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p); p.alignment = C; _valign(cs[0])
        _run(p, '№', pt=8, bold=True); _line_br(p); _run(p, 'п/п', pt=8, bold=True)
        _fill(cs[1],  'Команда',                                   align=C, pt=8, bold=True)
        _fill(cs[2],  'Номер',                                     align=C, pt=8, bold=True)
        _fill(cs[4],  'Фамилия, имя, отчество участников',         align=C, pt=8, bold=True)
        _fill(cs[5],  'Территория',                                align=C, pt=8, bold=True)
        _fill(cs[7],  'Очки Многоборье',                           align=C, pt=8, bold=True)
        _valign(cs[7], 'bottom')
        _fill(cs[10], 'Время\nпопытка 1',                        align=C, pt=8, bold=True)
        _fill(cs[11], 'Время\nпопытка 2',                        align=C, pt=8, bold=True)
        # All columns except Очки Многоборье span both header rows
        for cell in [cs[0], cs[1], cs[2], cs[4], cs[5], cs[10], cs[11]]:
            _vmerge_restart(cell)

        # Row 3 — header row 2 (Очки sub-split, Номер and Территория still merged)
        r = add_row(тbl); _row_h_min(r, 3.0)
        cs = list(r.cells)
        cs[2].merge(cs[3]);  _cw(cs[2], W_НОМ)
        cs[5].merge(cs[6]);  _cw(cs[5], W_ТЕРР)
        for i, w in enumerate(GRID):
            if i not in (2, 3, 5, 6):
                _cw(cs[i], w)
        border_row(r); _hdr_cell_margins(r)
        _vmerge_cont(cs[0])
        _vmerge_cont(cs[1])
        _vmerge_cont(cs[2])
        _vmerge_cont(cs[4])
        _vmerge_cont(cs[5])
        _fill(cs[7],  кол1_заг, align=C, pt=7, bold=True); _vert_cell(cs[7])
        _fill(cs[8],  кол2_заг, align=C, pt=7, bold=True); _vert_cell(cs[8])
        _fill(cs[9],  'Сумма',  align=C, pt=7, bold=True); _vert_cell(cs[9])
        _vmerge_cont(cs[10])
        _vmerge_cont(cs[11])

        # Section header row — all cols merged, h=0.6cm
        r = add_row(тbl); _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[NC - 1]); _cw(cs[0], TW)
        _borders(cs[0])
        _fill(cs[0], метка, align=C, pt=9, bold=True)

        # Data rows
        for пп, к in enumerate(команды, 1):
            r = add_row(тbl)
            cs = list(r.cells)
            cs[2].merge(cs[3]);  _cw(cs[2], W_НОМ)
            cs[5].merge(cs[6]);  _cw(cs[5], W_ТЕРР)
            for i, w in enumerate(GRID):
                if i not in (2, 3, 5, 6):
                    _cw(cs[i], w)
            border_row(r)

            спортсмены = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
            фио = ', '.join(
                с.фамилия_имя_отчество for с in спортсмены if с.фамилия_имя_отчество
            )

            s1 = _score(к, кол1_поле) if кол1_поле else ''
            s2 = _score(к, кол2_поле) if кол2_поле else ''
            sm = _sum_scores(к, мб_колонки)

            _fill(cs[0],  str(пп),                                        align=C)
            _fill(cs[1],  к.название or '—')
            _fill(cs[2],  str(к.нагрудный_номер or ''),                   align=C)
            _fill(cs[4],  фио)
            _fill(cs[5],  к.субъект.название if к.субъект else '—')
            _fill(cs[7],  s1,                                              align=C)
            _fill(cs[8],  s2,                                              align=C)
            _fill(cs[9],  sm,                                              align=C)
            _fill(cs[10], _time(к, 1),                                    align=C)
            _fill(cs[11], _time(к, 2),                                    align=C)

        _sep()

        # ── Table C: signatures (auto-sized rows) ─────────────────────────────
        tC = doc.add_table(rows=0, cols=2)
        _set_fixed(tC, [W_SIG1, W_SIG2])

        for pos, fio in (('Главный судья', сор.главный_судья or ''),
                         ('Главный секретарь', сор.главный_секретарь or '')):
            r = add_row(tC); _row_h_min(r, 1.2)
            cs = list(r.cells)
            _cw(cs[0], W_SIG1); _cw(cs[1], W_SIG2)
            _no_border(cs[0]); _no_border(cs[1])
            _fill(cs[0], pos, pt=10)
            _fill(cs[1], fio, pt=10)

    # ── document ──────────────────────────────────────────────────────────────

    doc  = Document()
    sect = doc.sections[0]
    sect.page_height   = Mm(297)
    sect.page_width    = Mm(210)
    sect.top_margin    = Mm(18)
    sect.bottom_margin = Mm(12)
    sect.left_margin   = Mm(7)
    sect.right_margin  = Mm(5)

    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    for i, (метка, команды, мб_колонки) in enumerate(секции):
        if i > 0:
            _page_br(doc)
        _add_header(doc)
        _add_table(doc, метка, команды, мб_колонки)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_результаты_слалом_docx(
    сор, кл, кол_ворот, итоговый,
    муж_рез, муж_без, муж_разр, муж_поп, муж_ш1, муж_ш2,
    жен_рез, жен_без, жен_разр, жен_поп, жен_ш1, жен_ш2,
    filename
):
    """Generate Slalom results protocol DOCX (preliminary with gate columns, or final)."""
    from docx.shared import Mm, Pt

    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER

    кол_ворот = max(кол_ворот, 1)

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

    def _borders(cell, sz=4):
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

    def _no_border(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')):
            pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')
            b.append(el)
        pr.append(b)

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

    def _vert_cell(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:textDirection')):
            pr.remove(old)
        td = OxmlElement('w:textDirection')
        td.set(qn('w:val'), 'btLr')
        va = pr.find(qn('w:vAlign'))
        if va is not None:
            va.addprevious(td)
        else:
            pr.append(td)
        _valign(cell, 'bottom')
        for p in cell.paragraphs:
            p.alignment = C

    def _vmerge_restart(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')):
            pr.remove(old)
        vm = OxmlElement('w:vMerge')
        vm.set(qn('w:val'), 'restart')
        pr.append(vm)

    def _vmerge_cont(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')):
            pr.remove(old)
        vm = OxmlElement('w:vMerge')
        pr.append(vm)
        for para in cell.paragraphs:
            para.clear()

    def _row_h(row, cm):
        val = round(cm * 1440 / 2.54)
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            row._tr.insert(0, trPr)
        trH = trPr.find(qn('w:trHeight'))
        if trH is None:
            trH = OxmlElement('w:trHeight')
            trPr.append(trH)
        trH.set(qn('w:val'), str(val))
        trH.set(qn('w:hRule'), 'exact')

    def _row_h_min(row, cm):
        val = round(cm * 1440 / 2.54)
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            row._tr.insert(0, trPr)
        trH = trPr.find(qn('w:trHeight'))
        if trH is None:
            trH = OxmlElement('w:trHeight')
            trPr.append(trH)
        trH.set(qn('w:val'), str(val))
        trH.set(qn('w:hRule'), 'atLeast')

    def _border_row(row, sz=4):
        for cell in row.cells:
            _borders(cell, sz)

    def _set_grid(tbl_obj, grid):
        tbl = tbl_obj._tbl
        for old in tbl.findall(qn('w:tblGrid')):
            tbl.remove(old)
        g = OxmlElement('w:tblGrid')
        for w in grid:
            col = OxmlElement('w:gridCol')
            col.set(qn('w:w'), str(w))
            g.append(col)
        tblPr = tbl.find(qn('w:tblPr'))
        tblPr.addnext(g) if tblPr is not None else tbl.insert(0, g)

    def _tbl_w(tbl_obj, tw):
        tbl = tbl_obj._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        el = tblPr.find(qn('w:tblW'))
        if el is None:
            el = OxmlElement('w:tblW')
            tblPr.append(el)
        el.set(qn('w:w'), str(tw))
        el.set(qn('w:type'), 'dxa')

    # ── data format helpers ────────────────────────────────────────────────────

    def _fmt_dur(val):
        if val is None:
            return '–'
        s = float(val)
        m = int(s // 60)
        sec = s % 60
        return f'{m:02d}:{sec:05.2f}'

    def _fmt_шт(val):
        if not val:
            return '00:00'
        v = int(val)
        return f'{v//60:02d}:{v%60:02d}'

    # ── document header ────────────────────────────────────────────────────────

    def _add_doc_header(doc):
        if сор.организаторы_блок:
            lines = [l.strip() for l in сор.организаторы_блок.split('\n')]
            p = _para(doc, align=C)
            for i, line in enumerate(lines):
                if i > 0:
                    _line_br(p)
                _run(p, line, pt=10, bold=True)
            _para(doc)
        _para(doc, сор.название or '', align=C, pt=12, bold=True)
        if итоговый and сор.номер_екп:
            _para(doc, f'ЕКП № {сор.номер_екп}', align=C, pt=11, bold=True)
        заг = 'ПРОТОКОЛ РЕЗУЛЬТАТОВ' if итоговый else 'ПРЕДВАРИТЕЛЬНЫЙ ПРОТОКОЛ РЕЗУЛЬТАТОВ'
        _para(doc, заг, align=C, pt=12, bold=True)
        if кл == 'R4':
            дисц = 'СЛАЛОМ – R4, код ВРВС 1740031811Я'
        elif кл == 'R6':
            дисц = 'СЛАЛОМ – R6, код ВРВС 1740041811Я'
        else:
            дисц = f'СЛАЛОМ{" – " + кл if кл else ""}'
        _para(doc, f'Спортивная дисциплина: {дисц}', align=C, pt=10, bold=True)

    # ── signatures ────────────────────────────────────────────────────────────

    def _add_sigs(doc):
        for pos, fio in (('Главный судья', сор.главный_судья or ''),
                         ('Главный секретарь', сор.главный_секретарь or '')):
            p = doc.add_paragraph()
            _sp(p)
            p.paragraph_format.space_before = Pt(12)
            _run(p, pos + '   ' + fio, pt=10, bold=True)

    # ══════════════════════════════════════════════════════════════════════════
    # PRELIMINARY PROTOCOL  (two rows per team, gate columns, vertical text)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_pred(doc):
        # Landscape A4 usable ≈ 16155 dxa
        FIXED_LEFT  = [300, 1700, 450, 2700, 1270, 300, 300]
        FIXED_RIGHT = [300, 600, 740, 800, 800, 350, 500, 800]
        FIXED_SUM   = sum(FIXED_LEFT) + sum(FIXED_RIGHT)
        gate_w      = max(250, (16155 - FIXED_SUM) // кол_ворот)
        GRID        = FIXED_LEFT + [gate_w] * кол_ворот + FIXED_RIGHT
        NC          = len(GRID)
        TW          = sum(GRID)

        I_ПОП   = 5
        I_СТАРТ = 6
        I_ВОР0  = 7
        I_ФИН   = 7  + кол_ворот
        I_ШТ    = 8  + кол_ворот
        I_ВР    = 9  + кол_ворот
        I_ИТ    = 10 + кол_ворот
        I_ЛУЧ   = 11 + кол_ворот
        I_МЕС   = 12 + кол_ворот
        I_ОЧК   = 13 + кол_ворот
        I_НОР   = 14 + кол_ворот

        ROWSPAN_H = [0, 1, 2, 3, 4, I_ПОП, I_ЛУЧ, I_МЕС, I_ОЧК, I_НОР]
        ROWSPAN_D = [0, 1, 2, 3, 4, I_ЛУЧ, I_МЕС, I_ОЧК, I_НОР]

        def _fill_поп(cs, рез, ш):
            ст = (рез.кол_штраф_с10 or 0) * 10 + (рез.кол_штраф_старт or 0) * 50 if рез else 0
            _fill(cs[I_СТАРТ], str(ст) if ст else '0', align=C)
            for g in range(1, кол_ворот + 1):
                v = ш.get(g, 0) if ш else 0
                _fill(cs[I_ВОР0 + g - 1], str(v) if v else '0', align=C)
            фш = (рез.кол_штраф_финиш or 0) * 50 if рез else 0
            _fill(cs[I_ФИН], str(фш) if фш else '0', align=C)
            if рез is None:
                _fill(cs[I_ШТ], '–', align=C); _fill(cs[I_ВР], '–', align=C); _fill(cs[I_ИТ], '–', align=C)
            elif рез.статус_попытки == 'НСТ':
                _fill(cs[I_ШТ], '–', align=C); _fill(cs[I_ВР], 'Не старт.', align=C); _fill(cs[I_ИТ], '–', align=C)
            elif рез.статус_попытки == 'НФН':
                _fill(cs[I_ШТ], '–', align=C); _fill(cs[I_ВР], 'Не фин.', align=C); _fill(cs[I_ИТ], '–', align=C)
            elif рез.статус_попытки == 'ПЛВ':
                _fill(cs[I_ШТ], '–', align=C); _fill(cs[I_ВР], 'Прев. ЛВ', align=C); _fill(cs[I_ИТ], '–', align=C)
            elif рез.статус_попытки == 'ДК':
                _fill(cs[I_ШТ], _fmt_шт(рез.штрафное_время), align=C)
                _fill(cs[I_ВР], _fmt_dur(рез.время_прохождения) if рез.время_прохождения else '–', align=C)
                _fill(cs[I_ИТ], 'ДК', align=C, bold=True)
            else:
                _fill(cs[I_ШТ], _fmt_шт(рез.штрафное_время), align=C)
                _fill(cs[I_ВР], _fmt_dur(рез.время_прохождения) if рез.время_прохождения else '–', align=C)
                _fill(cs[I_ИТ], _fmt_dur(рез.результат_попытки) if рез.результат_попытки else '–', align=C, bold=True)

        tbl = doc.add_table(rows=0, cols=NC)
        _set_grid(tbl, GRID)
        _tbl_w(tbl, TW)

        # Header row 1
        rh1 = tbl.add_row()
        _row_h_min(rh1, 3.5)
        cs = list(rh1.cells)
        for i, w in enumerate(GRID):
            _cw(cs[i], w)
        _border_row(rh1)
        cs[I_СТАРТ].merge(cs[I_ФИН])
        cs[I_ШТ].merge(cs[I_ИТ])
        p = cs[0].paragraphs[0]; p.clear(); _sp(p); p.alignment = C
        _run(p, '№', pt=8, bold=True); _line_br(p); _run(p, 'п/п', pt=8, bold=True)
        _valign(cs[0])
        _fill(cs[1], 'Команда', align=C, bold=True)
        _fill(cs[2], 'Номер', align=C, bold=True)
        _fill(cs[3], 'Фамилия, имя, отчество участников', align=C, bold=True)
        _fill(cs[4], 'Субъект РФ', align=C, bold=True)
        _fill(cs[I_ПОП], 'Попытка', align=C, bold=True); _vert_cell(cs[I_ПОП])
        _fill(cs[I_СТАРТ], 'Прохождение дистанции (штрафные баллы)', align=C, bold=True)
        _fill(cs[I_ШТ], 'Результат попытки', align=C, bold=True)
        _fill(cs[I_ЛУЧ], 'Результат лучшей попытки', align=C, bold=True)
        _fill(cs[I_МЕС], 'Место', align=C, bold=True); _vert_cell(cs[I_МЕС])
        _fill(cs[I_ОЧК], 'Очки', align=C, bold=True)
        _fill(cs[I_НОР], 'Выполнение разряда', align=C, bold=True); _vert_cell(cs[I_НОР])
        for col in ROWSPAN_H:
            _vmerge_restart(cs[col])

        # Header row 2 (vertical gate numbers)
        rh2 = tbl.add_row()
        _row_h_min(rh2, 3.0)
        cs2 = list(rh2.cells)
        for i, w in enumerate(GRID):
            _cw(cs2[i], w)
        _border_row(rh2)
        for col in ROWSPAN_H:
            _vmerge_cont(cs2[col])
        _fill(cs2[I_СТАРТ], 'Старт', align=C, bold=True);  _vert_cell(cs2[I_СТАРТ])
        for g in range(1, кол_ворот + 1):
            _fill(cs2[I_ВОР0 + g - 1], str(g), align=C, bold=True)
            _vert_cell(cs2[I_ВОР0 + g - 1])
        _fill(cs2[I_ФИН], 'Финиш', align=C, bold=True);   _vert_cell(cs2[I_ФИН])
        _fill(cs2[I_ШТ], 'Штраф', align=C, bold=True)
        _fill(cs2[I_ВР], 'Время', align=C, bold=True)
        _fill(cs2[I_ИТ], 'Итоговый', align=C, bold=True)

        def _add_section(рез_list, без_list, разряды, поп, ш1_d, ш2_d, заг_пол):
            if not рез_list and not без_list:
                return
            норм_флаг = (len(рез_list) + len(без_list)) >= 6
            rg = tbl.add_row(); _row_h(rg, 0.6)
            cg = list(rg.cells)
            cg[0].merge(cg[NC - 1])
            _border_row(rg)
            _fill(list(rg.cells)[0], заг_пол, align=C, pt=9, bold=True)

            idx = 0
            for рез in рез_list:
                к = рез.команда; idx += 1
                р1 = поп.get(к.id, {}).get(1)
                р2 = поп.get(к.id, {}).get(2)
                _add_team_rows(к, idx, р1, р2, ш1_d.get(к.id, {}), ш2_d.get(к.id, {}),
                               рез.лучший_результат, рез.место, рез.очки,
                               разряды.get(к.id, '–') if норм_флаг else '–')
            for рез_б in без_list:
                к = рез_б.команда; idx += 1
                р1 = поп.get(к.id, {}).get(1)
                р2 = поп.get(к.id, {}).get(2)
                _add_team_rows(к, idx, р1, р2, ш1_d.get(к.id, {}), ш2_d.get(к.id, {}),
                               None, None, None, '–')

        def _add_team_rows(к, пп, р1, р2, ш1, ш2, лучший, место, очки, норм):
            спортсмены = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
            фио = ', '.join(с.фамилия_имя_отчество for с in спортсмены if с.фамилия_имя_отчество)
            rd1 = tbl.add_row(); rd2 = tbl.add_row()
            _row_h(rd1, 0.55); _row_h(rd2, 0.55)
            cd1 = list(rd1.cells); cd2 = list(rd2.cells)
            for i, w in enumerate(GRID):
                _cw(cd1[i], w); _cw(cd2[i], w)
            _border_row(rd1); _border_row(rd2)
            for col in ROWSPAN_D:
                _vmerge_restart(cd1[col])
                _vmerge_cont(cd2[col])
            _fill(cd1[0], str(пп), align=C)
            _fill(cd1[1], к.название or '—')
            _fill(cd1[2], str(к.нагрудный_номер or ''), align=C)
            _fill(cd1[3], фио, pt=7)
            _fill(cd1[4], к.субъект.название if к.субъект else '—', pt=7)
            _fill(cd1[I_ПОП], '1', align=C)
            _fill(cd2[I_ПОП], '2', align=C)
            _fill_поп(cd1, р1, ш1)
            _fill_поп(cd2, р2, ш2)
            _fill(cd1[I_ЛУЧ], _fmt_dur(лучший), align=C, bold=True)
            _fill(cd1[I_МЕС], str(место) if место else '–', align=C, bold=True)
            _fill(cd1[I_ОЧК], str(int(очки)) if очки else '–', align=C)
            _fill(cd1[I_НОР], норм, align=C)

        _add_section(муж_рез, муж_без, муж_разр, муж_поп, муж_ш1, муж_ш2, 'Мужчины')
        _add_section(жен_рез, жен_без, жен_разр, жен_поп, жен_ш1, жен_ш2, 'Женщины')

    # ══════════════════════════════════════════════════════════════════════════
    # FINAL PROTOCOL  (one row per team, no gate columns)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_итог(doc):
        GRID = [300, 1700, 450, 2700, 1200,
                680, 520, 760,
                680, 520, 760,
                850, 360, 500, 850]
        NC = len(GRID)
        TW = sum(GRID)

        def _init_r(row, borders=True):
            cs = list(row.cells)
            for i, w in enumerate(GRID):
                _cw(cs[i], w)
            if borders:
                _border_row(row)
            return cs

        tbl = doc.add_table(rows=0, cols=NC)
        _set_grid(tbl, GRID)
        _tbl_w(tbl, TW)

        rh1 = tbl.add_row(); _row_h_min(rh1, 1.0)
        cs = _init_r(rh1)
        cs[5].merge(cs[7])
        cs[8].merge(cs[10])
        p = cs[0].paragraphs[0]; p.clear(); _sp(p); p.alignment = C
        _run(p, '№', pt=8, bold=True); _line_br(p); _run(p, 'п/п', pt=8, bold=True); _valign(cs[0])
        _fill(cs[1], 'Команда', align=C, bold=True)
        _fill(cs[2], 'Номер', align=C, bold=True)
        _fill(cs[3], 'Состав', align=C, bold=True)
        _fill(cs[4], 'Субъект РФ', align=C, bold=True)
        _fill(cs[5], '1-я попытка', align=C, bold=True)
        _fill(cs[8], '2-я попытка', align=C, bold=True)
        _fill(cs[11], 'Лучшая попытка', align=C, bold=True)
        _fill(cs[12], 'Место', align=C, bold=True)
        _fill(cs[13], 'Очки', align=C, bold=True)
        _fill(cs[14], 'Выполнение разряда', align=C, bold=True)
        for col in [0, 1, 2, 3, 4, 11, 12, 13, 14]:
            _vmerge_restart(cs[col])

        rh2 = tbl.add_row(); _row_h_min(rh2, 0.8)
        cs2 = _init_r(rh2)
        for col in [0, 1, 2, 3, 4, 11, 12, 13, 14]:
            _vmerge_cont(cs2[col])
        for i, lbl in enumerate(['Время', 'Штраф', 'Результат', 'Время', 'Штраф', 'Результат']):
            _fill(cs2[5 + i], lbl, align=C, bold=True)

        def _поп_cells(cs, рез, base):
            if рез is None:
                for i in range(3): _fill(cs[base + i], '–', align=C)
            elif рез.статус_попытки == 'НСТ':
                _fill(cs[base], 'Не старт.', align=C); _fill(cs[base+1], '–', align=C); _fill(cs[base+2], '–', align=C)
            elif рез.статус_попытки == 'НФН':
                _fill(cs[base], 'Не фин.', align=C); _fill(cs[base+1], '–', align=C); _fill(cs[base+2], '–', align=C)
            elif рез.статус_попытки == 'ПЛВ':
                _fill(cs[base], '–', align=C); _fill(cs[base+1], '–', align=C); _fill(cs[base+2], 'Прев. ЛВ', align=C)
            elif рез.статус_попытки == 'ДК':
                _fill(cs[base], _fmt_dur(рез.время_прохождения) if рез.время_прохождения else '–', align=C)
                _fill(cs[base+1], _fmt_шт(рез.штрафное_время), align=C)
                _fill(cs[base+2], 'ДК', align=C, bold=True)
            else:
                _fill(cs[base], _fmt_dur(рез.время_прохождения) if рез.время_прохождения else '–', align=C)
                _fill(cs[base+1], _fmt_шт(рез.штрафное_время), align=C)
                _fill(cs[base+2], _fmt_dur(рез.результат_попытки) if рез.результат_попытки else '–', align=C, bold=True)

        def _add_sect(рез_list, без_list, разряды, поп, заг_пол):
            if not рез_list and not без_list:
                return
            норм_флаг = (len(рез_list) + len(без_list)) >= 6
            rg = tbl.add_row(); _row_h(rg, 0.6)
            cg = list(rg.cells); cg[0].merge(cg[NC - 1])
            _border_row(rg)
            _fill(list(rg.cells)[0], заг_пол, align=C, pt=9, bold=True)
            idx = 0
            for is_без, entries in ((False, рез_list), (True, без_list)):
                for рез in entries:
                    к = рез.команда; idx += 1
                    р1 = поп.get(к.id, {}).get(1)
                    р2 = поп.get(к.id, {}).get(2)
                    спортсмены = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
                    состав = ', '.join(
                        f'{с.фамилия_имя_отчество} ({с.дата_рождения.strftime("%d.%m.%Y") if с.дата_рождения else "б/д"}, {с.разряд or "б/р"})'
                        for с in спортсмены
                    )
                    dr = tbl.add_row(); _row_h(dr, 0.55)
                    cs = _init_r(dr)
                    _fill(cs[0], str(idx), align=C)
                    _fill(cs[1], к.название or '—')
                    _fill(cs[2], str(к.нагрудный_номер or ''), align=C)
                    _fill(cs[3], состав, pt=7)
                    _fill(cs[4], к.субъект.название if к.субъект else '—', pt=7)
                    _поп_cells(cs, р1, 5)
                    _поп_cells(cs, р2, 8)
                    if is_без:
                        for i in range(11, 15): _fill(cs[i], '–', align=C)
                    else:
                        _fill(cs[11], _fmt_dur(рез.лучший_результат), align=C, bold=True)
                        _fill(cs[12], str(рез.место), align=C, bold=True)
                        _fill(cs[13], str(int(рез.очки)) if рез.очки else '–', align=C)
                        _fill(cs[14], разряды.get(к.id, '–') if норм_флаг else '–', align=C)

        _add_sect(муж_рез, муж_без, муж_разр, муж_поп, 'Мужчины')
        _add_sect(жен_рез, жен_без, жен_разр, жен_поп, 'Женщины')

    # ── Build document ────────────────────────────────────────────────────────

    doc = Document()
    sect = doc.sections[0]
    sect.page_height   = Mm(210)
    sect.page_width    = Mm(297)
    sect.top_margin    = Mm(12)
    sect.bottom_margin = Mm(10)
    sect.left_margin   = Mm(7)
    sect.right_margin  = Mm(5)

    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    _add_doc_header(doc)
    _para(doc)

    if итоговый:
        _build_итог(doc)
    else:
        _build_pred(doc)

    _add_sigs(doc)

    buf2 = BytesIO()
    doc.save(buf2)
    buf2.seek(0)
    return send_file(buf2, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_нагрудные_номера_docx(сор, команды, filename):
    """Generate bib-number issuance protocol DOCX matching the official sample layout.

    команды — all Команда objects for the competition, ordered by нагрудный_номер.
    Groups are derived from последовательность_видов or from actual data.
    """
    from docx.shared import Mm, Pt

    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT

    # Column widths (dxa) — 10 columns, total 15420 dxa matching the sample
    # №п/п, Команда, Территория, Тренер-представитель,
    # Выданы нагрудные номера, Получил подпись,
    # Сдан залог руб., Сданы нагрудные номера, Получил залог руб., Подпись
    GRID = [418, 2388, 2373, 2493, 1388, 1388, 1194, 1388, 1194, 1196]
    NC   = len(GRID)
    TW   = sum(GRID)   # 15420 dxa

    # Pre-computed merged widths for info-bar and signature rows
    W_DATE  = GRID[0] + GRID[1] + GRID[2]                          # cols 0-2
    W_MID   = GRID[3] + GRID[4] + GRID[5]                          # cols 3-5
    W_VENUE = GRID[6] + GRID[7] + GRID[8] + GRID[9]                # cols 6-9
    W_SIG_L = GRID[0] + GRID[1] + GRID[2] + GRID[3] + GRID[4]     # left sig col
    W_SIG_R = GRID[5] + GRID[6] + GRID[7] + GRID[8] + GRID[9]     # right sig col

    # ── micro-helpers ─────────────────────────────────────────────────────────

    def _sp(p):
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def _run(p, text, pt=10, bold=False):
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

    def _borders(cell, sz=3):
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

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')):
            pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')
            b.append(el)
        bot = OxmlElement('w:bottom')
        bot.set(qn('w:val'), 'single')
        bot.set(qn('w:sz'), str(sz))
        bot.set(qn('w:space'), '0')
        bot.set(qn('w:color'), 'auto')
        b.append(bot)
        pr.append(b)

    def _no_border(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')):
            pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')
            b.append(el)
        pr.append(b)

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

    def _fill(cell, text, align=None, pt=10, bold=False):
        p = cell.paragraphs[0]
        p.clear()
        _sp(p)
        if align is not None:
            p.alignment = align
        if text:
            _run(p, text, pt, bold)
        _valign(cell)

    def _set_grid(tbl_obj):
        tbl = tbl_obj._tbl
        for old in tbl.findall(qn('w:tblGrid')):
            tbl.remove(old)
        grid_el = OxmlElement('w:tblGrid')
        for w in GRID:
            col = OxmlElement('w:gridCol')
            col.set(qn('w:w'), str(w))
            grid_el.append(col)
        tblPr = tbl.find(qn('w:tblPr'))
        tblPr.addnext(grid_el) if tblPr is not None else tbl.insert(0, grid_el)

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

    def _row_h(row, cm):
        val = round(cm * 1440 / 2.54)
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            row._tr.insert(0, trPr)
        trH = trPr.find(qn('w:trHeight'))
        if trH is None:
            trH = OxmlElement('w:trHeight')
            trPr.append(trH)
        trH.set(qn('w:val'), str(val))
        trH.set(qn('w:hRule'), 'exact')

    def _add_row(tbl):
        return tbl.add_row()

    # ── group ordering helpers ────────────────────────────────────────────────

    КЛАСС_МЕТКА = {'R4': 'R4', 'R6': 'R6'}
    ПОЛ_МЕТКА   = {'м': 'Мужчины', 'ж': 'Женщины'}

    def _section_label(класс, пол):
        return f'Спринт – {КЛАСС_МЕТКА.get(класс, класс)}, {ПОЛ_МЕТКА.get(пол, пол)}'

    def _ordered_groups():
        """Return [(метка, [команды])] in the competition's programme sequence."""
        # Build a map of all teams by (класс_судна, пол)
        from collections import defaultdict
        bucket = defaultdict(list)
        for к in команды:
            bucket[(к.класс_судна, к.пол)].append(к)
        for v in bucket.values():
            v.sort(key=lambda к: к.нагрудный_номер or 0)

        # Determine ordering from последовательность_видов
        seq_raw = (сор.последовательность_видов or '').strip()
        if seq_raw:
            # "R4-м,R4-ж,R6-м,R6-ж" → [(R4,м),(R4,ж),(R6,м),(R6,ж)]
            order = []
            for token in seq_raw.split(','):
                parts = token.strip().split('-')
                if len(parts) == 2:
                    order.append((parts[0], parts[1]))
            # Append any keys from bucket not mentioned in sequence
            for key in bucket:
                if key not in order:
                    order.append(key)
        else:
            order = sorted(bucket.keys())

        result = []
        for key in order:
            if key in bucket and bucket[key]:
                result.append((_section_label(*key), bucket[key]))
        return result

    # ── document builders ─────────────────────────────────────────────────────

    def _add_header(doc):
        """Organizers block + competition name + protocol title."""
        if сор.организаторы_блок:
            lines = [l.strip() for l in сор.организаторы_блок.split('\n')]
            p = _para(doc, align=C)
            for i, line in enumerate(lines):
                if i > 0:
                    _line_br(p)
                qi = next((line.index(q) for q in ('"', '«') if q in line), None)
                if qi is None:
                    _run(p, line, pt=10)
                else:
                    before  = line[:qi].strip()
                    quoted  = line[qi:]
                    combined = (before + ' ' + quoted) if before else quoted
                    if len(combined) <= 80:
                        _run(p, combined, pt=10)
                    else:
                        if before:
                            _run(p, before, pt=10)
                            _line_br(p)
                        _run(p, quoted, pt=10)

        _para(doc, сор.название or '', align=C, pt=12, bold=True)
        _para(doc, 'Протокол выдачи нагрудных номеров', align=C, pt=12, bold=True)

    def _add_table(doc, группы):
        """One combined table: info bar + header + section blocks + signatures."""
        tbl = doc.add_table(rows=0, cols=NC)
        _tbl_w(tbl)
        _set_grid(tbl)

        # Row 0 — date / venue values with bottom underline
        r = _add_row(tbl); _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[2]); _cw(cs[0], W_DATE)
        cs[3].merge(cs[5]); _cw(cs[3], W_MID)
        cs[6].merge(cs[9]); _cw(cs[6], W_VENUE)
        _bottom_border(cs[0]); _no_border(cs[3]); _bottom_border(cs[6])
        _fill(cs[0], сор.даты_проведения  or '', align=C, pt=10, bold=True)
        _fill(cs[3], '')
        _fill(cs[6], сор.место_проведения or '', align=C, pt=10, bold=True)

        # Row 1 — date / venue labels
        r = _add_row(tbl); _row_h(r, 0.5)
        cs = list(r.cells)
        cs[0].merge(cs[2]); _cw(cs[0], W_DATE)
        cs[3].merge(cs[5]); _cw(cs[3], W_MID)
        cs[6].merge(cs[9]); _cw(cs[6], W_VENUE)
        _no_border(cs[0]); _no_border(cs[3]); _no_border(cs[6])
        _fill(cs[0], 'Дата проведения',  align=C, pt=10)
        _fill(cs[3], '')
        _fill(cs[6], 'Место проведения', align=C, pt=10)

        # Row 2 — column headers
        r = _add_row(tbl)
        cs = list(r.cells)
        for i, w in enumerate(GRID):
            _cw(cs[i], w)
            _borders(cs[i])
            _valign(cs[i])
        # №п/п with line break
        p = cs[0].paragraphs[0]; p.clear(); _sp(p); p.alignment = C
        _run(p, '№', pt=10, bold=True); _line_br(p); _run(p, 'п/п', pt=10, bold=True)
        _fill(cs[1], 'Команда',                    align=C, pt=10, bold=True)
        _fill(cs[2], 'Территория',                 align=C, pt=10, bold=True)
        _fill(cs[3], 'Тренер-представитель',        align=C, pt=10, bold=True)
        _fill(cs[4], 'Выданы нагрудные\nномера',   align=C, pt=10, bold=True)
        _fill(cs[5], 'Получил,\nподпись',           align=C, pt=10, bold=True)
        _fill(cs[6], 'Сдан залог,\nруб.',           align=C, pt=10, bold=True)
        _fill(cs[7], 'Сданы нагрудные\nномера',    align=C, pt=10, bold=True)
        _fill(cs[8], 'Получил залог,\nруб.',        align=C, pt=10, bold=True)
        _fill(cs[9], 'Подпись',                    align=C, pt=10, bold=True)

        # Section groups
        for метка, кмнды in группы:
            # Section header row
            r = _add_row(tbl); _row_h(r, 0.6)
            cs = list(r.cells)
            cs[0].merge(cs[NC - 1]); _cw(cs[0], TW)
            _borders(cs[0])
            _fill(cs[0], метка, align=C, pt=10, bold=True)

            # Data rows
            for пп, к in enumerate(кмнды, 1):
                r = _add_row(tbl)
                cs = list(r.cells)
                for i, w in enumerate(GRID):
                    _cw(cs[i], w)
                    _borders(cs[i])
                    _valign(cs[i])
                _fill(cs[0], str(пп),                                              align=C)
                _fill(cs[1], к.название or '—')
                _fill(cs[2], к.субъект.название     if к.субъект else '—')
                _fill(cs[3], к.субъект.представитель if к.субъект and к.субъект.представитель else '')
                _fill(cs[4], '')
                _fill(cs[5], '')
                _fill(cs[6], '')
                _fill(cs[7], '')
                _fill(cs[8], '')
                _fill(cs[9], '')

        # Signature rows — no borders
        for pos, fio in (('Главный судья',      сор.главный_судья      or ''),
                         ('Главный секретарь',  сор.главный_секретарь  or '')):
            r = _add_row(tbl); _row_h(r, 1.2)
            cs = list(r.cells)
            cs[0].merge(cs[4]); _cw(cs[0], W_SIG_L)
            cs[5].merge(cs[9]); _cw(cs[5], W_SIG_R)
            _no_border(cs[0]); _no_border(cs[5])
            _fill(cs[0], pos, pt=10)
            _fill(cs[5], fio, pt=10)

    # ── document assembly ─────────────────────────────────────────────────────

    doc  = Document()
    sect = doc.sections[0]
    sect.page_width     = Mm(297)
    sect.page_height    = Mm(210)
    sect.left_margin    = Mm(19)
    sect.right_margin   = Mm(13)
    sect.top_margin     = Mm(19)
    sect.bottom_margin  = Mm(19)

    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        pgSz.set(qn('w:orient'), 'landscape')

    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    _add_header(doc)
    _add_table(doc, _ordered_groups())

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)
