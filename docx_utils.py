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


def _apply_no_break(doc):
    """Запретить разрыв строк таблиц при переносе на следующую страницу."""
    for tbl in doc.tables:
        for row in tbl.rows:
            trPr = row._tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = OxmlElement('w:trPr')
                row._tr.insert(0, trPr)
            if trPr.find(qn('w:cantSplit')) is None:
                trPr.append(OxmlElement('w:cantSplit'))


def _apply_keep_next(doc):
    """Объединённая секция-заголовок всегда остаётся со следующей строкой данных."""
    for tbl in doc.tables:
        rows = list(tbl.rows)
        for i, row in enumerate(rows[:-1]):
            cells = row.cells
            if len(set(id(c) for c in cells)) == 1:
                for p in cells[0].paragraphs:
                    p.paragraph_format.keep_with_next = True

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
    pgSz = sec._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        pgSz.set(qn('w:orient'), 'portrait' if portrait else 'landscape')
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
    _apply_no_break(doc)
    _apply_keep_next(doc)
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
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_старт_спринт_docx(сор, группы, filename, портрет=True):
    """Generate Sprint start protocol DOCX.

    группы — list of sections groups; each group produces one complete protocol
    block (header + table). Typically two groups: R4 and R6.
    Each group is a list of (метка, команды) pairs.
    """
    from docx.shared import Mm, Pt

    TNR   = 'Times New Roman'
    C     = WD_ALIGN_PARAGRAPH.CENTER
    # Page text-area width (twips) — exact, used for all three tables
    TW = round(((210 if портрет else 297) - 15 - 10) * 1440 / 25.4)

    # Main-table column proportions 4/19/6/37/13/13/8
    # ФИО (37 %) is split evenly across two grid columns (3 and 4)
    _PCTS = [5, 19, 7, 18.5, 18.5, 13, 11, 8]   # sum = 100
    GRID  = [round(TW * p / 100) for p in _PCTS]
    GRID[-1] += TW - sum(GRID)                    # fix rounding remainder
    NC    = len(GRID)

    W01   = GRID[0] + GRID[1]
    W234  = GRID[2] + GRID[3] + GRID[4]
    W567  = GRID[5] + GRID[6] + GRID[7]
    W34   = GRID[3] + GRID[4]
    W0123 = GRID[0] + GRID[1] + GRID[2] + GRID[3]
    W4567 = GRID[4] + GRID[5] + GRID[6] + GRID[7]

    # Header table (дата/место) columns — derived from GRID for consistency
    HDR_MID = W234
    HDR_LOC = W567

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


    def _no_wrap(cell):
        pr = _tcpr(cell)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))

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
        for s in tblPr.findall(qn('w:tblStyle')):
            tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')):
            tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None:
            el = OxmlElement('w:tblW')
            tblPr.append(el)
        el.set(qn('w:w'), str(TW))
        el.set(qn('w:type'), 'dxa')
        ind = tblPr.find(qn('w:tblInd'))
        if ind is None:
            ind = OxmlElement('w:tblInd')
            tblPr.append(ind)
        ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None:
            lay = OxmlElement('w:tblLayout')
            tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')

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

        def _ls15(p):
            pPr = p._p.get_or_add_pPr()
            sp = pPr.find(qn('w:spacing'))
            if sp is None:
                sp = OxmlElement('w:spacing'); pPr.append(sp)
            sp.set(qn('w:line'), '360'); sp.set(qn('w:lineRule'), 'auto')

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
            # Blank gap after org block — exact 4 pt height
            p_gap = doc.add_paragraph()
            pPr4 = p_gap._p.get_or_add_pPr()
            sp4 = OxmlElement('w:spacing')
            sp4.set(qn('w:before'), '0'); sp4.set(qn('w:after'), '0')
            sp4.set(qn('w:line'), '80');  sp4.set(qn('w:lineRule'), 'exact')
            pPr4.append(sp4)
            r4 = p_gap.add_run(); r4.font.size = Pt(4)

        p1 = _para(doc, сор.название or '', align=C, pt=12, bold=True);              _ls15(p1)
        p2 = _para(doc, 'Стартовый протокол', align=C, pt=12, bold=True);            _ls15(p2)
        p3 = _para(doc, 'Группа спортивных дисциплин СПРИНТ', align=C, pt=10, bold=True); _ls15(p3)

    def _make_tbl(ncols):
        t = doc.add_table(rows=0, cols=ncols)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        el.set(qn('w:w'), str(TW)); el.set(qn('w:type'), 'dxa')
        ind = tblPr.find(qn('w:tblInd'))
        if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
        ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        return t

    def _set_grid_sp(tbl_obj, widths):
        tbl = tbl_obj._tbl
        old = tbl.find(qn('w:tblGrid'))
        if old is not None: tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in widths:
            gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); grid.append(gc)
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is not None: tblPr.addnext(grid)
        else: tbl.insert(0, grid)

    def _tiny_para_sp():
        p = doc.add_paragraph()
        p.style = 'No Spacing'
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None: sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0')
        sp.set(qn('w:line'), '20'); sp.set(qn('w:lineRule'), 'exact')

    def _add_table(doc, секции):
        """Add the start protocol table for one group of sections."""

        # ── Header table (date/location) ──────────────────────────────────────
        # W_DT/W_VN фиксированы с noWrap — ширина по содержимому без переносов
        W_DT = 2500   # достаточно для любой даты
        W_VN = 4500   # достаточно для большинства мест
        W_SP = TW - W_DT - W_VN
        hdr = _make_tbl(3)
        _set_grid_sp(hdr, [W_DT, W_SP, W_VN])

        # Row 0 — date/location values
        r = add_row(hdr); _row_h(r, 0.6)
        cs = list(r.cells)
        _cw(cs[0], W_DT); _cw(cs[1], W_SP); _cw(cs[2], W_VN)
        _bottom_border(cs[0]); _no_border(cs[1]); _bottom_border(cs[2])
        _no_wrap(cs[0]); _no_wrap(cs[2])
        _fill(cs[0], сор.даты_проведения  or '', align=C, pt=10, bold=True)
        _fill(cs[1], '')
        _fill(cs[2], сор.место_проведения or '', align=C, pt=10, bold=True)

        # Row 1 — date/location labels
        r = add_row(hdr); _row_h(r, 0.6)
        cs = list(r.cells)
        _cw(cs[0], W_DT); _cw(cs[1], W_SP); _cw(cs[2], W_VN)
        _no_border(cs[0]); _no_border(cs[1]); _no_border(cs[2])
        _no_wrap(cs[0]); _no_wrap(cs[2])
        _fill(cs[0], 'Дата проведения',  align=C, pt=8, bold=True)
        _fill(cs[1], '')
        _fill(cs[2], 'Место проведения', align=C, pt=8, bold=True)

        _tiny_para_sp()

        # ── Data table (column headers + sections) ────────────────────────────
        tbl = _make_tbl(NC)
        _set_grid(tbl)

        # Column headers: full borders
        r = add_row(tbl)
        cs = list(r.cells)
        cs[3].merge(cs[4])
        for i, w in enumerate(GRID):
            if i not in (3, 4):
                _cw(cs[i], w)
        _cw(cs[3], W34)
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
                cs[3].merge(cs[4])
                for i, w in enumerate(GRID):
                    if i not in (3, 4):
                        _cw(cs[i], w)
                _cw(cs[3], W34)
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

        # ── Подписи — отдельная таблица 1:2 ─────────────────────────────────
        _tiny_para_sp()
        W_SIG_L = TW // 3
        W_SIG_R = TW - W_SIG_L
        sig = _make_tbl(2)
        _set_grid_sp(sig, [W_SIG_L, W_SIG_R])
        for pos, fio in (('Главный судья', сор.главный_судья or ''),
                         ('Главный секретарь', сор.главный_секретарь or '')):
            r = add_row(sig); _row_h(r, 1.2)
            cs = list(r.cells)
            _cw(cs[0], W_SIG_L); _cw(cs[1], W_SIG_R)
            _no_border(cs[0]); _no_border(cs[1])
            _fill(cs[0], pos, pt=10)
            _fill(cs[1], fio, pt=10)

    # ── document ──────────────────────────────────────────────────────────────

    doc  = Document()
    sect = doc.sections[0]
    if портрет:
        sect.page_height   = Mm(297); sect.page_width    = Mm(210)
    else:
        sect.page_height   = Mm(210); sect.page_width    = Mm(297)
    sect.top_margin    = Mm(12)
    sect.bottom_margin = Mm(15)
    sect.left_margin   = Mm(15)
    sect.right_margin  = Mm(10)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if портрет:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')

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
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_старт_слалом_docx(сор, секции, filename, портрет=True):
    """Стартовый протокол Слалом — по форме Протокол_сл_старт.docx.

    секции: list of (метка, команды, мб_колонки).
      мб_колонки: list of {'заг': label, 'поле': field_name} — 0, 1 или 2 записи.
    Одна таблица на секцию (дата/место + шапка + данные + подписи).
    """
    from docx.shared import Mm, Pt

    TNR = 'Times New Roman'
    C = WD_ALIGN_PARAGRAPH.CENTER

    # ── Сетка столбцов 5/12/6/33/11/4/4/4/9/9 ───────────────────────────────
    _PCTS = [5, 12, 6, 33, 11, 4, 4, 4, 9, 9]   # sum=97
    _SUM_P = sum(_PCTS)
    # Текстовая область: 210/297 мм − 7.5 мм − 5 мм
    TW = round(((210 if портрет else 297) - 7.5 - 5) * 1440 / 25.4)
    GRID = [round(TW * p / _SUM_P) for p in _PCTS]
    GRID[-1] += TW - sum(GRID)
    NC = len(GRID)   # 10

    # Объединённые ширины
    W_МБ   = GRID[5] + GRID[6] + GRID[7]   # «Очки Многоборье» cols 5-7

    # ── Хелперы ──────────────────────────────────────────────────────────────

    def _sp(p):
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def _run(p, text, pt=8, bold=False):
        r = p.add_run(text)
        r.font.name = TNR; r.font.size = Pt(pt)
        if bold: r.font.bold = True
        return r

    def _para(doc, text='', align=None, pt=10, bold=False):
        p = doc.add_paragraph(); _sp(p)
        if align: p.alignment = align
        if text: _run(p, text, pt, bold)
        return p

    def _line_br(p):
        p.add_run()._r.append(OxmlElement('w:br'))

    def _page_br(doc):
        p = doc.add_paragraph(); _sp(p)
        br = OxmlElement('w:br'); br.set(qn('w:type'), 'page')
        p.add_run()._r.append(br)

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _borders(cell, sz=4, color='333333'):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')): pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0');    el.set(qn('w:color'), color)
            b.append(el)
        pr.append(b)

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')): pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0');  el.set(qn('w:color'), 'auto')
            b.append(el)
        bot = OxmlElement('w:bottom')
        bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), str(sz))
        bot.set(qn('w:space'), '0');    bot.set(qn('w:color'), 'auto')
        b.append(bot)
        pr.append(b)

    def _no_border(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')): pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0');  el.set(qn('w:color'), 'auto')
            b.append(el)
        pr.append(b)

    def _cw(cell, w):
        pr = _tcpr(cell)
        el = pr.find(qn('w:tcW'))
        if el is None: el = OxmlElement('w:tcW'); pr.append(el)
        el.set(qn('w:w'), str(w)); el.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        v = pr.find(qn('w:vAlign'))
        if v is None: v = OxmlElement('w:vAlign'); pr.append(v)
        v.set(qn('w:val'), val)

    def _fill(cell, text, align=None, pt=8, bold=False):
        p = cell.paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p)
        if align: p.alignment = align
        _valign(cell)
        if text: _run(p, text, pt, bold)

    def _row_h(row, cm):
        val = round(cm * 1440 / 2.54)
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None: trPr = OxmlElement('w:trPr'); row._tr.insert(0, trPr)
        trH = trPr.find(qn('w:trHeight'))
        if trH is None: trH = OxmlElement('w:trHeight'); trPr.append(trH)
        trH.set(qn('w:val'), str(val)); trH.set(qn('w:hRule'), 'exact')

    def _row_h_min(row, cm):
        val = round(cm * 1440 / 2.54)
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None: trPr = OxmlElement('w:trPr'); row._tr.insert(0, trPr)
        trH = trPr.find(qn('w:trHeight'))
        if trH is None: trH = OxmlElement('w:trHeight'); trPr.append(trH)
        trH.set(qn('w:val'), str(val)); trH.set(qn('w:hRule'), 'atLeast')

    def _vmerge_restart(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); vm.set(qn('w:val'), 'restart'); pr.append(vm)

    def _vmerge_cont(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); pr.append(vm)
        for para in cell.paragraphs: para.clear()

    def _vert_cell(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:textDirection')): pr.remove(old)
        td = OxmlElement('w:textDirection'); td.set(qn('w:val'), 'btLr')
        v = pr.find(qn('w:vAlign'))
        if v is not None: v.addprevious(td)
        else: pr.append(td)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))
        _valign(cell, 'center')
        for p in cell.paragraphs: p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    def _tbl_header_row(row):
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None: trPr = OxmlElement('w:trPr'); row._tr.insert(0, trPr)
        if trPr.find(qn('w:tblHeader')) is None:
            trPr.append(OxmlElement('w:tblHeader'))

    def _no_wrap(cell):
        pr = _tcpr(cell)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))

    def _auto_w(cell):
        pr = _tcpr(cell)
        el = pr.find(qn('w:tcW'))
        if el is None: el = OxmlElement('w:tcW'); pr.append(el)
        el.set(qn('w:w'), '0'); el.set(qn('w:type'), 'auto')

    # ── Скор/время ───────────────────────────────────────────────────────────

    def _fmt_score(v):
        if v is None or v == 0:
            return ''
        return str(int(v)) if float(v) == int(float(v)) else f'{v:.1f}'

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
        total = 0.0; has = False
        for кол in колонки:
            s = _score(к, кол['поле'])
            if s:
                try: total += float(s); has = True
                except ValueError: pass
        return _fmt_score(total) if has else ''

    def _time(к, попытка):
        рез = next((р for р in к.результат_слалом if р.номер_попытки == попытка), None)
        if рез and рез.результат_попытки is not None:
            s = float(рез.результат_попытки)
            m = int(s // 60); sec = s % 60
            return f'{m:02d}:{sec:05.2f}'
        return ''

    # ── Создать таблицу (fixed, TW, tblInd=0) ────────────────────────────────

    def _make_tbl():
        t = doc.add_table(rows=0, cols=NC)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None: tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        el.set(qn('w:w'), str(TW)); el.set(qn('w:type'), 'dxa')
        ind = tblPr.find(qn('w:tblInd'))
        if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
        ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        # tblGrid
        for old in tbl.findall(qn('w:tblGrid')): tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in GRID:
            gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); grid.append(gc)
        tblPr.addnext(grid)
        return t

    # ── Заголовочные параграфы (организаторы + название + "Стартовый протокол") ──

    def _add_header_paras():
        def _sp15(p):
            pPr = p._p.get_or_add_pPr()
            sp = pPr.find(qn('w:spacing'))
            if sp is None: sp = OxmlElement('w:spacing'); pPr.append(sp)
            sp.set(qn('w:line'), '276'); sp.set(qn('w:lineRule'), 'auto')
            sp.set(qn('w:after'), '200')

        if сор.организаторы_блок:
            p = _para(doc, align=C, pt=10, bold=True)
            lines = [l.strip() for l in сор.организаторы_блок.split('\n')]
            for i, line in enumerate(lines):
                if i > 0: _line_br(p)
                _run(p, line, pt=10, bold=True)
            _sp15(p)

        p = _para(doc, сор.название or '', align=C, pt=11, bold=True); _sp15(p)
        p = _para(doc, 'Стартовый протокол', align=C, pt=11, bold=True); _sp15(p)
        p = _para(doc, 'Группа спортивных дисциплин СЛАЛОМ', align=C, pt=10, bold=True); _sp15(p)

    # ── Разделитель 1 пт между таблицами ────────────────────────────────────

    def _tiny_para():
        p = doc.add_paragraph(); p.style = 'No Spacing'
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None: sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0')
        sp.set(qn('w:line'), '20'); sp.set(qn('w:lineRule'), 'exact')

    # ── Таблица дата/место (autofit, 3 cols) ─────────────────────────────────

    def _make_dv_tbl():
        t = doc.add_table(rows=0, cols=3)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None: tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        el.set(qn('w:w'), str(TW)); el.set(qn('w:type'), 'dxa')
        ind = tblPr.find(qn('w:tblInd'))
        if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
        ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'autofit')
        return t

    # ── Таблица подписей (fixed, 2 cols 1:2) ─────────────────────────────────

    def _make_sig_tbl():
        W_L = TW // 3; W_R = TW - W_L
        t = doc.add_table(rows=0, cols=2)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None: tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        el.set(qn('w:w'), str(TW)); el.set(qn('w:type'), 'dxa')
        ind = tblPr.find(qn('w:tblInd'))
        if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
        ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        for old in tbl.findall(qn('w:tblGrid')): tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in (W_L, W_R):
            gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); grid.append(gc)
        tblPr.addnext(grid)
        return t, W_L, W_R

    # ── Построить одну секцию ─────────────────────────────────────────────────

    def _build_section(метка, команды, мб_колонки):
        кол1_заг  = мб_колонки[0]['заг']        if len(мб_колонки) > 0 else ''
        кол2_заг  = мб_колонки[1]['заг']        if len(мб_колонки) > 1 else ''
        кол1_поле = мб_колонки[0].get('поле')   if len(мб_колонки) > 0 else None
        кол2_поле = мб_колонки[1].get('поле')   if len(мб_колонки) > 1 else None

        # ── Таблица 1: дата / место ──────────────────────────────────────────
        dv = _make_dv_tbl()
        r = dv.add_row(); _row_h(r, 0.6)
        cs = list(r.cells)
        _bottom_border(cs[0]); _no_border(cs[1]); _bottom_border(cs[2])
        _no_wrap(cs[0]); _auto_w(cs[0])
        _no_wrap(cs[2]); _auto_w(cs[2])
        _fill(cs[0], сор.даты_проведения  or '', align=C, pt=10, bold=True)
        _fill(cs[1], '')
        _fill(cs[2], сор.место_проведения or '', align=C, pt=10, bold=True)

        r = dv.add_row(); _row_h(r, 0.5)
        cs = list(r.cells)
        _no_border(cs[0]); _no_border(cs[1]); _no_border(cs[2])
        _no_wrap(cs[0]); _auto_w(cs[0])
        _no_wrap(cs[2]); _auto_w(cs[2])
        _fill(cs[0], 'Дата проведения',  align=C, pt=8, bold=True)
        _fill(cs[1], '')
        _fill(cs[2], 'Место проведения', align=C, pt=8, bold=True)

        _tiny_para()

        # ── Таблица 2: шапка + данные ────────────────────────────────────────
        tbl = _make_tbl()

        # Шапка строка 1
        r = tbl.add_row()
        _tbl_header_row(r)
        cs = list(r.cells)
        cs[5].merge(cs[7]); _cw(cs[5], W_МБ)
        for i, w in enumerate(GRID):
            if i not in (5, 6, 7):
                _cw(cs[i], w)
        for cell in [cs[0], cs[1], cs[2], cs[3], cs[4], cs[8], cs[9]]:
            _borders(cell); _vmerge_restart(cell)
        _borders(cs[5])
        p0 = cs[0].paragraphs[0]; p0.clear(); p0.style = 'No Spacing'; _sp(p0)
        p0.alignment = C; _valign(cs[0])
        _run(p0, '№', pt=8, bold=True); _line_br(p0); _run(p0, 'п/п', pt=8, bold=True)
        _fill(cs[1], 'Команда',                            align=C, pt=8, bold=True)
        _fill(cs[2], 'Номер',                              align=C, pt=8, bold=True)
        _fill(cs[3], 'Фамилия, имя, отчество участников',  align=C, pt=8, bold=True)
        _fill(cs[4], 'Территория',                         align=C, pt=8, bold=True)
        _fill(cs[5], 'Очки Многоборье',                    align=C, pt=8, bold=True)
        _fill(cs[8], 'Время\nстарта\nпопытка 1',           align=C, pt=8, bold=True)
        _fill(cs[9], 'Время\nстарта\nпопытка 2',           align=C, pt=8, bold=True)

        # Шапка строка 2 — вертикальный текст МБ
        r = tbl.add_row(); _row_h_min(r, 1690 * 2.54 / 1440)
        _tbl_header_row(r)
        cs = list(r.cells)
        for i, w in enumerate(GRID):
            _cw(cs[i], w)
        for cell in [cs[0], cs[1], cs[2], cs[3], cs[4], cs[8], cs[9]]:
            _borders(cell); _vmerge_cont(cell)
        for col_i, заг in ((5, кол1_заг), (6, кол2_заг), (7, 'Сумма')):
            _borders(cs[col_i])
            _fill(cs[col_i], заг, pt=7, bold=True)
            _vert_cell(cs[col_i])

        # Строка-заголовок секции
        r = tbl.add_row(); _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[NC - 1]); _cw(cs[0], TW)
        _borders(cs[0])
        _fill(cs[0], метка, align=C, pt=9, bold=True)

        # Строки данных
        for пп, к in enumerate(команды, 1):
            r = tbl.add_row()
            cs = list(r.cells)
            for i, w in enumerate(GRID):
                _cw(cs[i], w)
            for cell in cs:
                _borders(cell)
            спортсмены = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
            фио = ', '.join(с.фамилия_имя_отчество for с in спортсмены if с.фамилия_имя_отчество)
            s1 = _score(к, кол1_поле) if кол1_поле else ''
            s2 = _score(к, кол2_поле) if кол2_поле else ''
            sm = _sum_scores(к, мб_колонки)
            _fill(cs[0], str(пп),                                       align=C, pt=8)
            _fill(cs[1], к.название or '—',                                      pt=8)
            _fill(cs[2], str(к.нагрудный_номер or ''),                  align=C, pt=8)
            _fill(cs[3], фио,                                                     pt=8)
            _fill(cs[4], к.субъект.название if к.субъект else '—',               pt=8)
            _fill(cs[5], s1,                                             align=C, pt=8)
            _fill(cs[6], s2,                                             align=C, pt=8)
            _fill(cs[7], sm,                                             align=C, pt=8)
            _fill(cs[8], _time(к, 1),                                    align=C, pt=8)
            _fill(cs[9], _time(к, 2),                                    align=C, pt=8)

        _tiny_para()

        # ── Таблица 3: подписи (1:2) ─────────────────────────────────────────
        sig, W_L, W_R = _make_sig_tbl()
        for pos, fio in (('Главный судья',     сор.главный_судья     or ''),
                         ('Главный секретарь', сор.главный_секретарь or '')):
            r = sig.add_row(); _row_h_min(r, 1.2)
            cs = list(r.cells)
            _cw(cs[0], W_L); _cw(cs[1], W_R)
            _no_border(cs[0]); _no_border(cs[1])
            _fill(cs[0], pos, pt=10)
            _fill(cs[1], fio, pt=10)

    # ── Document ──────────────────────────────────────────────────────────────

    doc = Document()
    sect = doc.sections[0]
    if портрет:
        sect.page_width = Mm(210); sect.page_height = Mm(297)
    else:
        sect.page_width = Mm(297); sect.page_height = Mm(210)
    sect.top_margin    = Mm(18)
    sect.bottom_margin = Mm(12)
    sect.left_margin   = Mm(7.5)
    sect.right_margin  = Mm(5)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if портрет:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')

    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'): body.remove(child)

    first = True
    for метка, команды, мб_колонки in секции:
        if not команды:
            continue
        if not first:
            _page_br(doc)
        first = False
        _add_header_paras()
        _build_section(метка, команды, мб_колонки)

    buf = BytesIO()
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_старт_дг_docx(сор, секции, filename, портрет=True):
    """Стартовый протокол Длинной гонки — динамическая сетка, все МБ-колонки, Очередь.

    Реализован по образцу Протокол_дг_старт.docx.
    Страница: A4 портрет, left≈7.5 мм, right≈5 мм, top≈18 мм, bottom≈12 мм.
    TW = 11196 dxa (usable width).

    секции: list of (метка, команды, мб_колонки)
      мб_колонки: list of {'заг': label, 'поле': field_name|None}
        поле = 'очки_спринт' | 'очки_пс' | 'очки_слалом' | None (→ Сумма)
    """
    from docx.shared import Mm, Pt

    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER

    # Portrait A4: left=7.514mm, right=5.009mm → TW=11196 dxa
    # Landscape A4 with same proportional margins → TW=16157 dxa
    if портрет:
        TW = 11196
    else:
        TW = round((297 - 7.514 - 5.009) * 1440 / 25.4)

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
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), '000000')

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')
        el = b.find(qn('w:bottom'))
        if el is None:
            el = OxmlElement('w:bottom'); b.append(el)
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), '000000')

    def _no_border(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')

    def _cw(cell, w):
        pr = _tcpr(cell)
        cw = pr.find(qn('w:tcW'))
        if cw is None:
            cw = OxmlElement('w:tcW'); pr.append(cw)
        cw.set(qn('w:w'), str(w))
        cw.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        va = pr.find(qn('w:vAlign'))
        if va is None:
            va = OxmlElement('w:vAlign'); pr.append(va)
        va.set(qn('w:val'), val)

    def _fill(cell, text, align=None, pt=8, bold=False):
        for p in cell.paragraphs:
            p.clear()
        p = cell.paragraphs[0]
        p.style = 'No Spacing'
        _sp(p)
        if align:
            p.alignment = align
        _valign(cell)
        if text:
            _run(p, text, pt, bold)

    def _set_grid(tbl, grid_list):
        tbl_el = tbl._tbl
        old = tbl_el.find(qn('w:tblGrid'))
        if old is not None:
            tbl_el.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in grid_list:
            gc = OxmlElement('w:gridCol')
            gc.set(qn('w:w'), str(w))
            grid.append(gc)
        tblPr = tbl_el.find(qn('w:tblPr'))
        if tblPr is not None:
            tblPr.addnext(grid)
        else:
            tbl_el.insert(0, grid)

    def _tbl_w(tbl_obj):
        tbl = tbl_obj._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        el = tblPr.find(qn('w:tblW'))
        if el is None:
            el = OxmlElement('w:tblW'); tblPr.append(el)
        el.set(qn('w:w'), str(TW)); el.set(qn('w:type'), 'dxa')
        ind = tblPr.find(qn('w:tblInd'))
        if ind is None:
            ind = OxmlElement('w:tblInd'); tblPr.append(ind)
        ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
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
            trPr = OxmlElement('w:trPr'); row._tr.insert(0, trPr)
        trH = trPr.find(qn('w:trHeight'))
        if trH is None:
            trH = OxmlElement('w:trHeight'); trPr.append(trH)
        trH.set(qn('w:val'), str(val)); trH.set(qn('w:hRule'), 'exact')

    def _row_h_min(row, cm):
        val = round(cm * 1440 / 2.54)
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr'); row._tr.insert(0, trPr)
        trH = trPr.find(qn('w:trHeight'))
        if trH is None:
            trH = OxmlElement('w:trHeight'); trPr.append(trH)
        trH.set(qn('w:val'), str(val)); trH.set(qn('w:hRule'), 'atLeast')

    def border_row(row, sz=4):
        for cell in row.cells:
            _borders(cell, sz)

    def _vmerge_restart(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); vm.set(qn('w:val'), 'restart'); pr.append(vm)

    def _vmerge_cont(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); pr.append(vm)
        for para in cell.paragraphs: para.clear()

    def _vert_cell(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:textDirection')): pr.remove(old)
        td = OxmlElement('w:textDirection'); td.set(qn('w:val'), 'btLr')
        va = pr.find(qn('w:vAlign'))
        if va is not None:
            va.addprevious(td)
        else:
            pr.append(td)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))
        _valign(cell, 'center')
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    def _fmt_score(v):
        if v is None or v == 0:
            return ''
        return str(int(v)) if float(v) == int(v) else f'{v:.1f}'

    def _score(к, поле):
        рм = к.результат_многоборья
        if поле == 'очки_спринт':
            return _fmt_score(float(рм.очки_спринт or 0) if рм else None)
        elif поле == 'очки_пс':
            return _fmt_score(float(рм.очки_пс     or 0) if рм else None)
        elif поле == 'очки_слалом':
            return _fmt_score(float(рм.очки_слалом or 0) if рм else None)
        return ''

    def _sum_scores(к, колонки):
        total = 0.0; has = False
        for кол in колонки:
            if not кол.get('поле'):
                continue
            s = _score(к, кол['поле'])
            if s:
                try:
                    total += float(s); has = True
                except ValueError:
                    pass
        return _fmt_score(total) if has else ''

    def _time_старт(к):
        рез = к.результат_дг
        if рез and рез.время_старта is not None:
            s = float(рез.время_старта)
            h = int(s // 3600)
            m = int((s % 3600) // 60)
            sec = s % 60
            return f'{h:02d}:{m:02d}:{int(sec):02d}'
        return ''

    def _page_br(doc):
        p = doc.add_paragraph()
        _sp(p)
        p.style = 'No Spacing'
        r = p.add_run()
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        r._r.append(br)

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

        def _spacing_15(p):
            pPr = p._p.get_or_add_pPr()
            sp = pPr.find(qn('w:spacing'))
            if sp is None:
                sp = OxmlElement('w:spacing'); pPr.append(sp)
            sp.set(qn('w:line'), '360')
            sp.set(qn('w:lineRule'), 'auto')

        p4 = doc.add_paragraph(); p4.style = 'No Spacing'; _sp(p4); p4.add_run().font.size = Pt(4)
        p = _para(doc, сор.название or '', align=C, pt=12, bold=True); _spacing_15(p)
        p = _para(doc, 'Стартовый протокол', align=C, pt=12, bold=True); _spacing_15(p)
        p = _para(doc, 'Группа спортивных дисциплин ДЛИННАЯ ГОНКА', align=C, pt=10, bold=True); _spacing_15(p)

    def _add_table(doc, метка, команды, мб_колонки):
        """Одна секция (группа): дата/место + шапка + данные + подписи.

        Структура столбцов по образцу Протокол_дг_старт.docx:
          N=1: 10 cols, один широкий МБ-столбец (1038 dxa), заголовок «Очки Многоборье [дисц]»
          N>1: 5+N+2 cols, МБ-столбцы по 452 dxa каждый, btLr-вертикальный текст в строке 2
        """
        N = len(мб_колонки)

        # ── Фиксированные ширины (dxa) ─────────────────────────────────────────
        W_НПП = 410   # №п/п
        W_КОМ = 1580  # Команда
        W_НОМ = 699   # Номер (объединение двух подколонок 137+562)
        W_ОЧ  = 907   # Очередь выбора позиции
        W_ВС  = 915   # Время старта
        W_МБ  = round(TW * 5 / 100)  # 5% от TW ≈ 560 dxa — ширина одного МБ-столбца (btLr) при N>1

        if N == 1:
            # Один широкий МБ-столбец, Территория=1390(2cols), ФИО занимает остаток
            W_ТЕРР = 1390
            W_МБ1  = 1150   # единственный МБ-столбец широкий (+1% от TW)
            W_ФИО  = TW - W_НПП - W_КОМ - W_НОМ - W_ТЕРР - W_МБ1 - W_ОЧ - W_ВС
            # Разбивка для tblGrid (Номер→2 подкол, Территория→2 подкол)
            GRID = [W_НПП, W_КОМ, 137, 562, W_ФИО, 567, 823, W_МБ1, W_ОЧ, W_ВС]
            NC = len(GRID)
            I_НПП = 0; I_КОМ = 1; I_НОМ = 2; I_ФИО = 4; I_ТЕРР = 5
            I_МБ0 = 7; I_ОЧ  = 8; I_ВС  = 9
            # Индексы ячеек с объединением (gridSpan) в шапке
            I_НОМ_END  = 3   # Номер = cols 2-3
            I_ТЕРР_END = 6   # Территория = cols 5-6
        else:
            # N>1: МБ-столбцы по 452 каждый; Территория=1502(2cols) при N=3, или 1439(1col) при N=4+
            # По образцу: при N=3 Территория=1502 (2cols 997+505), ФИО=3827
            #             при N=4 Территория=1439 (1col), ФИО=3449, Команда разбита на 2 подкол
            # Для общего случая: используем 2-подкол Территорию при N≤3, 1-кол при N>3
            if N <= 3:
                W_ТЕРР = 1502
                W_ФИО  = TW - W_НПП - W_КОМ - W_НОМ - W_ТЕРР - N * W_МБ - W_ОЧ - W_ВС
                GRID = [W_НПП, W_КОМ, 137, 562, W_ФИО, 997, 505] + [W_МБ] * N + [W_ОЧ, W_ВС]
                NC = len(GRID)
                I_НПП = 0; I_КОМ = 1; I_НОМ = 2; I_ФИО = 4; I_ТЕРР = 5
                I_МБ0 = 7; I_ОЧ  = 7 + N; I_ВС  = 8 + N
                I_НОМ_END  = 3
                I_ТЕРР_END = 6
            else:
                # N=4+: простая структура, пропорции 5/12/6/34/11/4×N/6/9
                _PCTS = [5, 12, 7, 30, 11] + [5] * N + [8, 7]
                _SUM_P = sum(_PCTS)
                GRID = [round(TW * p / _SUM_P) for p in _PCTS]
                GRID[-1] += TW - sum(GRID)
                NC = len(GRID)   # 5+N+2
                I_НПП = 0; I_КОМ = 1; I_НОМ = 2; I_ФИО = 3; I_ТЕРР = 4
                I_МБ0 = 5; I_ОЧ  = 5 + N; I_ВС  = 6 + N
                I_НОМ_END  = None   # одиночная колонка
                I_ТЕРР_END = None   # одиночная колонка
                W_НПП  = GRID[0]; W_КОМ  = GRID[1]; W_НОМ = GRID[2]
                W_ФИО  = GRID[3]; W_ТЕРР = GRID[4]; W_МБ  = GRID[5]
                W_ОЧ   = GRID[I_ОЧ]; W_ВС = GRID[I_ВС]

        def _no_wrap(cell):
            pr = _tcpr(cell)
            if pr.find(qn('w:noWrap')) is None:
                pr.append(OxmlElement('w:noWrap'))

        def _make_tbl(ncols):
            """Создать fixed-таблицу TW dxa, без унаследованного стиля."""
            t = doc.add_table(rows=0, cols=ncols)
            tbl = t._tbl
            tblPr = tbl.find(qn('w:tblPr'))
            if tblPr is None:
                tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
            ts = tblPr.find(qn('w:tblStyle'))
            if ts is not None:
                tblPr.remove(ts)
            for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
            el = tblPr.find(qn('w:tblW'))
            if el is None:
                el = OxmlElement('w:tblW'); tblPr.append(el)
            el.set(qn('w:w'), str(TW)); el.set(qn('w:type'), 'dxa')
            ind = tblPr.find(qn('w:tblInd'))
            if ind is None:
                ind = OxmlElement('w:tblInd'); tblPr.append(ind)
            ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
            lay = tblPr.find(qn('w:tblLayout'))
            if lay is None:
                lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
            lay.set(qn('w:type'), 'fixed')
            return t

        def _set_tbl_grid(tbl_obj, widths):
            tbl = tbl_obj._tbl
            old = tbl.find(qn('w:tblGrid'))
            if old is not None: tbl.remove(old)
            grid = OxmlElement('w:tblGrid')
            for w in widths:
                gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); grid.append(gc)
            tblPr = tbl.find(qn('w:tblPr'))
            if tblPr is not None:
                tblPr.addnext(grid)
            else:
                tbl.insert(0, grid)

        def _tiny_para(doc):
            """1 pt разделитель между таблицами."""
            p = doc.add_paragraph()
            p.style = 'No Spacing'
            pPr = p._p.get_or_add_pPr()
            sp = pPr.find(qn('w:spacing'))
            if sp is None:
                sp = OxmlElement('w:spacing'); pPr.append(sp)
            sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0')
            sp.set(qn('w:line'), '20'); sp.set(qn('w:lineRule'), 'exact')

        def _tbl_hdr_row(row):
            """Пометить строку как tblHeader (повтор на каждой странице)."""
            trPr = row._tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = OxmlElement('w:trPr'); row._tr.insert(0, trPr)
            if trPr.find(qn('w:tblHeader')) is None:
                th = OxmlElement('w:tblHeader'); trPr.append(th)

        def _mb_margins(cell, lr=108):
            """Узкие отступы для btLr-ячеек МБ."""
            pr = _tcpr(cell)
            for old in pr.findall(qn('w:tcMar')): pr.remove(old)
            mar = OxmlElement('w:tcMar')
            for side in ('left', 'right'):
                el = OxmlElement(f'w:{side}')
                el.set(qn('w:w'), str(lr)); el.set(qn('w:type'), 'dxa')
                mar.append(el)
            pr.append(mar)

        def _hspan(cells, start, end):
            """Горизонтально объединить cells[start..end] через gridSpan (удалить satellite TC)."""
            c0 = cells[start]
            span = end - start + 1
            if span < 2:
                return
            pr = _tcpr(c0)
            gs = pr.find(qn('w:gridSpan'))
            if gs is None: gs = OxmlElement('w:gridSpan'); pr.append(gs)
            gs.set(qn('w:val'), str(span))
            for ci in range(start + 1, end + 1):
                tc = cells[ci]._tc
                parent = tc.getparent()
                if parent is not None:
                    parent.remove(tc)

        def _merge_span(cells, start, end, width):
            """Объединить ячейки start..end через gridSpan и установить ширину."""
            c0 = cells[start]
            tc0 = c0._tc
            tcPr = _tcpr(c0)
            span = end - start + 1
            if span > 1:
                gs = tcPr.find(qn('w:gridSpan'))
                if gs is None:
                    gs = OxmlElement('w:gridSpan'); tcPr.append(gs)
                gs.set(qn('w:val'), str(span))
                # Удалить лишние tc
                for ci in range(start + 1, end + 1):
                    tc = cells[ci]._tc
                    tc0.addnext(tc)   # сдвинуть к базовой tc
                    tc0.getparent().remove(tc)
            _cw(c0, width)

        # ── Таблица 1: Дата / Место (autofit, min-width ячейки, noWrap) ──────────

        def _make_dv_tbl():
            t = doc.add_table(rows=0, cols=3)
            tbl = t._tbl
            tblPr = tbl.find(qn('w:tblPr'))
            if tblPr is None: tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
            for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
            for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
            el = tblPr.find(qn('w:tblW'))
            if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
            el.set(qn('w:w'), str(TW)); el.set(qn('w:type'), 'dxa')
            ind = tblPr.find(qn('w:tblInd'))
            if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
            ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
            lay = tblPr.find(qn('w:tblLayout'))
            if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
            lay.set(qn('w:type'), 'autofit')
            return t

        def _auto_w(cell):
            pr = _tcpr(cell)
            el = pr.find(qn('w:tcW'))
            if el is None: el = OxmlElement('w:tcW'); pr.append(el)
            el.set(qn('w:w'), '0'); el.set(qn('w:type'), 'auto')

        hdr_tbl = _make_dv_tbl()

        r = hdr_tbl.add_row(); _row_h(r, 0.6)
        cs = list(r.cells)
        _bottom_border(cs[0]); _no_border(cs[1]); _bottom_border(cs[2])
        _no_wrap(cs[0]); _auto_w(cs[0])
        _no_wrap(cs[2]); _auto_w(cs[2])
        _fill(cs[0], сор.даты_проведения  or '', align=C, pt=10, bold=True)
        _fill(cs[1], '')
        _fill(cs[2], сор.место_проведения or '', align=C, pt=10, bold=True)

        r = hdr_tbl.add_row(); _row_h(r, 0.5)
        cs = list(r.cells)
        _no_border(cs[0]); _no_border(cs[1]); _no_border(cs[2])
        _no_wrap(cs[0]); _auto_w(cs[0])
        _no_wrap(cs[2]); _auto_w(cs[2])
        _fill(cs[0], 'Дата проведения',  align=C, pt=8, bold=True)
        _fill(cs[1], '')
        _fill(cs[2], 'Место проведения', align=C, pt=8, bold=True)

        _tiny_para(doc)

        # ── Таблица 2: шапка + секционная строка + данные ─────────────────────
        тbl = _make_tbl(NC)
        _set_tbl_grid(тbl, GRID)

        def _fill_fixed_hdr_cells(cs, with_borders=True, vmerge_mode='restart'):
            """Заполнить фиксированные ячейки шапки (№п/п, Команда, Номер, ФИО, Терр, Оч, Вс)."""
            if with_borders:
                for i in range(NC): _borders(cs[i], sz=6)

            # №п/п
            p = cs[I_НПП].paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p)
            p.alignment = C; _valign(cs[I_НПП])
            _cw(cs[I_НПП], W_НПП)
            if vmerge_mode == 'restart':
                _vmerge_restart(cs[I_НПП])
                _run(p, '№', pt=8, bold=True); _line_br(p); _run(p, 'п/п', pt=8, bold=True)
            else:
                _vmerge_cont(cs[I_НПП])

            _cw(cs[I_КОМ], W_КОМ)
            if vmerge_mode == 'restart':
                _vmerge_restart(cs[I_КОМ])
                _fill(cs[I_КОМ], 'Команда', align=C, pt=8, bold=True)
            else:
                _vmerge_cont(cs[I_КОМ])

            # Номер (span 2 cols если есть субколонка, иначе одиночная)
            if I_НОМ_END is not None:
                _hspan(cs, I_НОМ, I_НОМ_END)
            _cw(cs[I_НОМ], W_НОМ)
            if vmerge_mode == 'restart':
                _vmerge_restart(cs[I_НОМ])
                _fill(cs[I_НОМ], 'Номер', align=C, pt=8, bold=True)
            else:
                _vmerge_cont(cs[I_НОМ])

            # ФИО
            _cw(cs[I_ФИО], W_ФИО)
            if vmerge_mode == 'restart':
                _vmerge_restart(cs[I_ФИО])
                _fill(cs[I_ФИО], 'Фамилия, имя, отчество участников', align=C, pt=8, bold=True)
            else:
                _vmerge_cont(cs[I_ФИО])

            # Территория (span 2 cols или 1 col)
            if I_ТЕРР_END is not None:
                _hspan(cs, I_ТЕРР, I_ТЕРР_END)
            _cw(cs[I_ТЕРР], W_ТЕРР)
            if vmerge_mode == 'restart':
                _vmerge_restart(cs[I_ТЕРР])
                _fill(cs[I_ТЕРР], 'Территория', align=C, pt=8, bold=True)
            else:
                _vmerge_cont(cs[I_ТЕРР])

            # Очередь выбора позиции
            _cw(cs[I_ОЧ], W_ОЧ)
            if vmerge_mode == 'restart':
                _vmerge_restart(cs[I_ОЧ])
                p2 = cs[I_ОЧ].paragraphs[0]; p2.clear(); p2.style = 'No Spacing'; _sp(p2)
                p2.alignment = C; _valign(cs[I_ОЧ])
                _run(p2, 'Очередь', pt=8, bold=True); _line_br(p2)
                _run(p2, ' выбора', pt=8, bold=True); _line_br(p2)
                _run(p2, ' позиции', pt=8, bold=True)
            else:
                _vmerge_cont(cs[I_ОЧ])

            # Время старта
            _cw(cs[I_ВС], W_ВС)
            if vmerge_mode == 'restart':
                _vmerge_restart(cs[I_ВС])
                p3 = cs[I_ВС].paragraphs[0]; p3.clear(); p3.style = 'No Spacing'; _sp(p3)
                p3.alignment = C; _valign(cs[I_ВС])
                _run(p3, 'Время', pt=8, bold=True); _line_br(p3)
                _run(p3, 'старта', pt=8, bold=True)
            else:
                _vmerge_cont(cs[I_ВС])

        if N == 1:
            # Одна строка заголовка (нет второй строки vMerge)
            r = тbl.add_row(); cs = list(r.cells)
            _tbl_hdr_row(r)
            for i, w in enumerate(GRID): _cw(cs[i], w)
            for i in range(NC): _borders(cs[i], sz=6)

            p = cs[I_НПП].paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p)
            p.alignment = C; _valign(cs[I_НПП])
            _cw(cs[I_НПП], W_НПП)
            _run(p, '№', pt=8, bold=True); _line_br(p); _run(p, 'п/п', pt=8, bold=True)

            _cw(cs[I_КОМ], W_КОМ)
            _fill(cs[I_КОМ], 'Команда', align=C, pt=8, bold=True)

            if I_НОМ_END is not None:
                _hspan(cs, I_НОМ, I_НОМ_END)
            _cw(cs[I_НОМ], W_НОМ)
            _fill(cs[I_НОМ], 'Номер', align=C, pt=8, bold=True)

            _cw(cs[I_ФИО], W_ФИО)
            _fill(cs[I_ФИО], 'Фамилия, имя, отчество участников', align=C, pt=8, bold=True)

            if I_ТЕРР_END is not None:
                _hspan(cs, I_ТЕРР, I_ТЕРР_END)
            _cw(cs[I_ТЕРР], W_ТЕРР)
            _fill(cs[I_ТЕРР], 'Территория', align=C, pt=8, bold=True)

            # МБ-столбец: «Очки Многоборье [дисц]» — одной строкой
            _cw(cs[I_МБ0], W_МБ1)
            _fill(cs[I_МБ0],
                  'Очки Многоборье ' + (мб_колонки[0]['заг'] if мб_колонки else ''),
                  align=C, pt=8, bold=True)

            p2 = cs[I_ОЧ].paragraphs[0]; p2.clear(); p2.style = 'No Spacing'; _sp(p2)
            p2.alignment = C; _valign(cs[I_ОЧ])
            _cw(cs[I_ОЧ], W_ОЧ)
            _run(p2, 'Очередь', pt=8, bold=True); _line_br(p2)
            _run(p2, ' выбора', pt=8, bold=True); _line_br(p2)
            _run(p2, ' позиции', pt=8, bold=True)

            p3 = cs[I_ВС].paragraphs[0]; p3.clear(); p3.style = 'No Spacing'; _sp(p3)
            p3.alignment = C; _valign(cs[I_ВС])
            _cw(cs[I_ВС], W_ВС)
            _run(p3, 'Время', pt=8, bold=True); _line_br(p3)
            _run(p3, 'старта', pt=8, bold=True)

        else:
            # N>1: две строки заголовка с vMerge
            # Строка 1 (restart)
            r = тbl.add_row(); cs = list(r.cells)
            _tbl_hdr_row(r)
            for i, w in enumerate(GRID): _cw(cs[i], w)
            _fill_fixed_hdr_cells(cs, with_borders=True, vmerge_mode='restart')

            # МБ-строка 1: «Очки Многоборье» объединяет все N МБ-столбцов
            _hspan(cs, I_МБ0, I_МБ0 + N - 1)
            _cw(cs[I_МБ0], N * W_МБ)
            _fill(cs[I_МБ0], 'Очки Многоборье', align=C, pt=8, bold=True)
            _borders(cs[I_МБ0], sz=6)

            # Строка 2 (cont), высота 1772 dxa ≈ 3.12 cm
            r2 = тbl.add_row(); _row_h(r2, round(1772 * 2.54 / 1440, 2)); cs2 = list(r2.cells)
            _tbl_hdr_row(r2)
            for i, w in enumerate(GRID): _cw(cs2[i], w)
            _fill_fixed_hdr_cells(cs2, with_borders=True, vmerge_mode='cont')

            # МБ-подзаголовки — вертикальный текст btLr
            for j, кол in enumerate(мб_колонки):
                ci = I_МБ0 + j
                _cw(cs2[ci], W_МБ)
                _fill(cs2[ci], кол['заг'], align=C, pt=8, bold=True)
                _vert_cell(cs2[ci]); _mb_margins(cs2[ci], lr=28 if N <= 3 else 108)
                _borders(cs2[ci], sz=6)

        # ── Секционная строка ──────────────────────────────────────────────────
        r = тbl.add_row(); _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[NC - 1]); _cw(cs[0], TW)
        _borders(cs[0])
        _fill(cs[0], метка, align=C, pt=9, bold=True)

        # ── Строки данных ──────────────────────────────────────────────────────
        def border_row(row, sz=4):
            for cell in row.cells:
                _borders(cell, sz)

        for пп, к in enumerate(команды, 1):
            r = тbl.add_row(); cs = list(r.cells)
            border_row(r)
            спортсмены = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
            фио = ', '.join(с.фамилия_имя_отчество for с in спортсмены if с.фамилия_имя_отчество)

            # №п/п
            _cw(cs[I_НПП], W_НПП)
            _fill(cs[I_НПП], str(пп), align=C)

            # Команда
            _cw(cs[I_КОМ], W_КОМ)
            _fill(cs[I_КОМ], к.название or '—')

            # Номер
            if I_НОМ_END is not None:
                cs[I_НОМ].merge(cs[I_НОМ_END])
            _cw(cs[I_НОМ], W_НОМ)
            _fill(cs[I_НОМ], str(к.нагрудный_номер or ''), align=C)

            # ФИО
            _cw(cs[I_ФИО], W_ФИО)
            _fill(cs[I_ФИО], фио)

            # Территория
            if I_ТЕРР_END is not None:
                cs[I_ТЕРР].merge(cs[I_ТЕРР_END])
            _cw(cs[I_ТЕРР], W_ТЕРР)
            _fill(cs[I_ТЕРР], к.субъект.название if к.субъект else '—')

            # МБ-очки
            for j, кол in enumerate(мб_колонки):
                ci = I_МБ0 + j
                val = _sum_scores(к, мб_колонки) if not кол.get('поле') else _score(к, кол['поле'])
                _ш_мб = 1038 if N == 1 else W_МБ
                _cw(cs[ci], _ш_мб)
                _fill(cs[ci], val, align=C)

            # Очередь выбора позиции
            рез = к.результат_дг
            оч = str(пп)
            _cw(cs[I_ОЧ], W_ОЧ)
            _fill(cs[I_ОЧ], оч, align=C)

            # Время старта
            _cw(cs[I_ВС], W_ВС)
            _fill(cs[I_ВС], _time_старт(к), align=C)

        _tiny_para(doc)

        # ── Таблица 3: подписи ─────────────────────────────────────────────────
        # По образцу: gridCols=[3905, 7291], высота строки 663 dxa (exact)
        W_SIG_L = 3905
        W_SIG_R = TW - W_SIG_L   # 11196-3905=7291

        sig = _make_tbl(2)
        _set_tbl_grid(sig, [W_SIG_L, W_SIG_R])

        for pos, fio in (('Главный судья',    сор.главный_судья     or ''),
                         ('Главный секретарь', сор.главный_секретарь or '')):
            r = sig.add_row()
            # высота 663 dxa exact (по образцу)
            val663 = 663
            trPr = r._tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = OxmlElement('w:trPr'); r._tr.insert(0, trPr)
            trH = trPr.find(qn('w:trHeight'))
            if trH is None:
                trH = OxmlElement('w:trHeight'); trPr.append(trH)
            trH.set(qn('w:val'), str(val663)); trH.set(qn('w:hRule'), 'exact')
            cs = list(r.cells)
            _cw(cs[0], W_SIG_L); _cw(cs[1], W_SIG_R)
            _no_border(cs[0]); _no_border(cs[1])
            _fill(cs[0], pos, pt=10)
            _fill(cs[1], fio, pt=10)

    doc  = Document()
    sect = doc.sections[0]
    # Страница по образцу: A4 портрет, поля точно из образца (emu)
    if портрет:
        sect.page_width  = 7560310   # 210 мм
        sect.page_height = 10692130  # 297 мм
        sect.left_margin   = 270510  # 7.514 мм
        sect.right_margin  = 180340  # 5.009 мм
        sect.top_margin    = 647700  # 17.99 мм
        sect.bottom_margin = 431800  # 11.99 мм
    else:
        sect.page_width  = 10692130  # 297 мм
        sect.page_height = 7560310   # 210 мм
        sect.left_margin   = 270510
        sect.right_margin  = 180340
        sect.top_margin    = 647700
        sect.bottom_margin = 431800
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if портрет:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')

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
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_ход_пс_docx(сор, секции, filename, portrait=False, судьи=None):
    """Протокол хода Параллельного спринта.

    секции: list of {'метка': str, 'заезды': list}
    заезд: {'local': int, 'стадия': str, 'участники': [...], 'время_старта': time|None}
    участник: {'позиция': str, 'команда': obj|None, 'метка': str|None,
                'запись': obj|None, 'метка_стадии': str|None}
    """
    from docx.shared import Mm, Pt

    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER

    # 10-column grid (dxa)
    # A4 landscape L=15 R=10 → usable 272mm = 15420 dxa
    # A4 portrait  L=15 R=10 → usable 185mm = 10488 dxa  (proportional scale ≈ 0.68)
    _LANDSCAPE = [1200,  900, 1200, 2510, 2510, 1200, 1500, 1500, 1080, 1820]
    _PORTRAIT  = [ 900,  800, 1000, 1384, 1385,  816, 1021, 1021,  900, 1261]
    GRID = _PORTRAIT if portrait else _LANDSCAPE
    NC   = len(GRID)
    TW   = sum(GRID)
    PT   = 10 if portrait else 12

    I_СТ, I_ЗН, I_ПО = 0, 1, 2
    I_К1, I_К2, I_НО = 3, 4, 5
    I_ВС, I_ВФ, I_ШТ, I_МС = 6, 7, 8, 9

    W_КОМ  = GRID[I_К1] + GRID[I_К2]
    W_SIG_L = sum(GRID[:5])   # 9600
    W_SIG_R = TW - W_SIG_L   # 5820

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
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), '000000')

    def _no_border(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')

    def _cw(cell, w):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        tcW.set(qn('w:w'), str(w))
        tcW.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        va = pr.find(qn('w:vAlign'))
        if va is None:
            va = OxmlElement('w:vAlign'); pr.append(va)
        va.set(qn('w:val'), val)

    def _fill(cell, text, align=None, pt=8, bold=False):
        p = cell.paragraphs[0]
        p.clear()
        p.style = 'No Spacing'
        _sp(p)
        if align:
            p.alignment = align
        if text:
            _run(p, text, pt, bold)

    def _set_grid(tbl):
        tg = tbl._tbl.find(qn('w:tblGrid'))
        if tg is not None:
            tbl._tbl.remove(tg)
        tg = OxmlElement('w:tblGrid')
        for w in GRID:
            gc = OxmlElement('w:gridCol')
            gc.set(qn('w:w'), str(w))
            tg.append(gc)
        tbl._tbl.insert(0, tg)

    def _tbl_w(tbl):
        pr = tbl._tbl.find(qn('w:tblPr'))
        if pr is None:
            pr = OxmlElement('w:tblPr')
            tbl._tbl.insert(0, pr)
        tw = pr.find(qn('w:tblW'))
        if tw is None:
            tw = OxmlElement('w:tblW'); pr.append(tw)
        tw.set(qn('w:w'), str(TW))
        tw.set(qn('w:type'), 'dxa')
        tl = pr.find(qn('w:tblLayout'))
        if tl is None:
            tl = OxmlElement('w:tblLayout'); pr.append(tl)
        tl.set(qn('w:type'), 'fixed')
        ti = pr.find(qn('w:tblInd'))
        if ti is None:
            ti = OxmlElement('w:tblInd'); pr.append(ti)
        ti.set(qn('w:w'), '0')
        ti.set(qn('w:type'), 'dxa')

    def _row_h(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr')
            row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')):
            pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567)))
        rh.set(qn('w:hRule'), 'exact')
        pr.append(rh)

    def _row_h_min(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr')
            row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')):
            pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567)))
        rh.set(qn('w:hRule'), 'atLeast')
        pr.append(rh)

    def _vmerge_restart(cell):
        pr = _tcpr(cell)
        vm = OxmlElement('w:vMerge')
        vm.set(qn('w:val'), 'restart')
        pr.append(vm)

    def _vmerge_cont(cell):
        pr = _tcpr(cell)
        vm = OxmlElement('w:vMerge')
        pr.append(vm)

    # ── local time formatters ─────────────────────────────────────────────────

    def _fmt_time(t):
        if t is None:
            return ''
        t = float(t)
        h = int(t // 3600)
        rest = t % 3600
        m = int(rest // 60)
        sv = rest % 60
        si = int(sv)
        return f'{h:02d}:{m:02d}:{si:02d}'

    def _fmt_dur(secs):
        if secs is None:
            return ''
        secs = float(secs)
        m = int(secs // 60)
        sv = secs % 60
        si = int(sv)
        cc = min(99, round((sv - si) * 100))
        return f'{m}:{si:02d}.{cc:02d}'

    # ── header paragraphs ─────────────────────────────────────────────────────

    def _add_header(doc, дисциплина, пол_рус):
        _para(doc, сор.название, align=C, pt=PT, bold=True)
        _para(doc, 'ПРОТОКОЛ ХОДА СОРЕВНОВАНИЙ', align=C, pt=PT, bold=True)
        _para(doc, f'Спортивная дисциплина «{дисциплина}»', align=C, pt=PT, bold=True)
        _para(doc, пол_рус, align=C, pt=PT, bold=True)

    # ── table builder ─────────────────────────────────────────────────────────

    def _add_table(doc, заезды):
        tbl = doc.add_table(rows=0, cols=NC)
        _tbl_w(tbl); _set_grid(tbl)

        # ── Column header row ─────────────────────────────────────────────────
        r = tbl.add_row(); _row_h_min(r, 1.4)
        cs = list(r.cells)
        for i, w in enumerate(GRID):
            _cw(cs[i], w)
        _fill(cs[I_СТ], 'Стадия',                       align=C, pt=PT, bold=True)
        _fill(cs[I_ЗН], 'Заезд\n№',                     align=C, pt=PT, bold=True)
        _fill(cs[I_ПО], 'Место\nспринт,\nпозиция\nв заезде', align=C, pt=PT, bold=True)
        cs[I_К1].merge(cs[I_К2]); _cw(cs[I_К1], W_КОМ)
        _fill(cs[I_К],  'Команда',                        align=C, pt=PT, bold=True)
        _fill(cs[I_НО], 'Номер',                          align=C, pt=PT, bold=True)
        _fill(cs[I_ВС], 'Время\nстарта',                align=C, pt=PT, bold=True)
        _fill(cs[I_ВФ], 'Время\nфиниша',                align=C, pt=PT, bold=True)
        _fill(cs[I_ШТ], 'Штраф',                        align=C, pt=PT, bold=True)
        _fill(cs[I_МС], 'Место,\nстадия',                align=C, pt=PT, bold=True)
        for cell in cs:
            _borders(cell); _valign(cell)

        # ── Heat rows ─────────────────────────────────────────────────────────
        for заезд in заезды:
            вс = _fmt_time(заезд['время_старта'])
            for row_i, уч in enumerate(заезд['участники'][:2]):
                r = tbl.add_row(); _row_h_min(r, 0.7)
                cs = list(r.cells)
                for i, w in enumerate(GRID):
                    _cw(cs[i], w)
                for cell in cs:
                    _borders(cell); _valign(cell)

                # Col 0: Стадия (vmerge across 2 rows)
                if row_i == 0:
                    _fill(cs[I_СТ], заезд['стадия'], align=C, pt=PT, bold=False)
                    _vmerge_restart(cs[I_СТ])
                else:
                    _vmerge_cont(cs[I_СТ])

                # Col 1: Заезд № (vmerge across 2 rows)
                if row_i == 0:
                    _fill(cs[I_ЗН], str(заезд['local']), align=C, pt=PT, bold=False)
                    _vmerge_restart(cs[I_ЗН])
                else:
                    _vmerge_cont(cs[I_ЗН])

                # Col 2: Позиция
                _fill(cs[I_ПО], уч['позиция'], align=C, pt=PT)

                # Col 3+4: Команда (merged)
                cs[I_К1].merge(cs[I_К2]); _cw(cs[I_К1], W_КОМ)
                if уч['команда']:
                    _fill(cs[I_К1], уч['команда'].название or '—', pt=PT)

                # Col 5: Номер
                ном = str(уч['команда'].нагрудный_номер) if уч['команда'] else ''
                _fill(cs[I_НО], ном, align=C, pt=PT)

                # Col 6: Время старта (vmerge across 2 rows)
                if row_i == 0:
                    _fill(cs[I_ВС], вс, align=C, pt=PT)
                    _vmerge_restart(cs[I_ВС])
                else:
                    _vmerge_cont(cs[I_ВС])

                # Cols 7, 8, 9: empty — judges fill manually
                _fill(cs[I_ВФ], '', pt=PT)
                _fill(cs[I_ШТ], '', pt=PT)
                _fill(cs[I_МС], '', pt=PT)

        # ── Signature rows ────────────────────────────────────────────────────
        for pos, инфо in (судьи or []):
            r = tbl.add_row(); _row_h(r, 1.2)
            cs = list(r.cells)
            cs[0].merge(cs[4]); _cw(cs[0], W_SIG_L)
            cs[5].merge(cs[NC - 1]); _cw(cs[5], W_SIG_R)
            _no_border(cs[0]); _no_border(cs[5])
            _valign(cs[0]); _valign(cs[5])
            _fill(cs[0], pos, pt=PT)
            _fill(cs[5], инфо, pt=PT)

    # ── Document assembly ─────────────────────────────────────────────────────

    doc  = Document()
    sect = doc.sections[0]
    if portrait:
        sect.page_width    = Mm(210)
        sect.page_height   = Mm(297)
    else:
        sect.page_width    = Mm(297)
        sect.page_height   = Mm(210)
    sect.left_margin   = Mm(15)
    sect.right_margin  = Mm(10)
    sect.top_margin    = Mm(15)
    sect.bottom_margin = Mm(15)

    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if portrait:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')

    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    for i, секция in enumerate(секции):
        if i > 0:
            _page_br(doc)
        _add_header(doc, секция['дисциплина'], секция['пол_рус'])
        _add_table(doc, секция['заезды'])

    buf = BytesIO()
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_результаты_слалом_docx(
    сор, кл, кол_ворот, итоговый,
    муж_рез, муж_без, муж_разр, муж_поп, муж_ш1, муж_ш2,
    жен_рез, жен_без, жен_разр, жен_поп, жен_ш1, жен_ш2,
    filename, портрет=False
):
    """Generate Slalom results protocol DOCX (preliminary with gate columns, or final)."""
    from docx.shared import Mm, Pt

    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT

    кол_ворот = max(кол_ворот, 1)

    TW  = round((210 - 10 - 5) * 1440 / 25.4) if портрет else round((297 - 7 - 5) * 1440 / 25.4)
    MAR = 28
    if портрет:
        PT, PT_FIO, PT_HDR = 8, 7, 8
    else:
        PT, PT_FIO, PT_HDR = 10, 8, 10

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

    def _tcMar(cell, lr=28):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcMar')): pr.remove(old)
        mar = OxmlElement('w:tcMar')
        for side in ('left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:w'), str(lr)); el.set(qn('w:type'), 'dxa')
            mar.append(el)
        pr.append(mar)

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

    def _ls15(p):
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:line'), '360'); sp.set(qn('w:lineRule'), 'auto')

    def _no_border_cell(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')

    def _bottom_border_cell(cell, sz=6):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
        el = b.find(qn('w:bottom'))
        if el is None:
            el = OxmlElement('w:bottom'); b.append(el)
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _make_tbl_sl(ncols, fixed_tw, jc='center'):
        t = doc.add_table(rows=0, cols=ncols)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        el.set(qn('w:w'), str(fixed_tw)); el.set(qn('w:type'), 'dxa')
        ind = tblPr.find(qn('w:tblInd'))
        if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
        ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        jc_el = tblPr.find(qn('w:jc'))
        if jc_el is None: jc_el = OxmlElement('w:jc'); tblPr.append(jc_el)
        jc_el.set(qn('w:val'), jc)
        return t

    def _no_wrap_cell(cell):
        pr = _tcpr(cell)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))

    def _fill_sl(cell, text, align=None, pt=PT, bold=False):
        p = cell.paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p)
        if align: p.alignment = align
        text = str(text).replace('—', '–')
        if not text: return
        parts = str(text).split('\n')
        for i, part in enumerate(parts):
            r = p.add_run(part); r.font.name = TNR; r.font.size = Pt(pt)
            if bold: r.font.bold = True
            if i < len(parts) - 1:
                br = OxmlElement('w:br'); r._r.append(br)

    def _add_doc_header(doc):
        _hdr_pt = (10 if портрет else 12) if итоговый else (12 if портрет else 10)
        if сор.организаторы_блок:
            _para(doc, сор.организаторы_блок, align=C, pt=_hdr_pt, bold=True)
        p0 = _para(doc); p0.paragraph_format.line_spacing = Pt(4)
        _ls15(_para(doc, сор.название or '', align=C, pt=12, bold=True))
        if итоговый and сор.номер_екп:
            _ls15(_para(doc, f'ЕКП № {сор.номер_екп}', align=C, pt=_hdr_pt, bold=True))
        заг = 'ПРОТОКОЛ РЕЗУЛЬТАТОВ' if итоговый else 'ПРЕДВАРИТЕЛЬНЫЙ ПРОТОКОЛ РЕЗУЛЬТАТОВ'
        _ls15(_para(doc, заг, align=C, pt=_hdr_pt, bold=True))
        if кл == 'R4':
            дисц = 'СЛАЛОМ – R4, код ВРВС 1740031811Я'
        elif кл == 'R6':
            дисц = 'СЛАЛОМ – R6, код ВРВС 1740041811Я'
        else:
            дисц = f'СЛАЛОМ{" – " + кл if кл else ""}'
        _ls15(_para(doc, f'Спортивная дисциплина: {дисц}', align=C, pt=_hdr_pt, bold=True))

        # Дата/место block
        _cw_char = round(PT * 0.6 * 20)
        _w_dt  = max(100, len(сор.даты_проведения or 'Дата проведения') * _cw_char + 2 * MAR)
        _w_pl  = max(100, len(сор.место_проведения or 'Место проведения') * _cw_char + 2 * MAR)
        _w_mid = max(100, TW - _w_dt - _w_pl)
        dt = _make_tbl_sl(3, fixed_tw=TW)
        _set_grid(dt, [_w_dt, _w_mid, _w_pl])
        _dt_rh = lambda r: _row_h(r, 0.4)
        for idx, (лево, право, pt_v, bold_v) in enumerate([
            (сор.даты_проведения or '', сор.место_проведения or '', PT, True),
            ('Дата проведения', 'Место проведения', 8, True),
        ]):
            r = dt.add_row(); _dt_rh(r)
            cs = list(r.cells)
            for c in cs:
                _no_border_cell(c)
                _valign(c, 'bottom' if (not портрет and idx == 0) else ('top' if idx == 1 else 'center'))
            _no_wrap_cell(cs[0]); _no_wrap_cell(cs[2])
            if idx == 0:
                _bottom_border_cell(cs[0]); _bottom_border_cell(cs[2])
            _fill_sl(cs[0], лево, align=C, pt=pt_v, bold=bold_v)
            _fill_sl(cs[2], право, align=C, pt=pt_v, bold=bold_v)

    # ── signatures ────────────────────────────────────────────────────────────

    def _add_sigs(doc):
        w_l = TW * (33 if портрет else 24) // 100
        w_r = TW - w_l
        sig = _make_tbl_sl(2, fixed_tw=TW)
        _set_grid(sig, [w_l, w_r])
        for позиция, фио in (('Главный судья',    сор.главный_судья     or ''),
                              ('Главный секретарь', сор.главный_секретарь or '')):
            r = sig.add_row(); _row_h(r, 0.88)
            cs = list(r.cells)
            _cw(cs[0], w_l); _cw(cs[1], w_r)
            _no_border_cell(cs[0]); _no_border_cell(cs[1])
            _valign(cs[0]); _valign(cs[1])
            _fill_sl(cs[0], позиция, align=L, pt=PT, bold=True)
            _fill_sl(cs[1], фио,    align=L, pt=PT, bold=True)

    # ══════════════════════════════════════════════════════════════════════════
    # PRELIMINARY PROTOCOL  (two rows per team, gate columns, vertical text)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_pred(doc):
        снос = ((len(муж_рез) + len(муж_без) > 0 and len(муж_рез) + len(муж_без) < 6) or
                (len(жен_рез) + len(жен_без) > 0 and len(жен_рез) + len(жен_без) < 6))
        if портрет:
            _FL_PCT = [2.0, 9.8, 3.6, 21.3, 8.2, 1.9, 1.9]
            _FR_PCT = [1.9, 3.7, 4.3, 5.3, 4.9, 2.8, 3.3, 2.9]
        else:
            _FL_PCT = [1.8, 8.7, 3.3, 21.2, 6.2, 1.9, 1.9]
            _FR_PCT = [1.9, 3.7, 4.6, 5.0, 5.0, 2.2, 3.1, 3.6]
        _gate_pct = max(0.5, (100.0 - sum(_FL_PCT) - sum(_FR_PCT)) / кол_ворот)
        _pcts = _FL_PCT + [_gate_pct] * кол_ворот + _FR_PCT
        GRID  = [round(TW * p / 100) for p in _pcts]
        GRID[-1] += TW - sum(GRID)
        NC      = len(GRID)
        TW_пред = TW

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
            _fill(cs[I_СТАРТ], str(ст) if ст else '0', align=C, pt=PT)
            for g in range(1, кол_ворот + 1):
                v = ш.get(g, 0) if ш else 0
                _fill(cs[I_ВОР0 + g - 1], str(v) if v else '0', align=C, pt=PT)
            фш = (рез.кол_штраф_финиш or 0) * 50 if рез else 0
            _fill(cs[I_ФИН], str(фш) if фш else '0', align=C, pt=PT)
            if рез is None:
                _fill(cs[I_ШТ], '–', align=C, pt=PT); _fill(cs[I_ВР], '–', align=C, pt=PT); _fill(cs[I_ИТ], '–', align=C, pt=PT)
            elif рез.статус_попытки == 'НСТ':
                _fill(cs[I_ШТ], '–', align=C, pt=PT); _fill(cs[I_ВР], 'Не старт.', align=C, pt=PT); _fill(cs[I_ИТ], '–', align=C, pt=PT)
            elif рез.статус_попытки == 'НФН':
                _fill(cs[I_ШТ], '–', align=C, pt=PT); _fill(cs[I_ВР], 'Не фин.', align=C, pt=PT); _fill(cs[I_ИТ], '–', align=C, pt=PT)
            elif рез.статус_попытки == 'ПЛВ':
                _fill(cs[I_ШТ], '–', align=C, pt=PT); _fill(cs[I_ВР], 'Прев. ЛВ', align=C, pt=PT); _fill(cs[I_ИТ], '–', align=C, pt=PT)
            elif рез.статус_попытки == 'ДК':
                _fill(cs[I_ШТ], _fmt_шт(рез.штрафное_время), align=C, pt=PT)
                _fill(cs[I_ВР], _fmt_dur(рез.время_прохождения) if рез.время_прохождения else '–', align=C, pt=PT)
                _fill(cs[I_ИТ], 'ДК', align=C, pt=PT)
            else:
                _fill(cs[I_ШТ], _fmt_шт(рез.штрафное_время), align=C, pt=PT)
                _fill(cs[I_ВР], _fmt_dur(рез.время_прохождения) if рез.время_прохождения else '–', align=C, pt=PT)
                _fill(cs[I_ИТ], _fmt_dur(рез.результат_попытки) if рез.результат_попытки else '–', align=C, pt=PT)

        tbl = _make_tbl_sl(NC, fixed_tw=TW_пред)
        _set_grid(tbl, GRID)

        # Header row 1
        rh1 = tbl.add_row()
        _row_h_min(rh1, 0.5)
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
        _fill(cs[I_ПОП], 'Попытка', align=C, bold=True); _vert_cell(cs[I_ПОП]); _valign(cs[I_ПОП])
        _fill(cs[I_СТАРТ], 'Прохождение дистанции (штрафные баллы)', align=C, bold=True)
        _fill(cs[I_ШТ], 'Результат попытки', align=C, bold=True)
        _fill(cs[I_ЛУЧ], 'Результат лучшей попытки', align=C, bold=True)
        _fill(cs[I_МЕС], 'Место', align=C, bold=True); _vert_cell(cs[I_МЕС]); _valign(cs[I_МЕС])
        _fill(cs[I_ОЧК], 'Очки', align=C, bold=True)
        p_нор = cs[I_НОР].paragraphs[0]; p_нор.clear(); _sp(p_нор); p_нор.alignment = C
        _run(p_нор, 'Выполнение', pt=8, bold=True)
        _line_br(p_нор)
        _run(p_нор, 'разряда' + ('*' if снос else ''), pt=8, bold=True)
        _valign(cs[I_НОР]); _vert_cell(cs[I_НОР]); _valign(cs[I_НОР])
        for col in ROWSPAN_H:
            _vmerge_restart(cs[col])

        # Header row 2 (vertical gate numbers)
        rh2 = tbl.add_row()
        _row_h_min(rh2, 1.4)
        cs2 = list(rh2.cells)
        for i, w in enumerate(GRID):
            _cw(cs2[i], w)
        _border_row(rh2)
        for col in ROWSPAN_H:
            _vmerge_cont(cs2[col])
        _fill(cs2[I_СТАРТ], 'Старт', align=C, bold=True);  _vert_cell(cs2[I_СТАРТ]); _valign(cs2[I_СТАРТ])
        for g in range(1, кол_ворот + 1):
            _fill(cs2[I_ВОР0 + g - 1], str(g), align=C, bold=True)
        _fill(cs2[I_ФИН], 'Финиш', align=C, bold=True);   _vert_cell(cs2[I_ФИН]); _valign(cs2[I_ФИН])
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
            _fill(cd1[0], str(пп), align=C, pt=PT)
            _fill(cd1[1], к.название or '—', pt=7)
            _fill(cd1[2], str(к.нагрудный_номер or ''), align=C, pt=PT)
            _fill(cd1[3], фио, pt=7)
            _fill(cd1[4], к.субъект.название if к.субъект else '—', pt=7)
            _fill(cd1[I_ПОП], '1', align=C, pt=PT)
            _fill(cd2[I_ПОП], '2', align=C, pt=PT)
            _fill_поп(cd1, р1, ш1)
            _fill_поп(cd2, р2, ш2)
            _fill(cd1[I_ЛУЧ], _fmt_dur(лучший), align=C, pt=PT)
            _fill(cd1[I_МЕС], str(место) if место else '–', align=C, pt=PT)
            _fill(cd1[I_ОЧК], str(int(очки)) if очки else '–', align=C, pt=PT)
            _fill(cd1[I_НОР], норм, align=C, pt=8)

        _add_section(муж_рез, муж_без, муж_разр, муж_поп, муж_ш1, муж_ш2, 'Мужчины')
        _add_section(жен_рез, жен_без, жен_разр, жен_поп, жен_ш1, жен_ш2, 'Женщины')
        return снос

    # ══════════════════════════════════════════════════════════════════════════
    # FINAL PROTOCOL  (one row per team, no gate columns)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_итог(doc):
        снос = ((len(муж_рез) + len(муж_без) > 0 and len(муж_рез) + len(муж_без) < 6) or
                (len(жен_рез) + len(жен_без) > 0 and len(жен_рез) + len(жен_без) < 6))
        # Proportions from sample Протокол_результатов_сл_итог.docx
        _PCTS_ИТ = [2.8, 11.2, 4.3, 22.7, 7.2,   # НПП, Команда, Номер, Состав, Субъект
                    4.9, 4.7, 5.9,                   # 1-я: Время, Штраф, Результат
                    4.9, 4.7, 5.9,                   # 2-я: Время, Штраф, Результат
                    5.4, 4.2, 3.9, 7.1]              # Лучшая, Место, Очки, Разряд
        GRID_ИТ = [round(TW * p / 100) for p in _PCTS_ИТ]
        GRID_ИТ[-1] += TW - sum(GRID_ИТ)
        if портрет:
            # Col 14 (Разряд) btLr: column WIDTH = line width for rotated text.
            # "Выполнение" (10 chars × ~5pt) ≈ 50pt ≈ 1000 dxa + margins → use 1100 dxa.
            _W_RAZ = 1100
            GRID_ИТ[3] -= (_W_RAZ - GRID_ИТ[14])  # reduce Состав only
            GRID_ИТ[14] = _W_RAZ
        NC = len(GRID_ИТ)

        def _init_r(row, borders=True):
            cs = list(row.cells)
            for i, w in enumerate(GRID_ИТ):
                _cw(cs[i], w)
            if borders:
                _border_row(row)
            return cs

        def _init_r_mar(row):
            cs = _init_r(row)
            for c in cs:
                _valign(c)
                _tcMar(c, MAR)
            return cs

        tbl = doc.add_table(rows=0, cols=NC)
        _set_grid(tbl, GRID_ИТ)
        _tbl_w(tbl, TW)

        _HDR_H_ИТ = 0.85 if портрет else 1.0
        rh1 = tbl.add_row(); _row_h_min(rh1, _HDR_H_ИТ)
        cs = _init_r_mar(rh1)
        cs[5].merge(cs[7])
        cs[8].merge(cs[10])
        p = cs[0].paragraphs[0]; p.clear(); _sp(p); p.alignment = C
        _run(p, '№', pt=PT_HDR, bold=True); _line_br(p); _run(p, 'п/п', pt=PT_HDR, bold=True)
        _fill(cs[1],  'Команда',             align=C, pt=PT_HDR, bold=True)
        _fill(cs[2],  'Номер',               align=C, pt=PT_HDR, bold=True)
        _fill(cs[3],  'Состав',              align=C, pt=PT_HDR, bold=True)
        _fill(cs[4],  'Субъект РФ',          align=C, pt=PT_HDR, bold=True)
        _fill(cs[5],  '1-я попытка',         align=C, pt=PT_HDR, bold=True)
        _fill(cs[8],  '2-я попытка',         align=C, pt=PT_HDR, bold=True)
        _fill(cs[11], 'Лучшая попытка',      align=C, pt=PT_HDR, bold=True)
        _fill(cs[12], 'Место',               align=C, pt=PT_HDR, bold=True)
        _fill(cs[13], 'Очки',               align=C, pt=PT_HDR, bold=True)
        _fill(cs[14], 'Выполнение\nразряда' + ('*' if снос else ''),  align=C, pt=PT_HDR, bold=True)
        if портрет:
            for _c in (cs[2], cs[12], cs[14]):
                _vert_cell(_c)
                _valign(_c, 'center')
                for _p in _c.paragraphs:
                    _p.alignment = C
        for col in [0, 1, 2, 3, 4, 11, 12, 13, 14]:
            _vmerge_restart(cs[col])

        rh2 = tbl.add_row(); _row_h_min(rh2, _HDR_H_ИТ)
        cs2 = _init_r_mar(rh2)
        for col in [0, 1, 2, 3, 4, 11, 12, 13, 14]:
            _vmerge_cont(cs2[col])
        for i, lbl in enumerate(['Время', 'Штраф', 'Результат', 'Время', 'Штраф', 'Результат']):
            _fill(cs2[5 + i], lbl, align=C, pt=PT_HDR, bold=True)

        def _поп_cells(cs, рез, base):
            if рез is None:
                for i in range(3): _fill(cs[base + i], '–', align=C, pt=PT)
            elif рез.статус_попытки == 'НСТ':
                _fill(cs[base], 'Не старт.', align=C, pt=PT); _fill(cs[base+1], '–', align=C, pt=PT); _fill(cs[base+2], '–', align=C, pt=PT)
            elif рез.статус_попытки == 'НФН':
                _fill(cs[base], 'Не фин.', align=C, pt=PT); _fill(cs[base+1], '–', align=C, pt=PT); _fill(cs[base+2], '–', align=C, pt=PT)
            elif рез.статус_попытки == 'ПЛВ':
                _fill(cs[base], '–', align=C, pt=PT); _fill(cs[base+1], '–', align=C, pt=PT); _fill(cs[base+2], 'Прев. ЛВ', align=C, pt=PT)
            elif рез.статус_попытки == 'ДК':
                _fill(cs[base],   _fmt_dur(рез.время_прохождения) if рез.время_прохождения else '–', align=C, pt=PT)
                _fill(cs[base+1], _fmt_шт(рез.штрафное_время), align=C, pt=PT)
                _fill(cs[base+2], 'ДК', align=C, pt=PT)
            else:
                _fill(cs[base],   _fmt_dur(рез.время_прохождения) if рез.время_прохождения else '–', align=C, pt=PT)
                _fill(cs[base+1], _fmt_шт(рез.штрафное_время), align=C, pt=PT)
                _fill(cs[base+2], _fmt_dur(рез.результат_попытки) if рез.результат_попытки else '–', align=C, pt=PT)

        def _add_sect(рез_list, без_list, разряды, поп, заг_пол):
            if not рез_list and not без_list:
                return
            норм_флаг = (len(рез_list) + len(без_list)) >= 6
            rg = tbl.add_row(); _row_h(rg, 0.6)
            cg = list(rg.cells); cg[0].merge(cg[NC - 1])
            _border_row(rg)
            _fill(list(rg.cells)[0], заг_пол, align=C, pt=PT, bold=True)
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
                    dr = tbl.add_row(); _row_h_min(dr, 0.5)
                    cs = _init_r_mar(dr)
                    _fill(cs[0],  str(idx),                                    align=C, pt=PT)
                    _fill(cs[1],  к.название or '–',                       align=L, pt=PT)
                    _fill(cs[2],  str(к.нагрудный_номер or ''),                align=C, pt=PT)
                    _fill(cs[3],  состав,                                       align=L, pt=(6 if портрет else PT_FIO))
                    _fill(cs[4],  к.субъект.название if к.субъект else '–',    align=C, pt=PT)
                    _поп_cells(cs, р1, 5)
                    _поп_cells(cs, р2, 8)
                    if is_без:
                        for i in range(11, 15): _fill(cs[i], '–', align=C, pt=PT)
                    else:
                        _fill(cs[11], _fmt_dur(рез.лучший_результат),                align=C, pt=PT)
                        _fill(cs[12], str(рез.место),                                align=C, pt=PT)
                        _fill(cs[13], str(int(рез.очки)) if рез.очки else '–',       align=C, pt=PT)
                        _fill(cs[14], разряды.get(к.id, '–') if норм_флаг else '–', align=C, pt=PT)

        _add_sect(муж_рез, муж_без, муж_разр, муж_поп, 'Мужчины')
        _add_sect(жен_рез, жен_без, жен_разр, жен_поп, 'Женщины')
        if снос:
            _para(doc,
                  '* Разряды не присваиваются, т.к. количество участников в виде программы '
                  'недостаточно для присвоения в соответствии с положением о ЕВСК (пункт 25)',
                  align=L, pt=PT_FIO, bold=True)

    # ── Build document ────────────────────────────────────────────────────────

    doc = Document()
    sect = doc.sections[0]
    if портрет:
        sect.page_height = Mm(297); sect.page_width = Mm(210)
    else:
        sect.page_height = Mm(210); sect.page_width = Mm(297)
    if портрет:
        sect.top_margin    = Mm(15)
        sect.bottom_margin = Mm(15)
        sect.left_margin   = Mm(10)
        sect.right_margin  = Mm(5)
    else:
        sect.top_margin    = Mm(12)
        sect.bottom_margin = Mm(12)
        sect.left_margin   = Mm(7)
        sect.right_margin  = Mm(5)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if портрет:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(10)

    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    _add_doc_header(doc)
    _para(doc).paragraph_format.line_spacing = Pt(2 if портрет else 1)

    if итоговый:
        _build_итог(doc)
    else:
        снос_пред = _build_pred(doc)
        if снос_пред:
            _para(doc, '* Разряды не присваиваются, т.к. количество участников в виде программы '
                  'недостаточно для присвоения в соответствии с положением о ЕВСК (пункт 25)',
                  align=L, pt=8, bold=True)

    _add_sigs(doc)

    buf2 = BytesIO()
    if not итоговый and not портрет:
        for _tbl in doc.tables:
            for _row in _tbl.rows:
                for _cell in _row.cells:
                    _tcMar(_cell, MAR)
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf2)
    buf2.seek(0)
    return send_file(buf2, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_жеребьёвка_docx(сор, муж_р4, жен_р4, муж_р6, жен_р6, последовательность, filename):
    """Generate draw (жеребьёвка) protocol DOCX matching Протокол_жеребьёвки.docx sample."""
    from docx.shared import Mm, Pt

    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT

    # A4 portrait, L=15, R=10, T=12, B=22.5
    TW  = round((210 - 15 - 10) * 1440 / 25.4)
    MAR = 28
    PT  = 10

    _КАРТА = {'R4-м': муж_р4, 'R4-ж': жен_р4, 'R6-м': муж_р6, 'R6-ж': жен_р6}
    _МЕТКИ = {
        'R4-м': 'Спринт – R4, Мужчины', 'R4-ж': 'Спринт – R4, Женщины',
        'R6-м': 'Спринт – R6, Мужчины', 'R6-ж': 'Спринт – R6, Женщины',
    }
    виды = [к for к in последовательность if к in _КАРТА and _КАРТА[к]]

    # Column widths from sample proportions (6 logical cols)
    _PCTS = [4.9, 16.0, 7.6, 38.7, 17.3, 15.4]
    GRID  = [round(TW * p / 100) for p in _PCTS]
    GRID[-1] += TW - sum(GRID)
    NC = 6

    doc  = Document()
    sect = doc.sections[0]
    sect.page_width    = Mm(210)
    sect.page_height   = Mm(297)
    sect.left_margin   = Mm(15)
    sect.right_margin  = Mm(10)
    sect.top_margin    = Mm(12)
    sect.bottom_margin = Mm(22.5)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        pgSz.attrib.pop(qn('w:orient'), None)
    doc.styles['Normal'].font.name = TNR
    doc.styles['Normal'].font.size = Pt(PT)

    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    def _sp(p):
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def _run(p, text, pt=10, bold=False):
        r = p.add_run(text)
        r.font.name = TNR; r.font.size = Pt(pt)
        if bold: r.font.bold = True
        return r

    def _para(text='', align=None, pt=10, bold=False):
        p = doc.add_paragraph()
        _sp(p)
        if align: p.alignment = align
        if text: _run(p, text, pt, bold)
        return p

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _cw(cell, w):
        pr = _tcpr(cell)
        el = pr.find(qn('w:tcW'))
        if el is None: el = OxmlElement('w:tcW'); pr.append(el)
        el.set(qn('w:w'), str(w)); el.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        v = pr.find(qn('w:vAlign'))
        if v is None: v = OxmlElement('w:vAlign'); pr.append(v)
        v.set(qn('w:val'), val)

    def _borders(cell, sz=4):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')): pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), '000000')
            b.append(el)
        pr.append(b)

    def _no_border(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')): pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
            b.append(el)
        pr.append(b)

    def _bottom_border(cell, sz=6):
        _no_border(cell)
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        el = OxmlElement('w:bottom')
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0'); el.set(qn('w:color'), '000000')
        b.append(el)

    def _fill(cell, text, align=None, pt=10, bold=False):
        p = cell.paragraphs[0]; p.clear(); _sp(p)
        if align: p.alignment = align
        if text: _run(p, str(text), pt, bold)
        _valign(cell)

    def _row_h(row, cm):
        val = round(cm * 1440 / 2.54)
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None: trPr = OxmlElement('w:trPr'); row._tr.insert(0, trPr)
        trH = trPr.find(qn('w:trHeight'))
        if trH is None: trH = OxmlElement('w:trHeight'); trPr.append(trH)
        trH.set(qn('w:val'), str(val)); trH.set(qn('w:hRule'), 'atLeast')

    def _make_tbl(ncols):
        t = doc.add_table(rows=0, cols=ncols)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None: tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        el.set(qn('w:w'), str(TW)); el.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        return t

    def _set_grid(tbl, grid):
        t = tbl._tbl
        for old in t.findall(qn('w:tblGrid')): t.remove(old)
        g = OxmlElement('w:tblGrid')
        for w in grid:
            col = OxmlElement('w:gridCol'); col.set(qn('w:w'), str(w)); g.append(col)
        tblPr = t.find(qn('w:tblPr'))
        tblPr.addnext(g) if tblPr is not None else t.insert(0, g)

    # ── Document header ────────────────────────────────────────────────────────

    if сор.организаторы_блок:
        _para(сор.организаторы_блок, align=C, pt=10, bold=True)
    _para(сор.название or '', align=C, pt=12, bold=True)
    _para('Протокол жеребьёвки', align=C, pt=12, bold=True)
    _para('Группа спортивных дисциплин СПРИНТ', align=C, pt=10, bold=True)

    # Date / place table
    _cw_char = round(PT * 0.6 * 20)
    _w_dt  = max(100, len(сор.даты_проведения or 'Дата проведения') * _cw_char + 2 * MAR)
    _w_pl  = max(100, len(сор.место_проведения or 'Место проведения') * _cw_char + 2 * MAR)
    _w_mid = max(100, TW - _w_dt - _w_pl)
    dt = _make_tbl(3)
    _set_grid(dt, [_w_dt, _w_mid, _w_pl])
    for idx, (лево, право, pt_v, bold_v) in enumerate([
        (сор.даты_проведения or '', сор.место_проведения or '', 10, True),
        ('Дата проведения', 'Место проведения', 8, False),
    ]):
        r = dt.add_row()
        cs = list(r.cells)
        for c in cs: _no_border(c); _valign(c, 'center')
        if idx == 0:
            _bottom_border(cs[0]); _bottom_border(cs[2])
        _fill(cs[0], лево, align=C, pt=pt_v, bold=bold_v)
        _fill(cs[2], право, align=C, pt=pt_v, bold=bold_v)

    # ── Main table ─────────────────────────────────────────────────────────────

    tbl = _make_tbl(NC)
    _set_grid(tbl, GRID)

    # Header row
    rh = tbl.add_row()
    cs = list(rh.cells)
    for i, w in enumerate(GRID): _cw(cs[i], w); _borders(cs[i])
    _fill(cs[0], '№ п/п',  align=C, pt=8, bold=True)
    _fill(cs[1], 'Команда', align=C, pt=8, bold=True)
    _fill(cs[2], 'Номер',  align=C, pt=8, bold=True)
    _fill(cs[3], 'Фамилия, имя, отчество участников', align=C, pt=8, bold=True)
    _fill(cs[4], 'Территория', align=C, pt=8, bold=True)
    _fill(cs[5], 'Порядковый номер согласно жеребьёвке', align=C, pt=8, bold=True)

    for код in виды:
        # Section header row
        rs = tbl.add_row()
        cs_s = list(rs.cells)
        cs_s[0].merge(cs_s[NC - 1])
        for c in list(rs.cells): _borders(c)
        _fill(list(rs.cells)[0], _МЕТКИ[код], align=C, pt=9, bold=True)

        команды_вида = _КАРТА[код]
        for пп, к in enumerate(команды_вида, 1):
            спортсмены = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
            фио = ', '.join(с.фамилия_имя_отчество for с in спортсмены if с.фамилия_имя_отчество)
            rd = tbl.add_row()
            cd = list(rd.cells)
            for i, w in enumerate(GRID): _cw(cd[i], w); _borders(cd[i])
            _fill(cd[0], str(пп), align=C, pt=8)
            _fill(cd[1], к.название or '—', pt=8)
            _fill(cd[2], str(к.нагрудный_номер or ''), align=C, pt=8)
            _fill(cd[3], фио, pt=8)
            _fill(cd[4], к.субъект.название if к.субъект else '—', pt=8)
            _fill(cd[5], str(к.стартовый_номер_спринт or '—'), align=C, pt=8)

    # ── Signatures ─────────────────────────────────────────────────────────────

    sig = _make_tbl(2)
    w_l = TW * 43 // 100
    w_r = TW - w_l
    _set_grid(sig, [w_l, w_r])
    for позиция, фио in (('Главный судья', сор.главный_судья or ''),
                          ('Главный секретарь', сор.главный_секретарь or '')):
        rs = sig.add_row(); _row_h(rs, 1.5)
        cs = list(rs.cells)
        _cw(cs[0], w_l); _cw(cs[1], w_r)
        _no_border(cs[0]); _no_border(cs[1])
        _valign(cs[0]); _valign(cs[1])
        _fill(cs[0], позиция, align=L, pt=10, bold=False)
        _fill(cs[1], фио,    align=L, pt=10, bold=False)

    buf = BytesIO()
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


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
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_результаты_пред_спринт_docx(сор, муж_рез, муж_без, жен_рез, жен_без,
                                        муж_разр, жен_разр, класс, filename, портрет=False):
    """Предварительный протокол результатов Спринт — по образцу Протокол_результатов_сп_пред.docx."""
    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT
    R   = WD_ALIGN_PARAGRAPH.RIGHT

    # 13 grid columns: НПП/Команда/Номер/ФИО/Субъект/Старт/Финиш/РезШтраф/Время/Итоговый/Место/Очки/Разряд
    TW = round((210 - 15 - 15) * 1440 / 25.4) if портрет else round((297 - 15 - 10) * 1440 / 25.4)
    if портрет:
        _PCTS = [2.9, 12.1, 5.0, 19.6, 9.3, 5.5, 6.5, 6.5, 5.3, 7.6, 5.3, 5.3, 9.3]
    else:
        _PCTS = [3.1, 13.1, 4.7, 24.8, 10.7, 5.3, 5.3, 5.3, 5.3, 6.2, 4.5, 3.6, 8.0]
    GRID = [round(TW * p / 100) for p in _PCTS]
    GRID[-1] += TW - sum(GRID)
    NC = 13

    I_НПП, I_К,   I_НОМ, I_ФИО, I_СУБ = 0, 1, 2, 3, 4
    I_ШТС, I_ШТФ                        = 5, 6
    I_ШТИ, I_ВП,  I_ИТ                  = 7, 8, 9
    I_МЕ,  I_ОЧ,  I_РА                  = 10, 11, 12

    W_ПРОХ = GRID[I_ШТС] + GRID[I_ШТФ]
    W_РЕЗ  = GRID[I_ШТИ] + GRID[I_ВП] + GRID[I_ИТ]
    W_SIG_L = TW * 35 // 100
    W_SIG_R = TW - W_SIG_L



    if портрет:
        PT, PT_FIO, PT_HDR = 8, 7, 8
    else:
        PT, PT_FIO, PT_HDR = 10, 8, 10
    MAR = 28  # 0.05cm = ~28 dxa — отступ слева/справа во всех ячейках

    # ── helpers ──────────────────────────────────────────────────────────────
    def _fmt_dur(secs):
        if secs is None: return '–'
        secs = float(secs); neg = secs < 0; secs = abs(secs)
        m = int(secs // 60); sv = secs % 60; si = int(sv)
        cc = min(99, round((sv - si) * 100))
        return ('-' if neg else '') + f'{m}:{si:02d}.{cc:02d}'

    def _sp(p):
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def _run(p, text, pt=PT, bold=False):
        r = p.add_run(text); r.font.name = TNR; r.font.size = Pt(pt)
        if bold: r.font.bold = True
        return r

    def _para(doc, text='', align=None, pt=PT, bold=False):
        p = doc.add_paragraph(); _sp(p)
        if align: p.alignment = align
        if text: _run(p, text, pt, bold)
        return p

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _borders(cell, sz=4):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _no_border(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
        el = b.find(qn('w:bottom'))
        if el is None:
            el = OxmlElement('w:bottom'); b.append(el)
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _cw(cell, w):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        tcW.set(qn('w:w'), str(int(w))); tcW.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        va = pr.find(qn('w:vAlign'))
        if va is None:
            va = OxmlElement('w:vAlign'); pr.append(va)
        va.set(qn('w:val'), val)

    def _fill(cell, text, align=None, pt=PT, bold=False):
        p = cell.paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p)
        if align: p.alignment = align
        text = str(text).replace('—', '–')  # — → –
        if not text: return
        parts = str(text).split('\n')
        for i, part in enumerate(parts):
            r = p.add_run(part); r.font.name = TNR; r.font.size = Pt(pt)
            if bold: r.font.bold = True
            if i < len(parts) - 1:
                br = OxmlElement('w:br'); r._r.append(br)

    def _cw_pct(cell, dxa):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        pct = round(dxa * 5000 / TW)
        tcW.set(qn('w:w'), str(pct)); tcW.set(qn('w:type'), 'pct')

    def _make_tbl(ncols, fixed_tw=None, pct_w=5000, jc=None):
        t = doc.add_table(rows=0, cols=ncols)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        if fixed_tw:
            el.set(qn('w:w'), str(fixed_tw)); el.set(qn('w:type'), 'dxa')
        else:
            el.set(qn('w:w'), str(pct_w)); el.set(qn('w:type'), 'pct')
        if jc == 'center':
            # убрать tblInd чтобы jc=center работало
            for old_ind in tblPr.findall(qn('w:tblInd')): tblPr.remove(old_ind)
        else:
            ind = tblPr.find(qn('w:tblInd'))
            if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
            ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        if jc:
            jc_el = tblPr.find(qn('w:jc'))
            if jc_el is None: jc_el = OxmlElement('w:jc'); tblPr.append(jc_el)
            jc_el.set(qn('w:val'), jc)
        return t

    def _set_grid_sp(tbl_obj, widths):
        tbl = tbl_obj._tbl
        old = tbl.find(qn('w:tblGrid'))
        if old is not None: tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in widths:
            gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); grid.append(gc)
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is not None: tblPr.addnext(grid)
        else: tbl.insert(0, grid)

    def _row_h(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'exact')
        pr.append(rh)

    def _row_h_min(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'atLeast')
        pr.append(rh)

    def _vmerge_restart(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); vm.set(qn('w:val'), 'restart'); pr.append(vm)

    def _vmerge_cont(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); pr.append(vm)

    def _tcMar(cell, lr=28):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcMar')): pr.remove(old)
        mar = OxmlElement('w:tcMar')
        for side in ('left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:w'), str(lr)); el.set(qn('w:type'), 'dxa')
            mar.append(el)
        pr.append(mar)

    def _no_wrap_cell(cell):
        pr = _tcpr(cell)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))

    def _line_br(p):
        p.add_run()._r.append(OxmlElement('w:br'))

    def _vert_cell(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:textDirection')): pr.remove(old)
        td = OxmlElement('w:textDirection'); td.set(qn('w:val'), 'btLr'); pr.append(td)
        _valign(cell, 'center')
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def _ls15(p):
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:line'), '360'); sp.set(qn('w:lineRule'), 'auto')

    def _keep_next_p(p):
        pPr = p._p.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr'); p._p.insert(0, pPr)
        if pPr.find(qn('w:keepNext')) is None:
            pPr.append(OxmlElement('w:keepNext'))

    # ── flags ────────────────────────────────────────────────────────────────
    муж_норм = (len(муж_рез) + len(муж_без)) >= 6
    жен_норм = (len(жен_рез) + len(жен_без)) >= 6
    снос = ((len(муж_рез) + len(муж_без)) > 0 and not муж_норм) or \
           ((len(жен_рез) + len(жен_без)) > 0 and not жен_норм)

    _коды = {'R4': 'СПРИНТ – R4, код ВРВС 1740091811Я',
             'R6': 'СПРИНТ – R6, код ВРВС 1740101811Я'}
    дисциплина = _коды.get(класс, 'СПРИНТ' + (f' – {класс}' if класс else ''))

    # ── document ─────────────────────────────────────────────────────────────
    from docx.shared import Mm
    doc  = Document()
    sect = doc.sections[0]
    if портрет:
        sect.page_height   = Mm(297); sect.page_width    = Mm(210)
    else:
        sect.page_height   = Mm(210); sect.page_width    = Mm(297)
    sect.top_margin    = Mm(10)
    sect.bottom_margin = Mm(15)
    sect.left_margin   = Mm(15)
    sect.right_margin  = Mm(5) if портрет else Mm(10)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if портрет:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(10)
    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    _hdr_pt = 10 if портрет else 12
    if сор.организаторы_блок:
        _para(doc, сор.организаторы_блок, align=C, pt=_hdr_pt, bold=True)
    p0 = _para(doc); p0.paragraph_format.line_spacing = Pt(4)
    _ls15(_para(doc, сор.название or '', align=C, pt=12, bold=True))
    _ls15(_para(doc, 'ПРЕДВАРИТЕЛЬНЫЙ ПРОТОКОЛ РЕЗУЛЬТАТОВ', align=C, pt=_hdr_pt, bold=True))
    _ls15(_para(doc, f'Спортивная дисциплина: {дисциплина}', align=C, pt=_hdr_pt, bold=True))

    # ── таблица дата/место ───────────────────────────────────────────────────
    dt = _make_tbl(3, fixed_tw=TW, jc='center')
    # Ширина крайних колонок по содержимому: символ TNR ≈ PT*0.6пт, + 2×отступ
    _cw_char = round(PT * 0.6 * 20)
    _w_dt  = max(100, len(сор.даты_проведения or 'Дата проведения') * _cw_char + 2 * MAR)
    _w_pl  = max(100, len(сор.место_проведения or 'Место проведения') * _cw_char + 2 * MAR)
    _w_mid = max(100, TW - _w_dt - _w_pl)
    _set_grid_sp(dt, [_w_dt, _w_mid, _w_pl])
    _dt_rh = (lambda r: _row_h(r, 0.3)) if портрет else (lambda r: _row_h(r, 0.4))
    _dt_rows = [
        (сор.даты_проведения or '', сор.место_проведения or '', PT, True),
        ('Дата проведения', 'Место проведения', PT if not портрет else 7, not портрет),
    ]
    for idx, (лево, право, pt_v, bold_v) in enumerate(_dt_rows):
        r = dt.add_row(); _dt_rh(r)
        cs = list(r.cells)
        for c in cs:
            _no_border(c)
            if not портрет:
                _valign(c, 'bottom' if idx == 0 else 'top')
            else:
                _valign(c)
        _no_wrap_cell(cs[0])
        _no_wrap_cell(cs[2])
        if idx == 0:
            _bottom_border(cs[0])
            _bottom_border(cs[2])
        _fill(cs[0], лево, align=C, pt=pt_v, bold=bold_v)
        _fill(cs[2], право, align=C, pt=pt_v, bold=bold_v)

    # ── пустая строка перед таблицей (1пт) ──────────────────────────────────
    p = _para(doc)
    p.paragraph_format.line_spacing = Pt(1)

    # ── main table ───────────────────────────────────────────────────────────
    tbl = _make_tbl(NC, fixed_tw=TW, jc='center')
    _set_grid_sp(tbl, GRID)

    # row 0 — header part 1 (vMerge restart на одиночных; hspan на составных)
    r = tbl.add_row(); _row_h_min(r, 0.9 if портрет else 0.6)
    cs = list(r.cells)
    for i, w in enumerate(GRID): _cw_pct(cs[i], w)
    cs[I_ШТС].merge(cs[I_ШТФ]); _cw_pct(cs[I_ШТС], W_ПРОХ)
    cs[I_ШТИ].merge(cs[I_ИТ]);  _cw_pct(cs[I_ШТИ], W_РЕЗ)
    for i in (I_НПП, I_К, I_НОМ, I_ФИО, I_СУБ, I_МЕ, I_ОЧ, I_РА):
        _vmerge_restart(cs[i])
    _fill(cs[I_НПП], '№\nп/п',                                    align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_К],   'Команда',                                    align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_НОМ], 'Номер',                                      align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ФИО], 'Фамилия, имя, отчество участников',         align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_СУБ], 'Субъект РФ',                                align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ШТС], 'Штрафы',     align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ШТИ], 'Результат',  align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_МЕ],  'Место',                                      align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ОЧ],  'Очки',                                       align=C, pt=PT_HDR, bold=True)
    p_ра = cs[I_РА].paragraphs[0]; p_ра.clear(); _sp(p_ра); p_ра.alignment = C
    _run(p_ра, 'Выполнение', PT_HDR, True); _line_br(p_ра)
    _run(p_ра, 'разряда' + ('*' if снос else ''), PT_HDR, True)
    for i in range(NC):
        _borders(cs[i]); _valign(cs[i])
    _borders(cs[I_ШТС]); _borders(cs[I_ШТИ])
    for i in range(NC):
        _tcMar(cs[i], MAR); _no_wrap_cell(cs[i])
    if портрет:
        _vert_cell(cs[I_РА])

    # row 1 — header part 2 (vMerge continue; sub-headers Старт/Финиш/Штраф/Время/Итоговый)
    r = tbl.add_row(); _row_h_min(r, 0.9 if портрет else 0.6)
    cs = list(r.cells)
    for i, w in enumerate(GRID): _cw_pct(cs[i], w)
    for i in (I_НПП, I_К, I_НОМ, I_ФИО, I_СУБ, I_МЕ, I_ОЧ, I_РА):
        _vmerge_cont(cs[i])
    _fill(cs[I_ШТС], 'Старт',    align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ШТФ], 'Финиш',   align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ШТИ], 'Штраф',   align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ВП],  'Время',    align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ИТ],  'Итоговый', align=C, pt=PT_HDR, bold=True)
    for i in range(NC):
        _borders(cs[i]); _valign(cs[i])
    for i in range(NC):
        _tcMar(cs[i], MAR); _no_wrap_cell(cs[i])

    # ── section/data helpers ─────────────────────────────────────────────────
    def _section(label):
        r = tbl.add_row(); _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[NC - 1]); _cw_pct(cs[0], TW)
        _borders(cs[0]); _valign(cs[0])
        _fill(cs[0], label, align=C, pt=PT, bold=True)
        _keep_next_p(cs[0].paragraphs[-1])

    def _data_row(пп, рез, разряды, норм):
        from models import СТАТУС_ДК, СТАТУС_ДНС, СТАТУС_ДНФ, СТАТУС_ПМ, СТАТУС_ПЛВ
        к = рез.команда
        состав = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
        фио = ', '.join(с.фамилия_имя_отчество for с in состав) if состав else '–'
        шт_ст = (рез.штраф_с10 or 0) * 10 + (рез.штраф_с50 or 0) * 50
        шт_фн = (рез.штраф_ф50 or 0) * 50

        r = tbl.add_row(); _row_h_min(r, 0.5)
        cs = list(r.cells)
        for i, w in enumerate(GRID): _cw_pct(cs[i], w)
        for i in range(NC):
            _borders(cs[i]); _valign(cs[i]); _tcMar(cs[i], MAR)

        _fill(cs[I_НПП], str(пп),                        align=C, pt=PT)
        _fill(cs[I_К],   к.название or '–',              align=L, pt=PT)
        _fill(cs[I_НОМ], str(к.нагрудный_номер or ''),   align=C, pt=PT)
        _fill(cs[I_ФИО], фио,                             align=L, pt=PT_FIO)
        _fill(cs[I_СУБ], к.субъект.название if к.субъект else '–', align=C, pt=PT)

        ст = рез.статус
        if ст in (СТАТУС_ДНС, СТАТУС_ДНФ):
            _fill(cs[I_ШТС], 'Не старт.' if ст == СТАТУС_ДНС else '–', align=C, pt=PT)
            _fill(cs[I_ШТФ], 'Не фин.' if ст == СТАТУС_ДНФ else '–',  align=C, pt=PT)
            _fill(cs[I_ШТИ], '–', align=C, pt=PT)
            _fill(cs[I_ВП],  'Не старт.' if ст == СТАТУС_ДНС else 'Не фин.', align=C, pt=PT)
            _fill(cs[I_ИТ],  '–', align=C, pt=PT)
        elif ст == СТАТУС_ДК:
            _fill(cs[I_ШТС], 'ДК', align=C, pt=PT)
            _fill(cs[I_ШТФ], _fmt_dur(шт_фн) if шт_фн else '–', align=C, pt=PT)
            _fill(cs[I_ШТИ], _fmt_dur(рез.штрафное_время) or '0:00.00', align=C, pt=PT)
            _fill(cs[I_ВП],  _fmt_dur(рез.время_прохождения), align=C, pt=PT)
            _fill(cs[I_ИТ],  'ДК', align=C, pt=PT)
        else:
            _fill(cs[I_ШТС], _fmt_dur(шт_ст) if шт_ст else '–', align=C, pt=PT)
            фн_т = 'Посл. м.' if ст == СТАТУС_ПМ else (_fmt_dur(шт_фн) if шт_фн else '–')
            _fill(cs[I_ШТФ], фн_т, align=C, pt=PT)
            _fill(cs[I_ШТИ], _fmt_dur(рез.штрафное_время) or '0:00.00', align=C, pt=PT)
            _fill(cs[I_ВП],  _fmt_dur(рез.время_прохождения), align=C, pt=PT)
            if ст == СТАТУС_ПМ:    итог = 'Посл. м.'
            elif ст == СТАТУС_ПЛВ: итог = 'Прев. ЛВ'
            else:                   итог = _fmt_dur(рез.результат)
            _fill(cs[I_ИТ], итог, align=C, pt=PT)

        _fill(cs[I_МЕ], str(рез.место) if рез.место else '–', align=C, pt=PT)
        _fill(cs[I_ОЧ], str(int(рез.очки)) if рез.очки else '–', align=C, pt=PT)
        _fill(cs[I_РА], разряды.get(рез.команда_id, '–') if норм else '–', align=C, pt=PT)

    # ── data sections ────────────────────────────────────────────────────────
    _section('Мужчины')
    for пп, рез in enumerate(муж_рез + муж_без, 1):
        _data_row(пп, рез, муж_разр, муж_норм)

    _section('Женщины')
    for пп, рез in enumerate(жен_рез + жен_без, 1):
        _data_row(пп, рез, жен_разр, жен_норм)

    # ── footnote ─────────────────────────────────────────────────────────────
    if снос:
        _para(doc, '* Разряды не присваиваются, т.к. количество участников в виде программы '
              'недостаточно для присвоения в соответствии с положением о ЕВСК (пункт 25)',
              align=L, pt=PT_FIO, bold=True)

    # ── signature table (фикс. ширина = портрет, строго под основной таблицей) ──
    sig = _make_tbl(2, fixed_tw=TW, jc='center')
    _set_grid_sp(sig, [W_SIG_L, W_SIG_R])

    for позиция, фио in (('Главный судья',    сор.главный_судья     or ''),
                          ('Главный секретарь', сор.главный_секретарь or '')):
        r = sig.add_row(); _row_h(r, 0.88)
        cs = list(r.cells)
        _cw(cs[0], W_SIG_L); _cw(cs[1], W_SIG_R)
        _no_border(cs[0]); _no_border(cs[1])
        _valign(cs[0]); _valign(cs[1])
        _fill(cs[0], позиция, align=L, pt=PT, bold=True)
        _fill(cs[1], фио,    align=L, pt=PT, bold=True)

    # ── publish time ─────────────────────────────────────────────────────────
    from datetime import datetime
    now_str = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
    _para(doc, f'Время опубликования: {now_str}', align=L, pt=PT)

    buf = BytesIO()
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf); buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_результаты_пред_дг_docx(сор, муж_рез, муж_без, жен_рез, жен_без,
                                    муж_разр, жен_разр, класс, filename, портрет=False):
    """Предварительный протокол результатов Длинная гонка."""
    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT
    R   = WD_ALIGN_PARAGRAPH.RIGHT

    # 14 grid columns: НПП/Команда/Номер/ФИО/Субъект/Старт/ЛимВр/Финиш/РезШтраф/Время/Итоговый/Место/Очки/Разряд
    TW = round((210 - 15 - 10) * 1440 / 25.4) if портрет else round((297 - 15 - 10) * 1440 / 25.4)
    if портрет:
        _PCTS = [3.1, 13.1, 4.7, 21.3, 8.9, 5.3, 5.3, 5.3, 5.3, 5.3, 6.2, 4.4, 3.6, 8.0]
    else:
        _PCTS = [3.1, 13.1, 4.7, 21.3, 8.9, 5.3, 5.3, 5.3, 5.3, 5.3, 6.2, 4.4, 3.6, 8.0]
    GRID = [round(TW * p / 100) for p in _PCTS]
    GRID[-1] += TW - sum(GRID)
    NC = 14

    I_НПП, I_К,   I_НОМ, I_ФИО, I_СУБ = 0, 1, 2, 3, 4
    I_ШТС, I_ЛИМ, I_ШТФ               = 5, 6, 7
    I_ШТИ, I_ВП,  I_ИТ                  = 8, 9, 10
    I_МЕ,  I_ОЧ,  I_РА                  = 11, 12, 13

    W_ПРОХ = GRID[I_ШТС] + GRID[I_ЛИМ] + GRID[I_ШТФ]
    W_РЕЗ  = GRID[I_ШТИ] + GRID[I_ВП] + GRID[I_ИТ]
    W_SIG_L = TW * 35 // 100
    W_SIG_R = TW - W_SIG_L



    if портрет:
        PT, PT_FIO, PT_HDR = 8, 7, 8
    else:
        PT, PT_FIO, PT_HDR = 10, 8, 10
    MAR = 28  # 0.05cm = ~28 dxa — отступ слева/справа во всех ячейках

    # ── helpers ──────────────────────────────────────────────────────────────
    def _fmt_dur(secs):
        if secs is None: return '–'
        secs = float(secs); neg = secs < 0; secs = abs(secs)
        m = int(secs // 60); sv = secs % 60; si = int(sv)
        cc = min(99, round((sv - si) * 100))
        return ('-' if neg else '') + f'{m}:{si:02d}.{cc:02d}'

    def _sp(p):
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def _run(p, text, pt=PT, bold=False):
        r = p.add_run(text); r.font.name = TNR; r.font.size = Pt(pt)
        if bold: r.font.bold = True
        return r

    def _para(doc, text='', align=None, pt=PT, bold=False):
        p = doc.add_paragraph(); _sp(p)
        if align: p.alignment = align
        if text: _run(p, text, pt, bold)
        return p

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _borders(cell, sz=4):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _no_border(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
        el = b.find(qn('w:bottom'))
        if el is None:
            el = OxmlElement('w:bottom'); b.append(el)
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _cw(cell, w):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        tcW.set(qn('w:w'), str(int(w))); tcW.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        va = pr.find(qn('w:vAlign'))
        if va is None:
            va = OxmlElement('w:vAlign'); pr.append(va)
        va.set(qn('w:val'), val)

    def _fill(cell, text, align=None, pt=PT, bold=False):
        p = cell.paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p)
        if align: p.alignment = align
        if not text: return
        parts = str(text).split('\n')
        for i, part in enumerate(parts):
            r = p.add_run(part); r.font.name = TNR; r.font.size = Pt(pt)
            if bold: r.font.bold = True
            if i < len(parts) - 1:
                br = OxmlElement('w:br'); r._r.append(br)

    def _cw_pct(cell, dxa):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        pct = round(dxa * 5000 / TW)
        tcW.set(qn('w:w'), str(pct)); tcW.set(qn('w:type'), 'pct')

    def _make_tbl(ncols, fixed_tw=None, pct_w=5000, jc=None):
        t = doc.add_table(rows=0, cols=ncols)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        if fixed_tw:
            el.set(qn('w:w'), str(fixed_tw)); el.set(qn('w:type'), 'dxa')
        else:
            el.set(qn('w:w'), str(pct_w)); el.set(qn('w:type'), 'pct')
        if jc == 'center':
            # убрать tblInd чтобы jc=center работало
            for old_ind in tblPr.findall(qn('w:tblInd')): tblPr.remove(old_ind)
        else:
            ind = tblPr.find(qn('w:tblInd'))
            if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
            ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        if jc:
            jc_el = tblPr.find(qn('w:jc'))
            if jc_el is None: jc_el = OxmlElement('w:jc'); tblPr.append(jc_el)
            jc_el.set(qn('w:val'), jc)
        return t

    def _set_grid_sp(tbl_obj, widths):
        tbl = tbl_obj._tbl
        old = tbl.find(qn('w:tblGrid'))
        if old is not None: tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in widths:
            gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); grid.append(gc)
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is not None: tblPr.addnext(grid)
        else: tbl.insert(0, grid)

    def _row_h(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'exact')
        pr.append(rh)

    def _row_h_min(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'atLeast')
        pr.append(rh)

    def _vmerge_restart(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); vm.set(qn('w:val'), 'restart'); pr.append(vm)

    def _vmerge_cont(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); pr.append(vm)

    def _tcMar(cell, lr=28):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcMar')): pr.remove(old)
        mar = OxmlElement('w:tcMar')
        for side in ('left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:w'), str(lr)); el.set(qn('w:type'), 'dxa')
            mar.append(el)
        pr.append(mar)

    def _no_wrap_cell(cell):
        pr = _tcpr(cell)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))

    def _line_br(p):
        p.add_run()._r.append(OxmlElement('w:br'))

    def _vert_cell(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:textDirection')): pr.remove(old)
        td = OxmlElement('w:textDirection'); td.set(qn('w:val'), 'btLr'); pr.append(td)
        _valign(cell, 'center')
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def _ls15(p):
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:line'), '360'); sp.set(qn('w:lineRule'), 'auto')

    def _keep_next_p(p):
        pPr = p._p.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr'); p._p.insert(0, pPr)
        if pPr.find(qn('w:keepNext')) is None:
            pPr.append(OxmlElement('w:keepNext'))

    # ── flags ────────────────────────────────────────────────────────────────
    муж_норм = (len(муж_рез) + len(муж_без)) >= 6
    жен_норм = (len(жен_рез) + len(жен_без)) >= 6
    снос = ((len(муж_рез) + len(муж_без)) > 0 and not муж_норм) or \
           ((len(жен_рез) + len(жен_без)) > 0 and not жен_норм)

    _коды_дг = {'R4': '1740011811Я', 'R6': '1740021811Я'}
    _кл_дг = класс or сор.класс_судна or 'R4/R6'
    _код_дг = _коды_дг.get(_кл_дг, сор.код_соревнований or '')
    дисциплина = f'ДЛИННАЯ ГОНКА – {_кл_дг}, код ВРВС {_код_дг}'

    # ── document ─────────────────────────────────────────────────────────────
    from docx.shared import Mm
    doc  = Document()
    sect = doc.sections[0]
    if портрет:
        sect.page_height   = Mm(297); sect.page_width    = Mm(210)
    else:
        sect.page_height   = Mm(210); sect.page_width    = Mm(297)
    sect.top_margin    = Mm(15)
    sect.bottom_margin = Mm(15)
    sect.left_margin   = Mm(15)
    sect.right_margin  = Mm(10)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if портрет:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(10)
    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    _hdr_pt = 10 if портрет else 12
    if сор.организаторы_блок:
        _para(doc, сор.организаторы_блок, align=C, pt=_hdr_pt, bold=True)
    p0 = _para(doc); p0.paragraph_format.line_spacing = Pt(4)
    _ls15(_para(doc, сор.название or '', align=C, pt=12, bold=True))
    _ls15(_para(doc, 'ПРЕДВАРИТЕЛЬНЫЙ ПРОТОКОЛ РЕЗУЛЬТАТОВ', align=C, pt=_hdr_pt, bold=True))
    _ls15(_para(doc, f'Спортивная дисциплина: {дисциплина}', align=C, pt=_hdr_pt, bold=True))

    # ── таблица дата/место ───────────────────────────────────────────────────
    dt = _make_tbl(3, fixed_tw=TW, jc='center')
    # Ширина крайних колонок по содержимому: символ TNR ≈ PT*0.6пт, + 2×отступ
    _cw_char = round(PT * 0.6 * 20)
    _w_dt  = max(100, len(сор.даты_проведения or 'Дата проведения') * _cw_char + 2 * MAR)
    _w_pl  = max(100, len(сор.место_проведения or 'Место проведения') * _cw_char + 2 * MAR)
    _w_mid = max(100, TW - _w_dt - _w_pl)
    _set_grid_sp(dt, [_w_dt, _w_mid, _w_pl])
    _dt_rh = (lambda r: _row_h(r, 0.3)) if портрет else (lambda r: _row_h(r, 0.4))
    _dt_rows = [
        (сор.даты_проведения or '', сор.место_проведения or '', PT, True),
        ('Дата проведения', 'Место проведения', PT if not портрет else 7, not портрет),
    ]
    for idx, (лево, право, pt_v, bold_v) in enumerate(_dt_rows):
        r = dt.add_row(); _dt_rh(r)
        cs = list(r.cells)
        for c in cs:
            _no_border(c)
            if not портрет:
                _valign(c, 'bottom' if idx == 0 else 'top')
            else:
                _valign(c)
        _no_wrap_cell(cs[0])
        _no_wrap_cell(cs[2])
        if idx == 0:
            _bottom_border(cs[0])
            _bottom_border(cs[2])
        _fill(cs[0], лево, align=C, pt=pt_v, bold=bold_v)
        _fill(cs[2], право, align=C, pt=pt_v, bold=bold_v)

    # ── пустая строка перед таблицей (1пт) ──────────────────────────────────
    p = _para(doc)
    p.paragraph_format.line_spacing = Pt(1)

    # ── main table ───────────────────────────────────────────────────────────
    tbl = _make_tbl(NC, fixed_tw=TW, jc='center')
    _set_grid_sp(tbl, GRID)

    # row 0 — header part 1 (vMerge restart на одиночных; hspan на составных)
    r = tbl.add_row(); _row_h_min(r, 0.9 if портрет else 0.6)
    cs = list(r.cells)
    for i, w in enumerate(GRID): _cw_pct(cs[i], w)
    cs[I_ШТС].merge(cs[I_ШТФ]); _cw_pct(cs[I_ШТС], W_ПРОХ)
    cs[I_ШТИ].merge(cs[I_ИТ]);  _cw_pct(cs[I_ШТИ], W_РЕЗ)
    for i in (I_НПП, I_К, I_НОМ, I_ФИО, I_СУБ, I_МЕ, I_ОЧ, I_РА):
        _vmerge_restart(cs[i])
    _fill(cs[I_НПП], '№\nп/п',                                    align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_К],   'Команда',                                    align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_НОМ], 'Номер',                                      align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ФИО], 'Фамилия, имя, отчество участников',         align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_СУБ], 'Субъект РФ',                                align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ШТС], 'Штрафы',     align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ШТИ], 'Результат',  align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_МЕ],  'Место',                                      align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ОЧ],  'Очки',                                       align=C, pt=PT_HDR, bold=True)
    p_ра = cs[I_РА].paragraphs[0]; p_ра.clear(); _sp(p_ра); p_ра.alignment = C
    _run(p_ра, 'Выполнение', PT_HDR, True); _line_br(p_ра)
    _run(p_ра, 'разряда' + ('*' if снос else ''), PT_HDR, True)
    for i in range(NC):
        _borders(cs[i]); _valign(cs[i])
    _borders(cs[I_ШТС], sz=6); _borders(cs[I_ШТИ])
    for i in range(NC):
        _tcMar(cs[i], MAR); _no_wrap_cell(cs[i])
    if портрет:
        _vert_cell(cs[I_РА])

    # row 1 — header part 2 (vMerge continue; sub-headers Старт/Финиш/Штраф/Время/Итоговый)
    r = tbl.add_row(); _row_h_min(r, 0.9 if портрет else 0.6)
    cs = list(r.cells)
    for i, w in enumerate(GRID): _cw_pct(cs[i], w)
    for i in (I_НПП, I_К, I_НОМ, I_ФИО, I_СУБ, I_МЕ, I_ОЧ, I_РА):
        _vmerge_cont(cs[i])
    _fill(cs[I_ШТС], 'Старт',    align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ШТФ], 'Финиш',   align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ШТФ], 'Финиш',          align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ШТИ], 'Штраф',   align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ВП],  'Время',    align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ИТ],  'Итоговый', align=C, pt=PT_HDR, bold=True)
    for i in range(NC):
        _borders(cs[i]); _valign(cs[i])
    for i in range(NC):
        _tcMar(cs[i], MAR); _no_wrap_cell(cs[i])
    _borders(cs[I_ШТС], sz=3); _borders(cs[I_ШТФ], sz=3)

    # ── section/data helpers ─────────────────────────────────────────────────
    def _section(label):
        r = tbl.add_row(); _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[NC - 1]); _cw_pct(cs[0], TW)
        _borders(cs[0]); _valign(cs[0])
        _fill(cs[0], label, align=C, pt=PT, bold=True)
        _keep_next_p(cs[0].paragraphs[-1])

    def _data_row(пп, рез, разряды, норм, лим_вр=None):
        from models import СТАТУС_ДК, СТАТУС_ДНС, СТАТУС_ДНФ, СТАТУС_ПМ, СТАТУС_ПЛВ
        к = рез.команда
        состав = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
        фио = ', '.join(с.фамилия_имя_отчество for с in состав) if состав else '–'

        r = tbl.add_row(); _row_h_min(r, 0.5)
        cs = list(r.cells)
        for i, w in enumerate(GRID): _cw_pct(cs[i], w)
        for i in range(NC):
            _borders(cs[i]); _valign(cs[i]); _tcMar(cs[i], MAR)

        _borders(cs[I_ШТС], sz=3); _borders(cs[I_ЛИМ], sz=3); _borders(cs[I_ШТФ], sz=3)
        _fill(cs[I_НПП], str(пп),                        align=C, pt=PT)
        _fill(cs[I_К],   к.название or '–',              align=L, pt=PT)
        _fill(cs[I_НОМ], str(к.нагрудный_номер or ''),   align=C, pt=PT)
        _fill(cs[I_ФИО], фио,                             align=L, pt=PT_FIO)
        _fill(cs[I_СУБ], к.субъект.название if к.субъект else '–', align=C, pt=PT)
        _fill(cs[I_ЛИМ], _fmt_dur(лим_вр) if лим_вр else '–', align=C, pt=PT)

        ст = рез.статус
        if ст in (СТАТУС_ДНС, СТАТУС_ДНФ):
            _fill(cs[I_ШТС], 'Не старт.' if ст == СТАТУС_ДНС else '–', align=C, pt=PT)
            _fill(cs[I_ШТФ], '–', align=C, pt=PT)
            _fill(cs[I_ШТИ], '–', align=C, pt=PT)
            _fill(cs[I_ВП],  'Не старт.' if ст == СТАТУС_ДНС else 'Не фин.', align=C, pt=PT)
            _fill(cs[I_ИТ],  '–', align=C, pt=PT)
        elif ст == СТАТУС_ДК:
            _fill(cs[I_ШТС], 'ДК', align=C, pt=PT)
            _fill(cs[I_ШТФ], '–', align=C, pt=PT)
            _fill(cs[I_ШТИ], _fmt_dur(рез.штрафное_время) or '0:00.00', align=C, pt=PT)
            _fill(cs[I_ВП],  _fmt_dur(рез.время_прохождения), align=C, pt=PT)
            _fill(cs[I_ИТ],  'ДК', align=C, pt=PT)
        else:
            _fill(cs[I_ШТС], '–', align=C, pt=PT)
            _fill(cs[I_ШТФ], '–', align=C, pt=PT)
            _fill(cs[I_ШТИ], _fmt_dur(рез.штрафное_время) or '0:00.00', align=C, pt=PT)
            _fill(cs[I_ВП],  _fmt_dur(рез.время_прохождения), align=C, pt=PT)
            if ст == СТАТУС_ПМ:    итог = 'Посл. м.'
            elif ст == СТАТУС_ПЛВ: итог = 'Прев. ЛВ'
            else:                   итог = _fmt_dur(рез.результат)
            _fill(cs[I_ИТ], итог, align=C, pt=PT)

        _fill(cs[I_МЕ], str(рез.место) if рез.место else '–', align=C, pt=PT)
        _fill(cs[I_ОЧ], str(int(рез.очки)) if рез.очки else '–', align=C, pt=PT)
        _fill(cs[I_РА], разряды.get(рез.команда_id, '–') if норм else '–', align=C, pt=PT)

    # ── data sections ────────────────────────────────────────────────────────
    _section('Мужчины')
    for пп, рез in enumerate(муж_рез + муж_без, 1):
        _data_row(пп, рез, муж_разр, муж_норм, сор.лимит_время_муж)

    _section('Женщины')
    for пп, рез in enumerate(жен_рез + жен_без, 1):
        _data_row(пп, рез, жен_разр, жен_норм, сор.лимит_время_жен)

    # ── footnote ─────────────────────────────────────────────────────────────
    if снос:
        _para(doc, '* Разряды не присваиваются, т.к. количество участников в виде программы '
              'недостаточно для присвоения в соответствии с положением о ЕВСК (пункт 25)',
              align=L, pt=PT_FIO, bold=True)

    # ── signature table (фикс. ширина = портрет, строго под основной таблицей) ──
    sig = _make_tbl(2, fixed_tw=TW, jc='center')
    _set_grid_sp(sig, [W_SIG_L, W_SIG_R])

    for позиция, фио in (('Главный судья',    сор.главный_судья     or ''),
                          ('Главный секретарь', сор.главный_секретарь or '')):
        r = sig.add_row(); _row_h(r, 0.88)
        cs = list(r.cells)
        _cw(cs[0], W_SIG_L); _cw(cs[1], W_SIG_R)
        _no_border(cs[0]); _no_border(cs[1])
        _valign(cs[0]); _valign(cs[1])
        _fill(cs[0], позиция, align=L, pt=PT, bold=True)
        _fill(cs[1], фио,    align=L, pt=PT, bold=True)

    # ── publish time ─────────────────────────────────────────────────────────
    from datetime import datetime
    now_str = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
    _para(doc, f'Время опубликования: {now_str}', align=L, pt=PT)

    buf = BytesIO()
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf); buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)




def create_результаты_итог_спринт_docx(сор, муж_рез, муж_без, жен_рез, жен_без,
                                        муж_разр, жен_разр, класс, filename, портрет=False):
    """Итоговый протокол результатов Спринт — 11 колонок (без Старт/Финиш)."""
    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT
    R   = WD_ALIGN_PARAGRAPH.RIGHT

    # 11 columns: НПП/Команда/Номер/Состав/Субъект/Штраф/Время/Итоговый/Место/Очки/Разряд
    TW = round((210 - 15 - 10) * 1440 / 25.4) if портрет else round((297 - 15 - 10) * 1440 / 25.4)
    if портрет:
        _PCTS = [3, 12, 6, 29, 10, 6, 6, 8, 5, 5, 10]
    else:
        _PCTS = [4.4, 13.8, 5.5, 29.1, 10.2, 5.5, 5.7, 7.3, 5.1, 4.6, 8.8]
    GRID = [round(TW * p / 100) for p in _PCTS]
    GRID[-1] += TW - sum(GRID)
    NC = 11

    I_НПП, I_К,   I_НОМ, I_СОС, I_СУБ = 0, 1, 2, 3, 4
    I_ШТИ, I_ВП,  I_ИТ                  = 5, 6, 7
    I_МЕ,  I_ОЧ,  I_РА                  = 8, 9, 10

    W_РЕЗ  = GRID[I_ШТИ] + GRID[I_ВП] + GRID[I_ИТ]
    W_SIG_L = TW * 35 // 100
    W_SIG_R = TW - W_SIG_L

    if портрет:
        PT, PT_FIO, PT_HDR = 8, 7, 8
    else:
        PT, PT_FIO, PT_HDR = 10, 8, 10
    MAR = 28

    # ── helpers ──────────────────────────────────────────────────────────────
    def _fmt_dur(secs):
        if secs is None: return '–'
        secs = float(secs); neg = secs < 0; secs = abs(secs)
        m = int(secs // 60); sv = secs % 60; si = int(sv)
        cc = min(99, round((sv - si) * 100))
        return ('-' if neg else '') + f'{m}:{si:02d}.{cc:02d}'

    def _sp(p):
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def _run(p, text, pt=PT, bold=False):
        r = p.add_run(text); r.font.name = TNR; r.font.size = Pt(pt)
        if bold: r.font.bold = True
        return r

    def _para(doc, text='', align=None, pt=PT, bold=False):
        p = doc.add_paragraph(); _sp(p)
        if align: p.alignment = align
        if text: _run(p, text, pt, bold)
        return p

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _borders(cell, sz=4):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _no_border(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
        el = b.find(qn('w:bottom'))
        if el is None:
            el = OxmlElement('w:bottom'); b.append(el)
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _cw(cell, w):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        tcW.set(qn('w:w'), str(int(w))); tcW.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        va = pr.find(qn('w:vAlign'))
        if va is None:
            va = OxmlElement('w:vAlign'); pr.append(va)
        va.set(qn('w:val'), val)

    def _fill(cell, text, align=None, pt=PT, bold=False):
        p = cell.paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p)
        if align: p.alignment = align
        text = str(text).replace('—', '–')  # — → –
        if not text: return
        parts = str(text).split('\n')
        for i, part in enumerate(parts):
            r = p.add_run(part); r.font.name = TNR; r.font.size = Pt(pt)
            if bold: r.font.bold = True
            if i < len(parts) - 1:
                br = OxmlElement('w:br'); r._r.append(br)

    def _cw_pct(cell, dxa):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        pct = round(dxa * 5000 / TW)
        tcW.set(qn('w:w'), str(pct)); tcW.set(qn('w:type'), 'pct')

    def _make_tbl(ncols, fixed_tw=None, pct_w=5000, jc=None):
        t = doc.add_table(rows=0, cols=ncols)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        if fixed_tw:
            el.set(qn('w:w'), str(fixed_tw)); el.set(qn('w:type'), 'dxa')
        else:
            el.set(qn('w:w'), str(pct_w)); el.set(qn('w:type'), 'pct')
        if jc == 'center':
            for old_ind in tblPr.findall(qn('w:tblInd')): tblPr.remove(old_ind)
        else:
            ind = tblPr.find(qn('w:tblInd'))
            if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
            ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        if jc:
            jc_el = tblPr.find(qn('w:jc'))
            if jc_el is None: jc_el = OxmlElement('w:jc'); tblPr.append(jc_el)
            jc_el.set(qn('w:val'), jc)
        return t

    def _set_grid_sp(tbl_obj, widths):
        tbl = tbl_obj._tbl
        old = tbl.find(qn('w:tblGrid'))
        if old is not None: tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in widths:
            gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); grid.append(gc)
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is not None: tblPr.addnext(grid)
        else: tbl.insert(0, grid)

    def _row_h(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'exact')
        pr.append(rh)

    def _row_h_min(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'atLeast')
        pr.append(rh)

    def _vmerge_restart(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); vm.set(qn('w:val'), 'restart'); pr.append(vm)

    def _vmerge_cont(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); pr.append(vm)

    def _tcMar(cell, lr=28):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcMar')): pr.remove(old)
        mar = OxmlElement('w:tcMar')
        for side in ('left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:w'), str(lr)); el.set(qn('w:type'), 'dxa')
            mar.append(el)
        pr.append(mar)

    def _no_wrap_cell(cell):
        pr = _tcpr(cell)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))

    def _line_br(p):
        p.add_run()._r.append(OxmlElement('w:br'))

    def _vert_cell(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:textDirection')): pr.remove(old)
        td = OxmlElement('w:textDirection'); td.set(qn('w:val'), 'btLr'); pr.append(td)
        _valign(cell, 'center')
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def _ls15(p):
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:line'), '360'); sp.set(qn('w:lineRule'), 'auto')

    def _keep_next_p(p):
        pPr = p._p.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr'); p._p.insert(0, pPr)
        if pPr.find(qn('w:keepNext')) is None:
            pPr.append(OxmlElement('w:keepNext'))

    # ── flags ────────────────────────────────────────────────────────────────
    муж_норм = (len(муж_рез) + len(муж_без)) >= 6
    жен_норм = (len(жен_рез) + len(жен_без)) >= 6
    снос = ((len(муж_рез) + len(муж_без)) > 0 and not муж_норм) or \
           ((len(жен_рез) + len(жен_без)) > 0 and not жен_норм)

    _коды = {'R4': 'СПРИНТ – R4, код ВРВС 1740091811Я',
             'R6': 'СПРИНТ – R6, код ВРВС 1740101811Я'}
    дисциплина = _коды.get(класс, 'СПРИНТ' + (f' – {класс}' if класс else ''))

    # ── document ─────────────────────────────────────────────────────────────
    from docx.shared import Mm
    doc  = Document()
    sect = doc.sections[0]
    if портрет:
        sect.page_height   = Mm(297); sect.page_width    = Mm(210)
    else:
        sect.page_height   = Mm(210); sect.page_width    = Mm(297)
    sect.top_margin    = Mm(12.5)
    sect.bottom_margin = Mm(12.5)
    sect.left_margin   = Mm(15)
    sect.right_margin  = Mm(10)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if портрет:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(10)
    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    _hdr_pt = 10 if портрет else 12
    if сор.организаторы_блок:
        _para(doc, сор.организаторы_блок, align=C, pt=_hdr_pt, bold=True)
    p0 = _para(doc); p0.paragraph_format.line_spacing = Pt(4)
    _ls15(_para(doc, сор.название or '', align=C, pt=12, bold=True))
    _ls15(_para(doc, 'ПРОТОКОЛ РЕЗУЛЬТАТОВ', align=C, pt=_hdr_pt, bold=True))
    if сор.номер_екп:
        _ls15(_para(doc, f'ЕКП № {сор.номер_екп}', align=C, pt=_hdr_pt, bold=True))
    _ls15(_para(doc, f'Спортивная дисциплина: {дисциплина}', align=C, pt=_hdr_pt, bold=True))

    # ── таблица дата/место ───────────────────────────────────────────────────
    dt = _make_tbl(3, fixed_tw=TW, jc='center')
    _cw_char = round(PT * 0.6 * 20)
    _w_dt  = max(100, len(сор.даты_проведения or 'Дата проведения') * _cw_char + 2 * MAR)
    _w_pl  = max(100, len(сор.место_проведения or 'Место проведения') * _cw_char + 2 * MAR)
    _w_mid = max(100, TW - _w_dt - _w_pl)
    _set_grid_sp(dt, [_w_dt, _w_mid, _w_pl])
    _dt_rh = (lambda r: _row_h(r, 0.3)) if портрет else (lambda r: _row_h(r, 0.4))
    _dt_rows = [
        (сор.даты_проведения or '', сор.место_проведения or '', PT, True),
        ('Дата проведения', 'Место проведения', PT if not портрет else 7, not портрет),
    ]
    for idx, (лево, право, pt_v, bold_v) in enumerate(_dt_rows):
        r = dt.add_row(); _dt_rh(r)
        cs = list(r.cells)
        for c in cs:
            _no_border(c)
            if not портрет:
                _valign(c, 'bottom' if idx == 0 else 'top')
            else:
                _valign(c)
        _no_wrap_cell(cs[0])
        _no_wrap_cell(cs[2])
        if idx == 0:
            _bottom_border(cs[0])
            _bottom_border(cs[2])
        _fill(cs[0], лево, align=C, pt=pt_v, bold=bold_v)
        _fill(cs[2], право, align=C, pt=pt_v, bold=bold_v)

    # ── пустая строка перед таблицей (1пт) ──────────────────────────────────
    p = _para(doc)
    p.paragraph_format.line_spacing = Pt(1)

    # ── main table ───────────────────────────────────────────────────────────
    tbl = _make_tbl(NC, fixed_tw=TW, jc='center')
    _set_grid_sp(tbl, GRID)

    # header — одна строка
    r = tbl.add_row(); _row_h_min(r, 0.9 if портрет else 0.6)
    cs = list(r.cells)
    for i, w in enumerate(GRID): _cw_pct(cs[i], w)
    _fill(cs[I_НПП], '№\nп/п',      align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_К],   'Команда',      align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_НОМ], 'Номер',        align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_СОС], 'Состав',       align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_СУБ], 'Субъект РФ',  align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ШТИ], 'Штраф',       align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ВП],  'Время',        align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ИТ],  'Результат',    align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_МЕ],  'Место',        align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ОЧ],  'Очки',         align=C, pt=PT_HDR, bold=True)
    p_ра = cs[I_РА].paragraphs[0]; p_ра.clear(); _sp(p_ра); p_ра.alignment = C
    _run(p_ра, 'Выполнение', PT_HDR, True); _line_br(p_ра)
    _run(p_ра, 'разряда' + ('*' if снос else ''), PT_HDR, True)
    for i in range(NC):
        _borders(cs[i]); _valign(cs[i])
    for i in range(NC):
        _tcMar(cs[i], MAR); _no_wrap_cell(cs[i])

    # ── section/data helpers ─────────────────────────────────────────────────
    def _section(label):
        r = tbl.add_row(); _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[NC - 1]); _cw_pct(cs[0], TW)
        _borders(cs[0]); _valign(cs[0])
        _fill(cs[0], label, align=C, pt=PT, bold=True)
        _keep_next_p(cs[0].paragraphs[-1])

    def _data_row(пп, рез, разряды, норм):
        from models import СТАТУС_ДК, СТАТУС_ДНС, СТАТУС_ДНФ, СТАТУС_ПМ, СТАТУС_ПЛВ
        к = рез.команда
        состав = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
        фио = ', '.join(
            f'{с.фамилия_имя_отчество} ({с.дата_рождения.strftime("%d.%m.%Y") if с.дата_рождения else "б/д"}, {с.разряд or "б/р"})'
            for с in состав) if состав else '–'
        шт_ст = (рез.штраф_с10 or 0) * 10 + (рез.штраф_с50 or 0) * 50
        шт_фн = (рез.штраф_ф50 or 0) * 50

        r = tbl.add_row(); _row_h_min(r, 0.5)
        cs = list(r.cells)
        for i, w in enumerate(GRID): _cw_pct(cs[i], w)
        for i in range(NC):
            _borders(cs[i]); _valign(cs[i]); _tcMar(cs[i], MAR)

        _fill(cs[I_НПП], str(пп),                        align=C, pt=PT)
        _fill(cs[I_К],   к.название or '–',              align=L, pt=PT)
        _fill(cs[I_НОМ], str(к.нагрудный_номер or ''),   align=C, pt=PT)
        _fill(cs[I_СОС], фио,                             align=L, pt=PT_FIO)
        _fill(cs[I_СУБ], к.субъект.название if к.субъект else '–', align=C, pt=PT)

        ст = рез.статус
        шт_итог = (шт_ст or 0) + (шт_фн or 0)
        if ст in (СТАТУС_ДНС, СТАТУС_ДНФ):
            _fill(cs[I_ШТИ], '–', align=C, pt=PT)
            _fill(cs[I_ВП],  'Не старт.' if ст == СТАТУС_ДНС else 'Не фин.', align=C, pt=PT)
            _fill(cs[I_ИТ],  '–', align=C, pt=PT)
        elif ст == СТАТУС_ДК:
            _fill(cs[I_ШТИ], _fmt_dur(рез.штрафное_время) or '0:00.00', align=C, pt=PT)
            _fill(cs[I_ВП],  _fmt_dur(рез.время_прохождения), align=C, pt=PT)
            _fill(cs[I_ИТ],  'ДК', align=C, pt=PT)
        else:
            _fill(cs[I_ШТИ], _fmt_dur(рез.штрафное_время) or '0:00.00', align=C, pt=PT)
            _fill(cs[I_ВП],  _fmt_dur(рез.время_прохождения), align=C, pt=PT)
            if ст == СТАТУС_ПМ:    итог = 'Посл. м.'
            elif ст == СТАТУС_ПЛВ: итог = 'Прев. ЛВ'
            else:                   итог = _fmt_dur(рез.результат)
            _fill(cs[I_ИТ], итог, align=C, pt=PT)

        _fill(cs[I_МЕ], str(рез.место) if рез.место else '–', align=C, pt=PT)
        _fill(cs[I_ОЧ], str(int(рез.очки)) if рез.очки else '–', align=C, pt=PT)
        _fill(cs[I_РА], разряды.get(рез.команда_id, '–') if норм else '–', align=C, pt=PT)

    # ── data sections ────────────────────────────────────────────────────────
    _section('Мужчины')
    for пп, рез in enumerate(муж_рез + муж_без, 1):
        _data_row(пп, рез, муж_разр, муж_норм)

    _section('Женщины')
    for пп, рез in enumerate(жен_рез + жен_без, 1):
        _data_row(пп, рез, жен_разр, жен_норм)

    # ── footnote ─────────────────────────────────────────────────────────────
    if снос:
        _para(doc, '* Разряды не присваиваются, т.к. количество участников в виде программы '
              'недостаточно для присвоения в соответствии с положением о ЕВСК (пункт 25)',
              align=L, pt=PT_FIO, bold=True)

    # ── signature table ──────────────────────────────────────────────────────
    sig = _make_tbl(2, fixed_tw=TW, jc='center')
    _set_grid_sp(sig, [W_SIG_L, W_SIG_R])

    for позиция, фио in (('Главный судья',    сор.главный_судья     or ''),
                          ('Главный секретарь', сор.главный_секретарь or '')):
        r = sig.add_row(); _row_h(r, 0.88)
        cs = list(r.cells)
        _cw(cs[0], W_SIG_L); _cw(cs[1], W_SIG_R)
        _no_border(cs[0]); _no_border(cs[1])
        _valign(cs[0]); _valign(cs[1])
        _fill(cs[0], позиция, align=L, pt=PT, bold=True)
        _fill(cs[1], фио,    align=L, pt=PT, bold=True)


    buf = BytesIO()
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf); buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)



def create_результаты_итог_дг_docx(сор, муж_рез, муж_без, жен_рез, жен_без,
                                     муж_разр, жен_разр, класс, filename, портрет=False):
    """Итоговый протокол результатов Длинная гонка — 11 колонок."""
    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT

    TW = round((210 - 15 - 10) * 1440 / 25.4) if портрет else round((297 - 15 - 10) * 1440 / 25.4)
    if портрет:
        _PCTS = [3, 12, 6, 29, 10, 6, 6, 8, 5, 5, 10]
    else:
        _PCTS = [4.4, 13.8, 5.5, 29.1, 10.2, 5.5, 5.7, 7.3, 5.1, 4.6, 8.8]
    GRID = [round(TW * p / 100) for p in _PCTS]
    GRID[-1] += TW - sum(GRID)
    NC = 11

    I_НПП, I_К,   I_НОМ, I_СОС, I_СУБ = 0, 1, 2, 3, 4
    I_ШТИ, I_ВП,  I_ИТ                  = 5, 6, 7
    I_МЕ,  I_ОЧ,  I_РА                  = 8, 9, 10

    W_SIG_L = TW * 35 // 100
    W_SIG_R = TW - W_SIG_L


    if портрет:
        PT, PT_FIO, PT_HDR = 8, 7, 8
    else:
        PT, PT_FIO, PT_HDR = 10, 8, 10
    MAR = 28

    def _fmt_dur(secs):
        if secs is None: return '–'
        secs = float(secs); neg = secs < 0; secs = abs(secs)
        m = int(secs // 60); sv = secs % 60; si = int(sv)
        cc = min(99, round((sv - si) * 100))
        return ('-' if neg else '') + f'{m}:{si:02d}.{cc:02d}'

    def _sp(p):
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def _run(p, text, pt=PT, bold=False):
        r = p.add_run(text); r.font.name = TNR; r.font.size = Pt(pt)
        if bold: r.font.bold = True
        return r

    def _para(doc, text='', align=None, pt=PT, bold=False):
        p = doc.add_paragraph(); _sp(p)
        if align: p.alignment = align
        if text: _run(p, text, pt, bold)
        return p

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _borders(cell, sz=4):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _no_border(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
        el = b.find(qn('w:bottom'))
        if el is None:
            el = OxmlElement('w:bottom'); b.append(el)
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _cw(cell, w):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        tcW.set(qn('w:w'), str(int(w))); tcW.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        va = pr.find(qn('w:vAlign'))
        if va is None:
            va = OxmlElement('w:vAlign'); pr.append(va)
        va.set(qn('w:val'), val)

    def _fill(cell, text, align=None, pt=PT, bold=False):
        p = cell.paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p)
        if align: p.alignment = align
        text = str(text).replace('—', '–')
        if not text: return
        parts = str(text).split('\n')
        for i, part in enumerate(parts):
            r = p.add_run(part); r.font.name = TNR; r.font.size = Pt(pt)
            if bold: r.font.bold = True
            if i < len(parts) - 1:
                br = OxmlElement('w:br'); r._r.append(br)

    def _cw_pct(cell, dxa):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        pct = round(dxa * 5000 / TW)
        tcW.set(qn('w:w'), str(pct)); tcW.set(qn('w:type'), 'pct')

    def _make_tbl(ncols, fixed_tw=None, pct_w=5000, jc=None):
        t = doc.add_table(rows=0, cols=ncols)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        if fixed_tw:
            el.set(qn('w:w'), str(fixed_tw)); el.set(qn('w:type'), 'dxa')
        else:
            el.set(qn('w:w'), str(pct_w)); el.set(qn('w:type'), 'pct')
        if jc == 'center':
            for old_ind in tblPr.findall(qn('w:tblInd')): tblPr.remove(old_ind)
        else:
            ind = tblPr.find(qn('w:tblInd'))
            if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
            ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        if jc:
            jc_el = tblPr.find(qn('w:jc'))
            if jc_el is None: jc_el = OxmlElement('w:jc'); tblPr.append(jc_el)
            jc_el.set(qn('w:val'), jc)
        return t

    def _set_grid_sp(tbl_obj, widths):
        tbl = tbl_obj._tbl
        old = tbl.find(qn('w:tblGrid'))
        if old is not None: tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in widths:
            gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); grid.append(gc)
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is not None: tblPr.addnext(grid)
        else: tbl.insert(0, grid)

    def _row_h(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'exact')
        pr.append(rh)

    def _row_h_min(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'atLeast')
        pr.append(rh)

    def _tcMar(cell, lr=28):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcMar')): pr.remove(old)
        mar = OxmlElement('w:tcMar')
        for side in ('left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:w'), str(lr)); el.set(qn('w:type'), 'dxa')
            mar.append(el)
        pr.append(mar)

    def _no_wrap_cell(cell):
        pr = _tcpr(cell)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))

    def _line_br(p):
        p.add_run()._r.append(OxmlElement('w:br'))

    def _ls15(p):
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:line'), '360'); sp.set(qn('w:lineRule'), 'auto')

    def _keep_next_p(p):
        pPr = p._p.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr'); p._p.insert(0, pPr)
        if pPr.find(qn('w:keepNext')) is None:
            pPr.append(OxmlElement('w:keepNext'))

    муж_норм = (len(муж_рез) + len(муж_без)) >= 6
    жен_норм = (len(жен_рез) + len(жен_без)) >= 6
    снос = ((len(муж_рез) + len(муж_без)) > 0 and not муж_норм) or \
           ((len(жен_рез) + len(жен_без)) > 0 and not жен_норм)

    _коды_дг = {'R4': '1740011811Я', 'R6': '1740021811Я'}
    _кл_дг = класс or сор.класс_судна or 'R4/R6'
    _код_дг = _коды_дг.get(_кл_дг, сор.код_соревнований or '')
    дисциплина = f'ДЛИННАЯ ГОНКА – {_кл_дг}, код ВРВС {_код_дг}'

    from docx.shared import Mm
    doc  = Document()
    sect = doc.sections[0]
    if портрет:
        sect.page_height   = Mm(297); sect.page_width    = Mm(210)
    else:
        sect.page_height   = Mm(210); sect.page_width    = Mm(297)
    sect.top_margin    = Mm(12.5)
    sect.bottom_margin = Mm(12.5)
    sect.left_margin   = Mm(15)
    sect.right_margin  = Mm(10)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if портрет:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(10)
    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    _hdr_pt = 10 if портрет else 12
    if сор.организаторы_блок:
        _para(doc, сор.организаторы_блок, align=C, pt=_hdr_pt, bold=True)
    p0 = _para(doc); p0.paragraph_format.line_spacing = Pt(4)
    _ls15(_para(doc, сор.название or '', align=C, pt=12, bold=True))
    _ls15(_para(doc, 'ПРОТОКОЛ РЕЗУЛЬТАТОВ', align=C, pt=_hdr_pt, bold=True))
    if сор.номер_екп:
        _ls15(_para(doc, f'ЕКП № {сор.номер_екп}', align=C, pt=_hdr_pt, bold=True))
    _ls15(_para(doc, f'Спортивная дисциплина: {дисциплина}', align=C, pt=_hdr_pt, bold=True))

    dt = _make_tbl(3, fixed_tw=TW, jc='center')
    _cw_char = round(PT * 0.6 * 20)
    _w_dt  = max(100, len(сор.даты_проведения or 'Дата проведения') * _cw_char + 2 * MAR)
    _w_pl  = max(100, len(сор.место_проведения or 'Место проведения') * _cw_char + 2 * MAR)
    _w_mid = max(100, TW - _w_dt - _w_pl)
    _set_grid_sp(dt, [_w_dt, _w_mid, _w_pl])
    _dt_rh = (lambda r: _row_h(r, 0.3)) if портрет else (lambda r: _row_h(r, 0.4))
    _dt_rows = [
        (сор.даты_проведения or '', сор.место_проведения or '', PT, True),
        ('Дата проведения', 'Место проведения', PT if not портрет else 7, not портрет),
    ]
    for idx, (лево, право, pt_v, bold_v) in enumerate(_dt_rows):
        r = dt.add_row(); _dt_rh(r)
        cs = list(r.cells)
        for c in cs:
            _no_border(c)
            if not портрет:
                _valign(c, 'bottom' if idx == 0 else 'top')
            else:
                _valign(c)
        _no_wrap_cell(cs[0])
        _no_wrap_cell(cs[2])
        if idx == 0:
            _bottom_border(cs[0])
            _bottom_border(cs[2])
        _fill(cs[0], лево, align=C, pt=pt_v, bold=bold_v)
        _fill(cs[2], право, align=C, pt=pt_v, bold=bold_v)

    p = _para(doc)
    p.paragraph_format.line_spacing = Pt(1)

    tbl = _make_tbl(NC, fixed_tw=TW, jc='center')
    _set_grid_sp(tbl, GRID)

    r = tbl.add_row(); _row_h_min(r, 0.9 if портрет else 0.6)
    cs = list(r.cells)
    for i, w in enumerate(GRID): _cw_pct(cs[i], w)
    _fill(cs[I_НПП], '№\nп/п',      align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_К],   'Команда',           align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_НОМ], 'Номер',             align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_СОС], 'Состав',            align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_СУБ], 'Субъект РФ',        align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ШТИ], 'Штраф',            align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ВП],  'Время',             align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ИТ],  'Результат',         align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_МЕ],  'Место',             align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ОЧ],  'Очки',              align=C, pt=PT_HDR, bold=True)
    p_ра = cs[I_РА].paragraphs[0]; p_ра.clear(); _sp(p_ра); p_ра.alignment = C
    _run(p_ра, 'Выполнение', PT_HDR, True); _line_br(p_ра)
    _run(p_ра, 'разряда' + ('*' if снос else ''), PT_HDR, True)
    for i in range(NC):
        _borders(cs[i]); _valign(cs[i])
    for i in range(NC):
        _tcMar(cs[i], MAR); _no_wrap_cell(cs[i])

    def _section(label):
        r = tbl.add_row(); _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[NC - 1]); _cw_pct(cs[0], TW)
        _borders(cs[0]); _valign(cs[0])
        _fill(cs[0], label, align=C, pt=PT, bold=True)
        _keep_next_p(cs[0].paragraphs[-1])

    def _data_row(пп, рез, разряды, норм):
        from models import СТАТУС_ДК, СТАТУС_ДНС, СТАТУС_ДНФ, СТАТУС_ПМ, СТАТУС_ПЛВ
        к = рез.команда
        состав = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
        фио = ', '.join(
            f'{с.фамилия_имя_отчество} ({с.дата_рождения.strftime("%d.%m.%Y") if с.дата_рождения else "б/д"}, {с.разряд or "б/р"})'
            for с in состав) if состав else '–'

        r = tbl.add_row(); _row_h_min(r, 0.5)
        cs = list(r.cells)
        for i, w in enumerate(GRID): _cw_pct(cs[i], w)
        for i in range(NC):
            _borders(cs[i]); _valign(cs[i]); _tcMar(cs[i], MAR)

        _fill(cs[I_НПП], str(пп),                        align=C, pt=PT)
        _fill(cs[I_К],   к.название or '–',          align=L, pt=PT)
        _fill(cs[I_НОМ], str(к.нагрудный_номер or ''),   align=C, pt=PT)
        _fill(cs[I_СОС], фио,                             align=L, pt=PT_FIO)
        _fill(cs[I_СУБ], к.субъект.название if к.субъект else '–', align=C, pt=PT)

        ст = рез.статус
        if ст in (СТАТУС_ДНС, СТАТУС_ДНФ):
            _fill(cs[I_ШТИ], '–', align=C, pt=PT)
            _fill(cs[I_ВП],  'Не старт.' if ст == СТАТУС_ДНС else 'Не фин.', align=C, pt=PT)
            _fill(cs[I_ИТ],  '–', align=C, pt=PT)
        elif ст == СТАТУС_ДК:
            _fill(cs[I_ШТИ], _fmt_dur(рез.штрафное_время) or '0:00.00', align=C, pt=PT)
            _fill(cs[I_ВП],  _fmt_dur(рез.время_прохождения), align=C, pt=PT)
            _fill(cs[I_ИТ],  'ДК', align=C, pt=PT)
        else:
            _fill(cs[I_ШТИ], _fmt_dur(рез.штрафное_время) or '0:00.00', align=C, pt=PT)
            _fill(cs[I_ВП],  _fmt_dur(рез.время_прохождения), align=C, pt=PT)
            if ст == СТАТУС_ПМ:    итог = 'Посл. м.'
            elif ст == СТАТУС_ПЛВ: итог = 'Прев. ЛВ'
            else:                   итог = _fmt_dur(рез.результат)
            _fill(cs[I_ИТ], итог, align=C, pt=PT)

        _fill(cs[I_МЕ], str(рез.место) if рез.место else '–', align=C, pt=PT)
        _fill(cs[I_ОЧ], str(int(рез.очки)) if рез.очки else '–', align=C, pt=PT)
        _fill(cs[I_РА], разряды.get(рез.команда_id, '–') if норм else '–', align=C, pt=PT)

    _section('Мужчины')
    for пп, рез in enumerate(муж_рез + муж_без, 1):
        _data_row(пп, рез, муж_разр, муж_норм)

    _section('Женщины')
    for пп, рез in enumerate(жен_рез + жен_без, 1):
        _data_row(пп, рез, жен_разр, жен_норм)

    if снос:
        _para(doc, '* Разряды не присваиваются, т.к. количество участников в виде программы '
              'недостаточно для присвоения в соответствии с положением о ЕВСК (пункт 25)',
              align=L, pt=PT_FIO, bold=True)

    sig = _make_tbl(2, fixed_tw=TW, jc='center')
    _set_grid_sp(sig, [W_SIG_L, W_SIG_R])

    for позиция, фио in (('Главный судья',    сор.главный_судья     or ''),
                          ('Главный секретарь', сор.главный_секретарь or '')):
        r = sig.add_row(); _row_h(r, 0.88)
        cs = list(r.cells)
        _cw(cs[0], W_SIG_L); _cw(cs[1], W_SIG_R)
        _no_border(cs[0]); _no_border(cs[1])
        _valign(cs[0]); _valign(cs[1])
        _fill(cs[0], позиция, align=L, pt=PT, bold=True)
        _fill(cs[1], фио,    align=L, pt=PT, bold=True)

    buf = BytesIO()
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf); buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_результаты_итог_пс_docx(сор, муж_итог, жен_итог, муж_разр, жен_разр,
                                     класс, filename, портрет=False):
    """Итоговый протокол результатов Параллельный спринт — 8 колонок (без Штраф/Время/Результат)."""
    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT

    TW = round((210 - 10 - 5) * 1440 / 25.4) if портрет else round((297 - 10 - 5) * 1440 / 25.4)
    if портрет:
        _PCTS = [3, 12, 6, 46, 12, 6, 6, 9]   # sum=100
    else:
        _PCTS = [3, 12, 6, 49, 12, 5, 5, 8]   # sum=100
    GRID = [round(TW * p / 100) for p in _PCTS]
    GRID[-1] += TW - sum(GRID)
    NC = 8

    I_НПП, I_К, I_НОМ, I_СОС, I_СУБ, I_МЕ, I_ОЧ, I_РА = range(8)

    TW_SIG  = round((210 - 15 - 10) * 1440 / 25.4)
    W_SIG_L = TW_SIG * 35 // 100
    W_SIG_R = TW_SIG - W_SIG_L

    if портрет:
        PT, PT_FIO, PT_HDR = 8, 7, 8
    else:
        PT, PT_FIO, PT_HDR = 10, 8, 10
    MAR = 28

    def _sp(p):
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def _run(p, text, pt=PT, bold=False):
        r = p.add_run(text); r.font.name = TNR; r.font.size = Pt(pt)
        if bold: r.font.bold = True
        return r

    def _para(doc, text='', align=None, pt=PT, bold=False):
        p = doc.add_paragraph(); _sp(p)
        if align: p.alignment = align
        if text: _run(p, text, pt, bold)
        return p

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _borders(cell, sz=4):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _no_border(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
        el = b.find(qn('w:bottom'))
        if el is None:
            el = OxmlElement('w:bottom'); b.append(el)
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _cw(cell, w):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        tcW.set(qn('w:w'), str(int(w))); tcW.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        va = pr.find(qn('w:vAlign'))
        if va is None:
            va = OxmlElement('w:vAlign'); pr.append(va)
        va.set(qn('w:val'), val)

    def _fill(cell, text, align=None, pt=PT, bold=False):
        p = cell.paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p)
        if align: p.alignment = align
        text = str(text).replace('—', '–')
        if not text: return
        parts = str(text).split('\n')
        for i, part in enumerate(parts):
            r = p.add_run(part); r.font.name = TNR; r.font.size = Pt(pt)
            if bold: r.font.bold = True
            if i < len(parts) - 1:
                br = OxmlElement('w:br'); r._r.append(br)

    def _cw_pct(cell, dxa):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        pct = round(dxa * 5000 / TW)
        tcW.set(qn('w:w'), str(pct)); tcW.set(qn('w:type'), 'pct')

    def _make_tbl(ncols, fixed_tw=None, pct_w=5000, jc=None):
        t = doc.add_table(rows=0, cols=ncols)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        if fixed_tw:
            el.set(qn('w:w'), str(fixed_tw)); el.set(qn('w:type'), 'dxa')
        else:
            el.set(qn('w:w'), str(pct_w)); el.set(qn('w:type'), 'pct')
        if jc == 'center':
            for old_ind in tblPr.findall(qn('w:tblInd')): tblPr.remove(old_ind)
        else:
            ind = tblPr.find(qn('w:tblInd'))
            if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
            ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        if jc:
            jc_el = tblPr.find(qn('w:jc'))
            if jc_el is None: jc_el = OxmlElement('w:jc'); tblPr.append(jc_el)
            jc_el.set(qn('w:val'), jc)
        return t

    def _set_grid_sp(tbl_obj, widths):
        tbl = tbl_obj._tbl
        old = tbl.find(qn('w:tblGrid'))
        if old is not None: tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in widths:
            gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); grid.append(gc)
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is not None: tblPr.addnext(grid)
        else: tbl.insert(0, grid)

    def _row_h(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'exact')
        pr.append(rh)

    def _row_h_min(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'atLeast')
        pr.append(rh)

    def _tcMar(cell, lr=28):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcMar')): pr.remove(old)
        mar = OxmlElement('w:tcMar')
        for side in ('left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:w'), str(lr)); el.set(qn('w:type'), 'dxa')
            mar.append(el)
        pr.append(mar)

    def _no_wrap_cell(cell):
        pr = _tcpr(cell)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))

    def _line_br(p):
        p.add_run()._r.append(OxmlElement('w:br'))

    def _ls15(p):
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:line'), '360'); sp.set(qn('w:lineRule'), 'auto')

    def _keep_next_p(p):
        pPr = p._p.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr'); p._p.insert(0, pPr)
        if pPr.find(qn('w:keepNext')) is None:
            pPr.append(OxmlElement('w:keepNext'))

    муж_с = [р for р in муж_итог if р.итоговое_место]
    муж_без = [р for р in муж_итог if not р.итоговое_место]
    жен_с = [р for р in жен_итог if р.итоговое_место]
    жен_без = [р for р in жен_итог if not р.итоговое_место]

    муж_норм = (len(муж_с) + len(муж_без)) >= 6
    жен_норм = (len(жен_с) + len(жен_без)) >= 6
    снос = ((len(муж_с) + len(муж_без)) > 0 and not муж_норм) or \
           ((len(жен_с) + len(жен_без)) > 0 and not жен_норм)

    _коды_пс = {'R4': '1740051811Я', 'R6': '1740061811Я'}
    _кл_пс = класс or сор.класс_судна or 'R4/R6'
    _код_пс = _коды_пс.get(_кл_пс, сор.код_соревнований or '')
    дисциплина = f'ПАРАЛЛЕЛЬНЫЙ СПРИНТ – {_кл_пс}, код ВРВС {_код_пс}'

    from docx.shared import Mm
    doc  = Document()
    sect = doc.sections[0]
    if портрет:
        sect.page_height   = Mm(297); sect.page_width    = Mm(210)
    else:
        sect.page_height   = Mm(210); sect.page_width    = Mm(297)
    sect.top_margin    = Mm(15)
    sect.bottom_margin = Mm(15)
    sect.left_margin   = Mm(10)
    sect.right_margin  = Mm(5)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if портрет:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(10)
    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    _hdr_pt = 10 if портрет else 12
    if сор.организаторы_блок:
        _para(doc, сор.организаторы_блок, align=C, pt=_hdr_pt, bold=True)
    p0 = _para(doc); p0.paragraph_format.line_spacing = Pt(4)
    _ls15(_para(doc, сор.название or '', align=C, pt=12, bold=True))
    _ls15(_para(doc, 'ПРОТОКОЛ РЕЗУЛЬТАТОВ', align=C, pt=_hdr_pt, bold=True))
    if сор.номер_екп:
        _ls15(_para(doc, f'ЕКП № {сор.номер_екп}', align=C, pt=_hdr_pt, bold=True))
    _ls15(_para(doc, f'Спортивная дисциплина: {дисциплина}', align=C, pt=_hdr_pt, bold=True))

    dt = _make_tbl(3, fixed_tw=TW, jc='center')
    _cw_char = round(PT * 0.6 * 20)
    _w_dt  = max(100, len(сор.даты_проведения or 'Дата проведения') * _cw_char + 2 * MAR)
    _w_pl  = max(100, len(сор.место_проведения or 'Место проведения') * _cw_char + 2 * MAR)
    _w_mid = max(100, TW - _w_dt - _w_pl)
    _set_grid_sp(dt, [_w_dt, _w_mid, _w_pl])
    _dt_rh = (lambda r: _row_h(r, 0.3)) if портрет else (lambda r: _row_h(r, 0.4))
    _dt_rows = [
        (сор.даты_проведения or '', сор.место_проведения or '', PT, True),
        ('Дата проведения', 'Место проведения', PT if not портрет else 7, not портрет),
    ]
    for idx, (лево, право, pt_v, bold_v) in enumerate(_dt_rows):
        r = dt.add_row(); _dt_rh(r)
        cs = list(r.cells)
        for c in cs:
            _no_border(c)
            if not портрет:
                _valign(c, 'bottom' if idx == 0 else 'top')
            else:
                _valign(c)
        _no_wrap_cell(cs[0])
        _no_wrap_cell(cs[2])
        if idx == 0:
            _bottom_border(cs[0])
            _bottom_border(cs[2])
        _fill(cs[0], лево, align=C, pt=pt_v, bold=bold_v)
        _fill(cs[2], право, align=C, pt=pt_v, bold=bold_v)

    p = _para(doc)
    p.paragraph_format.line_spacing = Pt(1)

    tbl = _make_tbl(NC, fixed_tw=TW, jc='center')
    _set_grid_sp(tbl, GRID)

    r = tbl.add_row(); _row_h_min(r, 0.9 if портрет else 0.6)
    cs = list(r.cells)
    for i, w in enumerate(GRID): _cw_pct(cs[i], w)
    _fill(cs[I_НПП], '№\nп/п',     align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_К],   'Команда',     align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_НОМ], 'Номер',       align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_СОС], 'Состав',      align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_СУБ], 'Субъект РФ', align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_МЕ],  'Место',       align=C, pt=PT_HDR, bold=True)
    _fill(cs[I_ОЧ],  'Очки',        align=C, pt=PT_HDR, bold=True)
    p_ра = cs[I_РА].paragraphs[0]; p_ра.clear(); _sp(p_ра); p_ра.alignment = C
    _run(p_ра, 'Выполнение', PT_HDR, True); _line_br(p_ра)
    _run(p_ра, 'разряда' + ('*' if снос else ''), PT_HDR, True)
    for i in range(NC):
        _borders(cs[i]); _valign(cs[i])
    for i in range(NC):
        _tcMar(cs[i], MAR); _no_wrap_cell(cs[i])

    def _section(label):
        r = tbl.add_row(); _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[NC - 1]); _cw_pct(cs[0], TW)
        _borders(cs[0]); _valign(cs[0])
        _fill(cs[0], label, align=C, pt=PT, bold=True)
        _keep_next_p(cs[0].paragraphs[-1])

    def _data_row(пп, рез, разряды, норм):
        к = рез.команда
        состав = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
        фио = ', '.join(
            f'{с.фамилия_имя_отчество} ({с.дата_рождения.strftime("%d.%m.%Y") if с.дата_рождения else "б/д"}, {с.разряд or "б/р"})'
            for с in состав) if состав else '–'

        r = tbl.add_row(); _row_h_min(r, 0.5)
        cs = list(r.cells)
        for i, w in enumerate(GRID): _cw_pct(cs[i], w)
        for i in range(NC):
            _borders(cs[i]); _valign(cs[i]); _tcMar(cs[i], MAR)

        _fill(cs[I_НПП], str(пп),                        align=C, pt=PT)
        _fill(cs[I_К],   к.название or '–',          align=L, pt=PT)
        _fill(cs[I_НОМ], str(к.нагрудный_номер or ''),   align=C, pt=PT)
        _fill(cs[I_СОС], фио,                             align=L, pt=PT_FIO)
        _fill(cs[I_СУБ], к.субъект.название if к.субъект else '–', align=C, pt=PT)
        _fill(cs[I_МЕ],  str(рез.итоговое_место) if рез.итоговое_место else '–', align=C, pt=PT)
        _fill(cs[I_ОЧ],  str(int(рез.очки)) if рез.очки else '–', align=C, pt=PT)
        _fill(cs[I_РА],  разряды.get(рез.команда_id, '–') if норм else '–', align=C, pt=PT)

    _section('Мужчины')
    for пп, рез in enumerate(муж_с + муж_без, 1):
        _data_row(пп, рез, муж_разр, муж_норм)

    _section('Женщины')
    for пп, рез in enumerate(жен_с + жен_без, 1):
        _data_row(пп, рез, жен_разр, жен_норм)

    if снос:
        _para(doc, '* Разряды не присваиваются, т.к. количество участников в виде программы '
              'недостаточно для присвоения в соответствии с положением о ЕВСК (пункт 25)',
              align=L, pt=PT_FIO, bold=True)

    sig = _make_tbl(2, fixed_tw=TW, jc='center')
    _set_grid_sp(sig, [W_SIG_L, W_SIG_R])

    for позиция, фио in (('Главный судья',    сор.главный_судья     or ''),
                          ('Главный секретарь', сор.главный_секретарь or '')):
        r = sig.add_row(); _row_h(r, 0.88)
        cs = list(r.cells)
        _cw(cs[0], W_SIG_L); _cw(cs[1], W_SIG_R)
        _no_border(cs[0]); _no_border(cs[1])
        _valign(cs[0]); _valign(cs[1])
        _fill(cs[0], позиция, align=L, pt=PT, bold=True)
        _fill(cs[1], фио,    align=L, pt=PT, bold=True)

    buf = BytesIO()
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf); buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_результаты_пред_пс_docx(сор, муж_итог, жен_итог, муж_разр, жен_разр,
                                     муж_ст, жен_ст, класс, filename, портрет=False):
    """Предварительный протокол результатов Параллельный спринт."""
    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT

    TW = round((210 - 10 - 5) * 1440 / 25.4) if портрет else round((297 - 10 - 5) * 1440 / 25.4)
    if портрет:
        PT, PT_FIO, PT_HDR = 8, 6, 6
        PT_TXT = 6   # Команда, Номер, ФИО, Субъект в данных
    else:
        PT, PT_FIO, PT_HDR = 10, 8, 10
        PT_TXT = PT
    MAR = 28

    # Detect which stages are present
    _все_ст = set()
    for д in муж_ст.values(): _все_ст.update(д.keys())
    for д in жен_ст.values(): _все_ст.update(д.keys())
    ПОРЯДОК_СТАДИЙ_П = ['1/16', '1/8', '1/4', '1/2', 'финалы']
    есть_ст = {
        '1/16':   '1/16' in _все_ст,
        '1/8':    '1/8' in _все_ст,
        '1/4':    '1/4' in _все_ст,
        '1/2':    '1/2' in _все_ст,
        'финалы': 'Финал_А' in _все_ст or 'Финал_Б' in _все_ст,
    }
    ПОРЯДОК_ПОКАЗА = [ст for ст in ПОРЯДОК_СТАДИЙ_П if есть_ст[ст]]
    n_stages = len(ПОРЯДОК_ПОКАЗА)

    I_НПП, I_КОМ, I_НОМ, I_ФИО, I_СУБ, I_МСП = 0, 1, 2, 3, 4, 5
    I_СТА = 6 + 2 * n_stages
    I_ВРМ = 7 + 2 * n_stages
    I_МЕС = 8 + 2 * n_stages
    I_ОЧК = 9 + 2 * n_stages
    I_РАЗ = 10 + 2 * n_stages
    NC    = 11 + 2 * n_stages

    ROWSPAN_COLS = [I_НПП, I_КОМ, I_НОМ, I_ФИО, I_СУБ, I_МСП, I_МЕС, I_ОЧК, I_РАЗ]

    # Column width proportions
    if портрет:
        _P_НПП, _P_КОМ, _P_НОМ, _P_СУБ, _P_МСП = 3.0, 11.3, 3.9, 8.8, 3.9
        _P_ЗА, _P_РЕЗ = 3.9, 6.0
        _P_СТА, _P_ВРМ, _P_МЕС, _P_ОЧК, _P_РАЗ = 4.8, 6.1, 3.9, 3.9, 5.5
    else:
        _P_НПП, _P_КОМ, _P_НОМ, _P_СУБ, _P_МСП = 2.4, 9.9, 3.0, 8.2, 3.2
        _P_ЗА, _P_РЕЗ = 4.37, 6.6
        _P_СТА, _P_ВРМ, _P_МЕС, _P_ОЧК, _P_РАЗ = 5.2, 4.8, 4.6, 4.4, 4.1
    _fix  = _P_НПП + _P_КОМ + _P_НОМ + _P_СУБ + _P_МСП
    _end  = _P_СТА + _P_ВРМ + _P_МЕС + _P_ОЧК + _P_РАЗ
    _P_ФИО = max(5, 100 - _fix - _end - (_P_ЗА + _P_РЕЗ) * n_stages)
    _PCTS = [_P_НПП, _P_КОМ, _P_НОМ, _P_ФИО, _P_СУБ, _P_МСП]
    for _ in range(n_stages):
        _PCTS += [_P_ЗА, _P_РЕЗ]
    _PCTS += [_P_СТА, _P_ВРМ, _P_МЕС, _P_ОЧК, _P_РАЗ]
    _tot = sum(_PCTS)
    GRID = [round(TW * p / _tot) for p in _PCTS]
    GRID[-1] += TW - sum(GRID)

    W_SIG_L = TW * 35 // 100
    W_SIG_R = TW - W_SIG_L

    НАЗВАНИЯ_СТ = {
        '1/16': '1/16 финала', '1/8': '1/8 финала',
        '1/4': '¼ финала', '1/2': '½ финала', 'финалы': 'Финалы',
    }
    АББРЕВ_СТ = {
        'Финал_А': 'ФА', 'Финал_Б': 'ФБ',
        '1/2': '1/2', '1/4': '1/4', '1/8': '1/8', '1/16': '1/16',
    }

    def _sp(p):
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def _run(p, text, pt=PT, bold=False):
        r = p.add_run(text); r.font.name = TNR; r.font.size = Pt(pt)
        if bold: r.font.bold = True
        return r

    def _para(doc, text='', align=None, pt=PT, bold=False):
        p = doc.add_paragraph(); _sp(p)
        if align: p.alignment = align
        if text: _run(p, text, pt, bold)
        return p

    def _line_br(p):
        p.add_run()._r.append(OxmlElement('w:br'))

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _borders(cell, sz=4):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _no_border(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
        el = b.find(qn('w:bottom'))
        if el is None:
            el = OxmlElement('w:bottom'); b.append(el)
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _cw(cell, w):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        tcW.set(qn('w:w'), str(int(w))); tcW.set(qn('w:type'), 'dxa')

    def _cw_pct(cell, dxa):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        pct = round(dxa * 5000 / TW)
        tcW.set(qn('w:w'), str(pct)); tcW.set(qn('w:type'), 'pct')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        va = pr.find(qn('w:vAlign'))
        if va is None:
            va = OxmlElement('w:vAlign'); pr.append(va)
        va.set(qn('w:val'), val)

    def _fill(cell, text, align=None, pt=PT, bold=False):
        p = cell.paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p)
        if align: p.alignment = align
        text = str(text).replace('—', '–')
        if not text: return
        parts = str(text).split('\n')
        for i, part in enumerate(parts):
            r = p.add_run(part); r.font.name = TNR; r.font.size = Pt(pt)
            if bold: r.font.bold = True
            if i < len(parts) - 1:
                br = OxmlElement('w:br'); r._r.append(br)

    def _make_tbl(ncols, fixed_tw=None, jc=None):
        t = doc.add_table(rows=0, cols=ncols)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        if fixed_tw:
            el.set(qn('w:w'), str(fixed_tw)); el.set(qn('w:type'), 'dxa')
        else:
            el.set(qn('w:w'), '5000'); el.set(qn('w:type'), 'pct')
        ind = tblPr.find(qn('w:tblInd'))
        if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
        ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        if jc:
            jc_el = tblPr.find(qn('w:jc'))
            if jc_el is None: jc_el = OxmlElement('w:jc'); tblPr.append(jc_el)
            jc_el.set(qn('w:val'), jc)
        return t

    def _set_grid_sp(tbl_obj, widths):
        tbl = tbl_obj._tbl
        old = tbl.find(qn('w:tblGrid'))
        if old is not None: tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in widths:
            gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); grid.append(gc)
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is not None: tblPr.addnext(grid)
        else: tbl.insert(0, grid)

    def _row_h(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'exact')
        pr.append(rh)

    def _row_h_min(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'atLeast')
        pr.append(rh)

    def _tcMar(cell, lr=28):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcMar')): pr.remove(old)
        mar = OxmlElement('w:tcMar')
        for side in ('left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:w'), str(lr)); el.set(qn('w:type'), 'dxa')
            mar.append(el)
        pr.append(mar)

    def _no_wrap_cell(cell):
        pr = _tcpr(cell)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))

    def _ls15(p):
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:line'), '360'); sp.set(qn('w:lineRule'), 'auto')

    def _keep_next_p(p):
        pPr = p._p.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr'); p._p.insert(0, pPr)
        if pPr.find(qn('w:keepNext')) is None:
            pPr.append(OxmlElement('w:keepNext'))

    def _vmerge_restart(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge')
        vm.set(qn('w:val'), 'restart')
        pr.append(vm)

    def _vmerge_cont(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge')
        pr.append(vm)
        for para in cell.paragraphs:
            para.clear()

    def _fmt_dur(secs):
        if secs is None: return '–'
        secs = float(secs)
        if secs < 0: return '–'
        m = int(secs // 60); sv = secs % 60; si = int(sv)
        cc = min(99, round((sv - si) * 100))
        return f'{m}:{si:02d}.{cc:02d}'

    def _vert_cell(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:textDirection')): pr.remove(old)
        td = OxmlElement('w:textDirection'); td.set(qn('w:val'), 'btLr'); pr.append(td)
        _valign(cell, 'center')
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def _fmt_рез(з):
        if з is None: return '–'
        ст = з.статус
        if ст == 'ДК':  return 'ДК'
        if ст == 'НФН': return 'Не фин.'
        if ст == 'НСТ': return 'Не старт.'
        if ст == 'ПМ':  return 'Посл. м.'
        if ст == 'ПЛВ': return 'Прев. ЛВ'
        return _fmt_dur(з.результат)

    муж_с = [р for р in муж_итог if р.итоговое_место]
    муж_без = [р for р in муж_итог if not р.итоговое_место]
    жен_с = [р for р in жен_итог if р.итоговое_место]
    жен_без = [р for р in жен_итог if not р.итоговое_место]

    муж_норм = (len(муж_с) + len(муж_без)) >= 6
    жен_норм = (len(жен_с) + len(жен_без)) >= 6
    снос = ((len(муж_с) + len(муж_без)) > 0 and not муж_норм) or \
           ((len(жен_с) + len(жен_без)) > 0 and not жен_норм)

    _коды_пс = {'R4': '1740051811Я', 'R6': '1740061811Я'}
    _кл_пс = класс or сор.класс_судна or 'R4/R6'
    _код_пс = _коды_пс.get(_кл_пс, сор.код_соревнований or '')
    дисциплина = f'ПАРАЛЛЕЛЬНЫЙ СПРИНТ – {_кл_пс}, код ВРВС {_код_пс}'

    from docx.shared import Mm
    doc  = Document()
    sect = doc.sections[0]
    if портрет:
        sect.page_height   = Mm(297); sect.page_width    = Mm(210)
    else:
        sect.page_height   = Mm(210); sect.page_width    = Mm(297)
    sect.top_margin    = Mm(15)
    sect.bottom_margin = Mm(15)
    sect.left_margin   = Mm(10)
    sect.right_margin  = Mm(5)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if портрет:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(10)
    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    _hdr_pt  = 8  if портрет else 12
    _name_pt = 10 if портрет else 12
    if сор.организаторы_блок:
        _para(doc, сор.организаторы_блок, align=C, pt=_hdr_pt, bold=True)
    p0 = _para(doc); p0.paragraph_format.line_spacing = Pt(4)
    _ls15(_para(doc, сор.название or '', align=C, pt=_name_pt, bold=True))
    _ls15(_para(doc, 'ПРЕДВАРИТЕЛЬНЫЙ ПРОТОКОЛ РЕЗУЛЬТАТОВ', align=C, pt=_hdr_pt, bold=True))
    _ls15(_para(doc, f'Спортивная дисциплина: {дисциплина}', align=C, pt=_hdr_pt, bold=True))

    dt = _make_tbl(3, fixed_tw=TW, jc='center')
    _cw_char = round(PT * 0.6 * 20)
    _w_dt  = max(100, len(сор.даты_проведения or 'Дата проведения') * _cw_char + 2 * MAR)
    _w_pl  = max(100, len(сор.место_проведения or 'Место проведения') * _cw_char + 2 * MAR)
    _w_mid = max(100, TW - _w_dt - _w_pl)
    _set_grid_sp(dt, [_w_dt, _w_mid, _w_pl])
    _dt_rh = (lambda r: _row_h(r, 0.3)) if портрет else (lambda r: _row_h(r, 0.4))
    _dt_rows = [
        (сор.даты_проведения or '', сор.место_проведения or '', 6 if портрет else PT, True),
        ('Дата проведения', 'Место проведения', 6 if портрет else PT, not портрет),
    ]
    for idx, (лево, право, pt_v, bold_v) in enumerate(_dt_rows):
        r = dt.add_row(); _dt_rh(r)
        cs = list(r.cells)
        for c in cs:
            _no_border(c)
            _valign(c, 'bottom' if (not портрет and idx == 0) else ('top' if (not портрет and idx == 1) else 'center'))
        _no_wrap_cell(cs[0]); _no_wrap_cell(cs[2])
        if idx == 0:
            _bottom_border(cs[0]); _bottom_border(cs[2])
        _fill(cs[0], лево, align=C, pt=pt_v, bold=bold_v)
        _fill(cs[2], право, align=C, pt=pt_v, bold=bold_v)

    _para(doc).paragraph_format.line_spacing = Pt(1)

    tbl = _make_tbl(NC, fixed_tw=TW, jc='center')
    _set_grid_sp(tbl, GRID)

    # ── Header row 0 ──────────────────────────────────────────────────────────
    _HDR_H = 0.75 if портрет else 1.0
    r0 = tbl.add_row(); _row_h_min(r0, _HDR_H)
    cs0 = list(r0.cells)
    for i, w in enumerate(GRID): _cw_pct(cs0[i], w)

    # rowspan-2 single cols: set vMerge restart first
    for col in ROWSPAN_COLS:
        _vmerge_restart(cs0[col]); _borders(cs0[col]); _valign(cs0[col]); _tcMar(cs0[col], MAR)

    _fill(cs0[I_НПП], '№\nп/п',                            align=C, pt=PT_HDR, bold=True)
    _fill(cs0[I_КОМ], 'Команда',                           align=C, pt=PT_HDR, bold=True)
    _fill(cs0[I_НОМ], 'Номер',                             align=C, pt=PT_HDR, bold=True)
    _vert_cell(cs0[I_НОМ])
    _fill(cs0[I_ФИО], 'Фамилия, имя, отчество участников', align=C, pt=PT_HDR, bold=True)
    _fill(cs0[I_СУБ], 'Субъект РФ',                        align=C, pt=PT_HDR, bold=True)
    _fill(cs0[I_МСП], 'Место СПРИНТ',                      align=C, pt=PT_HDR, bold=True)
    _vert_cell(cs0[I_МСП])
    _fill(cs0[I_МЕС], 'Место',                             align=C, pt=PT_HDR, bold=True)
    _fill(cs0[I_ОЧК], 'Очки',                              align=C, pt=PT_HDR, bold=True)
    if портрет:
        _fill(cs0[I_РАЗ], 'Выполнение\nразряда' + ('*' if снос else ''), align=C, pt=PT_HDR, bold=True)
    else:
        _fill(cs0[I_РАЗ], 'Выполнение разряда' + ('*' if снос else ''), align=C, pt=PT_HDR, bold=True)
    _vert_cell(cs0[I_РАЗ])
    if портрет:
        _vert_cell(cs0[I_МЕС])

    # colspan-2 stage group headers
    for i, ст in enumerate(ПОРЯДОК_ПОКАЗА):
        c0 = cs0[6 + 2 * i]; c1 = cs0[7 + 2 * i]
        c0.merge(c1)
        _borders(cs0[6 + 2 * i]); _valign(cs0[6 + 2 * i]); _tcMar(cs0[6 + 2 * i], MAR)
        _fill(cs0[6 + 2 * i], НАЗВАНИЯ_СТ.get(ст, ст), align=C, pt=PT_HDR, bold=True)

    # colspan-2 итоговый result header
    cs0[I_СТА].merge(cs0[I_ВРМ])
    _borders(cs0[I_СТА]); _valign(cs0[I_СТА]); _tcMar(cs0[I_СТА], MAR)
    _fill(cs0[I_СТА], 'Итоговый результат', align=C, pt=PT_HDR, bold=True)

    # ── Header row 1 ──────────────────────────────────────────────────────────
    r1 = tbl.add_row(); _row_h_min(r1, _HDR_H)
    cs1 = list(r1.cells)
    for i, w in enumerate(GRID): _cw_pct(cs1[i], w)

    for col in ROWSPAN_COLS:
        _vmerge_cont(cs1[col]); _borders(cs1[col])

    for i, ст in enumerate(ПОРЯДОК_ПОКАЗА):
        заезд_лейб = 'Финал' if ст == 'финалы' else 'Заезд №'
        _borders(cs1[6 + 2 * i]); _valign(cs1[6 + 2 * i]); _tcMar(cs1[6 + 2 * i], MAR)
        _fill(cs1[6 + 2 * i], заезд_лейб, align=C, pt=PT_HDR, bold=True)
        _borders(cs1[7 + 2 * i]); _valign(cs1[7 + 2 * i]); _tcMar(cs1[7 + 2 * i], MAR)
        _fill(cs1[7 + 2 * i], 'Результат', align=C, pt=PT_HDR, bold=True)

    _borders(cs1[I_СТА]); _valign(cs1[I_СТА]); _tcMar(cs1[I_СТА], MAR)
    _fill(cs1[I_СТА], 'Стадия', align=C, pt=PT_HDR, bold=True)
    _borders(cs1[I_ВРМ]); _valign(cs1[I_ВРМ]); _tcMar(cs1[I_ВРМ], MAR)
    _fill(cs1[I_ВРМ], 'Время', align=C, pt=PT_HDR, bold=True)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _section(label):
        r = tbl.add_row(); _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[NC - 1]); _cw_pct(cs[0], TW)
        _borders(cs[0]); _valign(cs[0])
        _fill(cs[0], label, align=C, pt=PT, bold=True)
        _keep_next_p(cs[0].paragraphs[-1])

    def _data_row(пп, рез, ст_команд, разряды, норм):
        к = рез.команда
        состав = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
        фио = ', '.join(с.фамилия_имя_отчество for с in состав) if состав else '–'
        рез_сп = к.результат_спринт
        место_сп = str(int(рез_сп.место)) if рез_сп and рез_сп.место else '–'

        r = tbl.add_row(); _row_h_min(r, 0.5)
        cs = list(r.cells)
        for i, w in enumerate(GRID): _cw_pct(cs[i], w)
        for i in range(NC):
            _borders(cs[i]); _valign(cs[i]); _tcMar(cs[i], MAR)

        _fill(cs[I_НПП], str(пп),                              align=C, pt=PT)
        _fill(cs[I_КОМ], к.название or '–',               align=L, pt=PT_TXT)
        _fill(cs[I_НОМ], str(к.нагрудный_номер or ''),        align=C, pt=PT_TXT)
        _fill(cs[I_ФИО], фио,                                  align=L, pt=PT_FIO)
        _fill(cs[I_СУБ], к.субъект.название if к.субъект else '–', align=C, pt=PT_TXT)
        _fill(cs[I_МСП], место_сп,                             align=C, pt=PT)

        д = ст_команд.get(к.id, {})
        for i, ст in enumerate(ПОРЯДОК_ПОКАЗА):
            if ст == 'финалы':
                з_ф = д.get('Финал_А') or д.get('Финал_Б')
                заезд_текст = ('ФА' if 'Финал_А' in д else 'ФБ') if (д.get('Финал_А') or д.get('Финал_Б')) else '–'
                _fill(cs[6 + 2 * i], заезд_текст, align=C, pt=PT)
                _fill(cs[7 + 2 * i], _fmt_рез(з_ф), align=C, pt=PT)
            else:
                з_ст = д.get(ст)
                заезд_текст = str(з_ст.номер_заезда) if з_ст and з_ст.номер_заезда else '–'
                _fill(cs[6 + 2 * i], заезд_текст, align=C, pt=PT)
                _fill(cs[7 + 2 * i], _fmt_рез(з_ст), align=C, pt=PT)

        итог_стадия = АББРЕВ_СТ.get(рез.стадия, рез.стадия or '–') if рез.стадия else '–'
        _fill(cs[I_СТА], итог_стадия,                       align=C, pt=PT)
        _fill(cs[I_ВРМ], _fmt_рез(рез),                     align=C, pt=PT)
        _fill(cs[I_МЕС], str(рез.итоговое_место) if рез.итоговое_место else '–', align=C, pt=PT)
        _fill(cs[I_ОЧК], str(int(рез.очки)) if рез.очки else '–', align=C, pt=PT)
        _fill(cs[I_РАЗ], разряды.get(рез.команда_id, '–') if норм else '–', align=C, pt=PT)

    _section('Мужчины')
    for пп, рез in enumerate(муж_с + муж_без, 1):
        _data_row(пп, рез, муж_ст, муж_разр, муж_норм)

    _section('Женщины')
    for пп, рез in enumerate(жен_с + жен_без, 1):
        _data_row(пп, рез, жен_ст, жен_разр, жен_норм)

    if снос:
        _para(doc, '* Разряды не присваиваются, т.к. количество участников в виде программы '
              'недостаточно для присвоения в соответствии с положением о ЕВСК (пункт 25)',
              align=L, pt=PT_FIO, bold=True)

    sig = _make_tbl(2, fixed_tw=TW, jc='center')
    _set_grid_sp(sig, [W_SIG_L, W_SIG_R])
    for позиция, фио in (('Главный судья',    сор.главный_судья     or ''),
                          ('Главный секретарь', сор.главный_секретарь or '')):
        r = sig.add_row(); _row_h(r, 0.88)
        cs = list(r.cells)
        _cw(cs[0], W_SIG_L); _cw(cs[1], W_SIG_R)
        _no_border(cs[0]); _no_border(cs[1])
        _valign(cs[0]); _valign(cs[1])
        _fill(cs[0], позиция, align=L, pt=PT, bold=True)
        _fill(cs[1], фио,    align=L, pt=PT, bold=True)

    buf = BytesIO()
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf); buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_результаты_кз_docx(сор, итог, группы, таблица_зс, filename, портрет=False):
    """Протокол командного зачёта субъектов РФ."""
    ПОЛ_НАЗВ = {'м': 'Мужчины', 'ж': 'Женщины'}
    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT

    nГр = len(группы)
    TW = round((210 - 10 - 5) * 1440 / 25.4) if портрет else round((297 - 10 - 5) * 1440 / 25.4)

    _W_НПП = 3
    _W_СУБ = max(12, 38 - nГр * 6)
    _W_КОЛ = 6
    _W_БАЛ = 6
    _W_МЕ  = 4
    _weights = [_W_НПП, _W_СУБ] + [w for _ in группы for w in [_W_КОЛ, _W_БАЛ]] + [_W_КОЛ, _W_БАЛ, _W_МЕ]
    _total_W = sum(_weights)
    GRID = [round(TW * w / _total_W) for w in _weights]
    GRID[-1] += TW - sum(GRID)
    NC = len(GRID)

    I_НПП    = 0
    I_СУБ    = 1
    I_ИТ_КОЛ = 2 + 2 * nГр
    I_ИТ_БАЛ = 2 + 2 * nГр + 1
    I_МЕ     = 2 + 2 * nГр + 2

    TW_SIG  = round((210 - 15 - 10) * 1440 / 25.4)
    W_SIG_L = TW_SIG * 35 // 100
    W_SIG_R = TW_SIG - W_SIG_L

    if портрет:
        PT, PT_HDR = 8, 8
    else:
        PT, PT_HDR = 10, 10
    MAR = 28

    def _fmt_bal(v):
        s = f'{round(float(v), 2):.2f}'.rstrip('0').rstrip('.')
        return s if s else '0'

    def _sp(p):
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def _run(p, text, pt=PT, bold=False):
        r = p.add_run(text); r.font.name = TNR; r.font.size = Pt(pt)
        if bold: r.font.bold = True
        return r

    def _para(doc, text='', align=None, pt=PT, bold=False):
        p = doc.add_paragraph(); _sp(p)
        if align: p.alignment = align
        if text: _run(p, text, pt, bold)
        return p

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _borders(cell, sz=4):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _no_border(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
        el = b.find(qn('w:bottom'))
        if el is None:
            el = OxmlElement('w:bottom'); b.append(el)
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _cw(cell, w):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        tcW.set(qn('w:w'), str(int(w))); tcW.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        va = pr.find(qn('w:vAlign'))
        if va is None:
            va = OxmlElement('w:vAlign'); pr.append(va)
        va.set(qn('w:val'), val)

    def _fill(cell, text, align=None, pt=PT, bold=False):
        p = cell.paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p)
        if align: p.alignment = align
        text = str(text).replace('—', '–')
        if not text: return
        parts = str(text).split('\n')
        for i, part in enumerate(parts):
            r = p.add_run(part); r.font.name = TNR; r.font.size = Pt(pt)
            if bold: r.font.bold = True
            if i < len(parts) - 1:
                br = OxmlElement('w:br'); r._r.append(br)

    def _cw_pct(cell, dxa):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        pct = round(dxa * 5000 / TW)
        tcW.set(qn('w:w'), str(pct)); tcW.set(qn('w:type'), 'pct')

    def _make_tbl(ncols, fixed_tw=None, jc=None):
        t = doc.add_table(rows=0, cols=ncols)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        if fixed_tw:
            el.set(qn('w:w'), str(fixed_tw)); el.set(qn('w:type'), 'dxa')
        else:
            el.set(qn('w:w'), '5000'); el.set(qn('w:type'), 'pct')
        if jc == 'center':
            for old_ind in tblPr.findall(qn('w:tblInd')): tblPr.remove(old_ind)
        else:
            ind = tblPr.find(qn('w:tblInd'))
            if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
            ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        if jc:
            jc_el = tblPr.find(qn('w:jc'))
            if jc_el is None: jc_el = OxmlElement('w:jc'); tblPr.append(jc_el)
            jc_el.set(qn('w:val'), jc)
        return t

    def _set_grid_sp(tbl_obj, widths):
        tbl = tbl_obj._tbl
        old = tbl.find(qn('w:tblGrid'))
        if old is not None: tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in widths:
            gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); grid.append(gc)
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is not None: tblPr.addnext(grid)
        else: tbl.insert(0, grid)

    def _row_h(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'exact')
        pr.append(rh)

    def _row_h_min(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'atLeast')
        pr.append(rh)

    def _tcMar(cell, lr=28):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcMar')): pr.remove(old)
        mar = OxmlElement('w:tcMar')
        for side in ('left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:w'), str(lr)); el.set(qn('w:type'), 'dxa')
            mar.append(el)
        pr.append(mar)

    def _no_wrap_cell(cell):
        pr = _tcpr(cell)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))

    def _vmerge_restart(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); vm.set(qn('w:val'), 'restart'); pr.append(vm)

    def _vmerge_cont(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); pr.append(vm)

    def _ls15(p):
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:line'), '360'); sp.set(qn('w:lineRule'), 'auto')

    # ── document ──────────────────────────────────────────────────────────────
    from docx.shared import Mm
    doc  = Document()
    sect = doc.sections[0]
    if портрет:
        sect.page_height   = Mm(297); sect.page_width    = Mm(210)
    else:
        sect.page_height   = Mm(210); sect.page_width    = Mm(297)
    sect.top_margin    = Mm(15)
    sect.bottom_margin = Mm(15)
    sect.left_margin   = Mm(10)
    sect.right_margin  = Mm(5)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if портрет:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(10)
    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    # ── header ────────────────────────────────────────────────────────────────
    _hdr_pt = 10 if портрет else 12
    if сор.организаторы_блок:
        _para(doc, сор.организаторы_блок, align=C, pt=_hdr_pt, bold=True)
    p0 = _para(doc); p0.paragraph_format.line_spacing = Pt(4)
    _ls15(_para(doc, сор.название or '', align=C, pt=12, bold=True))
    _ls15(_para(doc, 'Командный зачет среди субъектов Российской Федерации',
                align=C, pt=_hdr_pt, bold=True))

    # ── дата/место ────────────────────────────────────────────────────────────
    dt = _make_tbl(3, fixed_tw=TW, jc='center')
    _cw_char = round(PT * 0.6 * 20)
    _w_dt  = max(100, len(сор.даты_проведения or 'Дата проведения') * _cw_char + 2 * MAR)
    _w_pl  = max(100, len(сор.место_проведения or 'Место проведения') * _cw_char + 2 * MAR)
    _w_mid = max(100, TW - _w_dt - _w_pl)
    _set_grid_sp(dt, [_w_dt, _w_mid, _w_pl])
    _dt_rh = (lambda r: _row_h(r, 0.3)) if портрет else (lambda r: _row_h(r, 0.4))
    for idx, (лево, право, pt_v, bold_v) in enumerate([
        (сор.даты_проведения or '', сор.место_проведения or '', PT, True),
        ('Дата проведения', 'Место проведения', PT if not портрет else 7, not портрет),
    ]):
        r = dt.add_row(); _dt_rh(r)
        cs = list(r.cells)
        for c in cs:
            _no_border(c)
            _valign(c, 'bottom' if (not портрет and idx == 0) else ('top' if (not портрет) else 'center'))
        _no_wrap_cell(cs[0]); _no_wrap_cell(cs[2])
        if idx == 0:
            _bottom_border(cs[0]); _bottom_border(cs[2])
        _fill(cs[0], лево, align=C, pt=pt_v, bold=bold_v)
        _fill(cs[2], право, align=C, pt=pt_v, bold=bold_v)

    p = _para(doc); p.paragraph_format.line_spacing = Pt(1)

    # ── main table ────────────────────────────────────────────────────────────
    tbl = _make_tbl(NC, fixed_tw=TW, jc='center')
    _set_grid_sp(tbl, GRID)

    # header row 0
    r0 = tbl.add_row(); _row_h_min(r0, 0.9 if портрет else 0.6)
    cs0 = list(r0.cells)
    for i, w in enumerate(GRID): _cw_pct(cs0[i], w)

    _vmerge_restart(cs0[I_НПП])
    _fill(cs0[I_НПП], '№\nп/п', align=C, pt=PT_HDR, bold=True)

    _vmerge_restart(cs0[I_СУБ])
    _fill(cs0[I_СУБ], 'Субъект Российской Федерации', align=C, pt=PT_HDR, bold=True)

    for gi, (пол, кл) in enumerate(группы):
        ci = 2 + 2 * gi
        cs0[ci].merge(cs0[ci + 1]); _cw_pct(cs0[ci], GRID[ci] + GRID[ci + 1])
        _fill(cs0[ci], f'{ПОЛ_НАЗВ[пол]} {кл}', align=C, pt=PT_HDR, bold=True)

    cs0[I_ИТ_КОЛ].merge(cs0[I_ИТ_БАЛ])
    _cw_pct(cs0[I_ИТ_КОЛ], GRID[I_ИТ_КОЛ] + GRID[I_ИТ_БАЛ])
    _fill(cs0[I_ИТ_КОЛ], 'ИТОГО', align=C, pt=PT_HDR, bold=True)

    _vmerge_restart(cs0[I_МЕ])
    _fill(cs0[I_МЕ], 'Место', align=C, pt=PT_HDR, bold=True)

    for i in range(NC):
        _borders(cs0[i]); _valign(cs0[i]); _tcMar(cs0[i], MAR)

    # header row 1
    r1 = tbl.add_row(); _row_h_min(r1, 0.9 if портрет else 0.6)
    cs1 = list(r1.cells)
    for i, w in enumerate(GRID): _cw_pct(cs1[i], w)

    _vmerge_cont(cs1[I_НПП])
    _vmerge_cont(cs1[I_СУБ])

    for gi in range(nГр):
        ci = 2 + 2 * gi
        _fill(cs1[ci],     'Кол-во\nэкипажей', align=C, pt=PT_HDR, bold=True)
        _fill(cs1[ci + 1], 'Баллы',             align=C, pt=PT_HDR, bold=True)

    _fill(cs1[I_ИТ_КОЛ], 'Кол-во\nэкипажей', align=C, pt=PT_HDR, bold=True)
    _fill(cs1[I_ИТ_БАЛ], 'Баллы',             align=C, pt=PT_HDR, bold=True)
    _vmerge_cont(cs1[I_МЕ])

    for i in range(NC):
        _borders(cs1[i]); _valign(cs1[i]); _tcMar(cs1[i], MAR)

    # data rows
    for пп, з in enumerate(итог, 1):
        r = tbl.add_row(); _row_h_min(r, 0.5)
        cs = list(r.cells)
        for i, w in enumerate(GRID): _cw_pct(cs[i], w)
        for i in range(NC):
            _borders(cs[i]); _valign(cs[i]); _tcMar(cs[i], MAR)

        _fill(cs[I_НПП], str(пп),                    align=C, pt=PT)
        _fill(cs[I_СУБ], з['субъект'].название or '–', align=L, pt=PT)

        for gi, (пол, кл) in enumerate(группы):
            ci = 2 + 2 * gi
            gd = з['группы'].get((пол, кл), {'кол': 0, 'баллы': 0.0})
            _fill(cs[ci],     str(gd['кол']) if gd['кол'] else '–', align=C, pt=PT)
            _fill(cs[ci + 1], _fmt_bal(gd['баллы']) if gd['баллы'] else '–', align=C, pt=PT)

        _fill(cs[I_ИТ_КОЛ], str(з['итого_кол']),            align=C, pt=PT)
        _fill(cs[I_ИТ_БАЛ], _fmt_bal(з['итого_баллы']),     align=C, pt=PT)
        _fill(cs[I_МЕ],      str(з['место']) if з.get('место') else '–', align=C, pt=PT)

    # ── signature table ───────────────────────────────────────────────────────
    _para(doc)
    sig = _make_tbl(2, fixed_tw=TW, jc='center')
    _set_grid_sp(sig, [W_SIG_L, W_SIG_R])
    for позиция, фио in (('Главный судья',    сор.главный_судья     or ''),
                          ('Главный секретарь', сор.главный_секретарь or '')):
        r = sig.add_row(); _row_h(r, 0.88)
        cs = list(r.cells)
        _cw(cs[0], W_SIG_L); _cw(cs[1], W_SIG_R)
        _no_border(cs[0]); _no_border(cs[1])
        _valign(cs[0]); _valign(cs[1])
        _fill(cs[0], позиция, align=L, pt=PT, bold=True)
        _fill(cs[1], фио,    align=L, pt=PT, bold=True)

    # ── coefficient table (таблица ЗС) ───────────────────────────────────────
    if таблица_зс:
        _para(doc)  # spacing

        кол_nums  = таблица_зс['rows']    # кол_субъектов values (e.g. [13,12,...,3])
        место_nums = таблица_зс['cols']   # место values (e.g. [1,2,3,4])
        row_labels = таблица_зс['row_labels']
        col_labels = таблица_зс['col_labels']
        data       = таблица_зс['data']

        # Cols in зс table: 1 (место label) + len(кол_nums) (one per кол_субъектов)
        nЗС = 1 + len(кол_nums)
        # Абсолютные ширины в dxa: первый и 2й/последний столбцы полные, средние — вдвое уже
        _W0  = round(45 * 1440 / 25.4)   # первый столбец (место): 45mm
        _WF  = round(18 * 1440 / 25.4)   # полный столбец (2й и последний): 18mm
        _WH  = round(10 * 1440 / 25.4)   # половинный столбец (средние): 10mm
        n_mid = max(0, len(кол_nums) - 2)
        if len(кол_nums) == 1:
            GRID_зс = [_W0, _WF]
        else:
            GRID_зс = [_W0, _WF] + [_WH] * n_mid + [_WF]
        TW_зс = sum(GRID_зс)

        # Заголовок — обычный параграф с правым отступом TW-TW_зс (в dxa через XML),
        # чтобы align=C центрировал текст по ширине ЗС-таблицы
        p_hdr = _para(doc, 'Таблица начисления баллов', align=C, pt=PT, bold=True)
        _pPr = p_hdr._p.get_or_add_pPr()
        _ind = _pPr.find(qn('w:ind'))
        if _ind is None:
            _ind = OxmlElement('w:ind'); _pPr.append(_ind)
        _ind.set(qn('w:right'), str(TW - TW_зс))

        зс = _make_tbl(nЗС, fixed_tw=TW_зс, jc='left')
        _set_grid_sp(зс, GRID_зс)

        # header row 0: col0 vMerge, cols 1..n merged with label
        rh0 = зс.add_row(); _row_h_min(rh0, 0.6)
        csh0 = list(rh0.cells)
        _cw(csh0[0], GRID_зс[0])
        _vmerge_restart(csh0[0]); _fill(csh0[0], 'Количество баллов за каждую дисциплину за занятое место', align=C, pt=PT, bold=True)
        csh0[1].merge(csh0[nЗС - 1])
        _cw(csh0[1], TW_зс - GRID_зс[0])
        _fill(csh0[1], 'Количество субъектов Российской Федерации', align=C, pt=PT, bold=True)
        for i in range(nЗС): _borders(csh0[i]); _valign(csh0[i]); _tcMar(csh0[i], MAR)

        # header row 1: col0 vMerge cont, cols 1..n кол_субъектов labels
        rh1 = зс.add_row(); _row_h_min(rh1, 0.6)
        csh1 = list(rh1.cells)
        _cw(csh1[0], GRID_зс[0]); _vmerge_cont(csh1[0])
        for j, кол in enumerate(кол_nums):
            _cw(csh1[j + 1], GRID_зс[j + 1])
            _fill(csh1[j + 1], row_labels[кол], align=C, pt=PT, bold=True)
        for i in range(nЗС): _borders(csh1[i]); _valign(csh1[i]); _tcMar(csh1[i], MAR)

        # data rows
        for место in место_nums:
            rd = зс.add_row(); _row_h_min(rd, 0.5)
            csd = list(rd.cells)
            for i in range(nЗС):
                _cw(csd[i], GRID_зс[i])
                _borders(csd[i]); _valign(csd[i]); _tcMar(csd[i], MAR)
            _fill(csd[0], col_labels[место], align=C, pt=PT, bold=True)
            for j, кол in enumerate(кол_nums):
                v = data.get((кол, место), '')
                _fill(csd[j + 1], _fmt_bal(v) if v != '' else '–', align=C, pt=PT)

    buf = BytesIO()
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf); buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_результаты_мн_docx(сор, муж_рез, жен_рез, муж_разр, жен_разр,
                                класс, filename, итоговый=True, портрет=False):
    """Протокол результатов Многоборье — 16 колонок."""
    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT

    TW = round((210 - 10 - 5) * 1440 / 25.4) if портрет else round((297 - 10 - 5) * 1440 / 25.4)
    if портрет:
        _PCTS = [3, 10, 6, 20, 10, 4.4, 4.4, 4.4, 4.4, 4.4, 4.4, 4.4, 4.4, 6, 4.9, 4.9]
    else:
        _PCTS = [2.9, 12.0, 3.5, 24.5, 7.7, 4.1, 4.1, 4.1, 4.1, 4.1, 4.1, 4.1, 4.1, 5.4, 4.7, 6.5]
    GRID = [round(TW * p / 100) for p in _PCTS]
    GRID[-1] += TW - sum(GRID)
    NC = 16

    I_НПП, I_К,    I_НОМ, I_СОС, I_СУБ  = 0, 1, 2, 3, 4
    I_СП_М, I_СП_О                         = 5, 6
    I_ПС_М, I_ПС_О                         = 7, 8
    I_СЛ_М, I_СЛ_О                         = 9, 10
    I_ДГ_М, I_ДГ_О                         = 11, 12
    I_СУМ, I_МЕ,  I_РА                    = 13, 14, 15

    W_SIG_L = TW * 27 // 100
    W_SIG_R = TW - W_SIG_L


    if портрет:
        PT, PT_FIO, PT_HDR = 8, 7, 8
    else:
        PT, PT_FIO, PT_HDR = 10, 8, 10
    MAR = 28

    def _sp(p):
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)

    def _run(p, text, pt=PT, bold=False):
        r = p.add_run(text); r.font.name = TNR; r.font.size = Pt(pt)
        if bold: r.font.bold = True
        return r

    def _para(doc, text='', align=None, pt=PT, bold=False):
        p = doc.add_paragraph(); _sp(p)
        if align: p.alignment = align
        if text: _run(p, text, pt, bold)
        return p

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _borders(cell, sz=4):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _no_border(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'right', 'insideH', 'insideV'):
            el = b.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}'); b.append(el)
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
        el = b.find(qn('w:bottom'))
        if el is None:
            el = OxmlElement('w:bottom'); b.append(el)
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')

    def _cw(cell, w):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        tcW.set(qn('w:w'), str(int(w))); tcW.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        va = pr.find(qn('w:vAlign'))
        if va is None:
            va = OxmlElement('w:vAlign'); pr.append(va)
        va.set(qn('w:val'), val)

    def _fill(cell, text, align=None, pt=PT, bold=False):
        p = cell.paragraphs[0]; p.clear(); p.style = 'No Spacing'; _sp(p)
        if align: p.alignment = align
        text = str(text).replace('—', '–')
        if not text: return
        parts = str(text).split('\n')
        for i, part in enumerate(parts):
            r = p.add_run(part); r.font.name = TNR; r.font.size = Pt(pt)
            if bold: r.font.bold = True
            if i < len(parts) - 1:
                br = OxmlElement('w:br'); r._r.append(br)

    def _cw_pct(cell, dxa):
        pr = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        pct = round(dxa * 5000 / TW)
        tcW.set(qn('w:w'), str(pct)); tcW.set(qn('w:type'), 'pct')

    def _make_tbl(ncols, fixed_tw=None, jc=None):
        t = doc.add_table(rows=0, cols=ncols)
        tbl = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        for tl in tblPr.findall(qn('w:tblLook')): tblPr.remove(tl)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        if fixed_tw:
            el.set(qn('w:w'), str(fixed_tw)); el.set(qn('w:type'), 'dxa')
        else:
            el.set(qn('w:w'), '5000'); el.set(qn('w:type'), 'pct')
        if jc == 'center':
            for old_ind in tblPr.findall(qn('w:tblInd')): tblPr.remove(old_ind)
        else:
            ind = tblPr.find(qn('w:tblInd'))
            if ind is None: ind = OxmlElement('w:tblInd'); tblPr.append(ind)
            ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        if jc:
            jc_el = tblPr.find(qn('w:jc'))
            if jc_el is None: jc_el = OxmlElement('w:jc'); tblPr.append(jc_el)
            jc_el.set(qn('w:val'), jc)
        return t

    def _set_grid_sp(tbl_obj, widths):
        tbl = tbl_obj._tbl
        old = tbl.find(qn('w:tblGrid'))
        if old is not None: tbl.remove(old)
        grid = OxmlElement('w:tblGrid')
        for w in widths:
            gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); grid.append(gc)
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is not None: tblPr.addnext(grid)
        else: tbl.insert(0, grid)

    def _row_h(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'exact')
        pr.append(rh)

    def _row_h_min(row, cm):
        pr = row._tr.find(qn('w:trPr'))
        if pr is None:
            pr = OxmlElement('w:trPr'); row._tr.insert(0, pr)
        for old in pr.findall(qn('w:trHeight')): pr.remove(old)
        rh = OxmlElement('w:trHeight')
        rh.set(qn('w:val'), str(int(cm * 567))); rh.set(qn('w:hRule'), 'atLeast')
        pr.append(rh)

    def _vmerge_restart(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); vm.set(qn('w:val'), 'restart'); pr.append(vm)

    def _vmerge_cont(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:vMerge')): pr.remove(old)
        vm = OxmlElement('w:vMerge'); pr.append(vm)

    def _tcMar(cell, lr=28):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcMar')): pr.remove(old)
        mar = OxmlElement('w:tcMar')
        for side in ('left', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:w'), str(lr)); el.set(qn('w:type'), 'dxa')
            mar.append(el)
        pr.append(mar)

    def _no_wrap_cell(cell):
        pr = _tcpr(cell)
        if pr.find(qn('w:noWrap')) is None:
            pr.append(OxmlElement('w:noWrap'))

    def _line_br(p):
        p.add_run()._r.append(OxmlElement('w:br'))


    def _vert_cell(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:textDirection')): pr.remove(old)
        td = OxmlElement('w:textDirection'); td.set(qn('w:val'), 'btLr'); pr.append(td)
        _valign(cell, 'center')
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER


    def _horiz_cell(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:textDirection')): pr.remove(old)

    def _ls15(p):
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:line'), '360'); sp.set(qn('w:lineRule'), 'auto')

    def _keep_next_p(p):
        pPr = p._p.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr'); p._p.insert(0, pPr)
        if pPr.find(qn('w:keepNext')) is None:
            pPr.append(OxmlElement('w:keepNext'))

    муж_норм = len(муж_рез) >= 6
    жен_норм = len(жен_рез) >= 6
    снос = (len(муж_рез) > 0 and not муж_норм) or (len(жен_рез) > 0 and not жен_норм)

    _коды_мн = {'R4': '1740071811Я', 'R6': '1740081811Я'}
    _кл = класс or сор.класс_судна or 'R4/R6'
    _код = _коды_мн.get(_кл, сор.код_соревнований or '')
    дисциплина = f'МНОГОБОРЬЕ – {_кл}, код ВРВС {_код}'

    # ── document ──────────────────────────────────────────────────────────────
    from docx.shared import Mm
    doc  = Document()
    sect = doc.sections[0]
    if портрет:
        sect.page_height   = Mm(297); sect.page_width    = Mm(210)
    else:
        sect.page_height   = Mm(210); sect.page_width    = Mm(297)
    sect.top_margin    = Mm(12)
    sect.bottom_margin = Mm(17.5)
    sect.left_margin   = Mm(10)
    sect.right_margin  = Mm(5)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        if портрет:
            pgSz.attrib.pop(qn('w:orient'), None)
        else:
            pgSz.set(qn('w:orient'), 'landscape')
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(10)
    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    # ── header ────────────────────────────────────────────────────────────────
    _hdr_pt = 10 if портрет else 12
    if сор.организаторы_блок:
        _para(doc, сор.организаторы_блок, align=C, pt=_hdr_pt, bold=True)
    p0 = _para(doc); p0.paragraph_format.line_spacing = Pt(4)
    _ls15(_para(doc, сор.название or '', align=C, pt=12, bold=True))
    заголовок = 'ПРОТОКОЛ РЕЗУЛЬТАТОВ' if итоговый else 'ПРЕДВАРИТЕЛЬНЫЙ ПРОТОКОЛ РЕЗУЛЬТАТОВ'
    _ls15(_para(doc, заголовок, align=C, pt=_hdr_pt, bold=True))
    if итоговый and сор.номер_екп:
        _ls15(_para(doc, f'ЕКП № {сор.номер_екп}', align=C, pt=_hdr_pt, bold=True))
    _ls15(_para(doc, f'Спортивная дисциплина: {дисциплина}', align=C, pt=_hdr_pt, bold=True))

    # ── дата/место ────────────────────────────────────────────────────────────
    dt = _make_tbl(3, fixed_tw=TW, jc='center')
    _cw_char = round(PT * 0.6 * 20)
    _w_dt  = max(100, len(сор.даты_проведения or 'Дата проведения') * _cw_char + 2 * MAR)
    _w_pl  = max(100, len(сор.место_проведения or 'Место проведения') * _cw_char + 2 * MAR)
    _w_mid = max(100, TW - _w_dt - _w_pl)
    _set_grid_sp(dt, [_w_dt, _w_mid, _w_pl])
    _dt_rh = (lambda r: _row_h(r, 0.3)) if портрет else (lambda r: _row_h(r, 0.4))
    for idx, (лево, право, pt_v, bold_v) in enumerate([
        (сор.даты_проведения or '', сор.место_проведения or '', PT, True),
        ('Дата проведения', 'Место проведения', PT if not портрет else 7, True),
    ]):
        r = dt.add_row(); _dt_rh(r)
        cs = list(r.cells)
        for c in cs:
            _no_border(c)
            _valign(c, 'bottom' if (not портрет and idx == 0) else ('top' if not портрет else 'center'))
        _no_wrap_cell(cs[0]); _no_wrap_cell(cs[2])
        if idx == 0:
            _bottom_border(cs[0]); _bottom_border(cs[2])
        _fill(cs[0], лево, align=C, pt=pt_v, bold=bold_v)
        _fill(cs[2], право, align=C, pt=pt_v, bold=bold_v)

    p = _para(doc); p.paragraph_format.line_spacing = Pt(1)

    # ── main table ────────────────────────────────────────────────────────────
    tbl = _make_tbl(NC, fixed_tw=TW, jc='center')
    _set_grid_sp(tbl, GRID)

    # header row 0
    r0 = tbl.add_row(); _row_h_min(r0, 0.9 if портрет else 0.6)
    cs0 = list(r0.cells)
    for i, w in enumerate(GRID): _cw_pct(cs0[i], w)

    for i in (I_НПП, I_К, I_НОМ, I_СОС, I_СУБ, I_СУМ, I_МЕ):
        _vmerge_restart(cs0[i])
    _fill(cs0[I_НПП], '№\nп/п',       align=C, pt=PT_HDR, bold=True)
    _fill(cs0[I_К],   'Команда',       align=C, pt=PT_HDR, bold=True)
    _fill(cs0[I_НОМ], 'Номер',         align=C, pt=PT_HDR, bold=True)
    _fill(cs0[I_СОС], 'Состав',        align=C, pt=PT_HDR, bold=True)
    _fill(cs0[I_СУБ], 'Субъект РФ',   align=C, pt=PT_HDR, bold=True)
    _fill(cs0[I_СУМ], 'Общая\nсумма', align=C, pt=PT_HDR, bold=True)
    _fill(cs0[I_МЕ],  'Место',         align=C, pt=PT_HDR, bold=True)

    _PT_ПС = 6.5 if портрет else 8
    for _disc in (
        (I_СП_М, I_СП_О, 'Спринт',              PT_HDR),
        (I_ПС_М, I_ПС_О, 'Параллельный\nспринт', _PT_ПС),
        (I_СЛ_М, I_СЛ_О, 'Слалом',              PT_HDR),
        (I_ДГ_М, I_ДГ_О, 'Длинная\nгонка',      PT_HDR),
    ):
        ci_м, ci_о, label, pt_disc = _disc
        cs0[ci_м].merge(cs0[ci_о])
        _cw_pct(cs0[ci_м], GRID[ci_м] + GRID[ci_о])
        _fill(cs0[ci_м], label, align=C, pt=pt_disc, bold=True)

    # Выполнение разряда — вертикальный текст в портрете, горизонтальный в альбоме
    _vmerge_restart(cs0[I_РА])
    if портрет:
        _vert_cell(cs0[I_РА])
        _fill(cs0[I_РА], 'Выполнение разряда' + ('*' if снос else ''), align=C, pt=PT_HDR, bold=True)
    else:
        _horiz_cell(cs0[I_РА])
        p_ра = cs0[I_РА].paragraphs[0]; p_ра.clear()
        _sp(p_ра); p_ра.alignment = C
        _run(p_ра, 'Выполнение', PT_HDR, True); _line_br(p_ра)
        _run(p_ра, 'разряда' + ('*' if снос else ''), PT_HDR, True)

    for i in range(NC):
        _borders(cs0[i]); _valign(cs0[i]); _tcMar(cs0[i], MAR)

    # header row 1
    r1 = tbl.add_row(); _row_h_min(r1, 0.9 if портрет else 0.6)
    cs1 = list(r1.cells)
    for i, w in enumerate(GRID): _cw_pct(cs1[i], w)

    for i in (I_НПП, I_К, I_НОМ, I_СОС, I_СУБ, I_СУМ, I_МЕ, I_РА):
        _vmerge_cont(cs1[i])
    if портрет:
        _vert_cell(cs1[I_РА])
    else:
        _horiz_cell(cs1[I_РА])

    _PT_DISC = 6.5 if портрет else PT_HDR
    for ci_м, ci_о in ((I_СП_М, I_СП_О), (I_ПС_М, I_ПС_О), (I_СЛ_М, I_СЛ_О), (I_ДГ_М, I_ДГ_О)):
        _fill(cs1[ci_м], 'Место', align=C, pt=_PT_DISC, bold=True)
        _fill(cs1[ci_о], 'Очки',  align=C, pt=_PT_DISC, bold=True)

    for i in range(NC):
        _borders(cs1[i]); _valign(cs1[i]); _tcMar(cs1[i], MAR)

    # ── section/data helpers ──────────────────────────────────────────────────
    def _section(label):
        r = tbl.add_row(); _row_h(r, 0.6)
        cs = list(r.cells)
        cs[0].merge(cs[NC - 1]); _cw_pct(cs[0], TW)
        _borders(cs[0]); _valign(cs[0])
        _fill(cs[0], label, align=C, pt=PT, bold=True)
        _keep_next_p(cs[0].paragraphs[-1])

    def _место_пс(к):
        итог = [р for р in к.результат_пс if р.итоговое_место is not None]
        return итог[0].итоговое_место if итог else None

    def _место_сл(к):
        итог = [р for р in к.результат_слалом if р.место is not None]
        return итог[0].место if итог else None

    def _data_row(пп, р, разряды, норм):
        к = р.команда
        состав = sorted(к.спортсмены, key=lambda с: (с.номер_в_экипаже or 0))
        фио = ', '.join(
            f'{с.фамилия_имя_отчество} ({с.дата_рождения.strftime("%d.%m.%Y") if с.дата_рождения else "б/д"}, {с.разряд or "б/р"})'
            for с in состав) if состав else '–'

        м_сп = к.результат_спринт.место if к.результат_спринт and к.результат_спринт.место else '–'
        м_пс = _место_пс(к) or '–'
        м_сл = _место_сл(к) or '–'
        м_дг = к.результат_дг.место if к.результат_дг and к.результат_дг.место else '–'

        def _очки(v): return str(int(v)) if v else '–'

        row = tbl.add_row(); _row_h_min(row, 0.5)
        cs = list(row.cells)
        for i, w in enumerate(GRID): _cw_pct(cs[i], w)
        for i in range(NC):
            _borders(cs[i]); _valign(cs[i]); _tcMar(cs[i], MAR)

        _fill(cs[I_НПП], str(пп),                        align=C, pt=PT)
        _fill(cs[I_К],   к.название or '–',          align=L, pt=PT)
        _fill(cs[I_НОМ], str(к.нагрудный_номер or ''),   align=C, pt=PT)
        _fill(cs[I_СОС], фио,                             align=L, pt=PT_FIO)
        _fill(cs[I_СУБ], к.субъект.название if к.субъект else '–', align=C, pt=PT)
        _fill(cs[I_СП_М], str(м_сп),                     align=C, pt=PT)
        _fill(cs[I_СП_О], _очки(р.очки_спринт),          align=C, pt=PT)
        _fill(cs[I_ПС_М], str(м_пс),                     align=C, pt=PT)
        _fill(cs[I_ПС_О], _очки(р.очки_пс),              align=C, pt=PT)
        _fill(cs[I_СЛ_М], str(м_сл),                     align=C, pt=PT)
        _fill(cs[I_СЛ_О], _очки(р.очки_слалом),          align=C, pt=PT)
        _fill(cs[I_ДГ_М], str(м_дг),                     align=C, pt=PT)
        _fill(cs[I_ДГ_О], _очки(р.очки_дг),              align=C, pt=PT)
        _fill(cs[I_СУМ], str(int(р.сумма_очков)) if р.сумма_очков else '–', align=C, pt=PT)
        _fill(cs[I_МЕ],  str(р.место) if р.место else '–', align=C, pt=PT)
        _fill(cs[I_РА],  разряды.get(к.id, '–') if норм else '–', align=C, pt=PT)

    _section('Мужчины')
    for пп, р in enumerate(муж_рез, 1):
        _data_row(пп, р, муж_разр, муж_норм)

    _section('Женщины')
    for пп, р in enumerate(жен_рез, 1):
        _data_row(пп, р, жен_разр, жен_норм)

    if снос:
        _para(doc, '* Разряды не присваиваются, т.к. количество участников в виде программы '
              'недостаточно для присвоения в соответствии с положением о ЕВСК (пункт 25)',
              align=L, pt=PT_FIO, bold=True)

    # ── signature table ───────────────────────────────────────────────────────
    sig = _make_tbl(2, fixed_tw=TW, jc='center')
    _set_grid_sp(sig, [W_SIG_L, W_SIG_R])
    for позиция, фио in (('Главный судья',    сор.главный_судья     or ''),
                          ('Главный секретарь', сор.главный_секретарь or '')):
        r = sig.add_row(); _row_h(r, 0.88)
        cs = list(r.cells)
        _cw(cs[0], W_SIG_L); _cw(cs[1], W_SIG_R)
        _no_border(cs[0]); _no_border(cs[1])
        _valign(cs[0]); _valign(cs[1])
        _fill(cs[0], позиция, align=L, pt=PT, bold=True)
        _fill(cs[1], фио,    align=L, pt=PT, bold=True)

    buf = BytesIO()
    _apply_no_break(doc)
    _apply_keep_next(doc)
    doc.save(buf); buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_регламент_docx(сор, программа, filename):
    """Регламент соревнований — A4 портрет, TNR 11pt."""
    from docx.shared import Mm

    PT  = 11
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT
    J   = WD_ALIGN_PARAGRAPH.JUSTIFY
    R   = WD_ALIGN_PARAGRAPH.RIGHT

    # Page geometry (dxa)
    PW  = round(210 * 1440 / 25.4)
    ML  = round(25  * 1440 / 25.4)
    MR  = round(15  * 1440 / 25.4)
    MT  = round(20  * 1440 / 25.4)
    MB  = round(20  * 1440 / 25.4)
    TW  = PW - ML - MR                  # ≈ 9638 dxa

    W_SIDE = round(TW * 0.47)           # each УТВЕРЖДАЮ column
    W_GAP  = TW - 2 * W_SIDE

    W_DATE = round(70 * 1440 / 72)
    W_TIME = round(80 * 1440 / 72)
    W_NAME = TW - W_DATE - W_TIME

    doc = Document()
    sect = doc.sections[0]
    sect.page_width    = Mm(210)
    sect.page_height   = Mm(297)
    sect.left_margin   = Mm(25)
    sect.right_margin  = Mm(15)
    sect.top_margin    = Mm(20)
    sect.bottom_margin = Mm(20)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        pgSz.attrib.pop(qn('w:orient'), None)

    for p in list(doc.paragraphs):
        p._element.getparent().remove(p._element)

    # ── helpers ──────────────────────────────────────────────────────────
    def _sp0(p):
        pPr = p._p.get_or_add_pPr()
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:before'), '0')
        sp.set(qn('w:after'),  '0')
        pPr.append(sp)

    def _ls115(p):
        pPr = p._p.get_or_add_pPr()
        sp  = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:line'),     str(round(1.15 * 240)))
        sp.set(qn('w:lineRule'), 'auto')

    def _spacing(p, before_pt=0, after_pt=0):
        pPr = p._p.get_or_add_pPr()
        sp  = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:before'),   str(round(before_pt * 20)))
        sp.set(qn('w:after'),    str(round(after_pt  * 20)))
        sp.set(qn('w:line'),     str(round(1.15 * 240)))
        sp.set(qn('w:lineRule'), 'auto')

    def _run_r(p, text, pt=PT, bold=False, char_sp_pt=0):
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(pt)
        run.font.bold = bool(bold)
        r   = run._r
        rPr = r.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr'); r.insert(0, rPr)
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts'); rPr.append(rFonts)
        for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia'):
            rFonts.set(qn(attr), 'Times New Roman')
        if char_sp_pt:
            csp = OxmlElement('w:spacing')
            csp.set(qn('w:val'), str(round(char_sp_pt * 20)))
            rPr.append(csp)
        return run

    def _line_br(p):
        p._p.append(OxmlElement('w:br'))

    def _para(text='', align=J, pt=PT, bold=False, indent_mm=0,
               before_pt=0, after_pt=0):
        p = doc.add_paragraph()
        _spacing(p, before_pt, after_pt)
        p.alignment = align
        if indent_mm:
            pPr = p._p.get_or_add_pPr()
            ind = OxmlElement('w:ind')
            ind.set(qn('w:firstLine'), str(round(indent_mm * 1440 / 25.4)))
            pPr.append(ind)
        if text:
            _run_r(p, text, pt=pt, bold=bold)
        return p

    def _heading(text, pt=PT, bold=True):
        p = _para(align=C, pt=pt, bold=bold, before_pt=8, after_pt=2)
        parts = text.split('\n')
        for i, part in enumerate(parts):
            _run_r(p, part, pt=pt, bold=bold)
            if i < len(parts) - 1:
                _line_br(p)
        return p

    def _indent(text, pt=PT, bold=False):
        return _para(text, align=J, pt=pt, bold=bold, indent_mm=10)

    def _indent_block(text, pt=PT):
        for line in (text or '').split('\n'):
            if not line.strip():
                continue
            p = doc.add_paragraph()
            _spacing(p)
            pPr = p._p.get_or_add_pPr()
            ind = OxmlElement('w:ind')
            ind.set(qn('w:left'), str(round(10 * 1440 / 25.4)))
            pPr.append(ind)
            p.alignment = J
            _run_r(p, line.strip(), pt=pt)

    def _list_dash(text, pt=PT):
        p = doc.add_paragraph()
        _spacing(p)
        pPr = p._p.get_or_add_pPr()
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'),    str(round(14 * 1440 / 25.4)))
        ind.set(qn('w:hanging'), str(round(4  * 1440 / 25.4)))
        pPr.append(ind)
        p.alignment = J
        _run_r(p, text, pt=pt)
        return p

    # ── cell helpers ─────────────────────────────────────────────────────
    def _tcpr(cell):
        tc   = cell._tc
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = OxmlElement('w:tcPr'); tc.insert(0, tcPr)
        return tcPr

    def _no_border(cell):
        pr = _tcpr(cell)
        tcB = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            tcB.append(el)
        pr.append(tcB)

    def _bottom_border(cell, sz=6):
        pr = _tcpr(cell)
        tcB = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            tcB.append(el)
        el = OxmlElement('w:bottom')
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0'); el.set(qn('w:color'), '000000')
        tcB.append(el)
        pr.append(tcB)

    def _ифио(фио):
        p = (фио or '').split()
        if len(p) >= 3:
            return f'{p[1][0]}. {p[2][0]}. {p[0]}'
        if len(p) == 2:
            return f'{p[1][0]}. {p[0]}'
        return фио or ''

    def _borders_all(cell, sz=4):
        pr = _tcpr(cell)
        tcB = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'),  str(sz))
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), '000000')
            tcB.append(el)
        pr.append(tcB)

    def _cw(cell, w):
        pr  = _tcpr(cell)
        tcW = pr.find(qn('w:tcW'))
        if tcW is None:
            tcW = OxmlElement('w:tcW'); pr.append(tcW)
        tcW.set(qn('w:w'), str(int(w)))
        tcW.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        vA = OxmlElement('w:vAlign')
        vA.set(qn('w:val'), val)
        pr.append(vA)

    def _fill(cell, text='', align=C, pt=PT, bold=False):
        for p in list(cell.paragraphs):
            p._element.getparent().remove(p._element)
        p = cell.add_paragraph()
        _sp0(p); _ls115(p)
        p.alignment = align
        if text:
            _run_r(p, text, pt=pt, bold=bold)
        return p

    def _fill_add(cell, text='', align=C, pt=PT, bold=False):
        p = cell.add_paragraph()
        _sp0(p); _ls115(p)
        p.alignment = align
        if text:
            _run_r(p, text, pt=pt, bold=bold)
        return p

    def _make_tbl(ncols, tw):
        tbl = doc.add_table(rows=0, cols=ncols)
        tPr = tbl._tbl.find(qn('w:tblPr'))
        if tPr is None:
            tPr = OxmlElement('w:tblPr'); tbl._tbl.insert(0, tPr)
        # remove default style / tblLook that override width
        for tag in ('w:tblStyle', 'w:tblLook'):
            for el in tPr.findall(qn(tag)):
                tPr.remove(el)
        jcEl = OxmlElement('w:jc'); jcEl.set(qn('w:val'), 'left'); tPr.append(jcEl)
        tblW = tPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = OxmlElement('w:tblW'); tPr.append(tblW)
        tblW.set(qn('w:w'), str(int(tw))); tblW.set(qn('w:type'), 'dxa')
        ind = tPr.find(qn('w:tblInd'))
        if ind is None:
            ind = OxmlElement('w:tblInd'); tPr.append(ind)
        ind.set(qn('w:w'), '0'); ind.set(qn('w:type'), 'dxa')
        lay = tPr.find(qn('w:tblLayout'))
        if lay is None:
            lay = OxmlElement('w:tblLayout'); tPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        tblBrd = OxmlElement('w:tblBorders')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{side}'); el.set(qn('w:val'), 'none')
            tblBrd.append(el)
        tPr.append(tblBrd)
        return tbl

    def _set_grid(tbl, widths):
        tblGrid = OxmlElement('w:tblGrid')
        for w in widths:
            gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(int(w)))
            tblGrid.append(gc)
        tbl._tbl.insert(1, tblGrid)

    def _row_h_min(row, cm):
        tr   = row._tr
        trPr = tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr'); tr.insert(0, trPr)
        trH = OxmlElement('w:trHeight')
        trH.set(qn('w:val'),   str(round(cm * 1440 / 2.54)))
        trH.set(qn('w:hRule'), 'atLeast')
        trPr.append(trH)

    # ════════════════════════════════════════════════════════════════════
    # УТВЕРЖДАЮ (3-column no-border table: left | gap | right)
    # ════════════════════════════════════════════════════════════════════
    utv = _make_tbl(3, TW)
    _set_grid(utv, [W_SIDE, W_GAP, W_SIDE])

    def _utv_block(cell, должность, название):
        _fill(cell, 'УТВЕРЖДАЮ', align=C, pt=PT, bold=True)
        if должность:
            _fill_add(cell, должность, align=C)
        if название and '«' in название:
            parts = название.split('«', 1)
            if parts[0].strip():
                _fill_add(cell, parts[0].strip(), align=C)
            _fill_add(cell, '«' + parts[1].strip(), align=C)
        elif название:
            _fill_add(cell, название, align=C)

    def _sig_block(cell, руководитель):
        _fill(cell, '', align=C)   # clear
        p = list(cell.paragraphs)[-1]
        _spacing(p, before_pt=6)
        p.alignment = R
        if руководитель:
            фио = руководитель.split()
            if len(фио) >= 3:
                иниц = f'{фио[1][0]}.{фио[2][0]}. {фио[0]}'
            else:
                иниц = руководитель
            _run_r(p, f'___________ {иниц}')
        else:
            _run_r(p, '___________ ___________')
        p2 = _fill_add(cell, '«____»___________202__г.', align=C)
        _spacing(p2, before_pt=4)

    r1 = utv.add_row()
    cs = list(r1.cells)
    for c in cs:
        _no_border(c)
    _cw(cs[0], W_SIDE); _cw(cs[1], W_GAP); _cw(cs[2], W_SIDE)
    _valign(cs[0], 'top'); _valign(cs[2], 'top')
    _utv_block(cs[0], сор.федерация_должность or '', сор.федерация_название or '')
    _utv_block(cs[2], сор.проводящая_орг_должность or '', сор.проводящая_орг_название or '')

    r2 = utv.add_row()
    cs2 = list(r2.cells)
    for c in cs2:
        _no_border(c)
    _cw(cs2[0], W_SIDE); _cw(cs2[1], W_GAP); _cw(cs2[2], W_SIDE)
    _valign(cs2[0], 'top'); _valign(cs2[2], 'top')
    _sig_block(cs2[0], сор.федерация_руководитель or '')
    _sig_block(cs2[2], сор.проводящая_орг_руководитель or '')

    # ════════════════════════════════════════════════════════════════════
    # РЕГЛАМЕНТ title
    # ════════════════════════════════════════════════════════════════════
    _para(before_pt=0, after_pt=10)   # spacer with margin-bottom=10pt after utv

    p_reg = _para(align=C, before_pt=0, after_pt=0)
    _run_r(p_reg, 'РЕГЛАМЕНТ', pt=14, bold=True, char_sp_pt=3)

    _para('проведения спортивных соревнований', align=C)
    if сор.название:
        _para(сор.название.lower(), align=C)
    if сор.возрастная_группа:
        _para(сор.возрастная_группа, align=C)
    if сор.номер_екп:
        _para(f'(С/М в ЕКП № {сор.номер_екп})', align=C)
    _para(after_pt=6)   # spacer

    # ════════════════════════════════════════════════════════════════════
    # 1. Общая информация
    # ════════════════════════════════════════════════════════════════════
    _heading('1. Общая информация')

    ekp_part = ''
    if сор.единый_кп:
        ekp_part = f', {сор.единый_кп}'
    elif сор.номер_екп:
        ekp_part = f', С/М в ЕКП № {сор.номер_екп}'

    _indent(f'1.1. {сор.название or "Соревнования по рафтингу"} проводятся в соответствии:')
    _indent(f'– с Единым календарным планом межрегиональных, всероссийских и международных '
            f'физкультурных мероприятий и спортивных мероприятий{ekp_part};')

    if сор.положение_о_соревнованиях:
        _indent(f'– с Положением о межрегиональных и всероссийских официальных спортивных '
                f'соревнованиях по рафтингу, {сор.положение_о_соревнованиях};')
    else:
        _indent('– с «Положением о спортивных соревнованиях по рафтингу», утверждённым '
                'Минспортом России и Федерацией рафтинга России;')

    pvt = '– с правилами вида спорта «рафтинг»'
    if сор.правила_вида_спорта:
        pvt += f', {сор.правила_вида_спорта}'
    _indent(pvt + ';')
    _indent('– с настоящим регламентом.')
    _indent(f'1.2. Сроки проведения спортивных соревнований: '
            f'{сор.даты_проведения or "___"}. '
            f'Место проведения: {сор.место_проведения or "___"}.')

    # ════════════════════════════════════════════════════════════════════
    # 2. Организаторы
    # ════════════════════════════════════════════════════════════════════
    _heading('2. Организаторы соревнований')
    _indent('2.1. Организаторами спортивных соревнований являются:')
    for орг in [сор.организатор_1, сор.организатор_2,
                сор.организатор_3, сор.организатор_4]:
        if орг:
            _indent(f'– {орг};')

    # ════════════════════════════════════════════════════════════════════
    # 3. Организация, проводящая соревнования
    # ════════════════════════════════════════════════════════════════════
    _heading('3. Организация, проводящая соревнования')
    _indent(f'3.1. Организацией, проводящей спортивные соревнования, является '
            f'{сор.проводящая_орг_название or "___"} (далее — Проводящая организация).')
    _indent('3.2. Непосредственное проведение спортивных соревнований возлагается на '
            'судейскую коллегию соревнований, утвержденную Проводящей организацией и '
            'согласованную ФРР.')
    судья = сор.главный_судья or '___'
    if сор.главный_судья_телефон:
        судья += f', телефон: {сор.главный_судья_телефон}'
    _indent(f'3.3. Главный судья соревнований {судья}.')
    _indent('3.4. Протоколы спортивных соревнований и другие документы судейской коллегии '
            'оформляются за подписью главного судьи и главного секретаря спортивных '
            'соревнований и заверяются печатью Проводящей организации.')

    # ════════════════════════════════════════════════════════════════════
    # 4. Распределение обязанностей
    # ════════════════════════════════════════════════════════════════════
    _heading('4. Распределение обязанностей организаторов\n'
             'и организации, проводящей соревнования')
    _indent('4.1. В соответствии с Положением организаторы определяют условия '
            'проведения спортивных соревнований.')
    _indent(f'4.2. {сор.федерация_название or "Федерация рафтинга России"} при организации '
            f'и проведении спортивных соревнований обеспечивает:')
    if сор.обязанности_федерации:
        for стр in сор.обязанности_федерации.split('\n'):
            if стр.strip():
                if стр.strip().startswith('–') or стр.strip().startswith('-'):
                    _list_dash(стр.strip())
                else:
                    _indent(стр.strip())
    else:
        _list_dash('– размещение информации о спортивных соревнованиях на сайте rusraftfed.ru;')
        _list_dash('– утверждение состава главной судейской коллегии соревнований;')
        _list_dash('– утверждение кандидатуры технического делегата ФРР.')
    _indent(f'4.3. {сор.проводящая_орг_название or "Проводящая организация"} при организации '
            f'и проведении спортивных соревнований обеспечивает:')
    if сор.обязанности_проводящей:
        for стр in сор.обязанности_проводящей.split('\n'):
            if стр.strip():
                if стр.strip().startswith('–') or стр.strip().startswith('-'):
                    _list_dash(стр.strip())
                else:
                    _indent(стр.strip())
    else:
        _list_dash('– формирование судейской коллегии и направление состава ГСК для '
                   'утверждения в ФРР;')
        _list_dash('– деятельность судейской коллегии во время проведения спортивных '
                   'соревнований;')
        _list_dash('– организацию материально-технического, транспортного и медицинского '
                   'обеспечения соревнований;')
        _list_dash('– организацию мер безопасности спортсменов и зрителей;')
        _list_dash('– организацию проживания и питания участников;')
        _list_dash('– направление в ФРР по итогам соревнований отчёта, итоговых протоколов '
                   'и справки о составе судейской коллегии.')

    # ════════════════════════════════════════════════════════════════════
    # 5. Программа соревнований
    # ════════════════════════════════════════════════════════════════════
    _heading('5. Программа соревнований')

    if программа:
        sched = _make_tbl(3, TW)
        _set_grid(sched, [W_DATE, W_TIME, W_NAME])

        hdr = sched.add_row()
        hcs = list(hdr.cells)
        _cw(hcs[0], W_DATE); _cw(hcs[1], W_TIME); _cw(hcs[2], W_NAME)
        for hc in hcs:
            _borders_all(hc); _valign(hc)
        _fill(hcs[0], 'Дата',         align=C, bold=True)
        _fill(hcs[1], 'Время',        align=C, bold=True)
        _fill(hcs[2], 'Наименование', align=C, bold=True)
        # keep header row with first data row; repeat on each page if table spans
        trPr = hdr._tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr'); hdr._tr.insert(0, trPr)
        if trPr.find(qn('w:tblHeader')) is None:
            trPr.append(OxmlElement('w:tblHeader'))
        for hc in hcs:
            for p in hc.paragraphs:
                p.paragraph_format.keep_with_next = True

        seen = {}
        for пункт in программа:
            dr  = sched.add_row(); _row_h_min(dr, 0.5)
            dcs = list(dr.cells)
            _cw(dcs[0], W_DATE); _cw(dcs[1], W_TIME); _cw(dcs[2], W_NAME)
            _borders_all(dcs[1]); _valign(dcs[1])
            _borders_all(dcs[2]); _valign(dcs[2])
            d = пункт.get('дата', '') or ''
            if d not in seen:
                seen[d] = True
                pr = _tcpr(dcs[0])
                vm = OxmlElement('w:vMerge'); vm.set(qn('w:val'), 'restart')
                pr.append(vm)
                _borders_all(dcs[0]); _valign(dcs[0])
                _fill(dcs[0], d, align=C)
            else:
                pr = _tcpr(dcs[0])
                vm = OxmlElement('w:vMerge'); pr.append(vm)
                _borders_all(dcs[0])
                _fill(dcs[0], '')
            _fill(dcs[1], пункт.get('время', '') or '',        align=L)
            _fill(dcs[2], пункт.get('наименование', '') or '', align=L)

    # ════════════════════════════════════════════════════════════════════
    # 6. Заявочный взнос
    # ════════════════════════════════════════════════════════════════════
    _heading('6. Дополнительные требования к участникам, заявочный взнос')
    _indent('6.1. Сборные команды субъектов Российской Федерации должны иметь единую '
            'парадную форму и флаг субъекта Российской Федерации.')
    _indent('6.2. Все члены сборных команд субъектов Российской Федерации, включённые '
            'в заявку, должны быть включены в действующие списки сборной команды субъекта '
            'Российской Федерации по рафтингу.')
    _indent('6.3. К соревнованиям допускаются сборные команды субъектов Российской '
            'Федерации, которые оплатили заявочный взнос за участие в спортивных '
            'соревнованиях.')
    _indent(f'6.4. Размер заявочного взноса за участие в спортивных соревнованиях в классе '
            f'судов R4: {сор.заявочный_взнос_r4 or "___"}, в классе судов R6: '
            f'{сор.заявочный_взнос or "___"}.')
    if сор.назначение_платежа:
        _indent(f'Назначение платежа: {сор.назначение_платежа}.')
    if сор.реквизиты_оплаты:
        _indent('6.5. Оплата взносов производится по безналичному расчёту по следующим '
                'реквизитам:')
        _indent_block(сор.реквизиты_оплаты)

    # ════════════════════════════════════════════════════════════════════
    # 7. Условия финансирования
    # ════════════════════════════════════════════════════════════════════
    _heading('7. Условия финансирования')
    _indent('7.1. Расходы по организации и проведению спортивных соревнований несут '
            'организаторы в соответствии с установленными полномочиями.')
    _indent('7.2. За счёт средств заявочных взносов Проводящая организация обеспечивает '
            'медицинское обеспечение, меры безопасности и материально-техническое '
            'обеспечение спортивных соревнований.')

    # ════════════════════════════════════════════════════════════════════
    # 8. Размещение и проживание
    # ════════════════════════════════════════════════════════════════════
    _heading('8. Размещение и проживание участников')
    п8 = 1
    if сор.проживание_описание:
        строки8 = [с.strip() for с in сор.проживание_описание.split('\n') if с.strip()]
        _indent(f'8.{п8}. {строки8[0]}')
        for стр in строки8[1:]:
            _indent(стр)
    else:
        _indent(f'8.{п8}. Размещение участников осуществляется самостоятельно.')
    п8 += 1
    _indent(f'8.{п8}. Питание участники организовывают самостоятельно.')
    п8 += 1
    if сор.инфраструктура_описание:
        строки8и = [с.strip() for с in сор.инфраструктура_описание.split('\n') if с.strip()]
        _indent(f'8.{п8}. {строки8и[0]}')
        for стр in строки8и[1:]:
            _indent(стр)

    # ════════════════════════════════════════════════════════════════════
    # 9. Награждение
    # ════════════════════════════════════════════════════════════════════
    _heading('9. Награждение')
    if сор.награждение_описание:
        строки9 = [с.strip() for с in сор.награждение_описание.split('\n') if с.strip()]
        _indent(f'9.1. {строки9[0]}')
        for стр in строки9[1:]:
            _indent(стр)
    else:
        _indent('9.1. Награждение победителей и призёров спортивных соревнований проводится '
                'в соответствии с Положением о соревнованиях.')

    # ════════════════════════════════════════════════════════════════════
    # 10. Контакты
    # ════════════════════════════════════════════════════════════════════
    _heading('10. Контакты организаторов')
    контакт = сор.главный_судья or '___'
    if сор.главный_судья_телефон:
        контакт += f', телефон: {сор.главный_судья_телефон}'
    if сор.проводящая_орг_email:
        контакт += f', email: {сор.проводящая_орг_email}'
    _indent(f'10.1. Вопросы, связанные с подачей заявок и условиями проведения '
            f'спортивных соревнований: {контакт}.')

    # ════════════════════════════════════════════════════════════════════
    # 11. Варианты проезда
    # ════════════════════════════════════════════════════════════════════
    _heading('11. Варианты проезда к месту проведения соревнований')
    _indent('11.1. Общественным транспортом:')
    for стр11 in (сор.проезд_общественный or '___').split('\n'):
        if стр11.strip():
            _indent(стр11.strip())
    _indent('11.2. Собственным автомобильным транспортом:')
    for стр11 in (сор.проезд_собственный or '___').split('\n'):
        if стр11.strip():
            _indent(стр11.strip())

    # ── Блок подписей ────────────────────────────────────────────────────────
    from datetime import date as _date
    _месяцы = {1:'января',2:'февраля',3:'марта',4:'апреля',5:'мая',6:'июня',
               7:'июля',8:'августа',9:'сентября',10:'октября',11:'ноября',12:'декабря'}
    W_POS   = round(TW * 0.47)
    W_BLANK = round(TW * 0.15)
    W_FIO   = TW - W_POS - W_BLANK

    _para(before_pt=12)
    sig = _make_tbl(3, TW)
    _set_grid(sig, [W_POS, W_BLANK, W_FIO])

    for должность, фио_raw in (
        ('Главный судья соревнований',     сор.главный_судья     or ''),
        ('Главный секретарь соревнований', сор.главный_секретарь or ''),
    ):
        # Строка 1: должность | пустая черта | И.О. Фамилия на черте
        r1 = sig.add_row(); _row_h_min(r1, 0.8)
        cs = list(r1.cells)
        _cw(cs[0], W_POS); _cw(cs[1], W_BLANK); _cw(cs[2], W_FIO)
        _no_border(cs[0]); _bottom_border(cs[1]); _bottom_border(cs[2])
        _valign(cs[0], 'bottom'); _valign(cs[1], 'bottom'); _valign(cs[2], 'bottom')
        _fill(cs[0], должность, align=L, pt=PT)
        _fill(cs[1], '',              align=C, pt=PT)
        _fill(cs[2], _ифио(фио_raw),  align=C, pt=PT)

        # Строка 2: пусто | (подпись) | (расшифровка подписи)
        r2 = sig.add_row(); _row_h_min(r2, 0.55)
        cs2 = list(r2.cells)
        _cw(cs2[0], W_POS); _cw(cs2[1], W_BLANK); _cw(cs2[2], W_FIO)
        _no_border(cs2[0]); _no_border(cs2[1]); _no_border(cs2[2])
        _fill(cs2[0], '',                      align=L, pt=8)
        _fill(cs2[1], '(подпись)',              align=C, pt=8)
        _fill(cs2[2], '(расшифровка подписи)',  align=C, pt=8)

    # Дата
    _сегодня = _date.today()
    _para(f'{_сегодня.day} {_месяцы[_сегодня.month]} {_сегодня.year} года',
          align=L, pt=PT, before_pt=10)

    buf = BytesIO()
    doc.save(buf); buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_гск_docx(сор, судьи, filename):
    """Generate ГСК composition Word document matching образцы_справок/Состав ГСК.docx."""
    from docx.shared import Mm, Pt

    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT
    R   = WD_ALIGN_PARAGRAPH.RIGHT

    # A4 portrait, L=10, R=5, T=15, B=15
    TW = round((210 - 10 - 5) * 1440 / 25.4)   # 11054 dxa

    # Column widths scaled from sample proportions (494:2762:3901:1719:2068)
    # Column widths: №=477, Должность=2664, ФИО=3024, Квалификация=2400, Субъект=1995
    GRID = [477, 2664, 3998, 1920, 1995]


    doc  = Document()
    sect = doc.sections[0]
    sect.page_width    = Mm(210)
    sect.page_height   = Mm(297)
    sect.left_margin   = Mm(10)
    sect.right_margin  = Mm(5)
    sect.top_margin    = Mm(15)
    sect.bottom_margin = Mm(15)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        pgSz.attrib.pop(qn('w:orient'), None)
    doc.styles['Normal'].font.name = TNR
    doc.styles['Normal'].font.size = Pt(10)

    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:p'):
            body.remove(child)

    def _sp(p, before=0, after=0):
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after  = Pt(after)

    def _run(p, text, pt=10, bold=False, underline=False):
        r = p.add_run(text)
        r.font.name = TNR; r.font.size = Pt(pt)
        if bold:      r.font.bold      = True
        if underline: r.font.underline = True
        return r

    def _para(text='', align=None, pt=10, bold=False, before=0, after=0):
        p = doc.add_paragraph()
        _sp(p, before, after)
        if align: p.alignment = align
        if text:  _run(p, text, pt, bold)
        return p

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _cw(cell, w):
        pr = _tcpr(cell)
        el = pr.find(qn('w:tcW'))
        if el is None: el = OxmlElement('w:tcW'); pr.append(el)
        el.set(qn('w:w'), str(w)); el.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        v  = pr.find(qn('w:vAlign'))
        if v is None: v = OxmlElement('w:vAlign'); pr.append(v)
        v.set(qn('w:val'), val)

    def _borders(cell, sz=4):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')): pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')
            b.append(el)
        pr.append(b)

    def _no_border(cell):
        pr = _tcpr(cell)
        for old in pr.findall(qn('w:tcBorders')): pr.remove(old)
        b = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
            b.append(el)
        pr.append(b)

    def _bottom_border(cell, sz=6):
        _no_border(cell)
        pr  = _tcpr(cell)
        b   = pr.find(qn('w:tcBorders'))
        el  = OxmlElement('w:bottom')
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')
        b.append(el)

    def _fill(cell, text, align=None, pt=10, bold=False):
        p = cell.paragraphs[0]; p.clear(); _sp(p)
        if align: p.alignment = align
        if text:  _run(p, str(text), pt, bold)
        _valign(cell)

    def _row_h(row, cm):
        val   = round(cm * 1440 / 2.54)
        trPr  = row._tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr'); row._tr.insert(0, trPr)
        trH = trPr.find(qn('w:trHeight'))
        if trH is None: trH = OxmlElement('w:trHeight'); trPr.append(trH)
        trH.set(qn('w:val'), str(val)); trH.set(qn('w:hRule'), 'atLeast')

    def _make_tbl(ncols, width=None):
        t    = doc.add_table(rows=0, cols=ncols)
        tbl  = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        el.set(qn('w:w'), str(width or TW)); el.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        return t

    def _set_grid(tbl, grid):
        t = tbl._tbl
        for old in t.findall(qn('w:tblGrid')): t.remove(old)
        g = OxmlElement('w:tblGrid')
        for w in grid:
            col = OxmlElement('w:gridCol'); col.set(qn('w:w'), str(w)); g.append(col)
        tblPr = t.find(qn('w:tblPr'))
        tblPr.addnext(g) if tblPr is not None else t.insert(0, g)

    def _abbr_surname_first(фио):
        """Фамилия И.О. (Собетов В. А.)"""
        p = (фио or '').split()
        if len(p) >= 3:
            return f"{p[0]} {p[1][0]}. {p[2][0]}."
        if len(p) == 2:
            return f"{p[0]} {p[1][0]}."
        return фио or ''

    def _indent(p, left, first=0):
        """Установить левый отступ параграфа (dxa)."""
        pPr = p._p.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr'); p._p.insert(0, pPr)
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind'); pPr.append(ind)
        ind.set(qn('w:left'), str(left))
        if first:
            ind.set(qn('w:firstLine'), str(first))

    # ── Шапка: организаторы ───────────────────────────────────────────────────

    if сор.организаторы_блок:
        lines = [l.strip() for l in сор.организаторы_блок.split('\n') if l.strip()]
        p_org = _para(align=C, pt=11)

        def _line_br_org(p):
            p.add_run()._r.append(OxmlElement('w:br'))

        for i, line in enumerate(lines):
            if i > 0:
                _line_br_org(p_org)
            qi = next((line.index(q) for q in ('"', '«') if q in line), None)
            if qi is None:
                _run(p_org, line, pt=11)
            else:
                before  = line[:qi].strip()
                quoted  = line[qi:]
                combined = (before + ' ' + quoted) if before else quoted
                if len(combined) <= 80:
                    _run(p_org, combined, pt=11)
                else:
                    if before:
                        _run(p_org, before, pt=11)
                        _line_br_org(p_org)
                    _run(p_org, quoted, pt=11)

    _para(сор.название or '', align=C, pt=13, bold=True, after=6)

    # ── Дата / место (таблица без рамок, нижняя черта) ───────────────────────

    MAR  = round(0.3 * 1440 / 2.54)   # ~170 dxa padding
    PT10 = round(10 * 0.6 * 20)       # примерная ширина символа 10pt в dxa
    _w_dt  = max(1500, len(сор.даты_проведения  or 'Дата проведения') * PT10 + 2 * MAR)
    _w_pl  = max(2000, len(сор.место_проведения or 'Место проведения') * PT10 + 2 * MAR)
    _w_mid = max(200, TW - _w_dt - _w_pl)
    dt = _make_tbl(3)
    _set_grid(dt, [_w_dt, _w_mid, _w_pl])
    for idx, (лево, право) in enumerate([
        (сор.даты_проведения  or '', сор.место_проведения or ''),
        ('Дата проведения', 'Место проведения'),
    ]):
        r = dt.add_row()
        cs = list(r.cells)
        va = 'bottom' if idx == 0 else 'top'
        for c in cs: _no_border(c); _valign(c, va)
        if idx == 0:
            _bottom_border(cs[0]); _bottom_border(cs[2])
        for cell, text in ((cs[0], лево), (cs[2], право)):
            p = cell.paragraphs[0]; p.clear(); _sp(p)
            pPr = p._p.get_or_add_pPr()
            pStyle = pPr.find(qn('w:pStyle'))
            if pStyle is None:
                pStyle = OxmlElement('w:pStyle'); pPr.insert(0, pStyle)
            pStyle.set(qn('w:val'), 'NoSpacing')
            p.alignment = C
            if text: _run(p, text, pt=11)
            _valign(cell, va)

    # ── УТВЕРЖДАЮ ─────────────────────────────────────────────────────────────
    # Блок начинается на 2/3 ширины страницы (w:left = TW*2/3)
    IND = round(TW * 2 // 3)  # ~7369 dxa, ~130мм от левого поля

    _para()
    _indent(_para('УТВЕРЖДАЮ', pt=13, bold=True, align=C), IND)

    if сор.проводящая_орг_должность:
        _indent(_para(сор.проводящая_орг_должность.strip(), pt=11, after=3, align=C), IND)

    if сор.проводящая_орг_название:
        for line in сор.проводящая_орг_название.split('\n'):
            if line.strip():
                _indent(_para(line.strip(), pt=11, after=3, align=C), IND)

    _para()
    руков = _abbr_surname_first(сор.проводящая_орг_руководитель or '')
    _indent(_para(руков, pt=11, align=R), IND)
    _para()
    _indent(_para('«____» ___________ 202__ г.', pt=11, align=C), IND)
    _para()

    # ── Заголовок ─────────────────────────────────────────────────────────────

    _para('Состав главной судейской коллегии', align=C, pt=13, bold=True)
    _para()
    _para()

    # ── Основная таблица ──────────────────────────────────────────────────────

    tbl = _make_tbl(5)
    _set_grid(tbl, GRID)

    # Заголовок таблицы
    rh = tbl.add_row(); _row_h(rh, 1.0)
    cs = list(rh.cells)
    for i, w in enumerate(GRID): _cw(cs[i], w); _borders(cs[i])
    _fill(cs[0], '№',                               align=C, pt=11, bold=True)
    _fill(cs[1], 'Должность',                       align=C, pt=11, bold=True)
    _fill(cs[2], 'Фамилия, имя, отчество',          align=C, pt=11, bold=True)
    _fill(cs[3], 'Квалификация спортивного судьи',  align=C, pt=11, bold=True)
    _fill(cs[4], 'Субъект Российской Федерации',    align=C, pt=11, bold=True)

    for пп, с in enumerate(судьи, 1):
        rd = tbl.add_row(); _row_h(rd, 1.0)
        cd = list(rd.cells)
        for i, w in enumerate(GRID): _cw(cd[i], w); _borders(cd[i])
        _fill(cd[0], str(пп),                         align=C, pt=11)
        _fill(cd[1], с.должность or '—',              pt=11)
        _fill(cd[2], с.фамилия_имя_отчество or '—',   pt=11)
        _fill(cd[3], с.квалификация or '—',           align=C, pt=11)
        _fill(cd[4], с.субъект_рф or '—',             align=C, pt=11)

    # ── Подпись: Главный судья ────────────────────────────────────────────────

    _para()
    _гс = next((с for с in судьи if с.должность == 'Главный судья'), None)
    _гс_фио = _гс.фамилия_имя_отчество if _гс else (сор.главный_судья or '')
    _гс_аббр = _abbr_surname_first(_гс_фио)

    w_l = TW * 43 // 100
    w_r = TW - w_l
    sig = _make_tbl(2)
    _set_grid(sig, [w_l, w_r])
    rs = sig.add_row(); _row_h(rs, 1.2)
    cs = list(rs.cells)
    _cw(cs[0], w_l); _cw(cs[1], w_r)
    _no_border(cs[0]); _no_border(cs[1])
    _valign(cs[0]); _valign(cs[1])
    _fill(cs[0], 'Главный судья', align=L, pt=11)
    _fill(cs[1], _гс_аббр,        align=L, pt=11)

    buf = BytesIO()
    _apply_no_break(doc)
    doc.save(buf); buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_минспорта_docx(сор, судьи, filename):
    """Список судейской коллегии по форме Минспорта — Portrait A4, L=30 R=15 T=7.5 B=10."""
    from docx.shared import Mm, Pt

    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT

    # A4 portrait, L=30, R=15, T=7.5, B=10
    TW   = round((210 - 30 - 15) * 1440 / 25.4)   # 9354 dxa
    GRID = [420, 3265, 2456, 1356, 1857]            # №, ФИО, Должность, Кв, Субъект

    doc  = Document()
    sect = doc.sections[0]
    sect.page_width    = Mm(210)
    sect.page_height   = Mm(297)
    sect.left_margin   = Mm(30)
    sect.right_margin  = Mm(15)
    sect.top_margin    = Mm(7)
    sect.bottom_margin = Mm(10)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        pgSz.attrib.pop(qn('w:orient'), None)
    doc.styles['Normal'].font.name = TNR
    doc.styles['Normal'].font.size = Pt(11)

    # ── хелперы ──────────────────────────────────────────────────────────────

    def _sp(p):
        pPr = p._p.get_or_add_pPr()
        for tag in ('w:spacing', 'w:contextualSpacing'):
            old = pPr.find(qn(tag))
            if old is not None: pPr.remove(old)
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0')
        pPr.append(sp)

    def _run(p, text, pt=11, bold=False):
        r = p.add_run(text)
        r.font.name = TNR
        r.font.size = Pt(pt)
        r.bold      = bold
        try:
            r._r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), TNR)
        except Exception:
            pass
        return r

    def _para(text='', align=L, pt=11, bold=False):
        p = doc.add_paragraph()
        _sp(p)
        p.alignment = align
        if text:
            _run(p, text, pt=pt, bold=bold)
        return p

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _borders(cell, sz=4):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')
            b.append(el)

    def _no_border(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
            b.append(el)

    def _cw(cell, w):
        pr = _tcpr(cell)
        tw = pr.find(qn('w:tcW'))
        if tw is None:
            tw = OxmlElement('w:tcW'); pr.append(tw)
        tw.set(qn('w:w'), str(w)); tw.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        va = pr.find(qn('w:vAlign'))
        if va is None:
            va = OxmlElement('w:vAlign'); pr.append(va)
        va.set(qn('w:val'), val)

    def _row_h(row, cm):
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr'); row._tr.insert(0, trPr)
        rh = trPr.find(qn('w:trHeight'))
        if rh is None:
            rh = OxmlElement('w:trHeight'); trPr.append(rh)
        rh.set(qn('w:val'), str(round(cm * 1440 / 2.54)))
        rh.set(qn('w:hRule'), 'atLeast')

    def _fill(cell, text, align=None, pt=11, bold=False):
        p = cell.paragraphs[0]; p.clear(); _sp(p)
        if align: p.alignment = align
        if text:  _run(p, str(text), pt, bold)
        _valign(cell)

    def _make_tbl(ncols):
        t    = doc.add_table(rows=0, cols=ncols)
        tbl  = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        el.set(qn('w:w'), str(TW)); el.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        return t

    def _set_grid(tbl, grid):
        t = tbl._tbl
        for old in t.findall(qn('w:tblGrid')): t.remove(old)
        g = OxmlElement('w:tblGrid')
        for w in grid:
            col = OxmlElement('w:gridCol'); col.set(qn('w:w'), str(w)); g.append(col)
        pr = t.find(qn('w:tblPr'))
        pr.addnext(g) if pr is not None else t.insert(0, g)

    def _abbr(фио):
        p = (фио or '').split()
        if len(p) >= 3: return f"{p[0]} {p[1][0]}. {p[2][0]}."
        if len(p) == 2: return f"{p[0]} {p[1][0]}."
        return фио or ''

    # ── Организаторы ──────────────────────────────────────────────────────────

    if сор.организаторы_блок:
        for line in сор.организаторы_блок.split('\n'):
            if line.strip():
                _para(line.strip(), align=C, pt=9, bold=True)

    _para()

    # ── Заголовок ─────────────────────────────────────────────────────────────

    _para('СПИСОК', align=C, pt=18, bold=True)
    _para('СУДЕЙСКОЙ КОЛЛЕГИИ ДЛЯ ПРОВЕДЕНИЯ', align=C, pt=14, bold=True)
    for line in (сор.название or '').split('\n'):
        if line.strip():
            _para(line.strip(), align=C, pt=14, bold=True)

    _para()

    # ── Дата и место ─────────────────────────────────────────────────────────

    date_place = ', '.join(x for x in (сор.даты_проведения or '', сор.место_проведения or '') if x)
    if date_place:
        _para(date_place, align=L, pt=10, bold=True)

    _para()
    _para()

    # ── Основная таблица ─────────────────────────────────────────────────────

    tbl = _make_tbl(5)
    _set_grid(tbl, GRID)

    rh = tbl.add_row(); _row_h(rh, 1.0)
    cs = list(rh.cells)
    for i, w in enumerate(GRID): _cw(cs[i], w); _borders(cs[i])
    _fill(cs[0], '№',                              align=C, pt=11, bold=True)
    _fill(cs[1], 'Фамилия, имя, отчество',         align=C, pt=11, bold=True)
    _fill(cs[2], 'Должность',                      align=C, pt=11, bold=True)
    _fill(cs[3], 'Квалификация спортивного судьи', align=C, pt=11, bold=True)
    _fill(cs[4], 'Субъект Российской Федерации',   align=C, pt=11, bold=True)

    for пп, с in enumerate(судьи, 1):
        rd = tbl.add_row(); _row_h(rd, 0.7)
        cd = list(rd.cells)
        for i, w in enumerate(GRID): _cw(cd[i], w); _borders(cd[i])
        _fill(cd[0], str(пп),                       align=C, pt=11)
        _fill(cd[1], с.фамилия_имя_отчество or '—', pt=11)
        _fill(cd[2], с.должность or '—',            pt=11)
        _fill(cd[3], с.квалификация or '—',         align=C, pt=11)
        _fill(cd[4], с.субъект_рф or '—',           align=C, pt=11)

    # ── Подписи ───────────────────────────────────────────────────────────────

    _para()

    w_l = TW * 45 // 100
    w_r = TW - w_l
    for pos, фио in (
        ('Главный судья',    сор.главный_судья     or ''),
        ('Главный секретарь', сор.главный_секретарь or ''),
    ):
        sig = _make_tbl(2)
        _set_grid(sig, [w_l, w_r])
        rs = sig.add_row(); _row_h(rs, 1.2)
        cs = list(rs.cells)
        _cw(cs[0], w_l); _cw(cs[1], w_r)
        _no_border(cs[0]); _no_border(cs[1])
        _valign(cs[0]); _valign(cs[1])
        _fill(cs[0], pos,       align=L, pt=10, bold=True)
        _fill(cs[1], _abbr(фио), align=L, pt=10, bold=True)

    buf = BytesIO()
    _apply_no_break(doc)
    doc.save(buf); buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)


def create_справка_docx(сор, судьи, стат, filename):
    """Справка о составе и квалификации судейской коллегии — Portrait A4, L=10 R=15 T=15 B=15."""
    from docx.shared import Mm, Pt

    TNR = 'Times New Roman'
    C   = WD_ALIGN_PARAGRAPH.CENTER
    L   = WD_ALIGN_PARAGRAPH.LEFT

    # A4 portrait, L=10, R=15, T=15, B=15
    TW   = round((210 - 10 - 15) * 1440 / 25.4)   # 10488 dxa
    GRID = [473, 2647, 3739, 1647, 1982]            # №, Должность, ФИО, Кв, Субъект
    INFO = [TW // 2, TW - TW // 2]                  # инфо-таблица: равные половины

    doc  = Document()
    sect = doc.sections[0]
    sect.page_width    = Mm(210)
    sect.page_height   = Mm(297)
    sect.left_margin   = Mm(10)
    sect.right_margin  = Mm(15)
    sect.top_margin    = Mm(15)
    sect.bottom_margin = Mm(15)
    pgSz = sect._sectPr.find(qn('w:pgSz'))
    if pgSz is not None:
        pgSz.attrib.pop(qn('w:orient'), None)
    doc.styles['Normal'].font.name = TNR
    doc.styles['Normal'].font.size = Pt(14)

    # ── хелперы ──────────────────────────────────────────────────────────────

    def _sp(p):
        pPr = p._p.get_or_add_pPr()
        for tag in ('w:spacing', 'w:contextualSpacing'):
            old = pPr.find(qn(tag))
            if old is not None: pPr.remove(old)
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0')
        pPr.append(sp)

    def _run(p, text, pt=14, bold=False):
        r = p.add_run(text)
        r.font.name = TNR
        r.font.size = Pt(pt)
        r.bold      = bold
        try:
            r._r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), TNR)
        except Exception:
            pass
        return r

    def _para(text='', align=L, pt=14, bold=False):
        p = doc.add_paragraph()
        _sp(p)
        p.alignment = align
        if text:
            _run(p, text, pt=pt, bold=bold)
        return p

    def _tcpr(cell):
        tc = cell._tc
        pr = tc.find(qn('w:tcPr'))
        if pr is None:
            pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
        return pr

    def _borders(cell, sz=4):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0');    el.set(qn('w:color'), '000000')
            b.append(el)

    def _no_border(cell):
        pr = _tcpr(cell)
        b = pr.find(qn('w:tcBorders'))
        if b is None:
            b = OxmlElement('w:tcBorders'); pr.append(b)
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
            b.append(el)

    def _cw(cell, w):
        pr = _tcpr(cell)
        tw = pr.find(qn('w:tcW'))
        if tw is None:
            tw = OxmlElement('w:tcW'); pr.append(tw)
        tw.set(qn('w:w'), str(w)); tw.set(qn('w:type'), 'dxa')

    def _valign(cell, val='center'):
        pr = _tcpr(cell)
        va = pr.find(qn('w:vAlign'))
        if va is None:
            va = OxmlElement('w:vAlign'); pr.append(va)
        va.set(qn('w:val'), val)

    def _row_h(row, cm):
        trPr = row._tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr'); row._tr.insert(0, trPr)
        rh = trPr.find(qn('w:trHeight'))
        if rh is None:
            rh = OxmlElement('w:trHeight'); trPr.append(rh)
        rh.set(qn('w:val'), str(round(cm * 1440 / 2.54)))
        rh.set(qn('w:hRule'), 'atLeast')

    def _fill(cell, text, align=None, pt=14, bold=False):
        p = cell.paragraphs[0]; p.clear(); _sp(p)
        if align: p.alignment = align
        if text:  _run(p, str(text), pt, bold)
        _valign(cell)

    def _make_tbl(ncols, tw=None):
        t    = doc.add_table(rows=0, cols=ncols)
        tbl  = t._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); tbl.insert(0, tblPr)
        for s in tblPr.findall(qn('w:tblStyle')): tblPr.remove(s)
        el = tblPr.find(qn('w:tblW'))
        if el is None: el = OxmlElement('w:tblW'); tblPr.append(el)
        el.set(qn('w:w'), str(tw or TW)); el.set(qn('w:type'), 'dxa')
        lay = tblPr.find(qn('w:tblLayout'))
        if lay is None: lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        return t

    def _set_grid(tbl, grid):
        t = tbl._tbl
        for old in t.findall(qn('w:tblGrid')): t.remove(old)
        g = OxmlElement('w:tblGrid')
        for w in grid:
            col = OxmlElement('w:gridCol'); col.set(qn('w:w'), str(w)); g.append(col)
        pr = t.find(qn('w:tblPr'))
        pr.addnext(g) if pr is not None else t.insert(0, g)

    def _abbr(фио):
        p = (фио or '').split()
        if len(p) >= 3: return f"{p[0]} {p[1][0]}. {p[2][0]}."
        if len(p) == 2: return f"{p[0]} {p[1][0]}."
        return фио or ''

    # ── Организаторы ──────────────────────────────────────────────────────────

    if сор.организаторы_блок:
        for line in сор.организаторы_блок.split('\n'):
            if line.strip():
                _para(line.strip(), align=C, pt=14)

    _para()
    _para()

    # ── Заголовок ─────────────────────────────────────────────────────────────

    _para('СПРАВКА', align=C, pt=14, bold=True)
    _para('о составе и квалификации судейской коллегии', align=C, pt=14, bold=True)

    _para()
    _para()

    # ── Инфо-таблица ─────────────────────────────────────────────────────────

    info = _make_tbl(2)
    _set_grid(info, INFO)
    for лейбл, значение in (
        ('Наименование соревнований', сор.название or ''),
        ('Дата проведения',           сор.даты_проведения or ''),
        ('Место проведения',          сор.место_проведения or ''),
    ):
        ri = info.add_row(); _row_h(ri, 0.8)
        cs = list(ri.cells)
        _cw(cs[0], INFO[0]); _cw(cs[1], INFO[1])
        _borders(cs[0]); _borders(cs[1])
        _fill(cs[0], лейбл,   align=L, pt=14, bold=True)
        _fill(cs[1], значение, align=L, pt=14)

    _para()

    # ── Статистика ────────────────────────────────────────────────────────────

    всего      = стат.get('всего', 0)
    иногородних = стат.get('иногородних', 0)
    по_квал    = стат.get('по_квал', {})

    _para(f'Всего спортивных судей {всего} человек, в том числе из других регионов {иногородних} человек.', pt=14)
    _para(f'Спортивных судей всероссийской категории (ССВК) – {по_квал.get("ССВК", 0)} человек.', pt=14)
    _para(f'Спортивных судей первой категории (СС1К) – {по_квал.get("СС1К", 0)} человек.', pt=14)
    _para(f'Спортивных судей второй категории (СС2К) – {по_квал.get("СС2К", 0)} человек.', pt=14)
    _para(f'Спортивных судей третьей категории (СС3К) – {по_квал.get("СС3К", 0)} человек.', pt=14)

    _para()
    _para()

    # ── Основная таблица ─────────────────────────────────────────────────────

    tbl = _make_tbl(5)
    _set_grid(tbl, GRID)

    rh = tbl.add_row(); _row_h(rh, 1.0)
    cs = list(rh.cells)
    for i, w in enumerate(GRID): _cw(cs[i], w); _borders(cs[i])
    _fill(cs[0], '№',                              align=C, pt=11, bold=True)
    _fill(cs[1], 'Должность',                      align=C, pt=11, bold=True)
    _fill(cs[2], 'Фамилия, имя, отчество',         align=C, pt=11, bold=True)
    _fill(cs[3], 'Квалификация спортивного судьи', align=C, pt=11, bold=True)
    _fill(cs[4], 'Субъект Российской Федерации',   align=C, pt=11, bold=True)

    for пп, с in enumerate(судьи, 1):
        rd = tbl.add_row(); _row_h(rd, 0.7)
        cd = list(rd.cells)
        for i, w in enumerate(GRID): _cw(cd[i], w); _borders(cd[i])
        _fill(cd[0], str(пп),                       align=C, pt=11)
        _fill(cd[1], с.должность or '—',            pt=11)
        _fill(cd[2], с.фамилия_имя_отчество or '—', pt=11)
        _fill(cd[3], с.квалификация or '—',         align=C, pt=11)
        _fill(cd[4], с.субъект_рф or '—',           align=C, pt=11)

    # ── Подпись: Главный судья ────────────────────────────────────────────────

    _para()

    _гс = next((с for с in судьи if с.должность == 'Главный судья'), None)
    _гс_фио = _гс.фамилия_имя_отчество if _гс else (сор.главный_судья or '')
    _гс_аббр = _abbr(_гс_фио)

    w_l = TW * 45 // 100
    w_r = TW - w_l
    sig = _make_tbl(2)
    _set_grid(sig, [w_l, w_r])
    rs = sig.add_row(); _row_h(rs, 1.2)
    cs = list(rs.cells)
    _cw(cs[0], w_l); _cw(cs[1], w_r)
    _no_border(cs[0]); _no_border(cs[1])
    _valign(cs[0]); _valign(cs[1])
    _fill(cs[0], 'Главный судья', align=L, pt=14)
    _fill(cs[1], _гс_аббр,        align=L, pt=14)

    buf = BytesIO()
    _apply_no_break(doc)
    doc.save(buf); buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)
