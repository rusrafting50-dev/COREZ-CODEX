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

def _post_process_date_place_tables(doc, border_size=6):
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

        # Label row — first and last TC: top align, center
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

def _merge_section_rows(doc, section_row_texts):
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
            # Bold the merged cell text
            for w_r in tcs[0].iter(qn('w:r')):
                rPr = w_r.find(qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr')
                    w_r.insert(0, rPr)
                if rPr.find(qn('w:b')) is None:
                    rPr.append(OxmlElement('w:b'))
                if rPr.find(qn('w:bCs')) is None:
                    rPr.append(OxmlElement('w:bCs'))


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
        for tr in trs:
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

def send_docx(html_string, filename, portrait=None, date_place_border_size=6):
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

    # 4. HTML → DOCX
    doc = _make_base_doc(portrait=portrait)
    HtmlToDocx().add_html_to_document(str(soup), doc)

    # 5. Apply column widths (page_width_cm already reflects orientation)
    _apply_table_widths(doc, table_widths, page_width_cm)

    # 5b. Restore vertical text lost by HtmlToDocx (writing-mode:vertical-*)
    _apply_vertical_text(doc, soup)

    # 5c. Re-merge section-row cells that HtmlToDocx expands into N empty cells
    _merge_section_rows(doc, section_row_texts)

    # 5d. Center all cells + No Spacing + Times New Roman (overridden where needed below)
    _center_all_cells(doc)

    # 5e. Left-align columns that should not be centered
    _left_align_data_columns(doc, {
        'Команда',
        'Фамилия, имя, отчество участников',
        'Территория',
    })

    # 5f. noWrap on specific header cells that must stay on one line
    _nowrap_cells_by_text(doc, {'Фамилия, имя, отчество участников'})

    # 6. Post-process: date/place table — bottom-align value cells, border, noWrap
    _post_process_date_place_tables(doc, border_size=date_place_border_size)

    # 6b. Signature tables: left-align + 2× row height
    _post_process_sig_tables(doc)

    # 7. Insert page break after each signature table
    _add_page_breaks_after_sig_tables(doc)

    # 8. Overwrite borders on data tables with uniform single 0.375pt lines
    _set_all_table_borders(doc, data_table_indices)

    # 9. Remove empty paragraphs HtmlToDocx leaves at the top of the document
    _remove_leading_empty_paragraphs(doc)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)
