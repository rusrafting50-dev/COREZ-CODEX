# docx_utils.py — HTML → DOCX export via python-docx + htmldocx (no external binaries)
import re
from io import BytesIO
from bs4 import BeautifulSoup, NavigableString
from docx import Document
from docx.shared import Pt, Cm
from htmldocx import HtmlToDocx
from flask import send_file

DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

# Tags that create their own block — children of these should NOT be wrapped in <p>
_BLOCK = frozenset({
    'div','p','h1','h2','h3','h4','h5','h6',
    'ul','ol','li','table','thead','tbody','tfoot','tr','th','td',
    'blockquote','section','article','header','footer','nav','figure',
})


def _make_base_doc():
    doc = Document()
    section = doc.sections[0]
    section.page_width    = Cm(29.7)
    section.page_height   = Cm(21.0)
    section.left_margin   = Cm(1.5)
    section.right_margin  = Cm(1.0)
    section.top_margin    = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal'].font.size = Pt(10)
    return doc


def _get_style(tag, prop):
    """Return a CSS property value from a tag's style attribute, or ''."""
    style = tag.get('style', '')
    m = re.search(r'(?:^|;)\s*' + re.escape(prop) + r'\s*:\s*([^;]+)', style)
    return m.group(1).strip() if m else ''


def _apply_bold(tag, soup):
    """If tag has font-weight:bold/700 in style, wrap its contents in <b>."""
    fw = _get_style(tag, 'font-weight')
    if fw in ('bold', '700', '800', '900'):
        # Wrap all children in <b>
        b = soup.new_tag('b')
        for child in list(tag.children):
            child.extract()
            b.append(child)
        tag.append(b)


def _pre_line_to_br(tag, soup):
    """Convert white-space:pre-line text nodes: replace \\n with <br>."""
    if _get_style(tag, 'white-space') == 'pre-line':
        for text_node in list(tag.find_all(string=True)):
            if '\n' not in str(text_node):
                continue
            parts = str(text_node).split('\n')
            # Build replacement fragment: text, <br>, text, <br>, ...
            fragment = []
            for i, part in enumerate(parts):
                fragment.append(NavigableString(part))
                if i < len(parts) - 1:
                    fragment.append(soup.new_tag('br'))
            # Replace the original text node with the fragment
            parent = text_node.parent
            pos = list(parent.children).index(text_node)
            text_node.extract()
            for j, node in enumerate(fragment):
                parent.insert(pos + j, node)


def _divs_to_paragraphs(soup):
    """
    Convert <div> elements to <p> so htmldocx treats each as a paragraph.
    Divs that contain block-level children are unwrapped (structural wrappers).
    Leaf divs (inline/text content only) become <p>.
    Also applies bold and pre-line handling.
    """
    # Iterate bottom-up so inner divs are processed before outer ones
    for div in reversed(soup.find_all('div')):
        _apply_bold(div, soup)
        _pre_line_to_br(div, soup)
        has_block = any(
            getattr(c, 'name', None) in _BLOCK
            for c in div.children
        )
        if has_block:
            div.unwrap()
        else:
            div.name = 'p'

    # Also apply bold to existing <p> tags
    for p in soup.find_all('p'):
        _apply_bold(p, soup)

    return soup


def _normalize_tables(soup):
    """
    Expand colspan cells and strip rowspan so htmldocx doesn't crash.
    Each cell with colspan=N becomes N individual cells.
    """
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


def send_docx(html_string, filename):
    """Strip nav, normalise HTML structure, convert to DOCX via htmldocx."""
    soup = BeautifulSoup(html_string, 'html.parser')

    # Remove non-print and CDN elements
    for el in soup.select('.no-print, link[rel="stylesheet"], script'):
        el.decompose()

    _divs_to_paragraphs(soup)
    _normalize_tables(soup)

    doc = _make_base_doc()
    HtmlToDocx().add_html_to_document(str(soup), doc)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)
