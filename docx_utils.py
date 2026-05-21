# docx_utils.py — HTML → DOCX export via python-docx + htmldocx (no external binaries)
from io import BytesIO
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Cm
from htmldocx import HtmlToDocx
from flask import send_file

DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'


def _make_base_doc():
    """Return a Document pre-configured with Times New Roman and A4 landscape margins."""
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


def _normalize_tables(soup):
    """
    Expand colspan cells and strip rowspan attrs so htmldocx doesn't crash.
    Each <th>/<td> with colspan=N becomes N individual cells.
    rowspan attributes are simply removed (cells don't span rows in DOCX output).
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
    """Strip .no-print elements, normalise tables, convert HTML→docx."""
    soup = BeautifulSoup(html_string, 'html.parser')

    for el in soup.select('.no-print'):
        el.decompose()
    for el in soup.select('link[rel="stylesheet"], script'):
        el.decompose()

    _normalize_tables(soup)

    doc = _make_base_doc()
    HtmlToDocx().add_html_to_document(str(soup), doc)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)
