# docx_utils.py — HTML → DOCX export via python-docx + htmldocx (no external binaries)
from io import BytesIO
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from htmldocx import HtmlToDocx
from flask import send_file

DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'


def _make_base_doc():
    """Return a Document pre-configured with Times New Roman and A4 landscape margins."""
    doc = Document()
    section = doc.sections[0]
    section.page_width   = Cm(29.7)
    section.page_height  = Cm(21.0)
    section.left_margin  = Cm(1.5)
    section.right_margin = Cm(1.0)
    section.top_margin   = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(10)
    return doc


def send_docx(html_string, filename):
    """Strip .no-print elements, convert HTML→docx, return Flask send_file response."""
    soup = BeautifulSoup(html_string, 'html.parser')

    for el in soup.select('.no-print'):
        el.decompose()
    for el in soup.select('link[rel="stylesheet"], script'):
        el.decompose()

    clean_html = str(soup)

    doc = _make_base_doc()
    parser = HtmlToDocx()
    parser.add_html_to_document(clean_html, doc)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)
