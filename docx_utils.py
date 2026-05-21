# docx_utils.py — utility for exporting rendered HTML to .docx
from io import BytesIO
from html2docx import html2docx
from bs4 import BeautifulSoup
from flask import send_file

DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'


def send_docx(html_string, filename):
    """Strip .no-print elements, convert HTML→docx, return Flask send_file response."""
    soup = BeautifulSoup(html_string, 'html.parser')
    for el in soup.select('.no-print'):
        el.decompose()
    clean_html = str(soup)
    buf = html2docx(clean_html, title=filename.replace('.docx', ''))
    buf.seek(0)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)
