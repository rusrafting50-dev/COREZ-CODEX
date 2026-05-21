# docx_utils.py — utility for exporting rendered HTML to .docx
from io import BytesIO
from html2docx import html2docx
from bs4 import BeautifulSoup


def html_to_docx_response(html_string, filename):
    """Strip no-print elements, convert HTML to docx, return (BytesIO, filename)."""
    soup = BeautifulSoup(html_string, 'html.parser')
    for el in soup.select('.no-print'):
        el.decompose()
    clean_html = str(soup)
    buf = html2docx(clean_html, title=filename.replace('.docx', ''))
    buf.seek(0)
    return buf, filename
