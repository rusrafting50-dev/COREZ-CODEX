# docx_utils.py — HTML → DOCX export via pandoc
import os
import tempfile
from bs4 import BeautifulSoup
from flask import send_file

DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
_REF_DOC  = os.path.join(os.path.dirname(__file__), 'pandoc_reference.docx')


def send_docx(html_string, filename):
    """Strip .no-print elements, convert HTML→docx via pandoc, return Flask send_file."""
    import pypandoc

    soup = BeautifulSoup(html_string, 'html.parser')

    # Remove nav/button elements not needed in print
    for el in soup.select('.no-print'):
        el.decompose()
    # Remove Bootstrap CDN link (not needed for conversion)
    for el in soup.select('link[rel="stylesheet"]'):
        el.decompose()
    for el in soup.select('script'):
        el.decompose()

    # Set explicit UTF-8 charset meta if missing
    head = soup.find('head')
    if head and not soup.find('meta', charset=True):
        meta = soup.new_tag('meta', charset='UTF-8')
        head.insert(0, meta)

    clean_html = str(soup)

    with tempfile.TemporaryDirectory() as td:
        out_path = os.path.join(td, filename)
        extra = ['--reference-doc', _REF_DOC] if os.path.exists(_REF_DOC) else []
        pypandoc.convert_text(
            clean_html,
            'docx',
            format='html',
            outputfile=out_path,
            extra_args=extra,
        )
        # Read into memory so temp dir can be cleaned up
        with open(out_path, 'rb') as f:
            data = f.read()

    from io import BytesIO
    buf = BytesIO(data)
    return send_file(buf, mimetype=DOCX_MIME, as_attachment=True, download_name=filename)
