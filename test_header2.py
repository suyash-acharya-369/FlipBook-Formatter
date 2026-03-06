from docx import Document
from docx.oxml.ns import qn
import sys

path = r"d:\DESKTOP\FORMATTING\Formatted Sample word file- DCA FUNDAMENTALS OF COMPUTER AND INFORMATION TECHNOLOGY - updated (03) (2).docx"
doc = Document(path)

for i, section in enumerate(doc.sections):
    print(f"\n--- Section {i} ---")
    headers = [section.header, section.first_page_header, section.even_page_header]
    for h in headers:
        if h is None: continue
        for p in h.paragraphs:
            for t_element in p._element.xpath('.//w:t'):
                print(f"Header Text Node: {repr(t_element.text)}")
