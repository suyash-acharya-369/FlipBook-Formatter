from docx import Document
from docx.oxml.ns import qn
import sys

path = r"d:\DESKTOP\FORMATTING\Formatted Sample word file- DCA FUNDAMENTALS OF COMPUTER AND INFORMATION TECHNOLOGY - updated (03) (2).docx"
doc = Document(path)

TARGET_STRING = "FUNDAMENTALS OF COMPUTER AND INFORMATION TECHNOLOGY"
NEW_STRING = "AUTOMOBILE ENGINEERING"

for section in doc.sections:
    for header in [section.header, section.first_page_header, section.even_page_header]:
        if header is None: continue
        for p in header.paragraphs:
            for t_elem in p._element.xpath('.//w:t'):
                if t_elem.text and TARGET_STRING in t_elem.text:
                    t_elem.text = t_elem.text.replace(TARGET_STRING, NEW_STRING.upper())

doc.save("test_header_replace.docx")
print("Saved test_header_replace.docx")
