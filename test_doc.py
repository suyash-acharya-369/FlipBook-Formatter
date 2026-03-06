import os
from docx import Document
from docx.oxml.ns import qn

doc_path = r"d:\DESKTOP\FORMATTING\13. Module-Automobile Engineering.docx"
doc = Document(doc_path)

with open('dump2.txt', 'w', encoding='utf-8') as f:
    body = doc.element.body
    for idx, child in enumerate(body):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        
        if tag == 'p':
            text = ''.join(t.text for t in child.findall('.//' + qn('w:t')) if t.text).strip()
            f.write(f"[{idx}] PARA: {repr(text[:120])}\n")
        elif tag == 'tbl':
            # Get first cell text
            first_cell = ''
            tcs = child.findall('.//' + qn('w:tc'))
            if tcs:
                first_cell = ''.join(t.text for t in tcs[0].findall('.//' + qn('w:t')) if t.text).strip()
            row_count = len(child.findall(qn('w:tr')))
            f.write(f"[{idx}] TABLE ({row_count} rows): first_cell={repr(first_cell[:80])}\n")
        elif tag == 'sectPr':
            f.write(f"[{idx}] SECTION_PROPS\n")
        else:
            f.write(f"[{idx}] {tag}\n")
    
    f.write(f"\nTotal children: {idx+1}\n")

print("Done. See dump2.txt")
