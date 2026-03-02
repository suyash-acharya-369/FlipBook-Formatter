import re
import io
import warnings
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from lxml import etree
import os
import tempfile

# Wrapped formatting logic for backend usage

ORANGE = RGBColor(227, 108, 10)
BLUE   = RGBColor(18, 67, 149)
BLACK  = RGBColor(29, 29, 27)
WHITE  = RGBColor(255, 255, 255)

h2_titles = {
    "objectives", "introduction", "programming", "steps in program development",
    "problem identification", "task analysis", "data analysis and input design",
    "output identification and specifications", "designing the solution",
    "decision tables", "algorithm", "data validation", "flowcharts",
    "coding the program", "debugging", "testing", "summary", "check your progress"
}

h3_titles = {
    "for source documents", "for files", "for processing", "for output",
    "printer and page layouts", "screen and page layouts",
    "efficient algorithms.", "approximate algorithm", "explanation",
    "refined algorithm", "types of errors", "syntax errors", "desk checking",
    "levels of testing", "difference between testing and debugging",
    "a. descriptive questions", "b. multiple choice questions (mcqs)",
    "answer table"
}

h4_titles = {
    "example 1.1", "example 1.2", "example 1.3: servicing a car",
    "example 1.4: mail ordering", "example 1.5",
    "advantages of using decision tables", "disadvantages",
    "where to use decision tables and where flowcharts?",
    "(1) input validation", "valid code", "valid character",
    "valid field size, sign, and composition", "valid transaction",
    "valid combinations of field", "missing date test", "check digit",
    "sequence test", "limit of reasonableness test"
}

CONTENT_WIDTH_CM = 21.0 - 2 * 2.54
CONTENT_WIDTH_EMU = int(CONTENT_WIDTH_CM * 360000)

MAX_WIDTH  = int(16 * 360000)
MAX_HEIGHT = int(20 * 360000)
MIN_SIZE   = int(1.0 * 360000)

MC_ALT = '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent'
MC_CHOICE = '{http://schemas.openxmlformats.org/markup-compatibility/2006}Choice'
WPS_WSP = '{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}wsp'
WPG_WGP = '{http://schemas.microsoft.com/office/word/2010/wordprocessingGroup}wgp'
WPC_WPC = '{http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas}wpc'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def set_font_xml(run, font_name):
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)

def set_spacing(para, before=0, after=0, line_mult=1.05, auto_before=False, auto_after=False):
    pPr = para._p.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        pPr.append(sp)
    for a in ['w:before','w:after','w:line','w:lineRule','w:beforeAutospacing','w:afterAutospacing']:
        k = qn(a)
        if k in sp.attrib:
            del sp.attrib[k]
    
    if auto_before:
        sp.set(qn('w:beforeAutospacing'), '1')
    else:
        sp.set(qn('w:before'), str(int(before * 20)))
        
    if auto_after:
        sp.set(qn('w:afterAutospacing'), '1')
    else:
        sp.set(qn('w:after'), str(int(after * 20)))
        
    if line_mult is not None:
        sp.set(qn('w:line'), str(int(line_mult * 240)))
        sp.set(qn('w:lineRule'), 'auto')

def set_shading(para, r, g, b):
    pPr = para._p.get_or_add_pPr()
    for old in pPr.findall(qn('w:shd')): pPr.remove(old)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), f'{r:02X}{g:02X}{b:02X}')
    pPr.append(shd)

def add_run(para, text, font='Times New Roman', size=16, bold=False, italic=False, color=None):
    run = para.add_run(text)
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color: run.font.color.rgb = color
    set_font_xml(run, font)
    return run

def is_shape_content(elem):
    if elem.findall('.//' + WPS_WSP): return True
    if elem.findall('.//' + WPG_WGP): return True
    if elem.findall('.//' + WPC_WPC): return True
    return False

def extract_safe_image(src_doc, inline_elem):
    if is_shape_content(inline_elem): return None
    extent = inline_elem.find(qn('wp:extent'))
    if extent is None: return None
    cx, cy = int(extent.get('cx', '0')), int(extent.get('cy', '0'))
    if cx > MAX_WIDTH and cy > MAX_HEIGHT: return None
    if cx < MIN_SIZE and cy < MIN_SIZE: return None
    blips = inline_elem.findall('.//{%s}blip' % A_NS)
    if not blips: return None
    r_embed = blips[0].get('{%s}embed' % R_NS)
    if not r_embed: return None
    try:
        part = src_doc.part.related_parts[r_embed]
        data = part.blob
        if part.content_type and not part.content_type.startswith('image/'): return None
        return (data, cx, cy)
    except: return None



def set_table_borders(table):
    tblPr = table._tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)
    old = tblPr.find(qn('w:tblBorders'))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement('w:tblBorders')
    for n in ['top','left','bottom','right','insideH','insideV']:
        b = OxmlElement(f'w:{n}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), '000000')
        borders.append(b)
    tblPr.append(borders)

def flatten_numbering(input_path, output_path):
    """Uses MS Word via COM to permanently convert list numbers/bullets into plain text"""
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        try:
            # Must use absolute paths for COM
            abs_in = os.path.abspath(input_path)
            abs_out = os.path.abspath(output_path)
            doc = word.Documents.Open(abs_in)
            doc.ConvertNumbersToText()
            doc.SaveAs2(abs_out, 16) # wdFormatDocumentDefault
            doc.Close()
        finally:
            word.Quit()
            pythoncom.CoUninitialize()
        return True
    except Exception as e:
        print(f"Failed to flatten numbering: {e}")
        return False

def format_document(input_path, output_path):
    print("PHASE 1: Extracting content from original raw file...")
    
    # Pre-process doc to flatten numbering and bullets to literal text
    flat_path = input_path + "_flat.docx"
    has_flat = flatten_numbering(input_path, flat_path)
    target_parse_path = flat_path if has_flat else input_path
    
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        raw = Document(target_parse_path)

    items = []
    toc_skip = True

    for child in raw.element.body:
        tag = child.tag
        if tag == qn('w:p'):
            texts = [t.text for t in child.findall('.//' + qn('w:t')) if t.text]
            raw_full = ''.join(texts).replace('\u2028', ' ').replace('\t', ' ').replace('\n', ' ')
            full_text = re.sub(r' +', ' ', raw_full).strip()
            
            safe_images = []
            for d in child.findall('.//' + qn('w:drawing')):
                for inl in d.findall(qn('wp:inline')):
                    res = extract_safe_image(raw, inl)
                    if res: safe_images.append(res)
            for ac in child.findall('.//' + MC_ALT):
                choice = ac.find(MC_CHOICE)
                if choice is not None:
                    for d in choice.findall('.//' + qn('w:drawing')):
                        for inl in d.findall(qn('wp:inline')):
                            res = extract_safe_image(raw, inl)
                            if res: safe_images.append(res)

            if full_text in ["Programming Concepts and Technique, Fox Pro", "Task Analysis-Decision Tables"] and not (full_text.isupper() and "TASK ANALYSIS-DECISION TABLES" in full_text):
                continue
                
            if not full_text and not safe_images: continue
            
            if toc_skip:
                if full_text == "FOXPRO":
                    items.append({'type': 'h1', 'text': 'CHAPTER-1\nFOXPRO', 'runs': []})
                    continue
                if full_text.lower() == "objectives":
                    toc_skip = False
                else:
                    if full_text.isupper() and not safe_images and full_text != "FOXPRO":
                        continue
                    
            runs = []
            for r in child.findall(qn('w:r')):
                rtext = ''.join(t.text for t in r.findall(qn('w:t')) if t.text)
                if not rtext: continue
                b, it = False, False
                rPr = r.find(qn('w:rPr'))
                if rPr is not None:
                    be = rPr.find(qn('w:b'))
                    if be is not None and be.get(qn('w:val')) in (None, 'true', '1', ''): b = True
                    ie = rPr.find(qn('w:i'))
                    if ie is not None and ie.get(qn('w:val')) in (None, 'true', '1', ''): it = True
                runs.append({'text': rtext, 'bold': b, 'italic': it})

            # --- Clean runs ---
            for rc in runs:
                rc['text'] = re.sub(r' +', ' ', rc['text'].replace('\u2028', ' ').replace('\t', ' ').replace('\n', ' '))
            for i in range(len(runs) - 1):
                if runs[i]['text'].endswith(' ') and runs[i+1]['text'].startswith(' '):
                    runs[i+1]['text'] = runs[i+1]['text'][1:]
            if runs:
                runs[0]['text'] = runs[0]['text'].lstrip()
                runs[-1]['text'] = runs[-1]['text'].rstrip()
            runs = [rc for rc in runs if rc['text']]
            # ------------------

            ctype = 'body'
            lower_text = full_text.lower()
            
            # Count leading numbers to determine heading depth: e.g. "1.1.2 " -> depth 3
            # Match formats like: "1.", "1.1", "1.1.1", "1.1.1." followed by space
            num_match = re.match(r'^((?:\d+\.)+)(?:\d+)?(?=\s|\b)', full_text)
            depth = 0
            if num_match:
                # e.g "1.1." count of dots gives depth
                depth = num_match.group(1).count('.')
                
            is_bold_para = any(r['bold'] for r in runs) and len(full_text) < 150
            
            if lower_text.startswith("q") and len(lower_text)>1 and lower_text[1].isdigit():
                is_bold_para = False
                
            if is_bold_para:
                if depth == 1: ctype = 'h2'
                elif depth == 2: ctype = 'h3'
                elif depth >= 3: ctype = 'h4'
                # Fallback to text matching if no numbers
                elif lower_text in h2_titles: ctype = 'h2'
                elif lower_text in h3_titles: ctype = 'h3'
                elif lower_text in h4_titles: ctype = 'h4'
                elif re.match(r'^fig(ure)?[\s:\.\-]', lower_text, re.IGNORECASE): ctype = 'fig'
                else: ctype = 'h3'  # fallback
            
            if safe_images and not full_text: ctype = 'img'
            if ctype == 'body' and re.match(r'^fig(ure)?[\s:\.\-]', lower_text, re.IGNORECASE): ctype = 'fig'

            items.append({'type': ctype, 'text': full_text, 'runs': runs, 'images': safe_images})
            
        elif tag == qn('w:tbl') and not toc_skip:
            rows = []
            for tr in child.findall(qn('w:tr')):
                cells = []
                for tc in tr.findall(qn('w:tc')):
                    ct = ''.join(t.text for t in tc.findall('.//' + qn('w:t')) if t.text).strip()
                    cells.append(ct)
                rows.append(cells)
            if rows: items.append({'type': 'table', 'rows': rows})

    print("PHASE 2: Creating new document...")
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(2.54), Cm(2.54)
    sec.left_margin, sec.right_margin = Cm(1.91), Cm(1.91)

    for hf in [sec.header, sec.footer]:
        for child in list(hf._element): hf._element.remove(child)
        etree.SubElement(hf._element, qn('w:p'))

    print("PHASE 3: Rebuilding...")
    for item in items:
        ct = item['type']
        
        if ct == 'h1':
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_shading(p, 227, 108, 10)
            
            parts = item['text'].split('\n')
            for i, part_text in enumerate(parts):
                if i > 0:
                    p.add_run().add_break()
                add_run(p, part_text, 'Montserrat', 35, bold=True, color=WHITE)
            set_spacing(p, auto_before=True, auto_after=True, line_mult=None)
            
        elif ct in ['h2', 'h3', 'h4']:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            color = ORANGE if ct=='h2' else (BLUE if ct=='h3' else BLACK)
            size = 30 if ct=='h2' else (20 if ct=='h3' else 18)
            
            fmt_text = item['text'].title()
            
            add_run(p, fmt_text, 'Times New Roman', size, bold=True, color=color)
            if ct in ['h2', 'h3']:
                set_spacing(p, auto_before=True, auto_after=True, line_mult=None)
            else:
                set_spacing(p, before=12, after=12, line_mult=1.05)
            
        elif ct == 'fig':
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(p, item['text'], 'Times New Roman', 14, bold=True, italic=True, color=BLACK)
            set_spacing(p, before=0, after=10, line_mult=1.05)
            
        elif ct == 'img':
            for data, w, h in item.get('images', []):
                try:
                    if w > 0 and w != CONTENT_WIDTH_EMU:
                        ratio = CONTENT_WIDTH_EMU / w
                        w, h = CONTENT_WIDTH_EMU, int(h * ratio)
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.add_run().add_picture(io.BytesIO(data), width=Emu(w))
                    set_spacing(p, before=0, after=10, line_mult=1.05)
                except: pass
                
        elif ct == 'body':
            for data, w, h in item.get('images', []):
                try:
                    if w > 0 and w != CONTENT_WIDTH_EMU:
                        ratio = CONTENT_WIDTH_EMU / w
                        w, h = CONTENT_WIDTH_EMU, int(h * ratio)
                    pi = doc.add_paragraph()
                    pi.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    pi.add_run().add_picture(io.BytesIO(data), width=Emu(w))
                    set_spacing(pi, before=0, after=10, line_mult=1.05)
                except: pass
            if item['text']:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                if item['runs']:
                    for r in item['runs']:
                        add_run(p, r['text'], 'Times New Roman', 16, bold=r['bold'], italic=r['italic'], color=BLACK)
                else:
                    add_run(p, item['text'], 'Times New Roman', 16, color=BLACK)
                set_spacing(p, before=0, after=6, line_mult=1.05)
                
        elif ct == 'table':
            rows = item['rows']
            max_cols = max(len(r) for r in rows)
            tbl = doc.add_table(rows=len(rows), cols=max_cols)
            try:
                tbl.style = 'Table Grid'
            except: pass
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            for ri, row_data in enumerate(rows):
                for ci, cell_text in enumerate(row_data):
                    if ci < max_cols:
                        cell = tbl.cell(ri, ci)
                        cell.paragraphs[0].text = ''
                        p = cell.paragraphs[0]
                        add_run(p, cell_text, 'Times New Roman', 12, bold=(ri==0), color=BLACK)
                        set_spacing(p, before=0, after=0, line_mult=1.0)

    for t in doc.tables:
        set_table_borders(t)

    print(f"Saving to {output_path}")
    doc.save(output_path)
    
    # Cleanup temp flattened file
    if has_flat:
        try: os.remove(flat_path)
        except: pass
        
    print("DONE!")

