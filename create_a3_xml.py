"""
Create A3 landscape diploma by directly manipulating docx XML.
Source: user's .docx (DrawingML + VML fallback).
Uses pre-verified absolute positions (from COM analysis) to avoid
paragraph-relative offset ambiguities.

A3 page 1 LEFT  = source page 1
A3 page 1 RIGHT = source page 2  (shifted right by A4 width)
A3 page 2 LEFT  = source page 3
"""
import zipfile, shutil, re, sys, io
from copy import deepcopy
from lxml import etree

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SRC = r'c:\Users\User\Downloads\Gabi-diplomi\New folder\диплома државна матура (1).docx'
OUT = r'c:\Users\User\Downloads\Gabi-diplomi\диплома_А3_v2.docx'

# Namespaces
W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
V  = 'urn:schemas-microsoft-com:vml'

# Paragraph indices (0-based) containing shapes, by source page
# (discovered by examining the user's docx XML)
PAGE1_PARAS = {10, 13, 29, 32, 34}
PAGE2_PARAS = {44, 48, 50, 52, 54, 56, 60, 65, 70, 72, 74, 78, 80, 88}
PAGE3_PARAS = {91}

# Pre-computed final positions on A3 (from COM analysis of source shapes).
# Indexed 1-30, matching the document order of wp:anchor elements collected
# from PAGE1_PARAS (11 shapes), PAGE2_PARAS (18 shapes), PAGE3_PARAS (1 shape).
# (left_pt, top_pt, a3_page)
ANCHOR_A3_POS = [
    None,               # 0 – placeholder (list is 1-based)
    ( 68.4,  33.1, 1),  # 1  → src-page-1 shape (school name)
    (401.7,  30.8, 1),  # 2  → src-page-1 shape (number/date)
    (106.9,  27.5, 1),  # 3  → src-page-1 shape (city)
    ( 62.3,  39.9, 1),  # 4  → src-page-1 shape (student name)
    (416.0, 479.6, 1),  # 5  → src-page-1 (DOB)
    (130.6, 481.9, 1),  # 6  → src-page-1 (father)
    (388.5, 512.7, 1),  # 7  → src-page-1 (born city)
    (104.2, 511.0, 1),  # 8  → src-page-1 (residence)
    (122.8, 534.7, 1),  # 9  → src-page-1 (country)
    (124.5, 558.3, 1),  # 10 → src-page-1 (nationality)
    (281.8, 543.5, 1),  # 11 → src-page-1 (empty field)
    (882.1,  30.9, 1),  # 12 → src-page-2 group (border/line)
    (643.4, 171.6, 1),  # 13 → src-page-2 (Macedonian language)
    (855.2, 233.2, 1),  # 14 → src-page-2 (grade Одличен 5)
    (1036.7,233.2, 1),  # 15 → src-page-2 (score 93.55)
    (640.1, 231.6, 1),  # 16 → src-page-2 (English language)
    (651.1,  36.7, 1),  # 17 → src-page-2 group (border)
    (650.0,  35.5, 1),  # 18 → src-page-2 group (border)
    (653.3,  33.2, 1),  # 19 → src-page-2 group (border)
    (650.5,  62.3, 1),  # 20 → src-page-2 group (border)
    (657.2, 393.9, 1),  # 21 → src-page-2 (school Lady Diana)
    (1011.9,463.7, 1),  # 22 → src-page-2 (grade Одличен)
    (743.0, 533.0, 1),  # 23 → src-page-2 (гимназиско)
    (722.6, 557.2, 1),  # 24 → src-page-2 (Лични услуги)
    (738.5, 582.5, 1),  # 25 → src-page-2 (техничар...)
    (829.3, 638.6, 1),  # 26 → src-page-2 (number 08-51)
    (822.7, 665.5, 1),  # 27 → src-page-2 (Bitola date)
    (945.9,  32.0, 1),  # 28 → src-page-2 (Ристо Грујовски)
    (634.6,  33.2, 1),  # 29 → src-page-2 (Весна Христовска)
    ( 80.5, 597.3, 2),  # 30 → src-page-3 (stamp/verification) → A3 page 2
]

# -------------------------------------------------------------------------
# Load source docx
# -------------------------------------------------------------------------
with zipfile.ZipFile(SRC) as z:
    doc_xml_bytes = z.read('word/document.xml')

tree = etree.fromstring(doc_xml_bytes)
body = tree.find(f'{{{W}}}body')
paragraphs = body.findall(f'{{{W}}}p')
print(f'Total paragraphs: {len(paragraphs)}')

# -------------------------------------------------------------------------
# Collect drawing runs in document order
# -------------------------------------------------------------------------
def get_anchor_runs(para):
    return [r for r in para.findall(f'{{{W}}}r')
            if r.find(f'.//{{{WP}}}anchor') is not None]

p1_runs, p2_runs, p3_runs = [], [], []
for idx, para in enumerate(paragraphs):
    if idx in PAGE1_PARAS:
        p1_runs.extend(get_anchor_runs(para))
    elif idx in PAGE2_PARAS:
        p2_runs.extend(get_anchor_runs(para))
    elif idx in PAGE3_PARAS:
        p3_runs.extend(get_anchor_runs(para))

all_runs = p1_runs + p2_runs + p3_runs
print(f'Runs → page1:{len(p1_runs)}  page2:{len(p2_runs)}  page3:{len(p3_runs)}  total:{len(all_runs)}')

# -------------------------------------------------------------------------
# Apply positions to each run
# -------------------------------------------------------------------------
def set_anchor_pos(run, left_pt, top_pt):
    """Set DrawingML anchor to page-relative (left_pt, top_pt) and sync VML."""
    # --- DrawingML ---
    anchor = run.find(f'.//{{{WP}}}anchor')
    if anchor is None:
        return

    posH = anchor.find(f'{{{WP}}}positionH')
    posV = anchor.find(f'{{{WP}}}positionV')
    if posH is None or posV is None:
        return

    posH.set('relativeFrom', 'page')
    h_off = posH.find(f'{{{WP}}}posOffset')
    if h_off is None:
        h_off = etree.SubElement(posH, f'{{{WP}}}posOffset')
    h_off.text = str(int(round(left_pt * 12700)))

    posV.set('relativeFrom', 'page')
    v_off = posV.find(f'{{{WP}}}posOffset')
    if v_off is None:
        v_off = etree.SubElement(posV, f'{{{WP}}}posOffset')
    v_off.text = str(int(round(top_pt * 12700)))

    # --- VML fallback (v:shape or v:group, top-level only) ---
    for vml in run.findall(f'.//{{{V}}}shape') + run.findall(f'.//{{{V}}}group'):
        parent = vml.getparent()
        if parent is not None and parent.tag == f'{{{V}}}group':
            continue  # sub-shape, skip
        style = vml.get('style', '')
        if 'margin-left' not in style and 'margin-top' not in style:
            continue
        style = re.sub(r'margin-left:[^;\"]+', f'margin-left:{left_pt:.3f}pt', style)
        style = re.sub(r'margin-top:[^;\"]+',  f'margin-top:{top_pt:.3f}pt',  style)
        style = re.sub(r'mso-position-horizontal-relative:[^;\"]+',
                       'mso-position-horizontal-relative:page', style)
        style = re.sub(r'mso-position-vertical-relative:[^;\"]+',
                       'mso-position-vertical-relative:page',   style)
        if 'mso-position-horizontal-relative' not in style:
            style += ';mso-position-horizontal-relative:page'
        if 'mso-position-vertical-relative' not in style:
            style += ';mso-position-vertical-relative:page'
        vml.set('style', style)

print('\nApplying A3 positions:')
for i, run in enumerate(all_runs, start=1):
    left_pt, top_pt, a3_page = ANCHOR_A3_POS[i]
    run_copy = deepcopy(run)
    set_anchor_pos(run_copy, left_pt, top_pt)
    all_runs[i-1] = (run_copy, a3_page)
    print(f'  Anchor {i:2d}: left={left_pt:.1f}pt  top={top_pt:.1f}pt  A3-page={a3_page}')

a3p1_runs = [r for r, pg in all_runs if pg == 1]
a3p2_runs = [r for r, pg in all_runs if pg == 2]

# -------------------------------------------------------------------------
# Build new document body
# -------------------------------------------------------------------------
all_nsmap = {}
for elem in tree.iter():
    for k, v in elem.nsmap.items():
        all_nsmap[k] = v

new_body = etree.Element(f'{{{W}}}body', nsmap=all_nsmap)

# Para 1: all A3 page-1 shapes (source pages 1 + 2)
para1 = etree.SubElement(new_body, f'{{{W}}}p')
for r in a3p1_runs:
    para1.append(r)

# Para 2: page break → creates A3 page 2
para2 = etree.SubElement(new_body, f'{{{W}}}p')
pPr2 = etree.SubElement(para2, f'{{{W}}}pPr')
etree.SubElement(pPr2, f'{{{W}}}pageBreakBefore')

# Para 3: A3 page-2 shapes (source page 3)
para3 = etree.SubElement(new_body, f'{{{W}}}p')
for r in a3p2_runs:
    para3.append(r)

# Para 4: required trailing paragraph
etree.SubElement(new_body, f'{{{W}}}p')

# sectPr: A3 landscape, ~1 mm margins
# A3 landscape: 420 mm × 297 mm = 23811 × 16838 twips
sectPr = etree.SubElement(new_body, f'{{{W}}}sectPr')
pgSz = etree.SubElement(sectPr, f'{{{W}}}pgSz')
pgSz.set(f'{{{W}}}w', '23811')
pgSz.set(f'{{{W}}}h', '16838')
pgSz.set(f'{{{W}}}orient', 'landscape')
pgMar = etree.SubElement(sectPr, f'{{{W}}}pgMar')
for attr in ('top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'):
    pgMar.set(f'{{{W}}}{attr}', '57' if attr not in ('header', 'footer', 'gutter') else '0')

# -------------------------------------------------------------------------
# Wrap in w:document and serialise
# -------------------------------------------------------------------------
new_doc_elem = etree.Element(f'{{{W}}}document', nsmap=all_nsmap)
new_doc_elem.append(new_body)
new_xml = etree.tostring(new_doc_elem, xml_declaration=True,
                          encoding='UTF-8', standalone=True)

# -------------------------------------------------------------------------
# Write output docx
# -------------------------------------------------------------------------
with zipfile.ZipFile(SRC, 'r') as zin, \
     zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'word/document.xml':
            data = new_xml
        zout.writestr(item, data)

print(f'\nSaved: {OUT}')
