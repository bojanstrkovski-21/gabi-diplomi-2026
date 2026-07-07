"""
Create A3 landscape diploma by directly manipulating docx XML.
Source: "училишна стручна матура.docx" (DrawingML + VML fallback).
Uses pre-verified absolute positions (from COM analysis + XML text matching)
to avoid paragraph-relative offset ambiguities.

A3 page 1 LEFT  = source page 1
A3 page 1 RIGHT = source page 2  (shifted right by A4 width)
A3 page 2 LEFT  = source page 3
"""
import zipfile, shutil, re, sys, io
from copy import deepcopy
from lxml import etree

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SRC = r'f:\Gabi-diplomi\New folder\училишна стручна матура.docx'
OUT = r'f:\Gabi-diplomi\училишна_стручна_матура_А3.docx'

# Namespaces
W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
V  = 'urn:schemas-microsoft-com:vml'

# Paragraph indices (0-based) containing shapes, by source page
# (discovered by parsing document.xml and cross-checking text against
#  a Word COM shape dump: page, RelativeHorizontalPosition, Left/Top)
PAGE1_PARAS = {9, 12, 25}
PAGE2_PARAS = {37, 38, 44, 50, 52, 57, 64, 66, 72, 74, 82}
PAGE3_PARAS = {126, 128}

# Pre-computed final positions on A3 (from COM analysis of source shapes,
# matched to document order via unique textbox text).
# Indexed 1-26, matching the document order of wp:anchor elements collected
# from PAGE1_PARAS (11 shapes), PAGE2_PARAS (13 shapes), PAGE3_PARAS (2 shapes).
# (left_pt, top_pt, a3_page)
ANCHOR_A3_POS = [
    None,               # 0 - placeholder (list is 1-based)
    ( 62.85,  33.7, 1),  # 1  -> school name
    (408.35,  39.1, 1),  # 2  -> "1"
    (461.65,  39.6, 1),  # 3  -> "26"
    ( 90.95,  33.6, 1),  # 4  -> city (Битола, header)
    (127.80, 542.3, 1),  # 5  -> nationality/citizenship
    (128.40, 518.7, 1),  # 6  -> country (Република Северна Македонија)
    (413.80, 456.0, 1),  # 7  -> DOB
    (131.60, 456.0, 1),  # 8  -> father's name (Феим)
    (389.00, 490.1, 1),  # 9  -> born city
    (100.80, 486.2, 1),  # 10 -> residence city
    ( 79.45,  74.2, 1),  # 11 -> student full name
    (963.45, 190.9, 1),  # 12 -> grade (Доволен 2) - Macedonian language
    (647.85, 193.1, 1),  # 13 -> Macedonian language course name
    (965.75, 283.3, 1),  # 14 -> grade (Доволен 2) - mining course
    (651.65, 283.3, 1),  # 15 -> Добивање на минерални суровини
    (649.45, 379.5, 1),  # 16 -> (empty field)
    (647.25, 386.1, 1),  # 17 -> Пресметка на рудни резерви
    (980.55, 461.5, 1),  # 18 -> grade (Мн.добар 4)
    (725.35, 561.0, 1),  # 19 -> Геологија, рударство и металургија
    (741.85, 582.5, 1),  # 20 -> геолошко, рударски техничар
    (827.15, 665.5, 1),  # 21 -> number 08-53/1
    (871.15, 692.0, 1),  # 22 -> Bitola date
    (640.70,  47.4, 1),  # 23 -> Љупчо Поповски
    (968.40,  46.8, 1),  # 24 -> Ристо Грујовски
    (100.25,  33.7, 2),  # 25 -> 113-2800 / date -> A3 page 2
    (112.35,  36.4, 2),  # 26 -> Министерство за образование и наука -> A3 page 2
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
print(f'Runs -> page1:{len(p1_runs)}  page2:{len(p2_runs)}  page3:{len(p3_runs)}  total:{len(all_runs)}')

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

# Para 2: page break -> creates A3 page 2
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
# A3 landscape: 420 mm x 297 mm = 23811 x 16838 twips
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
