"""
Create A3 landscape diploma from "drzavna strucna.docx" by directly
manipulating docx XML.

Unlike the previous two diplomas (which are built almost entirely from
absolute-positioned floating text boxes), this source document is
mostly normal flowing text + one table, split across 3 existing Word
sections (one per source page) plus 2 stray trailing paragraphs
("ДИПЛОМА" / "ЗА ДРЖАВНА СТРУЧНА МАТУРА") that begin a 4th, unwanted
page and are dropped per user instruction.

Layout (matches the two earlier diplomas):
  A3 Sheet 1 LEFT  = source page 1 (personal data)
  A3 Sheet 1 RIGHT = source page 2 (grades table + signatures)
  A3 Sheet 2 LEFT  = source page 3 (verification footer)
  A3 Sheet 2 RIGHT = empty

Technique:
  - Source page1+page2 are merged into ONE new section using a
    genuine two-column w:cols layout with UNEQUAL column widths/gutter
    computed from each source page's own original margins, so each
    column reproduces its original page's usable text width exactly
    (verified algebraically: results in a clean +595pt shift,
    identical to the "shift by one A4 width" rule used previously).
  - The one floating shape (signature "Директор" text box, anchored
    RelativeHorizontalPosition=Page) is shifted +595pt in both the
    DrawingML anchor and the VML fallback, matching the column2 shift.
    Its vertical position is RelativeVerticalPosition=Paragraph, so it
    tracks its anchor paragraph automatically and needs no Y change.
  - Page1's own top margin (1518 twips) is smaller than page2's
    (2574 twips); since a section can only have ONE top margin, the
    smaller one is used and the difference (1056 twips) is added as
    extra "space before" on page2's first paragraph, reproducing
    page2's original vertical start position without touching the
    shape's Y (which rides along via paragraph-relative anchoring).
  - Source page3 is placed unchanged on its own final section, just
    widened to A3 with a larger right margin so its usable text width
    (9129 twips) is preserved and it only occupies the left half.
"""
import zipfile, shutil, sys, io
from lxml import etree

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SRC = r'd:\My Backups\Bojan\gabi-diplomi-2026\New folder\drzavna strucna.docx'
OUT = r'd:\My Backups\Bojan\gabi-diplomi-2026\диплома_државна_стручна_матура_А3.docx'

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'

SHIFT_PT = 595.0
SHIFT_EMU = int(round(SHIFT_PT * 12700))  # 7,556,500

shutil.copyfile(SRC, OUT)

with zipfile.ZipFile(SRC) as zin:
    doc_xml = zin.read('word/document.xml')

root = etree.fromstring(doc_xml)
body = root.find(W + 'body')
children = list(body)

# --- sanity: confirm expected paragraphs by text before mutating ---
def ptext(el):
    return ''.join(t.text or '' for t in el.iter(W + 't'))

assert ptext(children[12]).strip() == 'држава државјанство', ptext(children[12])
assert ptext(children[13]).startswith('На државната стручна матура'), ptext(children[13])
assert 'Директор' in ptext(children[26]), ptext(children[26])
assert ptext(children[27]).strip() == 'М. П.', ptext(children[27])
assert ptext(children[28]).startswith('Училиштето е верифицирано'), ptext(children[28])
assert ptext(children[34]).strip() == 'ДИПЛОМА', ptext(children[34])
assert ptext(children[35]).strip() == 'ЗА ДРЖАВНА СТРУЧНА МАТУРА', ptext(children[35])
assert etree.QName(children[36]).localname == 'sectPr'

# 1. Remove the section break at the end of page 1 (para 12) so page1
#    and page2 flow into a single new merged section.
p12_ppr = children[12].find(W + 'pPr')
p12_sectpr = p12_ppr.find(W + 'sectPr')
p12_ppr.remove(p12_sectpr)

# 2. Para 13 (first paragraph of page 2): add a column break so its
#    content starts in column 2, and add extra top spacing (1056
#    twips = 2574-1518) to reproduce page2's original top margin.
p13 = children[13]
p13_ppr = p13.find(W + 'pPr')
spacing = p13_ppr.find(W + 'spacing')
before = int(spacing.get(W + 'before'))
spacing.set(W + 'before', str(before + 1056))

br_run = etree.SubElement(p13, W + 'r')
br = etree.SubElement(br_run, W + 'br')
br.set(W + 'type', 'column')
# move the new run to be the first child (after pPr)
p13.remove(br_run)
p13.insert(list(p13).index(p13_ppr) + 1, br_run)

# 3. Shape "Директор" (para 26): shift +595pt in DrawingML anchor and
#    VML fallback (X only; Y is paragraph-relative and needs no change).
p26 = children[26]
posH_offset = p26.find('.//' + WP + 'positionH/' + WP + 'posOffset')
old_emu = int(posH_offset.text)
new_emu = old_emu + SHIFT_EMU
posH_offset.text = str(new_emu)

V = 'urn:schemas-microsoft-com:vml'
vshape = p26.find('.//{%s}shape' % V)
style = vshape.get('style')
new_pt = new_emu / 12700.0
import re
style_new, n = re.subn(
    r'margin-left:[0-9.]+pt',
    'margin-left:%.2fpt' % new_pt,
    style,
)
assert n == 1, style
vshape.set('style', style_new)

# 4. Para 27 ("М. П.", end of page 2): replace its sectPr with the new
#    merged section-1 sectPr: A3 landscape, unequal 2-column layout
#    reproducing page1's and page2's own margins exactly.
p27_ppr = children[27].find(W + 'pPr')
p27_sectpr = p27_ppr.find(W + 'sectPr')

pgSz = p27_sectpr.find(W + 'pgSz')
pgSz.set(W + 'w', '23800')
pgSz.set(W + 'h', '16840')
pgSz.set(W + 'orient', 'landscape')

pgMar = p27_sectpr.find(W + 'pgMar')
pgMar.set(W + 'top', '1518')     # page1's own (smaller of the two)
pgMar.set(W + 'right', '1431')   # page2's own right margin
pgMar.set(W + 'bottom', '776')   # page2's own (smaller of the two)
pgMar.set(W + 'left', '1296')    # page1's own left margin

cols = p27_sectpr.find(W + 'cols')
for child in list(cols):
    cols.remove(child)
for attr in list(cols.attrib):
    del cols.attrib[attr]
cols.set(W + 'num', '2')
cols.set(W + 'equalWidth', '0')
col1 = etree.SubElement(cols, W + 'col')
col1.set(W + 'w', '9158')
col1.set(W + 'space', '2757')
col2 = etree.SubElement(cols, W + 'col')
col2.set(W + 'w', '9158')

# 5. Drop the stray trailing paragraphs (unwanted 4th page).
body.remove(children[35])
body.remove(children[34])

# 6. Final body sectPr (section 3 = page 3 only): widen to A3, push
#    content to the left half by enlarging the right margin while
#    keeping the original usable width (9129 twips) intact.
final_sectpr = children[36]
pgSz3 = final_sectpr.find(W + 'pgSz')
pgSz3.set(W + 'w', '23800')
pgSz3.set(W + 'h', '16840')
pgSz3.set(W + 'orient', 'landscape')
pgMar3 = final_sectpr.find(W + 'pgMar')
pgMar3.set(W + 'right', '13343')  # left(1328) + 9129 usable + 13343 = 23800

new_doc_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

# --- write back into a copy of the source docx ---
with zipfile.ZipFile(SRC) as zin, zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'word/document.xml':
            data = new_doc_xml
        zout.writestr(item, data)

print('Wrote', OUT)
