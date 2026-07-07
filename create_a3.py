import win32com.client
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

doc_path = r'c:\Users\User\Downloads\Gabi-diplomi\диплома државна матура (1).doc'
out_path = r'c:\Users\User\Downloads\Gabi-diplomi\диплома_А3.docx'

# WD/MSO constants
wdRelHPos_Page = 1
wdRelHPos_Margin = 2
wdRelVPos_Page = 1
wdRelVPos_Margin = 2
wdOrientLandscape = 1
wdCollapseStart = 1
wdCollapseEnd = 0
wdPageBreak = 7
wdFormatDocumentDefault = 16
wdActiveEndPageNumber = 3

# Source document measurements (points)
A4_WIDTH = 595.35
A4_HEIGHT = 842.0
SRC_LEFT_MARGIN = 89.85
SRC_TOP_MARGIN = 27.0

# A3 landscape dimensions (points): 420mm x 297mm
A3_WIDTH = 1190.55
A3_HEIGHT = 841.89

print("Starting Word automation...")
word = win32com.client.Dispatch('Word.Application')
word.Visible = False

try:
    print("Opening source document...")
    src_doc = word.Documents.Open(doc_path)

    print("Creating new A3 document...")
    new_doc = word.Documents.Add()

    # Configure A3 landscape page
    ps = new_doc.PageSetup
    ps.PageWidth = A3_WIDTH
    ps.PageHeight = A3_HEIGHT
    ps.LeftMargin = 1
    ps.RightMargin = 1
    ps.TopMargin = 1
    ps.BottomMargin = 1

    # Verify page dimensions were set
    print(f"New doc page: {ps.PageWidth:.2f} x {ps.PageHeight:.2f}")

    # Add a page break to create the second A3 page
    # Para 1 -> A3 page 1, Para last -> A3 page 2
    rng = new_doc.Range()
    rng.Collapse(wdCollapseEnd)
    rng.InsertBreak(wdPageBreak)

    para1_range = new_doc.Paragraphs(1).Range
    para2_range = new_doc.Paragraphs(new_doc.Paragraphs.Count).Range

    print(f"Source shapes: {src_doc.Shapes.Count}")
    print("Processing shapes...\n")

    shapes_data = []
    for i in range(1, src_doc.Shapes.Count + 1):
        s = src_doc.Shapes(i)
        page = s.Anchor.Information(wdActiveEndPageNumber)
        shapes_data.append({
            'index': i,
            'page': page,
            'rel_h': s.RelativeHorizontalPosition,
            'rel_v': s.RelativeVerticalPosition,
            'left': s.Left,
            'top': s.Top,
            'width': s.Width,
            'height': s.Height,
            'type': s.Type,
        })
        try:
            text = s.TextFrame.TextRange.Text[:40] if s.TextFrame.HasText else ''
        except:
            text = ''
        print(f"  Shape {i} page={page} type={s.Type} left={s.Left:.1f} top={s.Top:.1f} text={repr(text[:40])}")

    print("\nCopying shapes to A3 document...")

    for sd in shapes_data:
        i = sd['index']
        s = src_doc.Shapes(i)
        page = sd['page']

        # Convert to absolute position from page edge
        if sd['rel_h'] == wdRelHPos_Margin:
            abs_left = SRC_LEFT_MARGIN + sd['left']
        else:
            abs_left = sd['left']

        if sd['rel_v'] == wdRelVPos_Margin:
            abs_top = SRC_TOP_MARGIN + sd['top']
        else:
            abs_top = sd['top']

        # Placement on A3:
        # Source page 1 -> left half of A3 page 1 (no x offset)
        # Source page 2 -> right half of A3 page 1 (x offset = A4_WIDTH)
        # Source page 3 -> left half of A3 page 2 (no x offset)
        if page == 1:
            new_left = abs_left
            use_para2 = False
        elif page == 2:
            new_left = abs_left + A4_WIDTH
            use_para2 = False
        else:  # page 3
            new_left = abs_left
            use_para2 = True

        new_top = abs_top

        # Set anchor position in new_doc before switching to src_doc
        new_doc.Activate()
        if use_para2:
            para2_range.Select()
        else:
            para1_range.Select()
        word.Selection.Collapse(wdCollapseStart)

        # Copy shape from source
        src_doc.Activate()
        s.Select()
        word.Selection.Copy()

        # Paste into new_doc at the set selection position
        new_doc.Activate()
        word.Selection.Paste()

        # Adjust the pasted shape position
        if new_doc.Shapes.Count > 0:
            pasted = new_doc.Shapes(new_doc.Shapes.Count)
            pasted.RelativeHorizontalPosition = wdRelHPos_Page
            pasted.RelativeVerticalPosition = wdRelVPos_Page
            pasted.Left = new_left
            pasted.Top = new_top
            pasted.Width = sd['width']
            pasted.Height = sd['height']
            print(f"  Placed shape {i} (src page {page}) at left={new_left:.1f} top={new_top:.1f}")

    print(f"\nSaving to: {out_path}")
    new_doc.SaveAs2(out_path, wdFormatDocumentDefault)
    new_doc.Close()
    src_doc.Close(False)
    print("Done!")

except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
    try:
        new_doc.Close(False)
    except:
        pass
    try:
        src_doc.Close(False)
    except:
        pass

finally:
    try:
        word.Quit()
    except:
        pass
