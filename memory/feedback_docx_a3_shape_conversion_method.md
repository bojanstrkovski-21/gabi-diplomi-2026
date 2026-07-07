---
name: feedback-docx-a3-shape-conversion-method
description: Validated technique for repositioning floating docx text-box shapes onto a new A3 layout without breaking editability
metadata:
  type: feedback
---

When converting a Word .docx with many floating/anchored text-box shapes onto a new page layout (e.g. A4→A3 side-by-side), do NOT use Word COM copy-paste automation to move shapes between documents — it was tried first (`create_a3.py`) and was unreliable/broken. Directly editing `word/document.xml` inside the .docx zip with `lxml` is the reliable approach.

**Why:** COM copy-paste loses fidelity on complex shapes (grouped border lines, mixed DrawingML/VML shapes) and is slow/fragile. Direct XML manipulation preserves every shape exactly and is deterministic.

**How to apply** (the working recipe, used successfully twice):
1. Open the source `.doc`/`.docx` via `win32com.client.Dispatch('Word.Application')` and dump every shape's `Anchor.Information(wdActiveEndPageNumber)` (source page), `RelativeHorizontalPosition`/`RelativeVerticalPosition`, `Left`/`Top`, and `TextFrame.TextRange.Text` (for identification).
2. Separately parse `word/document.xml` with `lxml`, walk paragraphs in order, and collect all `w:r` runs containing a `wp:anchor` (DrawingML). Extract each anchor's text via nested `w:t` elements — this gives the true document order, which COM's `Shapes` collection does NOT reliably match.
3. Match each XML-order shape to its COM record by exact text. Duplicate text (e.g. two shapes both saying "Битола") gets disambiguated by cross-checking spatial position/context (e.g. header cluster vs. footer cluster) — always sanity-check ambiguous matches this way rather than assuming COM order.
4. Compute each shape's true page-absolute position: if `RelativeHorizontalPosition/Vertical == wdRelHPos_Margin`, add the page margin (source doc margins: left=89.85pt, top=27pt in both diplomas processed so far); otherwise the COM value is already page-relative.
5. Map source page → A3 target page/offset: source page 1 → A3 page 1 unchanged; source page 2 → A3 page 1 shifted right by one A4 width (595.35pt); source page 3 → A3 page 2 unchanged. This works because A4 portrait height (842pt/297mm) equals A3 landscape height (297mm) — the vertical axis needs no scaling, only horizontal.
6. In the XML, set both DrawingML (`wp:anchor/wp:positionH|V` → `relativeFrom="page"` + `posOffset` in EMU = pt×12700) and the VML fallback (`v:shape`/`v:group` inline `style` margin-left/margin-top, `mso-position-*-relative:page`) to keep old Word render paths consistent.
7. Rebuild the document body as: para with all A3-page-1 shapes → para with `pageBreakBefore` → para with all A3-page-2 shapes → empty trailing para → `sectPr` set to A3 landscape (23811×16838 twips, ~1mm/57-twip margins).
8. **Always verify the output** by reopening it via Word COM and checking page count, shape count, and each shape's resulting page/position matches expectations before calling the task done.

This whole method is reusable as-is for the next diploma document ([[project-diploma-a3-conversion]]) — only the per-document paragraph indices and position table change.
