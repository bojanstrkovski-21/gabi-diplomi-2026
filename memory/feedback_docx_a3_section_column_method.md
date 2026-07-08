---
name: feedback-docx-a3-section-column-method
description: Column-based technique for A3-izing a docx that is mostly normal flowing text/tables (few or no floating shapes), as an alternative to shape-by-shape repositioning
metadata:
  type: feedback
---

Not every diploma source document is built from dozens of floating text boxes like the first two ([[feedback-docx-a3-shape-conversion-method]]). Some (e.g. `drzavna strucna.docx`) are almost entirely normal flowing paragraphs/tables split across pre-existing Word sections (one section per source page), with at most one or two floating shapes. Trying to force the shape-repositioning method onto this kind of document would be pointless — there's nothing to reposition, the content just flows.

**Why:** discovered when the 3rd diploma template turned out to have only 1 floating shape total vs. 26-30 in the first two. Forcing every document into the "dump shapes via COM, match by text" playbook would waste effort — always check document structure first (see "How to apply").

**How to apply:**
1. Before picking a technique, check the source's actual structure: `doc.Shapes.Count` via COM, and whether `word/document.xml` has multiple `w:sectPr` elements (one per paragraph carrying a section break) vs. one shape-heavy single section. Few/no shapes + multiple sections → use this column method. Many floating shapes → use [[feedback-docx-a3-shape-conversion-method]] instead.
2. Each source page's own section already carries its own `w:pgMar` (margins can legitimately differ slightly page-to-page — don't assume they're uniform even within one document).
3. To combine two source pages side-by-side on one A3 sheet: delete the section break between them (remove the `w:sectPr` from the first page's last paragraph) so they become one section, then give that merged section a genuine 2-column `w:cols` with **unequal** column widths computed from each page's own original left/right margins: `col1_width = pageWidth - left1 - right1`, `gutter = right1 + left2`, `col2_width = pageWidth - left2 - right2`. This reproduces each page's original usable text width exactly and is algebraically identical to "shift page 2 by one full page-width" — verify the arithmetic by hand before trusting it.
4. **Critical gotcha:** `w:cols` needs BOTH `w:num="2"` AND `w:equalWidth="0"` with two `<w:col>` children. Omitting `w:num` makes Word silently treat the section as single-column even though the unequal `<w:col>` children are present and valid — and a `w:br type="column"` in a 1-column section renders as a **page break**, silently producing one extra page. Always verify column count via `doc.Sections(n).PageSetup.TextColumns.Count` after building, not just page count.
5. Insert `<w:br w:type="column"/>` as the first run of the second page's first paragraph to force it into column 2.
6. If the two merged pages have different original top margins, a section can only have one — use the smaller value as the section's top margin, and add the difference as extra "space before" on the second page's first paragraph to reproduce its original vertical start position.
7. Any floating shape inside the merged section: check `RelativeHorizontalPosition`/`RelativeVerticalPosition` via COM. If horizontal is `Page`-relative, shift its X by the same page-width amount as the column shift (both in the DrawingML `wp:anchor posOffset` and the VML fallback `v:shape` style `margin-left`). If vertical is `Paragraph`-relative, it auto-follows its anchor paragraph's reflow — no Y edit needed at all.
8. A source page that stays alone on its own A3 sheet (nothing beside it) doesn't need columns — just widen `pgSz` to A3 and enlarge the *unused* side's margin (right margin, if the content stays on the left) while keeping the original margin on the content side, so the usable text width is unchanged and content doesn't shift.
9. Verify the same way as the shape method: reopen via Word COM, check page count, shape count/position, and `TextColumns.Count`/widths — then render to PDF (`doc.ExportAsFixedFormat(path, 17)` + PyMuPDF `fitz` to rasterize pages) and visually inspect before calling it done.

This is reusable for future diploma documents that turn out to be mostly flowing text — check structure first ([[project-diploma-a3-conversion]]), pick shape-method or column-method accordingly.
