---
name: project-diploma-a3-conversion
description: Converting Macedonian diploma .docx files from 3-page A4 portrait to 2-page A3 landscape, preserving editable text boxes
metadata:
  type: project
---

Working directory `f:\Gabi-diplomi` (not a git repo) contains multiple Macedonian diploma/certificate documents that need to be converted from 3-page A4 portrait layout into 2-page A3 landscape layout, laid out as two A4 pages side-by-side per A3 sheet (source page 1 = A3 p1 left, source page 2 = A3 p1 right, source page 3 = A3 p2 left, A3 p2 right empty). All floating text-box shapes must remain live/editable in Word after conversion — not flattened to images.

Completed so far:
- `диплома државна матура (1).docx` → `диплома_А3_v2.docx` (30 shapes, script: `create_a3_xml.py`)
- `училишна стручна матура.docx` → `училишна_стручна_матура_А3.docx` (26 shapes, script: `create_a3_strucna.py`)

Not yet converted: `Diploma Drzavna strucna matura IV1.docx` (in `New folder\`) — likely the next target if this work resumes.

**Why:** user (Gabi's collaborator) wants printable A3 versions of these diploma documents that combine multiple A4 pages onto fewer physical A3 sheets while keeping all fields editable.

**How to apply:** See [[feedback-docx-a3-shape-conversion-method]] for the validated technical approach. Each new source document needs its own per-shape position table since paragraph indices and shape positions differ per file — reuse the method, not literal numbers, from the completed scripts. `NOTES.txt` in the project root documents both conversions in detail and should be extended for each new document processed.
