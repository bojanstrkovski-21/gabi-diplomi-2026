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
- `drzavna strucna.docx` (in `New folder\`) → `диплома_државна_стручна_матура_А3.docx` (script: `create_a3_drzavna_strucna.py`). This source was structurally different — normal flowing text + 1 table across 3 pre-existing sections, only 1 floating shape total, plus a stray unwanted 4th page (dropped per user instruction, keeping the same 3-source-page → 2-A3-sheet scheme as the other two). See [[feedback-docx-a3-section-column-method]] for the column-based technique used instead of shape-by-shape repositioning.

Not yet converted: none identified. `New folder\Diploma Drzavna strucna  matura  IV1.docx` was inspected and is NOT a template to convert — it's a filled-in real-student data overlay (2 pages, 15 shapes, no printed labels), a different kind of artifact entirely.

**Why:** user (Gabi's collaborator) wants printable A3 versions of these diploma documents that combine multiple A4 pages onto fewer physical A3 sheets while keeping all fields editable.

**How to apply:** See [[feedback-docx-a3-shape-conversion-method]] for the validated shape-repositioning approach (used when the source is built from many floating text boxes), and [[feedback-docx-a3-section-column-method]] for the column-based approach (used when the source is normal flowing text/tables with few or no floating shapes). Before starting a new document, check via Word COM/lxml which situation applies — don't assume it matches the previous file. `NOTES.txt` in the project root documents all conversions in detail and should be extended for each new document processed.
