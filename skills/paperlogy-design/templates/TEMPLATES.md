---
name: deck-replica
description: Create portable editable PPTX from a bundled catalog of 189 Paperology reference templates, or inspect and replicate authorized PPTX/PDF. Use when the user wants source-layout fidelity rather than a generic new style.
---
# Deck replica v0.3 portable

Read `references/PORTABILITY.md` for extraction and tested limitations.

## Route
- Bundled template reuse: read `references/catalog.json`, select by category/original_page and review its preview. Never require the original 2GB PPTX or a D: path.
- Source documents not in this pack: inspect their actual objects and pages; this catalog is a design resource, not evidence about another input.
- PDF visual copy may preserve appearance but is not editable reconstruction. Do not call a screenshot slide editable.

## Bundled execution
Python dependencies: lxml, Pillow. Optional validation: python-pptx, PyMuPDF; PDF rendering requires LibreOffice or PowerPoint. Never assume a renderer/font exists.
Run `python scripts/use_template.py --id paperology-050 --out NEW.pptx` to retain original text and native objects.
For changed content, write JSON mapping exact shape IDs to replacement strings, then run `python scripts/use_template.py --id paperology-050 --replace changes.json --out NEW.pptx`.
Inspect catalog `text_slots` before writing changes. Unknown shape IDs fail. Omitted slots retain text. Use --font only when an explicit substitution is wanted and disclose it.
Each invocation emits one editable slide and evidence JSON. To compose multi-slide output, use a PPTX-aware merger with dependency remapping; do not concatenate ZIP/XML parts naively. Current command is intentionally single-template.

## Fidelity boundaries
The five packs are static copies of 189 selected source slides. Native text, vector shapes, tables/charts where native, image crops, groups and background geometry remain. Photos and chart screenshots stay images, not newly editable graphs.
Videos/audio, animations, transitions, notes and external links are removed in copies. Video poster images remain. Raster dimensions are capped at 1920 px; JPEG quality88. This saves size but sacrifices original media playback and high-resolution detail. Original files are untouched.
No font files are included. Original font declarations may substitute in cloud. Fonts and third-party photos/logos require appropriate rights for redistribution. Do not present private access as a public redistribution license.
Changed text uses each paragraph's initial run style, so mixed inline styling may not survive. Unchanged reference mode retains original runs. Avoid claims of pixel-perfect typography.

## Verification before delivery
Run `python scripts/validate_pack.py OUTPUT.pptx` for ZIP relationships, missing rIds, media and native text/shape inventory. Render with an available office renderer; inspect all generated pages for truncation/overlap and compare intended text. Report renderer and font substitution. PDF text differences alone cannot prove source mismatch (outlines/off-canvas text exist).
Read `references/validation.json` for the tested build and `references/catalog.json` for source page provenance. If a new change breaks validation, fix or report it; do not silently rasterize to pass.
