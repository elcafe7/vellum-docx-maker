# Vellum DOCX Maker

Generate a Vellum-ready manuscript DOCX from the prepared Seneca edition in this folder.

The script builds a single print/import manuscript with:

- Part 1: English
- Part 2: Latin
- Chapter 1 to 6 as real `Heading 1` breaks
- `Heading 2` available for subheads
- Verse passages preserved as indented, italic blocks
- No title page or TOC inside the manuscript

## What it uses

- `build_vellum_docx.py` reads the finished bilingual HTML edition in this directory
- It emits `seneca-on-providence-vellum.docx`
- The DOCX is intended for Vellum import and ordinary print workflows

## Build

```bash
python3 build_vellum_docx.py
```

## Notes

This repo is intentionally small. The script is the deliverable; the surrounding files are the source edition and QA artifacts used to produce it.
