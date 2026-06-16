# Seneca: On Providence

Vellum-ready manuscript tooling for Seneca's *De Providentia* / *On Providence*.

This repo packages a small Python builder that turns the bilingual edition into a single DOCX for print and Vellum import. The manuscript is shaped as:

- Part 1: English
- Part 2: Latin
- Chapter 1 to 6 as real `Heading 1` breaks
- Subheads available for `Heading 2+`
- Verse passages left-indented and italicized
- No title page or TOC inside the source manuscript

## Files

- [`build_vellum_docx.py`](./build_vellum_docx.py): builds the Vellum DOCX from the prepared HTML edition
- [`seneca-on-providence.html`](./seneca-on-providence.html): the bilingual reader edition used as source
- [`seneca-on-providence-vellum.docx`](./seneca-on-providence-vellum.docx): the finished manuscript for import or print

## Build

```bash
python3 build_vellum_docx.py
```

## Source

The English translation is Aubrey Stewart's 1900 public-domain version. The Latin text is from The Latin Library.
