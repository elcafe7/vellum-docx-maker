# Vellum DOCX Maker

Convert Markdown manuscripts into Vellum-friendly DOCX files.

The script is opinionated for long-form manuscripts:

- the first `#` line becomes a centered title line
- `##` starts a new chapter by default
- `###` and deeper become subheads
- Blockquotes become indented italic paragraphs
- Bullets and numbered items become real Word lists
- Fenced code blocks become monospaced blocks
- Page breaks are inserted before each new chapter after the first

## Usage

```bash
python3 build_vellum_docx.py input.md -o output.docx --title "Book Title" --author "Author Name"
```

You can pass more than one Markdown file, and the script will append them in order.

## Output

The generated DOCX is meant for Vellum import and print production. Keep the Markdown structurally clean: use headings for sections, avoid decorative layout, and let the manuscript carry the structure.
