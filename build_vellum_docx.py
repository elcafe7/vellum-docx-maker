#!/usr/bin/env python3
from pathlib import Path

from lxml import html
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt


ROOT = Path(__file__).resolve().parent
TITLE = "Seneca: On Providence"
AUTHOR = "Seneca"
TRANSLATOR = "Aubrey Stewart"


def load_volume(article):
    label = article.get("data-volume")
    lang = article.get("lang")
    title = article.xpath('.//header[contains(@class, "volume-header")]//h2/text()')
    title = title[0].strip() if title else ""
    sections = []
    for section in article.xpath('./section[contains(@class, "chapter")]'):
        roman = section.xpath('.//div[contains(@class, "chapter-mark")]//strong/text()')
        roman = roman[0].strip() if roman else ""
        blocks = []
        body = section.xpath('.//div[contains(@class, "chapter-body")]')
        if body:
            for node in body[0]:
                tag = node.tag.lower() if isinstance(node.tag, str) else ""
                if tag == "p":
                    text = node.text_content().strip()
                    if text:
                        blocks.append(("p", text))
                elif tag == "blockquote":
                    verse_lines = [line.strip() for line in node.text_content().splitlines() if line.strip()]
                    if verse_lines:
                        blocks.append(("v", verse_lines))
        sections.append({"roman": roman, "blocks": blocks})
    return {"number": label, "lang": lang, "title": title, "sections": sections}


def add_title_style(doc: Document):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Georgia"
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.25

    for name in ("Heading 1", "Heading 2"):
        style = styles[name]
        style.font.name = "Georgia"

    h1 = styles["Heading 1"]
    h1.font.size = Pt(14)
    h1.font.bold = True
    h1.paragraph_format.space_before = Pt(18)
    h1.paragraph_format.space_after = Pt(8)

    h2 = styles["Heading 2"]
    h2.font.size = Pt(12)
    h2.font.bold = True
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(4)


def add_metadata(doc: Document):
    props = doc.core_properties
    props.title = TITLE
    props.author = AUTHOR
    props.subject = "English and Latin manuscript prepared for Vellum import"
    props.comments = f"Volume 1 English translation by {TRANSLATOR}; Volume 2 Latin text."
    props.keywords = "Vellum, DOCX, Seneca, De Providentia, English, Latin"


def add_part(doc: Document, title: str):
    p = doc.add_paragraph(style="Heading 1")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(title)


def add_chapter(doc: Document, title: str):
    p = doc.add_paragraph(style="Heading 1")
    p.add_run(title)


def add_body_paragraph(doc: Document, text: str):
    p = doc.add_paragraph(style="Normal")
    # Preserve typographic quotes and em dashes from the source, but keep the
    # manuscript structurally plain for Vellum.
    p.add_run(text)


def add_verse_block(doc: Document, lines):
    for idx, line in enumerate(lines):
        p = doc.add_paragraph(style="Normal")
        p.paragraph_format.left_indent = Inches(0.6)
        p.paragraph_format.right_indent = Inches(0.6)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(line)
        r.italic = True
        if idx < len(lines) - 1:
            r.add_break()


def build_volume(doc: Document, volume):
    add_part(doc, f"Part {volume['number']}: {volume['title']}")
    for idx, section in enumerate(volume["sections"], start=1):
        add_chapter(doc, f"Chapter {idx}: {section['roman']}")
        for kind, value in section["blocks"]:
            if kind == "p":
                add_body_paragraph(doc, value)
            else:
                add_verse_block(doc, value)


def main():
    doc = Document()
    add_title_style(doc)
    add_metadata(doc)

    section = doc.sections[0]
    section.top_margin = Inches(0.85)
    section.bottom_margin = Inches(0.85)
    section.left_margin = Inches(0.95)
    section.right_margin = Inches(0.95)

    tree = html.parse(str(ROOT / "seneca-on-providence.html"))
    for article in tree.xpath('//article[contains(@class, "volume")]'):
        build_volume(doc, load_volume(article))

    out = ROOT / "seneca-on-providence-vellum.docx"
    doc.save(out)
    print(out)


if __name__ == "__main__":
    main()
