"""Gabungkan naskah markdown menjadi satu file Word (.docx) bergaya.

Pemakaian:  python3 build.py
Hasil:      Rahasia-Panggul-Perkasa.docx (di folder yang sama)
"""
import pathlib
import re
import subprocess
import tempfile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = pathlib.Path(__file__).parent
NASKAH = ROOT / "naskah"
OUTPUT = ROOT / "Rahasia-Panggul-Perkasa.docx"

HIJAU = RGBColor(0x1B, 0x6B, 0x3A)
ORANYE = RGBColor(0xD9, 0x74, 0x1C)
TEKS = RGBColor(0x22, 0x22, 0x22)
FONT_TEKS = "Georgia"
FONT_JUDUL = "Arial"
FONT_ARAB = "Traditional Arabic"

PAGEBREAK = '```{=openxml}\n<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n```'


def set_fonts(style, latin, size=None, bold=None, color=None, italic=None):
    style.font.name = latin
    if size:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if italic is not None:
        style.font.italic = italic
    if color is not None:
        style.font.color.rgb = color
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        rfonts.set(qn(attr), latin)
    rfonts.set(qn("w:cs"), FONT_ARAB)
    for tag in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        rfonts.attrib.pop(qn(tag), None)
    szcs = rpr.find(qn("w:szCs"))
    if szcs is None:
        szcs = OxmlElement("w:szCs")
        rpr.append(szcs)
    szcs.set(qn("w:val"), str(int((size or 11) * 2 + 6)))


def shade_paragraph_style(style, fill, border_color):
    ppr = style.element.get_or_add_pPr()
    for tag in ("w:pBdr", "w:shd"):
        old = ppr.find(qn(tag))
        if old is not None:
            ppr.remove(old)
    bdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "24")
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), border_color)
    bdr.append(left)
    ppr.append(bdr)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    ppr.append(shd)


def build_reference(path):
    default = subprocess.run(
        ["pandoc", "--print-default-data-file", "reference.docx"],
        check=True, capture_output=True,
    ).stdout
    path.write_bytes(default)
    doc = Document(path)

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.5)
    sec.top_margin = sec.bottom_margin = Cm(2.5)

    styles = {s.name: s for s in doc.styles}
    st = styles
    for name in ("Normal", "Body Text", "First Paragraph", "Compact"):
        if name in st:
            set_fonts(st[name], FONT_TEKS, 11, color=TEKS)
            pf = st[name].paragraph_format
            pf.space_after = Pt(6)
            pf.line_spacing = 1.25

    set_fonts(st["Title"], FONT_JUDUL, 30, bold=True, color=HIJAU)
    st["Title"].paragraph_format.space_before = Pt(180)
    st["Title"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_fonts(st["Subtitle"], FONT_JUDUL, 15, bold=False, color=ORANYE, italic=True)
    st["Subtitle"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if "Author" in st:
        set_fonts(st["Author"], FONT_JUDUL, 12, color=TEKS)
        st["Author"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        st["Author"].paragraph_format.space_before = Pt(36)

    heads = {
        "Heading 1": (20, HIJAU, 0, 14),
        "Heading 2": (14, ORANYE, 14, 6),
        "Heading 3": (12, HIJAU, 10, 4),
    }
    for name, (size, color, before, after) in heads.items():
        set_fonts(st[name], FONT_JUDUL, size, bold=True, color=color)
        pf = st[name].paragraph_format
        pf.space_before, pf.space_after = Pt(before), Pt(after)
        pf.keep_with_next = True

    if "Block Text" in st:
        bt = st["Block Text"]
        set_fonts(bt, FONT_TEKS, 10.5, color=TEKS, italic=False)
        bt.paragraph_format.left_indent = Cm(0.4)
        bt.paragraph_format.right_indent = Cm(0.2)
        bt.paragraph_format.space_after = Pt(0)
        bt.paragraph_format.space_before = Pt(0)
        shade_paragraph_style(bt, "EEF6EF", "1B6B3A")

    doc.save(path)


def combine():
    parts = []
    for md in sorted(NASKAH.glob("*.md")):
        parts.append(md.read_text(encoding="utf-8"))
    text = "\n\n".join(parts)
    # Beri jarak antar-kotak berurutan agar tidak menyatu.
    text = re.sub(r"(^>[^\n]*\n)\n(?=> )", "\\1\n\u00a0\n\n", text, flags=re.M)
    return text.replace("<<<PAGEBREAK>>>", PAGEBREAK)


def main():
    with tempfile.TemporaryDirectory() as tmp:
        tmp = pathlib.Path(tmp)
        ref = tmp / "reference.docx"
        build_reference(ref)
        src = tmp / "naskah.md"
        src.write_text(combine(), encoding="utf-8")
        subprocess.run(
            ["pandoc", str(src), "-f", "markdown+raw_attribute",
             "-o", str(OUTPUT), "--reference-doc", str(ref), "--columns=40"],
            check=True,
        )
    print(f"OK: {OUTPUT}")


if __name__ == "__main__":
    main()
