"""Synthetic, manually scoped bibliography adaptation; not an automatic editor.

Usage: python tests/check_bibliography_template.py NEW_OUTPUT_DIRECTORY
Requires python-docx and lxml. Word opening and visual comparison are separate.
"""

import re
import sys
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from lxml import etree as E


def text(node):
    return "".join(node.xpath(".//w:t/text()", namespaces=NS))


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def fixture(path, is_reference):
    doc = Document()
    doc.core_properties.author = "Synthetic bibliography test"
    doc.core_properties.last_modified_by = ""
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.25
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin = section.right_margin = Cm(2.5)
    section.top_margin = section.bottom_margin = Cm(2.5)
    doc.add_paragraph("Fictional bibliography formatting fixture")
    p = doc.add_paragraph("Protected body citation")
    p.add_run("[1]").font.superscript = True
    p.add_run(". Preserve this text, the formula and the final paragraph.")
    math = OxmlElement("m:oMath")
    mr = OxmlElement("m:r")
    mt = OxmlElement("m:t")
    mt.text = "x = 2"
    mr.append(mt)
    math.append(mr)
    doc.add_paragraph()._p.append(math)
    heading = doc.add_paragraph("5. References")
    heading.paragraph_format.space_before = Pt(8 if is_reference else 4)
    heading.paragraph_format.space_after = Pt(8 if is_reference else 4)
    heading.paragraph_format.keep_with_next = True
    heading.runs[0].font.size = Pt(14 if is_reference else 12)
    heading.runs[0].font.bold = True
    heading.runs[0].font.color.rgb = RGBColor(0, 0, 0)
    for i in range(1, 13):
        # Deliberately fictional identities and reserved example-domain URLs.
        author = "Mori A B, Van Example C D, et al."
        title = f"Fictional study {i}: controlled comparison of calibration methods under changing operating conditions"
        if i % 3 == 1:
            entry = f"{author} {title}[J]. Example Journal, 2025, 7(2): 10-19. DOI: 10.0000/fictional-{i}."
        elif i % 3 == 2:
            entry = f"{author} {title}[EB/OL]. Example Archive: demo-{i}v2, 2025[2026-01-02]. https://example.org/preprints/demo-{i}v2."
        else:
            entry = f"{author} {title}[C/OL]//Example Workshop. 2025[2026-01-02]. https://example.org/workshop/demo-{i}. DOI: 10.0000/fictional-{i}."
        p = doc.add_paragraph()
        p.add_run(f"[{i}]\t")
        p.add_run(entry)
        f = p.paragraph_format
        f.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if is_reference else WD_ALIGN_PARAGRAPH.LEFT
        f.left_indent = Cm(0.85 if is_reference else 1.1)
        f.first_line_indent = -f.left_indent
        f.tab_stops.add_tab_stop(f.left_indent)
        f.line_spacing = Pt(18) if is_reference else 1.25
        f.keep_together = False if is_reference else True
        f.widow_control = True
    doc.add_paragraph("Protected end paragraph; not a reference.")
    doc.save(path)


def main():
    out = Path(sys.argv[1]).resolve()
    out.mkdir(parents=True, exist_ok=False)
    fixture(out / "reference.docx", True)
    fixture(out / "before.docx", False)
    with ZipFile(out / "before.docx") as source, ZipFile(out / "reference.docx") as reference:
        old = E.fromstring(source.read("word/document.xml"))
        root = deepcopy(old)
        template = E.fromstring(reference.read("word/document.xml"))
        ps = root.findall(".//w:p", NS)
        originals = old.findall(".//w:p", NS)
        refs = template.findall(".//w:p", NS)
        start = next(i for i, p in enumerate(ps) if text(p) == "5. References")
        for i in range(start, start + 13):
            ps[i].replace(ps[i].find(qn("w:pPr")), deepcopy(refs[i].find(qn("w:pPr"))))
            if i == start:
                run = ps[i].find("w:r", NS)
                run.replace(run.find(qn("w:rPr")), deepcopy(refs[i].find("w:r/w:rPr", NS)))
                continue
            node = ps[i].findall("w:r/w:t", NS)[1]
            value = node.text
            author, rest = value.split(". ", 1)
            author = re.sub(r"\b[A-Z](?: [A-Z])+\b", lambda m: m[0].replace(" ", ""), author)
            node.text = author + ". " + rest
            if "[J]." in value:
                node.text = re.sub(r" DOI: 10\.0000/fictional-\d+\.$", "", node.text)
            assert "Van Example CD" in node.text
            assert re.findall(r"https?://\S+", value) == re.findall(r"https?://\S+", node.text)
            assert value.split(". ", 1)[1].split(" DOI: ")[0] == node.text.split(". ", 1)[1].split(" DOI: ")[0]
            if "[J]." not in value:
                assert value.split(". ", 1)[1] == node.text.split(". ", 1)[1]
        payload = E.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
        with ZipFile(out / "after.docx", "w") as target:
            for item in source.infolist():
                target.writestr(deepcopy(item), payload if item.filename == "word/document.xml" else source.read(item.filename))
        with ZipFile(out / "after.docx") as target:
            assert target.testzip() is None
            assert source.namelist() == target.namelist()
            assert all(source.read(n) == target.read(n) for n in source.namelist() if n != "word/document.xml")
            saved = E.fromstring(target.read("word/document.xml"))
        saved_ps = saved.findall(".//w:p", NS)
        for i in range(start, start + 13):
            node = saved_ps[i]
            node.getparent().replace(node, deepcopy(originals[i]))
        assert E.tostring(saved) == E.tostring(old)
        assert [text(p).split("]", 1)[0] for p in ps[start+1:start+13]] == [f"[{i}" for i in range(1, 13)]
    Document(out / "after.docx")
    print("PASS: 12 identities and numbers; author order and compound surname retained")
    print("PASS: 4 optional journal DOI displays removed; all URLs, versions and other DOI retained")
    print("PASS: reference-derived paragraph settings; only bibliography heading and entries changed")
    print("PASS: body superscript, editable math, end paragraph and other ZIP parts unchanged")
    print("NOT VERIFIED BY SCRIPT: visual matching, Word opening, metadata truth, automatic selection")
    print(out)


if __name__ == "__main__":
    main()
