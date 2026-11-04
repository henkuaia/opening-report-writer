"""Synthetic, manually scoped citation edit; not an automatic citation engine.

Usage: python tests/check_citation_word.py OUTPUT_DIRECTORY
Requires python-docx and lxml. Rendering/visual review is a separate check.
"""

import re
import sys
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from lxml import etree


def expand(group):
    """Expand already-reviewed numeric citation groups, never identify citations."""
    result = set()
    for bracket in re.findall(r"\[([\d,-]+)\]", group):
        for token in bracket.split(","):
            ends = [int(value) for value in token.split("-")]
            result.update(range(ends[0], ends[-1] + 1))
    return result


def run_node(text):
    run = OxmlElement("w:r")
    value = OxmlElement("w:t")
    value.text = text
    run.append(value)
    return run


def main():
    output = Path(sys.argv[1]).resolve()
    output.mkdir(parents=True, exist_ok=False)
    before = output / "before.docx"
    after = output / "after.docx"
    doc = Document()
    doc.core_properties.author = "Synthetic citation test"
    doc.core_properties.last_modified_by = ""
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
    normal.paragraph_format.line_spacing = 1.25
    normal.paragraph_format.space_after = Pt(6)
    doc.styles["Title"].font.color.rgb = RGBColor(0, 0, 0)
    doc.styles["Title"].font.size = Pt(18)
    for border in doc.styles["Title"].element.xpath("./w:pPr/w:pBdr"):
        border.getparent().remove(border)
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(2)
    section.left_margin = section.right_margin = Cm(2.5)
    doc.add_paragraph("虚构引文格式测试", "Title")
    doc.add_paragraph("所有材料均为虚构，仅验证引用格式，不是研究证据。")
    selected = []

    def citation(paragraph, text, replacement):
        run = paragraph.add_run(text)
        selected.append((run._r, replacement))

    p = doc.add_paragraph("甲、乙、丙均以校准窗口的中位数为固定参照")
    citation(p, "[1][2][3]", "[1-3]")
    p.add_run("。乙仅在恒速工况报告误报减少")
    citation(p, "[2]", "[2]")
    p.add_run("；丙仅比较存储开销。")
    p = doc.add_paragraph("同组非连续来源")
    citation(p, "[2][5][7]", "[2,5,7]")
    p.add_run("，不能补入缺号。混合来源")
    citation(p, "[1][2][3][7]", "[1-3,7]")
    p.add_run("。")
    doc.add_paragraph("公式（2）中的向量[1,2]、区间[2,7]和图3不是引文。")
    p = doc.add_paragraph("可编辑公式：")
    math = OxmlElement("m:oMath")
    mr = OxmlElement("m:r")
    mt = OxmlElement("m:t")
    mt.text = "x = 2"
    mr.append(mt)
    math.append(mr)
    p._p.append(math)
    p = doc.add_paragraph("交叉引用域")
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), " REF ref2 \\h ")
    cached = run_node("[2]")
    field.append(cached)
    p._p.append(field)
    selected.append((cached, "[2]"))
    p.add_run("。内部超链接")
    link = OxmlElement("w:hyperlink")
    link.set(qn("w:anchor"), "ref3")
    linked = run_node("[3]")
    link.append(linked)
    p._p.append(link)
    selected.append((linked, "[3]"))
    p.add_run("。")
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    p = table.cell(0, 0).paragraphs[0]
    p.add_run("表内同组引文")
    citation(p, "[1][2][3]", "[1-3]")
    p.add_run("。")
    p = doc.add_paragraph("表1 虚构材料对照")
    citation(p, "[1][3]", "[1,3]")
    p.add_run("。")
    doc.add_paragraph("虚构来源编号表（保留既有编号，不重排）")
    for number, label in enumerate("甲乙丙丁戊己庚", 1):
        p = doc.add_paragraph()
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(number))
        start.set(qn("w:name"), f"ref{number}")
        p._p.append(start)
        p.add_run(f"[{number}]")
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(number))
        p._p.append(end)
        p.add_run(f" 材料{label}，虚构测试标识，非正式题录。")
    all_runs = list(doc.element.xpath(".//w:r"))
    plan = [(all_runs.index(run), value) for run, value in selected]
    doc.save(before)

    with ZipFile(before) as source:
        original = etree.fromstring(source.read("word/document.xml"))
        edited = deepcopy(original)
        ns = {"w": qn("w:r").split("}")[0][1:]}
        runs = edited.xpath(".//w:r", namespaces=ns)
        for index, replacement in plan:
            run = runs[index]
            text = run.find(qn("w:t"))
            assert expand(text.text) == expand(replacement)
            text.text = replacement
            properties = run.find(qn("w:rPr"))
            if properties is None:
                properties = OxmlElement("w:rPr")
                run.insert(0, properties)
            vertical = OxmlElement("w:vertAlign")
            vertical.set(qn("w:val"), "superscript")
            properties.append(vertical)
        field = edited.find(".//w:fldSimple", ns)
        original_instruction = field.get(qn("w:instr"))
        field.set(qn("w:instr"), original_instruction + "\\* MERGEFORMAT ")
        with ZipFile(after, "w") as target:
            for entry in source.infolist():
                data = source.read(entry.filename)
                if entry.filename == "word/document.xml":
                    data = etree.tostring(edited, xml_declaration=True, encoding="UTF-8", standalone=True)
                target.writestr(entry, data)

    with ZipFile(before) as source, ZipFile(after) as target:
        assert source.namelist() == target.namelist()
        for name in source.namelist():
            if name != "word/document.xml":
                assert source.read(name) == target.read(name), name
        saved = etree.fromstring(target.read("word/document.xml"))
    old_runs = original.xpath(".//w:r", namespaces=ns)
    saved_runs = saved.xpath(".//w:r", namespaces=ns)
    for index, replacement in plan:
        run = saved_runs[index]
        assert run.find(qn("w:t")).text == replacement
        assert run.find("w:rPr/w:vertAlign", ns).get(qn("w:val")) == "superscript"
        restored = deepcopy(run)
        restored.find(qn("w:t")).text = old_runs[index].find(qn("w:t")).text
        props = restored.find(qn("w:rPr"))
        props.remove(props.find(qn("w:vertAlign")))
        if old_runs[index].find(qn("w:rPr")) is None:
            restored.remove(props)
        assert etree.tostring(restored, method="c14n", exclusive=True) == etree.tostring(
            old_runs[index], method="c14n", exclusive=True
        )
        run.getparent().replace(run, deepcopy(old_runs[index]))
    saved_field = saved.find(".//w:fldSimple", ns)
    assert saved_field.get(qn("w:instr")) == original_instruction + "\\* MERGEFORMAT "
    saved_field.set(qn("w:instr"), original_instruction)
    assert etree.tostring(saved) == etree.tostring(original)
    assert expand("[1,3,5]") != expand("[1-5]")
    assert expand("[2,5,7]") != expand("[2-7]")
    Document(after)
    print(f"PASS: {len(plan)} reviewed citation groups; exact source sets preserved")
    print("PASS: citation text/superscript and field format switch only; REF target unchanged")
    print("PASS: all other XML/ZIP parts, bookmarks, links, math and bibliography unchanged")
    print("NOT VERIFIED BY SCRIPT: evidence support, automatic citation detection, rendering, Word opening")
    print(output)


if __name__ == "__main__":
    main()
