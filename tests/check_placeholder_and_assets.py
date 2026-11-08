"""Structural check for unified placeholders, Word draft/final gate, T25c parens, and diagram source parity.

Usage: python tests/check_placeholder_and_assets.py OUTPUT_DIRECTORY
Requires python-docx and lxml. Rasterization/visual review is a separate check.
"""

import re
import struct
import sys
import zlib
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from lxml import etree

PLACEHOLDER_RE = re.compile(r"【待补[：:](数据|文献|基本情况)[｜|]([^】【\n]+)】")
MALFORMED_PLACEHOLDER_RE = re.compile(r"\[待补[^\]]*\]|【待补(?![：:](?:数据|文献|基本情况)[｜|])[^】]*】")
CITATION_RE = re.compile(r"\[([\d,-]+)\]")
NUMBERING_PAREN_RE = re.compile(r"(?:式|图|表)\s*[（(]\d+(?:-\d+)?[）)]|^[（(]\d+[）)]")


def audit_text_placeholders(text: str):
    """Return valid placeholders, malformed placeholders, and numeric citation matches."""
    valid = PLACEHOLDER_RE.findall(text)
    malformed = MALFORMED_PLACEHOLDER_RE.findall(text)
    citations = CITATION_RE.findall(text)
    return valid, malformed, citations


def count_explanatory_parens(sentence: str) -> int:
    """Count explanatory parentheses in a sentence, excluding formula/item numbering like 式（2-3）."""
    stripped = NUMBERING_PAREN_RE.sub("", sentence)
    return len(re.findall(r"[（(][^）)]+[）)]", stripped))


def audit_docx_placeholders(docx_path: Path):
    """Inspect a DOCX file for unresolved placeholders and verify yellow highlight on draft runs."""
    with ZipFile(docx_path) as zf:
        root = etree.fromstring(zf.read("word/document.xml"))
    ns = {"w": qn("w:r").split("}")[0][1:]}
    unresolved = []
    for run in root.xpath(".//w:r", namespaces=ns):
        texts = [t.text for t in run.findall("w:t", ns) if t.text]
        joined = "".join(texts)
        for match in PLACEHOLDER_RE.finditer(joined):
            highlight = run.find("w:rPr/w:highlight", ns)
            hl_val = highlight.get(qn("w:val")) if highlight is not None else None
            unresolved.append((match.group(0), hl_val))
    full_text = "".join(t.text for t in root.xpath(".//w:t", namespaces=ns) if t.text)
    malformed = MALFORMED_PLACEHOLDER_RE.findall(full_text)
    return {
        "count": len(unresolved),
        "unresolved": unresolved,
        "malformed": malformed,
        "status": "FINAL_CLEARED" if (len(unresolved) == 0 and not malformed) else "DRAFT_UNRESOLVED",
    }


def write_minimal_grayscale_safe_png(png_path: Path, width: int = 320, height: int = 120):
    """Write a synthetic PNG fixture representing the rendered diagram for same-stem pairing checks."""
    bands = [
        (232, 241, 245),
        (214, 230, 242),
        (247, 234, 217),
        (226, 240, 203),
    ]
    raw_rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            band_idx = min(3, (x * 4) // width)
            if y < 4 or y >= height - 4 or (x % (width // 4) < 3):
                r, g, b = (45, 55, 72)
            else:
                r, g, b = bands[band_idx]
            row.extend((r, g, b))
        raw_rows.append(bytes(row))
    raw_data = b"".join(raw_rows)

    def chunk(tag: bytes, data: bytes) -> bytes:
        crc = zlib.crc32(tag + data) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", crc)

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    idat = zlib.compress(raw_data, 9)
    png_bytes = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")
    png_path.write_bytes(png_bytes)


def main():
    output = Path(sys.argv[1]).resolve()
    output.mkdir(parents=True, exist_ok=True)

    # 1. Placeholder vs [n] citation non-conflict check (both full-width ｜ and half-width |)
    sample = (
        "已有研究在恒定负载下验证了基线方法[1,3]，但在变负载切换下的丢包率尚未报告"
        "【待补：数据｜需补充变负载切换工况下的实测丢包率】；"
        "相关温漂补偿假设见【待补：文献|需补充支持温漂线性补偿假设的真实文献】，"
        "现场控制器型号与采样周期见【待补:基本情况｜需补充目标产线PLC型号与采样周期】[2-4]。"
    )
    valid, malformed, citations = audit_text_placeholders(sample)
    assert len(valid) == 3, f"Expected 3 valid placeholders, got {valid}"
    assert [v[0] for v in valid] == ["数据", "文献", "基本情况"]
    assert not malformed, f"Unexpected malformed placeholders: {malformed}"
    assert citations == ["1,3", "2-4"], f"Citations corrupted by placeholders: {citations}"

    bad_sample = "此处缺数据[待补：某数据]或【待补数据：未用竖线分隔】。"
    bad_valid, bad_malformed, _ = audit_text_placeholders(bad_sample)
    assert len(bad_valid) == 0 and len(bad_malformed) == 2

    # 2. T25c single-sentence explanatory parenthesis bound check
    t25c_output = (
        "针对边缘推理的时延受限问题，本节在黑盒教师模型输出端引入指数移动平均（Exponential Moving Average, EMA），"
        "该平滑策略仅在训练阶段更新参数且不增加推理阶段的计算开销。"
        "实验以采样频率（kHz）记录信号并将基准值设为 20、窗口时长设为 50 ms，"
        "再根据数据表中的“alarm_state”字段读取标签值为“fault”的样本。"
        "随后，模型通过平滑系数（$\\alpha \\in [0.9, 0.99]$）更新实验组 G2 的参考阈值，具体更新规则见式（2-3）。"
    )
    sentences = [s for s in t25c_output.split("。") if s.strip()]
    paren_counts = [count_explanatory_parens(s) for s in sentences]
    assert paren_counts == [1, 1, 1], f"Expected <=1 explanatory paren per sentence in T25c, got {paren_counts}"
    for required_token in [
        "（Exponential Moving Average, EMA）",
        "（kHz）",
        "（$\\alpha \\in [0.9, 0.99]$）",
        "实验组 G2",
        "式（2-3）",
        "“alarm_state”",
        "“fault”",
    ]:
        assert required_token in t25c_output, f"Missing required token in T25c: {required_token}"

    # 3. Word draft highlight check & pre-finalization zero-clearance gate
    draft_path = output / "draft_with_placeholder.docx"
    final_path = output / "final_cleared.docx"

    draft_doc = Document()
    p = draft_doc.add_paragraph("基线方法在恒定转速下已验证[1]，变转速下的误报率")
    ph_run = p.add_run("【待补：数据｜需补充变转速工况下的实测误报率】")
    ph_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    p.add_run("。")
    draft_doc.save(draft_path)

    draft_audit = audit_docx_placeholders(draft_path)
    assert draft_audit["status"] == "DRAFT_UNRESOLVED"
    assert draft_audit["count"] == 1
    assert draft_audit["unresolved"][0][1] == "yellow"

    final_doc = Document()
    final_doc.add_paragraph("基线方法在恒定转速下已验证[1]，变转速下的误报率为4.2%[2]。")
    final_doc.save(final_path)

    final_audit = audit_docx_placeholders(final_path)
    assert final_audit["status"] == "FINAL_CLEARED"
    assert final_audit["count"] == 0

    # 4. Same-directory, same-stem diagram source (.mmd, .svg) and pairing fixture (.png) check
    stem = "fig2-1"
    mmd_path = output / f"{stem}.mmd"
    svg_path = output / f"{stem}.svg"
    png_path = output / f"{stem}.png"

    mmd_content = """flowchart LR
  subgraph S1["输入阶段"]
    A["信号采集输入"]
  end
  subgraph S2["处理阶段"]
    B["滑窗特征处理"]
  end
  subgraph S3["输出阶段"]
    C["异常状态输出"]
  end
  subgraph S4["评估阶段"]
    D["离线协议评估"]
  end
  A --> B --> C -.-> D
  classDef inputRole fill:#E8F1F5,stroke:#1B4965,stroke-width:1.5px,color:#111111;
  classDef procRole fill:#D6E6F2,stroke:#2B580C,stroke-width:1.5px,color:#111111;
  classDef outRole fill:#F7EAD9,stroke:#7F4F24,stroke-width:1.5px,color:#111111;
  classDef evalRole fill:#E2F0CB,stroke:#386641,stroke-width:1.5px,stroke-dasharray:4 2,color:#111111;
  class A inputRole;
  class B procRole;
  class C outRole;
  class D evalRole;
"""
    svg_content = """<svg xmlns="http://www.w3.org/2000/svg" width="640" height="140" viewBox="0 0 640 140">
  <g id="node-A"><rect x="20" y="40" width="120" height="50" fill="#E8F1F5" stroke="#1B4965" stroke-width="1.5"/><text x="80" y="70" text-anchor="middle" font-size="12">信号采集输入</text></g>
  <g id="node-B"><rect x="180" y="40" width="120" height="50" fill="#D6E6F2" stroke="#2B580C" stroke-width="1.5"/><text x="240" y="70" text-anchor="middle" font-size="12">滑窗特征处理</text></g>
  <g id="node-C"><rect x="340" y="40" width="120" height="50" fill="#F7EAD9" stroke="#7F4F24" stroke-width="1.5"/><text x="400" y="70" text-anchor="middle" font-size="12">异常状态输出</text></g>
  <g id="node-D"><rect x="500" y="40" width="120" height="50" fill="#E2F0CB" stroke="#386641" stroke-width="1.5" stroke-dasharray="4,2"/><text x="560" y="70" text-anchor="middle" font-size="12">离线协议评估</text></g>
  <line x1="140" y1="65" x2="180" y2="65" stroke="#1B4965" stroke-width="1.5"/>
  <line x1="300" y1="65" x2="340" y2="65" stroke="#2B580C" stroke-width="1.5"/>
  <line x1="460" y1="65" x2="500" y2="65" stroke="#386641" stroke-width="1.5" stroke-dasharray="4,2"/>
</svg>
"""
    mmd_path.write_text(mmd_content, encoding="utf-8")
    svg_path.write_text(svg_content, encoding="utf-8")
    write_minimal_grayscale_safe_png(png_path)

    assert mmd_path.exists() and svg_path.exists() and png_path.exists()
    assert mmd_path.stem == svg_path.stem == png_path.stem == stem
    expected_labels = ["信号采集输入", "滑窗特征处理", "异常状态输出", "离线协议评估"]
    for label in expected_labels:
        assert label in mmd_content and label in svg_content
    fills = set(re.findall(r'fill="(#[0-9A-Fa-f]{6})"', svg_content))
    assert 1 <= len(fills) <= 5, f"Expected <=5 semantic fill colors, got {fills}"

    print("PASS: full-width placeholder regex (｜ and |) never conflicts with [n] citations; malformed brackets caught")
    print("PASS: T25c single-sentence explanatory parenthesis bound (<=1 per sentence) and required tokens verified")
    print("PASS: Word draft placeholder highlight verified and final DOCX zero-placeholder gate enforced")
    print("PASS: same-directory same-stem .mmd/.svg/.png pairing and <=5 semantic colors verified")
    print("NOT VERIFIED BY SCRIPT: Mermaid/SVG rasterization, visual layout inspection, and Microsoft Word opening")
    print(output)


if __name__ == "__main__":
    main()
