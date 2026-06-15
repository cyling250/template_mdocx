"""通用辅助：XML 操作、OMath 公式、字体、边框、书签ID"""
from lxml import etree
from docx.shared import Pt
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_bm_id_gen = [0]


def next_bookmark_id():
    _bm_id_gen[0] += 1
    return _bm_id_gen[0]


# ═══════════════ 段落样式 ═══════════════

def set_para_style(paragraph, style_id):
    """直接设置段落样式 ID（绕过 python-docx 单元格 bug）。"""
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        pPr = parse_xml(f'<w:pPr {nsdecls("w")}/>')
        paragraph._element.insert(0, pPr)
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        pStyle = parse_xml(f'<w:pStyle {nsdecls("w")} w:val="{style_id}"/>')
        pPr.insert(0, pStyle)
    else:
        pStyle.set(qn("w:val"), style_id)


# ═══════════════ OMath 公式 ═══════════════

def add_omath_inline(para, latex):
    """行内 OMath（LaTeX 文本，用户手动渲染）。"""
    run_el = etree.SubElement(para._element, f'{{{W_NS}}}r')
    om = etree.SubElement(run_el, f'{{{MATH_NS}}}oMath')
    r = etree.SubElement(om, f'{{{MATH_NS}}}r')
    rPr = etree.SubElement(r, f'{{{W_NS}}}rPr')
    rf = etree.SubElement(rPr, f'{{{W_NS}}}rFonts')
    rf.set(f'{{{W_NS}}}ascii', 'Cambria Math')
    rf.set(f'{{{W_NS}}}hAnsi', 'Cambria Math')
    t = etree.SubElement(r, f'{{{MATH_NS}}}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = latex


def add_omath_display(para, latex):
    """块级 OMath（LaTeX 文本，用户手动渲染）。"""
    op = etree.SubElement(para._element, f'{{{MATH_NS}}}oMathPara')
    om = etree.SubElement(op, f'{{{MATH_NS}}}oMath')
    r = etree.SubElement(om, f'{{{MATH_NS}}}r')
    rPr = etree.SubElement(r, f'{{{W_NS}}}rPr')
    rf = etree.SubElement(rPr, f'{{{W_NS}}}rFonts')
    rf.set(f'{{{W_NS}}}ascii', 'Cambria Math')
    rf.set(f'{{{W_NS}}}hAnsi', 'Cambria Math')
    t = etree.SubElement(r, f'{{{MATH_NS}}}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = latex


# ═══════════════ 字体 ═══════════════

def apply_font(run, en="Times New Roman", cn="宋体", sz=12, bold=False, italic=False):
    """设置 run 的中英文字体、字号、粗斜体。"""
    run.font.name = en; run.font.size = Pt(sz); run.bold = bold; run.italic = italic
    r = run._element; rPr = r.find(qn("w:rPr"))
    if rPr is None: rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>'); r.insert(0, rPr)
    rf = rPr.find(qn("w:rFonts"))
    if rf is None: rf = parse_xml(f'<w:rFonts {nsdecls("w")}/>'); rPr.insert(0, rf)
    rf.set(qn("w:eastAsia"), cn); rf.set(qn("w:ascii"), en); rf.set(qn("w:hAnsi"), en)


# ═══════════════ 单元格边框 ═══════════════

def set_cell_border(cell, **kwargs):
    """设置单元格边框（无参数则清空边框）。"""
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    b = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for edge, val in kwargs.items():
        b.append(parse_xml(
            f'<w:{edge} {nsdecls("w")} '
            f'w:val="{val.get("val","single")}" '
            f'w:sz="{val.get("sz",4)}" w:space="0" '
            f'w:color="{val.get("color","000000")}"/>'))
    old = tcPr.find(qn("w:tcBorders"))
    if old is not None: tcPr.remove(old)
    tcPr.append(b)
