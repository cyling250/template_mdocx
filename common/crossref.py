"""交叉引用：WORD 字段 + 题注编号"""
from lxml import etree
from .utils import W_NS, next_bookmark_id


def make_complex_field(paragraph, parts, vert_align=None):
    """创建复杂字段。vert_align='superscript' → 上标。"""
    for typ, val in parts:
        r = etree.SubElement(paragraph._element, f'{{{W_NS}}}r')
        rPr = etree.SubElement(r, f'{{{W_NS}}}rPr')
        if vert_align:
            etree.SubElement(rPr, f'{{{W_NS}}}vertAlign').set(f'{{{W_NS}}}val', vert_align)
        if typ == 'fldChar':
            etree.SubElement(r, f'{{{W_NS}}}fldChar').set(
                f'{{{W_NS}}}fldCharType', val)
        elif typ == 'instr':
            it = etree.SubElement(r, f'{{{W_NS}}}instrText')
            it.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            it.text = val
        elif typ == 'text':
            etree.SubElement(rPr, f'{{{W_NS}}}noProof')
            t = etree.SubElement(r, f'{{{W_NS}}}t')
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = val


def add_seq_numbering(paragraph, caption_type, styleref_level=1):
    """在段落中插入 STYLEREF + SEQ 自动编号。"""
    # 如: '表 STYLEREF 1 \s - SEQ 表 \* ARABIC \s 1'
    etree.SubElement(etree.SubElement(paragraph._element, f'{{{W_NS}}}r'),
                     f'{{{W_NS}}}t').text = caption_type
    etree.SubElement(etree.SubElement(paragraph._element, f'{{{W_NS}}}r'),
                     f'{{{W_NS}}}t').text = ' '
    make_complex_field(paragraph, [
        ('fldChar', 'begin'),
        ('instr', f' STYLEREF {styleref_level} \\s '),
        ('fldChar', 'separate'), ('text', '1'), ('fldChar', 'end'),
    ])
    etree.SubElement(etree.SubElement(paragraph._element, f'{{{W_NS}}}r'),
                     f'{{{W_NS}}}t').text = '\u2013'
    make_complex_field(paragraph, [
        ('fldChar', 'begin'),
        ('instr', f' SEQ {caption_type} \\* ARABIC \\s 1 '),
        ('fldChar', 'separate'), ('text', '1'), ('fldChar', 'end'),
    ])


def add_ref_field(paragraph, bookmark_name, number_only=False, superscript=False):
    """插入 REF 字段。number_only: \n 取段落编号；\# "0" 格式化为纯数字（剥离方括号等）。"""
    switches = r'\n \# "0" \h' if number_only else r'\h'
    make_complex_field(paragraph, [
        ('fldChar', 'begin'),
        ('instr', f' REF {bookmark_name} {switches} '),
        ('fldChar', 'separate'), ('text', '?'), ('fldChar', 'end'),
    ], vert_align='superscript' if superscript else None)


def add_caption(doc, caption_type, bookmark_name, desc="",
                caption_style_name="Caption", styleref_level=1):
    """插入题注段落：书签仅包围编号，描述文字在外。

    Args:
        caption_type: "图" / "表"
        bookmark_name: 书签名
        desc: 题注描述文字
        caption_style_name: 题注样式名
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    p = doc.add_paragraph()
    try:
        p.style = doc.styles[caption_style_name]
    except (KeyError, AttributeError):
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    bm_id = next_bookmark_id()
    bm_s = etree.SubElement(p._element, f'{{{W_NS}}}bookmarkStart')
    bm_s.set(f'{{{W_NS}}}id', str(bm_id))
    bm_s.set(f'{{{W_NS}}}name', bookmark_name)
    add_seq_numbering(p, caption_type, styleref_level)
    bm_e = etree.SubElement(p._element, f'{{{W_NS}}}bookmarkEnd')
    bm_e.set(f'{{{W_NS}}}id', str(bm_id))

    if desc:
        r = etree.SubElement(p._element, f'{{{W_NS}}}r')
        rPr = etree.SubElement(r, f'{{{W_NS}}}rPr')
        rf = etree.SubElement(rPr, f'{{{W_NS}}}rFonts')
        rf.set(f'{{{W_NS}}}eastAsia', '宋体')
        etree.SubElement(r, f'{{{W_NS}}}t').text = f' {desc}'
    return p
