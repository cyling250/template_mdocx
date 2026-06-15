"""参考文献处理：BIB 解析 → 格式化器接口 + 书目列表渲染"""
import os, re
from lxml import etree
from docx.shared import Pt, Cm
from .utils import W_NS, next_bookmark_id, apply_font


# ═══════════════ BIB 解析 ═══════════════

def parse_bib(path: str) -> list[dict]:
    """解析 .bib 文件，返回条目列表 [{type, key, fields}]。"""
    if not path or not os.path.exists(path):
        return []
    with open(path, 'r', encoding='utf-8') as f:
        text = f.read()

    entries = []
    pattern = re.compile(r'@(\w+)\{([^,]+),\s*((?:(?!\n\s*@).)*)', re.DOTALL)
    for m in pattern.finditer(text):
        entry_type = m.group(1)
        cite_key = m.group(2).strip()
        body = m.group(3)
        fields = {}
        for fm in re.finditer(r'(\w+)\s*=\s*\{([^}]*)\}', body):
            fields[fm.group(1).lower()] = fm.group(2).strip()
        entries.append({"type": entry_type, "key": cite_key, "fields": fields})
    return entries


# ═══════════════ 格式化器接口 ═══════════════

def collapse_ranges(nums: list[int]) -> list:
    """工具函数：折叠连续数字。 [1,2,3,5,8] → [(1,3), 5, 8]"""
    if not nums: return []
    nums = sorted(nums)
    result = []; start = prev = nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
        else:
            result.append((start, prev) if start != prev else start)
            start = prev = n
    result.append((start, prev) if start != prev else start)
    return result


class BibFormatter:
    """BIB 条目格式化接口。在具体的 thesis 包中实现（如 hust_thesis/formatters.py）。

    Usage:
        class MyFormatter(BibFormatter):
            def format(self, entry): return "作者. 标题[J]. 期刊, 年."
    """

    def format(self, entry: dict) -> str:
        raise NotImplementedError

    @staticmethod
    def _fmt_authors(authors_str: str, max_display: int = 3, et_al: str = "等") -> str:
        """作者格式化辅助：LastName, First → First Last。"""
        if not authors_str: return ""
        alist = [a.strip() for a in authors_str.replace("\n", " ").split(" and ")]
        fmt = []
        for a in alist:
            if not a: continue
            parts = a.split(",")
            fmt.append(parts[1].strip() + " " + parts[0].strip() if len(parts) >= 2 else a.strip())
        return ", ".join(fmt) if len(fmt) <= max_display else fmt[0] + " " + et_al


class CitationFormatter:
    """正文引用格式化器接口。在具体的 thesis 包中实现。

    段落渲染器识别的标记（由 format_keys 返回）:
      @@CT@@bookmark@@   → REF 字段（仅编号）
      @@CTX@@text@@       → 引用附属文本
      其他纯文本           → 普通 run

    属性:
      ref_superscript: REF 字段是否上标（华科: True, 作者-年份: False）
    """
    ref_superscript = False

    def format_keys(self, keys: list[str], cite_map: dict[str, str]) -> str:
        """将 @cite{key1,key2} 转为段落渲染标记字符串。"""
        raise NotImplementedError


# ═══════════════ 书目列表渲染（通用） ═══════════════

def render_bibliography(doc, entries: list[dict],
                        bibliography_style_name: str = "Bibliography",
                        formatter: BibFormatter | None = None):
    """将 BIB 条目列表渲染为书目段落（Bibliography 样式 + 书签）。

    Args:
        doc: python-docx Document
        entries: parse_bib 返回的条目列表
        bibliography_style_name: 书目样式名
        formatter: BibFormatter 实例
    Returns:
        cite_map: {cite_key: bookmark_name} 用于正文交叉引用
    """
    cite_map = {}
    for idx, entry in enumerate(entries, 1):
        bib_bm = f"bib_{idx}"
        cite_map[entry["key"]] = bib_bm

        formatted = formatter.format(entry) if formatter else entry["key"]
        bp = doc.add_paragraph()
        bp.paragraph_format.first_line_indent = Pt(0)
        bp.paragraph_format.left_indent = Cm(1.0)
        bp.paragraph_format.hanging_indent = Cm(1.0)
        try:
            bp.style = doc.styles[bibliography_style_name]
        except (KeyError, AttributeError):
            pass

        # 书签包围格式化文字；编号由 Bibliography 样式自动生成，引用用 REF \n 获取编号
        bm_id = next_bookmark_id()
        bm_s = etree.SubElement(bp._element, f'{{{W_NS}}}bookmarkStart')
        bm_s.set(f'{{{W_NS}}}id', str(bm_id))
        bm_s.set(f'{{{W_NS}}}name', bib_bm)
        r_txt = bp.add_run(formatted)
        apply_font(r_txt, sz=10.5)
        bm_e = etree.SubElement(bp._element, f'{{{W_NS}}}bookmarkEnd')
        bm_e.set(f'{{{W_NS}}}id', str(bm_id))

    return cite_map
