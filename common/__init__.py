"""
common — 模板无关的 Word 格式化工具集
======================================
按 Word 排版难点组织：
  figures    — 图片 + 题注
  tables     — 三线表 + 单元格样式 + 题注
  formulas   — 块公式 + 行内公式 + 编号
  references — BIB 解析 → 多标准格式化 + 书目列表
  crossref   — 图/表/公式/文献 交叉引用
  utils      — XML/OMath/字体/边框 辅助
  parser     — MD + BIB 文本解析
"""
from .utils import (
    MATH_NS, W_NS, next_bookmark_id,
    set_para_style, apply_font, set_cell_border,
    add_omath_inline, add_omath_display,
)
from .crossref import (
    make_complex_field, add_seq_numbering, add_ref_field, add_caption,
)
from .figures import make_image_block
from .tables import make_three_line_table
from .formulas import make_formula_block
from .references import parse_bib, BibFormatter, CitationFormatter, render_bibliography, collapse_ranges
from .parser import parse_md, RE_FIGURE, RE_TABLE, RE_FORMULA, RE_REF, RE_BIB, \
    RE_HEADING, RE_MD_IMG, RE_CITE, RE_INLINE
