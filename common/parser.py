"""MD + BIB 文本解析器"""
import re

# ── 正则 ──
RE_FIGURE  = re.compile(r'@figure\[([^\]]*)\]\{([^}]*)\}')
RE_TABLE   = re.compile(r'@table\[([^\]]*)\]\{([^}]*)\}')
RE_FORMULA = re.compile(r'@formula\[([^\]]*)\]')
RE_REF     = re.compile(r'@ref\{([^}]*)\}')
RE_BIB     = re.compile(r'@bibliography\s*(?:\{([^}]*)\})?', re.IGNORECASE)
RE_HEADING = re.compile(r'^(#{1,4})\s+(.+)$')
RE_MD_IMG  = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')
RE_CITE    = re.compile(r'@cite\{([^}]+)\}')
RE_INLINE  = re.compile(r'(\*\*[^*]+\*\*|\*[^*]+\*|\$[^$]+\$)')


def parse_md(md_text: str) -> list[dict]:
    """将 MD 文本解析为抽象语法块列表。"""
    blocks = []; lines = md_text.splitlines(); i = 0
    while i < len(lines):
        l = lines[i]

        m = RE_FIGURE.match(l.strip())
        if m:
            src = None
            if i + 1 < len(lines):
                im = RE_MD_IMG.match(lines[i + 1].strip())
                if im: src = im.group(2); i += 1
            blocks.append({"type": "figure", "label": m.group(1),
                           "caption": m.group(2), "src": src})
            i += 1; continue

        m = RE_BIB.match(l.strip())
        if m:
            bib_path_md = m.group(1) if m and m.group(1) else ""
            blocks.append({"type": "bib", "path_hint": bib_path_md})
            i += 1; continue

        m = RE_FORMULA.match(l.strip())
        if m:
            eq_label = m.group(1); i += 1
            eq = []
            while i < len(lines) and not lines[i].strip().startswith("$$"):
                i += 1
            if i < len(lines):
                if lines[i].strip() == "$$":
                    i += 1
                    while i < len(lines) and lines[i].strip() != "$$":
                        eq.append(lines[i]); i += 1
                    i += 1
                else:
                    eq = [lines[i].strip().strip("$")]; i += 1
            t = "\n".join(eq).strip()
            if t: blocks.append({"type": "formula_with_label",
                                 "label": eq_label, "text": t})
            continue

        if l.strip().startswith("$$"):
            i += 1
            while i < len(lines) and lines[i].strip() != "$$":
                i += 1
            i += 1
            blocks.append({"type": "formula_missing_label"})
            continue

        m = RE_HEADING.match(l)
        if m:
            blocks.append({"type": "heading", "level": len(m.group(1)),
                           "text": m.group(2).strip()})
            i += 1; continue

        m = RE_TABLE.match(l.strip())
        if m:
            label = m.group(1); desc = m.group(2); i += 1
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                if re.match(r"^\|[\s\-:]+(\|[\s\-:]+)*\|$", lines[i].strip()):
                    i += 1; continue
                rows.append([c.strip() for c in
                             lines[i].strip().strip("|").split("|")])
                i += 1
            blocks.append({"type": "table_with_name", "label": label,
                           "desc": desc, "rows": rows})
            continue

        if l.strip().startswith("|") and l.strip().endswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                if re.match(r"^\|[\s\-:]+(\|[\s\-:]+)*\|$", lines[i].strip()):
                    i += 1; continue
                rows.append([c.strip() for c in
                             lines[i].strip().strip("|").split("|")])
                i += 1
            blocks.append({"type": "table", "rows": rows})
            continue

        if not l.strip(): i += 1; continue

        pl = [l]
        while i + 1 < len(lines):
            n = lines[i + 1]
            if (not n.strip() or
                n.strip().startswith(("|", "#", "@figure", "@table",
                                       "@formula", "@bibliography", "$$")) or
                RE_FIGURE.match(n.strip()) or RE_TABLE.match(n.strip()) or
                RE_FORMULA.match(n.strip()) or RE_HEADING.match(n)):
                break
            pl.append(n); i += 1
        t = "\n".join(pl).strip()
        if t: blocks.append({"type": "para", "text": t})
        i += 1
    return blocks
