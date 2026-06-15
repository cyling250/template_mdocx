"""华科博士论文格式化器 — 所有 HUST 特有逻辑在此实现"""
from common.references import BibFormatter, CitationFormatter, collapse_ranges


class HustBibFormatter(BibFormatter):
    """华科博士论文参考文献 GB/T 7714 格式。"""

    def format(self, entry: dict) -> str:
        f = entry["fields"]
        authors = self._fmt_authors(f.get("author", ""))
        title = f.get("title", "")
        year = f.get("year", "")
        journal = f.get("journal", "") or f.get("booktitle", "") or f.get("school", "")
        volume = f.get("volume", "")
        number = f.get("number", "")
        pages = f.get("pages", "")
        publisher = f.get("publisher", "")

        t = entry["type"]
        if t in ("article",):
            parts = [authors, f"{title}[J]." if title else "", journal,
                     f", {volume}{f', {number}' if number else ''}" if volume else "",
                     f": {pages}" if pages else "",
                     f", {year}." if year else ""]
        elif t in ("inproceedings", "conference"):
            parts = [authors, f"{title}[C]." if title else "", journal,
                     publisher, f": {pages}" if pages else "",
                     f", {year}." if year else ""]
        elif t in ("phdthesis", "mastersthesis"):
            parts = [authors, f"{title}[D]." if title else "", journal,
                     f", {year}." if year else ""]
        elif t == "book":
            parts = [authors, f"{title}[M]." if title else "", publisher,
                     f", {year}." if year else ""]
        else:
            parts = [authors, title, journal, f", {year}." if year else ""]
        return " ".join(p for p in parts if p)


class HustCitationFormatter(CitationFormatter):
    """华科格式：右上角角标 [1-3,5]，连续用连接线。

    段落渲染器识别的标记:
      @@CT@@bib_N@@   → REF bib_N（上标，仅编号）
      @@CTX@@text@@    → 上标附属文本（- , 等）
    """
    ref_superscript = True

    def format_keys(self, keys: list[str], cite_map: dict[str, str]) -> str:
        nums = []
        for k in keys:
            bm = cite_map.get(k)
            if bm and bm.startswith("bib_"):
                nums.append(int(bm[4:]))
        if not nums:
            return "[?]"

        ranges = collapse_ranges(nums)
        parts = []
        for r in ranges:
            if isinstance(r, tuple):
                parts.append(
                    f"@@CT@@bib_{r[0]}@@@@CTX@@-@@@@CT@@bib_{r[1]}@@")
            else:
                parts.append(f"@@CT@@bib_{r}@@")
        return "@@CTX@@[@@" + "@@CTX@@,@@".join(parts) + "@@CTX@@]@@"
