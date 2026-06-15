"""开发测试：用 docs/test_coverage.md 生成 docs/test_output.docx"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from hust_thesis import generate_thesis

base = os.path.dirname(__file__)
docs_dir = os.path.join(base, "docs")
with open(os.path.join(docs_dir, "test_coverage.md"), "r", encoding="utf-8") as f:
    md = f.read()

result = generate_thesis(
    md_content=md,
    output_path=os.path.join(docs_dir, "test_output.docx"),
    bib_path=os.path.join(docs_dir, "refs.bib"),
)
print(f"生成成功: {result}")
