"""
template_mdocx — Markdown + 模板 → Word 论文
============================================
工具: help_md / example / generate_hust_thesis
"""
import asyncio, json, logging, os
from mcp.server import Server
from mcp.server.stdio import stdio_server
import mcp.types as types

from hust_thesis import generate_thesis

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("template_mdocx")
server = Server("template_mdocx")


# ═══════════════════ 工具列表 ═══════════════════

@server.list_tools()
async def handle_list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="help_md",
            description="返回扩展 Markdown 语法完整参考。生成论文前先调用此工具获取语法说明。",
            inputSchema={"type": "object", "properties": {}},
        ),
        types.Tool(
            name="example",
            description="返回示例 BIB 文献库和示例 MD 论文内容。",
            inputSchema={"type": "object", "properties": {}},
        ),
        types.Tool(
            name="generate_hust_thesis",
            description="使用 HUST 博士论文模板生成论文。\n"
                        "- #/## 标题  - **粗体**/*斜体*  - $行内公式$\n"
                        "- @figure[标签]{描述} → 图 + 题注\n"
                        "- @table[标签]{表名} → 三线表 + 题注\n"
                        "- @formula[标签] → 块公式 + 编号\n"
                        "- @ref{标签} → 交叉引用\n"
                        "- @cite{key1,key2} → 文献引用\n"
                        "- @bibliography{refs.bib} → 参考文献列表",
            inputSchema={
                "type": "object",
                "properties": {
                    "md_content":    {"type": "string", "description": "扩展 Markdown 内容"},
                    "template_path": {"type": "string", "description": "模板 .docx 路径"},
                    "output_path":   {"type": "string", "description": "输出路径（可选）"},
                    "bib_path":      {"type": "string", "description": "BIB 文件路径（可选）"},
                },
                "required": ["md_content", "template_path"],
            },
        ),
    ]


# ═══════════════════ 路由 ═══════════════════

@server.call_tool()
async def handle_call_tool(name: str, arguments: dict | None) -> list[types.TextContent]:
    if not arguments:
        raise ValueError("缺少参数")
    try:
        if name == "help_md":                return await _help_md()
        elif name == "example":              return await _example()
        elif name == "generate_hust_thesis": return await _gen(arguments)
        else: raise ValueError(f"未知工具: {name}")
    except Exception as e:
        logger.exception(f"工具 {name} 失败")
        return [types.TextContent(type="text",
                text=json.dumps({"success": False, "error": str(e)}, ensure_ascii=False))]


# ═══════════════════ 处理器 ═══════════════════

async def _help_md():
    help_text = (
        "# 扩展 Markdown 语法参考\n\n"
        "## 标题\n"
        "使用 # ## ### #### 表示一至四级标题。\n\n"
        "## 行内格式\n"
        "- **粗体**  `**粗体**`\n"
        "- *斜体*    `*斜体*`\n"
        "- 行内公式  `$E=mc^2$`\n\n"
        "## 图片\n"
        "@figure[标签]{图片描述文字}\n"
        "![alt](图片路径)\n\n"
        "## 表格\n"
        "@table[标签]{表格名称}\n"
        "| 列1 | 列2 | 列3 |\n"
        "| --- | --- | --- |\n"
        "| 数据 | 数据 | 数据 |\n\n"
        "## 块公式\n"
        "@formula[标签]\n"
        "$$\n"
        "公式内容\n"
        "$$\n\n"
        "## 交叉引用\n"
        "- @ref{标签}  引用图/表/公式\n"
        "- @cite{key}  引用文献，多个用逗号分隔\n\n"
        "## 文献引用\n"
        "- @cite{key1,key2}  正文中引用文献\n"
        "- @bibliography{refs.bib}  在文末插入参考文献列表\n\n"
        "## 规则\n"
        "1. 每个表格和图片必须有 @table/@figure 声明（含标签和名称）\n"
        "2. 每个块公式必须有 @formula 声明\n"
        "3. 交叉引用 @ref{标签} 的标签必须与声明的标签一致\n"
        "4. 文献引用 @cite{key} 的 key 必须存在于 BIB 文件中"
    )
    return [types.TextContent(type="text",
            text=json.dumps({"success": True, "help": help_text}, ensure_ascii=False))]


async def _example():
    base = os.path.dirname(__file__)
    def _read(name):
        p = os.path.join(base, "docs", name)
        return open(p, encoding="utf-8").read() if os.path.exists(p) else f"({name} 不存在)"
    return [types.TextContent(type="text", text=json.dumps({
        "success": True,
        "example_bib": _read("refs.bib"),
        "example_md":  _read("test_coverage.md"),
    }, ensure_ascii=False))]


async def _gen(args):
    md = args["md_content"]; tp = args["template_path"]
    if not md: raise ValueError("md_content 为空")
    if not tp or not os.path.exists(tp): raise FileNotFoundError(f"模板不存在: {tp}")
    op = args.get("output_path") or os.path.join(
        os.path.dirname(tp),
        f"{os.path.splitext(os.path.basename(tp))[0]}_thesis.docx")
    bp = args.get("bib_path")
    r = generate_thesis(md, op, template_path=tp, bib_path=bp)
    return [types.TextContent(type="text",
            text=json.dumps({"success": True, "output_path": r, "message": f"论文已生成: {r}"}, ensure_ascii=False))]


# ═══════════════════ 入口 ═══════════════════

async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())

if __name__ == "__main__":
    asyncio.run(main())
