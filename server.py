"""
template_mdocx — Markdown + 模板 → Word 论文
============================================
工具: help_md / example / generate / docx_to_md
"""
import asyncio, json, logging, os
from mcp.server import Server
from mcp.server.stdio import stdio_server
import mcp.types as types

from hust_thesis import generate_thesis
from utils import docx_to_md

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("template_mdocx")
server = Server("template_mdocx")

# ── 项目路径 ──
_PROJECT_DIR = os.path.dirname(__file__)

# ── 内置模板注册表 ──
# 添加新模板时只需在此注册，无需新增工具
_BUILTIN_TEMPLATES = {
    "hust_thesis": {
        "name": "华科（HUST）博士论文模板",
        "description": "适用于华中科技大学博士/硕士/学士学位论文。标题自动编号，支持图、表、公式的自动编号与交叉引用。",
        "template_docx": os.path.join(_PROJECT_DIR, "hust_thesis", "template.docx"),
        "rules_path": os.path.join(_PROJECT_DIR, "hust_thesis", "RULES.md"),
        "example_path": os.path.join(_PROJECT_DIR, "hust_thesis", "EXAMPLE.md"),
    },
}

_GLOBAL_RULES_PATH = os.path.join(_PROJECT_DIR, "RULES.md")


# ═══════════════════ 工具列表 ═══════════════════

@server.list_tools()
async def handle_list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="help_md",
            description="获取 MCP 服务完整说明，包括内置模板列表、可用工具列表。",
            inputSchema={"type": "object", "properties": {}},
        ),
        types.Tool(
            name="example",
            description="获取指定模板的全局规则、局部规则和写作示例。通过 template 参数指定模板名称（如 hust_thesis），返回 global_rules、template_rules、example_md 等字段。",
            inputSchema={
                "type": "object",
                "properties": {
                    "template": {
                        "type": "string",
                        "description": "模板名称，当前支持: hust_thesis",
                    },
                },
                "required": ["template"],
            },
        ),
        types.Tool(
            name="generate",
            description="使用指定模板生成 Word 文档。\n"
                        "通过 template 指定模板名称，kwargs 传入模板特定的参数（如 bib_path 等）。",
            inputSchema={
                "type": "object",
                "properties": {
                    "template":   {"type": "string", "description": "模板名称，如 hust_thesis"},
                    "md_path":    {"type": "string", "description": "扩展 Markdown 文件路径"},
                    "output_path": {"type": "string", "description": "输出 .docx 路径（可选，默认与 MD 同目录）"},
                    "kwargs":     {"type": "object", "description": "模板特定参数（可选），如 {\"bib_path\": \"...\", \"template_path\": \"...\"}"},
                },
                "required": ["template", "md_path"],
            },
        ),
        types.Tool(
            name="docx_to_md",
            description="将 .docx 文档转换为 Markdown（无需格式参考，自动提取标题、表格、图片、公式、题注）。",
            inputSchema={
                "type": "object",
                "properties": {
                    "docx_path":  {"type": "string", "description": "输入 .docx 文件路径"},
                    "output_md":  {"type": "string", "description": "输出 .md 文件路径"},
                    "image_dir_name": {"type": "string", "description": "图片目录名称（相对于输出 md 所在目录，默认 images）"},
                },
                "required": ["docx_path", "output_md"],
            },
        ),
    ]


# ═══════════════════ 路由 ═══════════════════

@server.call_tool()
async def handle_call_tool(name: str, arguments: dict | None) -> list[types.TextContent]:
    args = arguments or {}
    try:
        if name == "help_md":     return await _help_md()
        elif name == "example":   return await _example(args)
        elif name == "generate":  return await _generate(args)
        elif name == "docx_to_md": return await _d2m(args)
        else: raise ValueError(f"未知工具: {name}")
    except Exception as e:
        logger.exception(f"工具 {name} 失败")
        return [types.TextContent(type="text",
                text=json.dumps({"success": False, "error": str(e)}, ensure_ascii=False))]


# ═══════════════════ 文件读取辅助 ═══════════════════

def _read_file(path: str) -> str:
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    return ""


# ═══════════════════ 处理器 ═══════════════════

async def _help_md():
    lines = ["# template_mdocx — Markdown 转 Word 文档服务\n"]
    lines.append("## 服务简介\n")
    lines.append("本服务将**扩展 Markdown** 文件生成为格式完整的 Word 文档。")
    lines.append("支持图、表、公式的自动编号与交叉引用，参考文献采用 GB/T 7714 格式。\n")
    lines.append("## 内置模板\n")
    for key, info in _BUILTIN_TEMPLATES.items():
        lines.append(f"- `{key}`：{info['name']} — {info['description']}")
    lines.append("\n## 可用工具\n")
    lines.append("- `example`：获取指定模板的规则和示例（需传 template 参数）")
    lines.append("- `generate`：生成 Word 文档（需传 template + md_path）")
    lines.append("- `docx_to_md`：将已有 .docx 转回 Markdown")
    return [types.TextContent(type="text",
            text=json.dumps({"success": True, "help": "\n".join(lines)}, ensure_ascii=False))]


async def _example(args):
    template = args.get("template", "").strip()
    if not template:
        return [types.TextContent(type="text",
                text=json.dumps({"success": False, "error": "请指定 template 参数"}, ensure_ascii=False))]
    if template not in _BUILTIN_TEMPLATES:
        supported = list(_BUILTIN_TEMPLATES.keys())
        return [types.TextContent(type="text",
                text=json.dumps({"success": False, "error": f"未知模板: {template}，当前支持: {supported}"}, ensure_ascii=False))]

    info = _BUILTIN_TEMPLATES[template]
    result = {
        "success": True,
        "template": template,
        "template_name": info["name"],
        "template_docx_path": info["template_docx"],
        # 全局规则
        "global_rules": _read_file(_GLOBAL_RULES_PATH),
        # 模板局部规则
        "template_rules": _read_file(info["rules_path"]),
        # 示例 MD
        "example_md": _read_file(info["example_path"]),
        # 示例 BIB
        "example_bib": _read_file(os.path.join(_PROJECT_DIR, "docs", "refs.bib")),
        "note": "1. 按照 example_md 中的格式编写您的 MD 文件\n"
                "2. 调用 generate 工具生成 Word\n"
                "3. template 参数传模板名称，kwargs 可传 bib_path 等",
    }
    return [types.TextContent(type="text",
            text=json.dumps(result, ensure_ascii=False))]


async def _generate(args):
    template = args.get("template", "").strip()
    mp = args.get("md_path", "").strip()

    if template not in _BUILTIN_TEMPLATES:
        supported = list(_BUILTIN_TEMPLATES.keys())
        raise ValueError(f"未知模板: {template}，当前支持: {supported}")

    if not mp or not os.path.exists(mp):
        raise FileNotFoundError(f"MD 文件不存在: {mp}")
    with open(mp, "r", encoding="utf-8") as f:
        md = f.read()

    kwargs = args.get("kwargs") or {}
    output_path = args.get("output_path") or os.path.join(
        os.path.dirname(mp),
        f"{os.path.splitext(os.path.basename(mp))[0]}.docx")

    # 根据 template 分发到具体生成器
    if template == "hust_thesis":
        tp = kwargs.get("template_path") or _BUILTIN_TEMPLATES["hust_thesis"]["template_docx"]
        bp = kwargs.get("bib_path")
        if not os.path.exists(tp):
            raise FileNotFoundError(f"模板不存在: {tp}")
        r = generate_thesis(md, output_path, template_path=tp, bib_path=bp)

    # 后续添加新模板时在此增加 elif 分支

    return [types.TextContent(type="text",
            text=json.dumps({"success": True, "output_path": r, "message": f"文档已生成: {r}"}, ensure_ascii=False))]


async def _d2m(args):
    docx_path = args["docx_path"]
    output_md = args["output_md"]
    img_dir   = args.get("image_dir_name", "images")
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"文件不存在: {docx_path}")
    r = docx_to_md(docx_path, output_md, image_dir_name=img_dir)
    return [types.TextContent(type="text",
            text=json.dumps({"success": True, "output_path": r, "message": f"转换完成: {r}"}, ensure_ascii=False))]


# ═══════════════════ 入口 ═══════════════════

async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())

if __name__ == "__main__":
    asyncio.run(main())
