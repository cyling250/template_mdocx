# template_mdocx

**版本**: 0.1.0

一个基于 [Model Context Protocol (MCP)](https://modelcontextprotocol.io) 的服务，将 **扩展 Markdown** 内容转换为格式完整的 **Word (.docx)** 学术论文。内置华中科技大学（HUST）博士论文模板支持，生成符合学校格式要求的论文文档。

## 功能特性

- **扩展 Markdown 解析** — 在标准 Markdown 基础上扩展了图片、表格、公式、交叉引用、文献引用等学术写作要素
- **自动编号** — 图题注、表题注、公式编号、章节编号均由 Word 域代码驱动，支持自动更新
- **三线表** — 自动生成学术论文标准三线表
- **公式支持** — 行内公式（`$...$`）和块公式（`$$...$$`）均使用 Word OMath 格式，可在 Word 中编辑
- **交叉引用** — `@ref{标签}` 语法引用图、表、公式，生成 Word 交叉引用域
- **文献引用** — 支持 BibTeX 格式文献库，自动生成 GB/T 7714 格式参考文献列表，引用标记为右上角上标格式
- **MCP 协议** — 通过标准 MCP stdio 传输，可集成到任何支持 MCP 的 AI 客户端（如 Claude Desktop、Cursor 等）

## 架构

```
template_mdocx/
├── server.py               ★ MCP 服务入口，暴露三个工具
├── build.py                  开发辅助脚本，用于本地测试生成
├── pyproject.toml             uv 项目配置与依赖声明
├── docs/                      示例文件（BibTeX 文献库 & Markdown 论文）
│
├── common/                   模板无关的 Word 格式化工具集
│   ├── parser.py             MD + BIB 文本解析器
│   ├── utils.py              XML、OMath、字体、边框辅助函数
│   ├── crossref.py           Word 域代码（STYLEREF、SEQ、REF）、书签、题注
│   ├── figures.py            图片嵌入 + 题注
│   ├── tables.py             三线表生成 + 题注
│   ├── formulas.py           块公式（三列无边框表格）+ 编号
│   └── references.py         BibTeX 解析、格式化器接口、参考文献渲染
│
└── hust_thesis/              HUST 特定论文模块
    ├── generator.py          论文生成主流程（解析 → 校验 → 渲染 → 保存）
    ├── formatters.py         文献格式化器（GB/T 7714）和引用格式化器（上标）
    └── template.docx         默认 HUST 博士论文 Word 模板
```

## 快速开始

### 环境要求

- Python >= 3.12
- [uv](https://docs.astral.sh/uv/)（包管理器）

### 安装

```bash
# 克隆或进入项目目录
cd template_mdocx

# uv 自动创建虚拟环境并安装依赖
uv sync
```

### 本地测试

```bash
uv run build.py
```

此命令读取 `docs/test_coverage.md` 和 `docs/refs.bib`，生成 `docs/test_output.docx`。

## MCP 集成

本服务使用 **stdio 传输**，可集成到任何支持 MCP 的客户端中。

### Claude Desktop 配置

在 Claude Desktop 的 MCP 配置文件中添加：

**Windows (`%AppData%\Claude\claude_desktop_config.json`)**:

```json
{
  "mcpServers": {
    "template_mdocx": {
      "command": "uv",
      "args": ["run", "--directory", "F:\\template_mdocx", "server.py"]
    }
  }
}
```

**macOS / Linux (`~/.config/Claude/claude_desktop_config.json`)**:

```json
{
  "mcpServers": {
    "template_mdocx": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/template_mdocx", "server.py"]
    }
  }
}
```

### Cursor 配置

在 Cursor 的 MCP 配置中添加：

```json
{
  "mcpServers": {
    "template_mdocx": {
      "command": "uv",
      "args": ["run", "--directory", "F:\\template_mdocx", "server.py"]
    }
  }
}
```

## MCP 工具说明

本服务暴露三个 MCP 工具：

### 1. `help_md`

返回扩展 Markdown 语法完整参考。

- **输入**: 无
- **输出**: JSON 字符串，包含 `success` 和 `help` 字段

**调用示例**:

```json
{
  "success": true,
  "help": "# 扩展 Markdown 语法参考\n\n..."
}
```

### 2. `example`

返回示例 BibTeX 文献库（`refs.bib`）和示例 Markdown 论文内容（`test_coverage.md`）。

- **输入**: 无
- **输出**: JSON 字符串，包含 `success`、`example_bib` 和 `example_md` 字段

**调用示例**:

```json
{
  "success": true,
  "example_bib": "@article{...}\n...",
  "example_md": "# 论文标题\n\n..."
}
```

### 3. `generate_hust_thesis`

核心工具：将扩展 Markdown 内容生成为格式化 Word 论文。

| 参数            | 类型   | 必填 | 说明                           |
| --------------- | ------ | ---- | ------------------------------ |
| `md_content`    | string | 是   | 扩展 Markdown 格式的论文内容   |
| `template_path` | string | 是   | Word 模板文件（.docx）的路径   |
| `output_path`   | string | 否   | 输出文件路径，不指定则自动生成 |
| `bib_path`      | string | 否   | BibTeX 文献库文件（.bib）路径  |

**输出**: JSON 字符串，包含 `success`、`output_path` 和 `message` 字段。

**调用示例**:

```json
{
  "success": true,
  "output_path": "F:\\template\\template_thesis.docx",
  "message": "论文已生成: F:\\template\\template_thesis.docx"
}
```

## 扩展 Markdown 语法

本服务使用一套扩展 Markdown 语法，在标准 Markdown 基础上增加了学术论文所需的元素。

### 标题

```
# 一级标题（通常对应论文第 X 章）
## 二级标题
### 三级标题
#### 四级标题
```

### 行内格式

| 语法           | 效果                   |
| -------------- | ---------------------- |
| `**粗体文字**` | **粗体文字**           |
| `*斜体文字*`   | _斜体文字_             |
| `$E=mc^2$`     | 行内公式（Word OMath） |

### 图片

```
@figure[fig:标签]{图片描述文字}
![替代文本](图片路径)
```

- `@figure` 声明图片标签和题注文字
- 紧随其后的 `![...](...)` 指定图片路径
- 支持相对路径和绝对路径

### 表格

```
@table[tbl:标签]{表格名称}
| 列1 | 列2 | 列3 |
| --- | --- | --- |
| 数据 | 数据 | 数据 |
```

- `@table` 声明表格标签和名称
- 紧随其后的 Markdown 表格将被转换为三线表
- 表格名称自动编号为"表 X-Y"

### 块公式

```
@formula[eq:标签]
$$
公式内容
$$
```

- `@formula` 声明公式标签
- `$$...$$` 包围 LaTeX 公式内容
- 公式自动编号为"(X-Y)"，右对齐

### 交叉引用

```
@ref{fig:架构图}
@ref{tbl:对比表}
@ref{eq:损失函数}
```

- `@ref{标签}` 引用已声明的图、表或公式
- 生成 Word 交叉引用域，可自动更新编号

### 文献引用

```
正文中引用 @cite{key1,key2}
```

- `@cite{key1,key2}` 在正文中引用文献
- 多个文献用逗号分隔，自动格式化为 `[1-3,5]` 风格
- 引用显示为右上角上标

### 参考文献列表

```
@bibliography{refs.bib}
```

- 在文末插入参考文献列表
- 参数为 BibTeX 文件路径
- 自动生成 GB/T 7714 格式的参考文献列表

### 语法规则

1. 每个表格和图片必须有 `@table` / `@figure` 声明（含标签和名称）
2. 每个块公式必须有 `@formula` 声明
3. 交叉引用 `@ref{标签}` 的标签必须与声明的标签一致
4. 文献引用 `@cite{key}` 的 key 必须存在于 BIB 文件中

## 开发指南

### 项目设置

```bash
# 克隆仓库
git clone <repo-url>
cd template_mdocx

# uv 自动创建虚拟环境并安装所有依赖
uv sync
```

### 添加新模板

1. 在项目根目录创建新模块目录，如 `my_university/`
2. 参考 `hust_thesis/generator.py` 实现生成器
3. 实现 `formatters.py` 中的文献格式化接口
4. 在 `server.py` 中注册新工具

### 扩展语法解析

编辑 `common/parser.py`，添加新的正则表达式和 AST 节点类型：

```python
# 添加新的正则模式
RE_NEW_ELEMENT = re.compile(r'@new\[label\]{content}')

# 在 parse_md 函数中添加处理逻辑
```

### 构建测试

```bash
# build.py 是本地测试入口
uv run build.py
```

### 代码结构

```
common/          # 核心库，模板无关
  ├── parser.py       — 文本解析层
  ├── figures.py      — 图片处理
  ├── tables.py       — 表格处理
  ├── formulas.py     — 公式处理
  ├── references.py   — 文献处理
  ├── crossref.py     — 交叉引用
  └── utils.py        — 通用工具

hust_thesis/     # HUST 模板实现
  ├── generator.py    — 论文生成主流程
  ├── formatters.py   — 文献格式化
  └── template.docx   — Word 模板文件
```

## 依赖

| 包                                                   | 版本     | 用途                            |
| ---------------------------------------------------- | -------- | ------------------------------- |
| [mcp](https://pypi.org/project/mcp/)                 | >= 1.0.0 | Model Context Protocol 服务框架 |
| [python-docx](https://pypi.org/project/python-docx/) | >= 1.1.0 | Word 文档读写                   |
| [lxml](https://pypi.org/project/lxml/)               | >= 5.0.0 | XML 处理（python-docx 依赖）    |

## 许可证

[MIT](LICENSE) (请根据实际情况替换)
