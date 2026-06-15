# template_mdocx

**版本**: 0.1.0

一个基于 [Model Context Protocol (MCP)](https://modelcontextprotocol.io) 的服务，将 **扩展 Markdown** 内容转换为格式完整的 **Word (.docx)** 学术论文。内置华中科技大学（HUST）博士论文模板支持，生成符合学校格式要求的论文文档。

## 效果展示

![生成效果](docs/image.png)

## 功能特性

- **扩展 Markdown 解析** — 在标准 Markdown 基础上扩展了图片、表格、公式、交叉引用、文献引用等学术写作要素
- **自动编号** — 图题注、表题注、公式编号、章节编号均由 Word 域代码驱动，支持自动更新
- **三线表** — 自动生成学术论文标准三线表
- **公式支持** — 行内公式（`$...$`）和块公式（`$$...$$`）均使用 Word OMath 格式，可在 Word 中编辑
- **交叉引用** — `@ref{标签}` 语法引用图、表、公式，生成 Word 交叉引用域
- **文献引用** — 支持 BibTeX 格式文献库，自动生成 GB/T 7714 格式参考文献列表，引用标记为右上角上标格式
- **MCP 协议** — 通过标准 MCP stdio 传输，可集成到任何支持 MCP 的 AI 客户端（如 Claude Desktop、Cursor 等）

> **模板可定制**：标题样式映射和参考文献格式均可自由更改，方便适配其他 Word 模板。作者作为华科校友，使用华科博士论文模板作为测试基准。

## 架构

```
template_mdocx/
├── server.py               ★ MCP 服务入口，暴露三个工具
├── build.py                  开发辅助脚本，用于本地测试生成
├── pyproject.toml             uv 项目配置与依赖声明
├── docs/                     示例文件（BibTeX 文献库、Markdown 论文、效果图）
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
cd template_mdocx
uv sync
```

### 本地测试

```bash
uv run build.py
```

此命令读取 `docs/test_coverage.md` 和 `docs/refs.bib`，生成 `docs/test_output.docx`。

### 生成后操作（重要）

生成的 `.docx` 文件还需要在 Word 中完成两步手动操作：

**第一步 — 渲染公式**：按 `Alt + F11` 打开 VBA 编辑器，`Insert → Module`，粘贴以下代码并运行（`F5`）：

```vba
Sub ConvertAllEquations()
    Dim eq As OMath
    For Each eq In ActiveDocument.OMaths
        On Error Resume Next
        eq.BuildUp
        On Error GoTo 0
    Next
    MsgBox "公式转换完成。", vbInformation
End Sub
```

**第二步 — 更新域**：`Ctrl + A` 全选，然后按 `F9` 更新所有域（编号、交叉引用等）。

> 提示：如果模板 `.docx` 的 VBA 宏功能被禁用，可以将模板另存为 `.dotm` 格式并启用宏。

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

### `help_md`

返回扩展 Markdown 语法完整参考。生成论文前先调用此工具获取语法说明。

- **输入**: 无

### `example`

返回示例 BibTeX 文献库（`refs.bib`）和示例 Markdown 论文内容（`test_coverage.md`）。

- **输入**: 无

### `generate_hust_thesis`

核心工具：将扩展 Markdown 内容生成为格式化 Word 论文。

| 参数            | 类型   | 必填 | 说明                           |
| --------------- | ------ | ---- | ------------------------------ |
| `md_content`    | string | 是   | 扩展 Markdown 格式的论文内容   |
| `template_path` | string | 是   | Word 模板文件（.docx）的路径   |
| `output_path`   | string | 否   | 输出文件路径，不指定则自动生成 |
| `bib_path`      | string | 否   | BibTeX 文献库文件（.bib）路径  |

## 扩展 Markdown 语法

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

## 使用 AI 生成 Markdown 论文文档

将此提示词发送给 AI（Claude、ChatGPT 等），即可生成符合本服务解析要求的 Markdown 论文内容：

---

```
你是一名学术论文撰写助手。请根据我的要求，生成一篇用扩展 Markdown 格式编写的学术论文。

## 语法规则（必须严格遵守）

1. **标题**：使用 #、##、###、#### 表示一至四级标题。# 对应"第X章"级别。
2. **粗体/斜体**：**粗体** / *斜体*。
3. **行内公式**：$LaTeX公式$。
4. **图片**：@figure[fig:标签]{图注描述} 后紧跟 ![alt](图片路径)。
5. **表格**：@table[tbl:标签]{表格名称} 后紧跟 Markdown 表格（三线表）。
6. **块公式**：@formula[eq:标签] 后紧跟 $$...$$。
7. **交叉引用**：@ref{标签}，标签须与 @figure/@table/@formula 中声明的标签一致。
8. **文献引用**：@cite{key1,key2}，key 须存在于 BIB 文件中。
9. **参考文献列表**：@bibliography{bib文件名.bib}。

## 对应关系说明

- fig 前缀用于图片（@figure[fig:xxx]）
- tbl 前缀用于表格（@table[tbl:xxx]）
- eq 前缀用于公式（@formula[eq:xxx]）
- @ref{fig:xxx} / @ref{tbl:xxx} / @ref{eq:xxx} 分别引用上述对象

## 请按以下步骤生成

1. 先确定论文标题和章节结构
2. 逐节撰写内容，每章一级标题后表格和公式编号会重新计数
3. 确保每个 @figure/@table/@formula 声明都有唯一的标签
4. 确保 @ref 引用的标签都已声明
5. 确保 @cite 引用的文献 key 与 BIB 文件中的一致

## 参考示例

@figure[fig:arch]{模型架构图}
![architecture](images/model.png)

如@ref{fig:arch}所示，本文提出的架构包含三个模块...

@table[tbl:compare]{不同方法性能对比}
| 方法 | 准确率 | 参数量 |
| --- | --- | --- |
| ResNet | 76.1% | 25.6M |
| 本文方法 | 78.5% | 32.1M |

实验结果如@ref{tbl:compare}所示。优化目标为：

@formula[eq:opt]
$$
L = \frac{1}{N} \sum_{i=1}^{N} (y_i - \hat{y}_i)^2
$$

相关研究@cite{resnet2016,transformer2017}表明，深度学习方法效果显著。

@bibliography{refs.bib}
```

---

> 只要 MD 文档严格按照上述语法编写，一键转换即可产出高质量 Word 论文。

## 开发指南

### 项目设置

```bash
git clone <repo-url>
cd template_mdocx
uv sync
```

### Vibecoding 二次开发 — 适配你自己的模板

本项目采用**模板无关**架构，适配新模板只需三步：

**1. 创建新模板模块**

```bash
mkdir my_university
touch my_university/__init__.py
touch my_university/generator.py
touch my_university/formatters.py
```

**2. 实现生成器和格式化器**

参考 [hust_thesis/generator.py](hust_thesis/generator.py) 的实现：

```python
# my_university/generator.py
# 核心：定义样式映射 + 编排渲染逻辑
_STYLE_MAP = {
    "body": "Normal",
    "heading_1": "正文",
    "heading_2": "标题 2",
    "caption": "题注",
    "bibliography": "参考文献",
}
_TABLE_CELL_STYLE = "表格正文"
```

参考 [hust_thesis/formatters.py](hust_thesis/formatters.py) 实现文献格式：

```python
# my_university/formatters.py
class MyBibFormatter(BibFormatter):
    def format(self, entry):
        # 实现你的参考文献格式（APA / MLA / 自定义）
        ...
```

**3. 在 server.py 中注册新工具**

```python
from my_university import generate_thesis as gen_my

@server.call_tool()
async def handle_call_tool(name, arguments):
    if name == "generate_my_thesis":
        return await _gen(args, gen_my)
```

Vibecoding 开发时，可以用自然语言描述需求让 AI 完成上述代码，参考华科模板的实现作为范例。

### 构建测试

```bash
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

hust_thesis/     # HUST 模板实现（供二次开发参考）
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

MIT
