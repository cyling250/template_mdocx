---
name: "template-mdocx"
description: "Write Markdown, generate formatted Word documents via the template_mdocx MCP server. Invoke when user needs to create or edit academic papers, theses, reports, or any formatted Word document from Markdown content. MCP server located at E:\\mcp-server\\template_mdocx."
---

# Template MDocx — Markdown to Word Document Generator

This skill uses the **template_mdocx** MCP server to convert extended Markdown into formatted Word documents (`.docx`), or convert existing `.docx` files back to Markdown.

## MCP Tools

| Tool                                      | Description                                             |
| ----------------------------------------- | ------------------------------------------------------- |
| `mcp_template_mdocx_help_md`              | Get the full extended Markdown syntax reference         |
| `mcp_template_mdocx_example`              | Get example `.bib` and Markdown content                 |
| `mcp_template_mdocx_generate_hust_thesis` | Generate Word document from Markdown + template `.docx` |
| `mcp_template_mdocx_docx_to_md`           | Convert `.docx` back to Markdown                        |

### Tool Parameters

**generate_hust_thesis:**
- `md_path` (required) — Path to the extended Markdown file
- `template_path` (required) — Path to the template `.docx` file
- `output_path` (optional) — Output `.docx` path
- `bib_path` (optional) — Path to `.bib` bibliography file

**docx_to_md:**

- `docx_path` (required) — Input `.docx` file path
- `output_md` (required) — Output `.md` file path
- `image_dir_name` (optional) — Image directory name (default: `images`)

## Extended Markdown Syntax

### Headings

```
# 一级标题（第 X 章）
## 二级标题
### 三级标题
#### 四级标题
```

### Inline Formatting

```
**粗体** *斜体* $行内公式$
```

### Figures

```
@figure[fig:标签]{图注描述}
![替代文字](图片路径.png)
```

- `@figure` 必须紧接在 `![alt](path)` 之前
- 标签前缀使用 `fig:`

### Tables

```
@table[tbl:标签]{表名}
| 列1 | 列2 | 列3 |
| --- | --- | --- |
| 内容 | 内容 | 内容 |
```

- `@table` 必须紧接在 Markdown 表格之前
- 标签前缀使用 `tbl:`
- 生成三线表格式

### Block Formulas

```
@formula[eq:标签]
$$
\mathcal{L} = \frac{1}{N} \sum (y - \hat{y})^2
$$
```

- `@formula` 必须紧接在 `$$...$$` 之前
- 标签前缀使用 `eq:`
- 公式自动编号

### Cross-References

```
如 @ref{fig:架构图} 所示……结果见 @ref{tbl:对比}……目标函数见 @ref{eq:损失函数}。
```

- `@ref{标签}` 的标签必须与声明的标签一致

### Citation References

```
相关研究 @cite{resnet2016,transformer2017} 表明……
```

- `@cite{key1,key2}` 的 key 必须存在于 `.bib` 文件中
- 多个引用用逗号分隔
- 生成 `[1-3,5]` 风格上标引用

### Bibliography

```
@bibliography{refs.bib}
```

- 放在文档末尾
- 引用列表采用 GB/T 7714 格式

## Four Golden Rules

1. **每个图/表必须有** `@figure[标签]{描述}` / `@table[标签]{表名}` 声明紧接在前
2. **每个块公式必须有** `@formula[标签]` 声明紧接在 `$$...$$` 之前
3. **`@ref{标签}` 的标签必须** 与声明的标签完全一致
4. **`@cite{key}` 的 key 必须** 存在于 `.bib` 文件中

## Tag Prefix Convention

- `fig:` — figures/images
- `tbl:` — tables
- `eq:` — formulas/equations

## Post-Generation Steps (in Word)

After opening the generated `.docx` in Word:

1. **Formula rendering**: Press `Alt+F11` → Insert → Module, paste the VBA macro and run `F5`
2. **Refresh numbering**: `Ctrl+A` then `F9` to update all cross-references and numbering
