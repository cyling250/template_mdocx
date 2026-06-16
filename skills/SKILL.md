---
name: "template-mdocx"
description: "使用 template_mdocx MCP 服务将扩展 Markdown 转换为格式完整的 Word 文档，或将已有 .docx 转换回 Markdown。适用于学术论文、报告、计算书等场景。"
---

# Template MDocx — Markdown 转 Word 文档工具

本 skill 通过 **template_mdocx** MCP 服务，将扩展 Markdown 生成为格式完整的 Word 文档。

## MCP 工具

| 工具 | 说明 |
|------|------|
| `mcp_template_mdocx_help_md` | 获取完整的扩展 Markdown 语法参考 |
| `mcp_template_mdocx_example` | 获取示例 `.bib` 和 Markdown 内容 |
| `mcp_template_mdocx_generate_hust_thesis` | 根据 Markdown + 模板生成 Word 文档 |
| `mcp_template_mdocx_docx_to_md` | 将 `.docx` 转换回 Markdown |

### 工具参数

**generate_hust_thesis（生成文档）：**
- `md_path`（必填）— 扩展 Markdown 文件路径
- `template_path`（必填）— 模板 `.docx` 文件路径
- `output_path`（可选）— 输出 `.docx` 路径
- `bib_path`（可选）— `.bib` 参考文献文件路径

**docx_to_md（转换文档）：**
- `docx_path`（必填）— 输入 `.docx` 文件路径
- `output_md`（必填）— 输出 `.md` 文件路径
- `image_dir_name`（可选）— 图片目录名（默认 `images`）

## 标题规则

不同模板的标题规则见对应文件：

| 模板 | 标题规则文件 |
|------|-------------|
| `hust_thesis` | `hust_thesis/TITLE_RULES.md`（MCP 项目内路径） |

## 扩展 Markdown 语法

### 行内格式
```
**粗体** *斜体* $行内公式$
```

### 图片
```
@figure[fig:标签]{图注描述}
![替代文字](图片路径.png)
```
- `@figure` 必须紧接在 `![alt](path)` 之前
- 标签前缀使用 `fig:`

### 表格
```
@table[tbl:标签]{表名}
| 列1 | 列2 | 列3 |
| --- | --- | --- |
| 内容 | 内容 | 内容 |
```
- `@table` 必须紧接在 Markdown 表格之前
- 标签前缀使用 `tbl:`，生成三线表格式

### 块公式
```
@formula[eq:标签]
$$
公式内容
$$
```
- `@formula` 必须紧接在 `$$...$$` 之前
- 标签前缀使用 `eq:`，公式自动编号

### 交叉引用
```
如 @ref{fig:架构图} 所示……结果见 @ref{tbl:对比}……目标函数见 @ref{eq:损失函数}。
```
- `@ref{标签}` 的标签必须与声明的标签一致

### 文献引用
```
相关研究 @cite{resnet2016,transformer2017} 表明……
```
- `@cite{key1,key2}` 的 key 必须存在于 `.bib` 文件中
- 多个引用用逗号分隔，生成 `[1-3,5]` 风格上标引用

### 参考文献列表
```
@bibliography{refs.bib}
```
- 放在文档末尾，采用 GB/T 7714 格式

## 四大规则

1. **每个图/表必须有** `@figure[标签]{描述}` / `@table[标签]{表名}` 声明紧接在前
2. **每个块公式必须有** `@formula[标签]` 声明紧接在 `$$...$$` 之前
3. **`@ref{标签}` 的标签必须** 与声明的标签完全一致
4. **`@cite{key}` 的 key 必须** 存在于 `.bib` 文件中

## 标签前缀约定

- `fig:` — 图片
- `tbl:` — 表格
- `eq:` — 公式

## 生成后的操作（在 Word 中）

1. **公式渲染**：`Alt+F11` → 插入 → 模块，粘贴 VBA 宏并运行 `F5`
2. **刷新编号**：`Ctrl+A` 然后 `F9`，更新所有交叉引用和编号
