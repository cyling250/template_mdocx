# template_mdocx · 写 Markdown，出 Word

> **让 template_mdocx 把你从重复而繁杂的排版中解放出来，你只管专注于内容。**

把写作变成聊天——你对 AI 说人话，AI 输出 Markdown，本工具一键导出为格式完美的 Word 文档。

**版本**: 0.1.0

### 不止是论文

本 MCP 服务的核心能力是 **Markdown + Word 模板 → 格式化文档**。虽然当前以华科博士论文作为测试基准，但只要你的 Word 模板样式命名清晰、排版不过于复杂（非双栏、非复杂图文混排），几乎任何规范性文档都能无缝适配：

| 文档类型        | 关键元素                                       | 适用度 |
| --------------- | ---------------------------------------------- | ------ |
| 学位论文        | 章节标题、三线表、公式编号、交叉引用、参考文献 | 完美   |
| 项目报告        | 标题层级、图表、公式、引用                     | 完美   |
| 公文 / 红头文件 | 标题、段落、编号、落款                         | 良好   |
| 技术手册        | 多级标题、表格、代码块                         | 良好   |
| 申报材料        | 表格、段落、编号列表                           | 良好   |
| 规章制度        | 条款编号、层级标题                             | 良好   |

> 只要你的 Word 模板有规范样式名，MCP 就能匹配上。换个模板就是换种文档类型。

### 诚邀 Fork 共建

学术论文只是冰山一角。如果你有项目报告模板、公文模板、申报材料模板……欢迎 **Fork 本项目**，参考 `hust_thesis/` 的设计范式，用 **Vibecoding** 填充属于你自己领域的新模板。把话告诉 AI，让它帮你适配——就跟配 MCP 服务一样简单。大家一起把"写 Markdown 出 Word"这件事覆盖到更多场景。

---

## 效果一览

![生成效果](docs/image.png)

---

## 不会配环境？把这段话丢给你的 AI

你不需要自己动手安装和配置。直接打开 **Cursor / Claude Desktop / 任何支持 MCP 的 AI 工具**，把下面这段话粘贴给它，它会帮你自动完成下载、安装、配置全部流程：

---

> 帮我安装 template_mdocx 这个 MCP 服务，仓库地址是 https://github.com/cyling250/template_mdocx。请按以下步骤操作：\
>
> 1. 把仓库 clone 到本地\
> 2. 确保电脑装了 Python 3.12+ 和 uv 包管理器\
> 3. 进入项目目录，运行 `uv sync` 安装依赖\
> 4. 把这个 MCP 服务配置到我当前使用的 AI 客户端里（如果是 Cursor 就配到 Cursor，如果是 Claude Desktop 就配到 Claude Desktop），配置内容如下：\
>
> ````json\
> {\
>   "mcpServers": {\
>     "template_mdocx": {\
>       "command": "uv",\
>       "args": ["run", "--directory", "项目所在路径", "server.py"]\
>     }\
>   }\
> }\
> ```\
> 5. 安装完成后，跟我说一声，并告诉我接下来怎么用它写论文。
> ````

---

就这一句话，AI 帮你配好一切。配完之后你就能直接对话生成论文了。

---

## 它能做什么？

| 你做的事                   | 效果                     |
| -------------------------- | ------------------------ |
| 跟 AI 聊天，写出论文内容   | —                        |
| AI 按约定格式输出 Markdown | —                        |
| 一键调用本工具             | 输出格式完整的 Word 论文 |
| 在 Word 中运行一个宏       | 所有公式渲染完成         |
| `Ctrl+A` 然后 `F9`         | 编号、引用全部更新       |

生成的 Word 包含：**自动编号的标题、三线表格、公式编号、图题注、交叉引用、GB/T 7714 参考文献、`[1-3,5]` 风格上标引用**。

---

## 生成后还有两件小事

生成的 `.docx` 用 Word 打开后，还有两个步骤（30 秒搞定）：

### 第一步：公式渲染

按 `Alt+F11` → `插入` → `模块`，粘贴这段代码，按 `F5` 运行：

```vba
Sub ConvertAllEquations()
    Dim eq As OMath
    For Each eq In ActiveDocument.OMaths
        On Error Resume Next
        eq.BuildUp
        On Error GoTo 0
    Next
    MsgBox "公式转换完成！", vbInformation
End Sub
```

### 第二步：刷新编号

`Ctrl+A` 全选，然后按 `F9`。所有图、表、公式、参考文献的编号和交叉引用就都刷好了。

---

## 模板可以换吗？可以！

虽然内置的是华科博士论文模板（作者是华科学生），但你完全可以换成自己的模板：

- 把样式名从 `Heading 1` 改成 `标题 1`，从 `Caption` 改成 `题注`……改几行配置就行
- 参考文献格式也可以改：从 GB/T 7714 换成 APA、MLA，或者你自定义的格式
- 具体改法看下面「二次开发」部分，或者直接告诉 AI：「帮我把华科的样式映射改成北航的」

---

## 手把手配置参考（如果你想自己来）

### Claude Desktop

**Windows** (`%AppData%\Claude\claude_desktop_config.json`)：

```json
{
  "mcpServers": {
    "template_mdocx": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "F:\\你的路径\\template_mdocx",
        "server.py"
      ]
    }
  }
}
```

**macOS / Linux** (`~/.config/Claude/claude_desktop_config.json`)：

```json
{
  "mcpServers": {
    "template_mdocx": {
      "command": "uv",
      "args": ["run", "--directory", "/你的路径/template_mdocx", "server.py"]
    }
  }
}
```

### Cursor

```json
{
  "mcpServers": {
    "template_mdocx": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "F:\\你的路径\\template_mdocx",
        "server.py"
      ]
    }
  }
}
```

### 手动安装

```bash
# 1. 克隆仓库
git clone https://github.com/cyling250/template_mdocx
cd template_mdocx

# 2. 安装依赖（需要 Python 3.12+ 和 uv）
uv sync

# 3. 测试一下
uv run build.py
```

---

## Markdown 写作规范

跟 AI 说"按下面格式写论文"，它输出的内容就能直接转换：

### 标题

各模板的标题规则见对应文件：

| 模板 | 标题规则 |
|------|----------|
| `hust_thesis` | [hust_thesis/TITLE_RULES.md](hust_thesis/TITLE_RULES.md) |

### 粗体 / 斜体 / 行内公式

```
**粗体**    *斜体*    $E=mc^2$
```

### 插图

```
@figure[fig:arch]{模型架构图}
![替代文字](图片路径.png)
```

### 表格

```
@table[tbl:compare]{不同方法性能对比}
| 方法 | 准确率 | 参数量 |
| --- | --- | --- |
| ResNet | 76.1% | 25.6M |
| 本文方法 | 78.5% | 32.1M |
```

### 块公式

```
@formula[eq:loss]
$$
\mathcal{L} = \frac{1}{N} \sum (y - \hat{y})^2
$$
```

### 交叉引用

```
如@ref{fig:arch}所示……结果见@ref{tbl:compare}……目标函数见@ref{eq:loss}。
```

### 参考文献引用

```
相关研究@cite{resnet2016,transformer2017}表明……
```

### 参考文献列表

```
@bibliography{refs.bib}
```

### 四条铁律

1. 每个图、表必须有 `@figure[标签]{描述}` / `@table[标签]{表名}` 声明
2. 每个块公式必须有 `@formula[标签]` 声明
3. `@ref{标签}` 的标签必须和声明的一致
4. `@cite{key}` 的 key 必须存在于 `.bib` 文件中

---

## 让 AI 帮你写论文的提示词

把这段话发给 AI，它就会按上述规范写论文：

```
你是一名学术论文撰写助手。请用扩展 Markdown 格式撰写学术论文，严格遵守以下规则：

【语法规则】
- 标题：#、##、###、####（# 对应"第X章"）
- 粗体/斜体：**粗体** / *斜体*
- 行内公式：$LaTeX$
- 图片：@figure[fig:标签]{图注} 后紧接 ![alt](路径)
- 表格：@table[tbl:标签]{表名} 后紧接 Markdown 表格
- 块公式：@formula[eq:标签] 后紧接 $$...$$
- 交叉引用：@ref{标签}（引 fig/tbl/eq）
- 文献引用：@cite{key1,key2}
- 文献列表：@bibliography{文件名.bib}

【标签前缀约定】
fig: 图片 / tbl: 表格 / eq: 公式

【生成要求】
1. 先确定章节结构
2. 逐节撰写，确保 @figure/@table/@formula 标签唯一
3. @ref 引用必须能在文中找到声明
4. @cite 的 key 必须与 BIB 文件一致
```

---

## 二次开发：适配你自己的模板

这个项目是**模板无关**的——华科的模板只是一个示例。你可以用 Vibecoding 的方式适配任何学校的论文模板。

**把这个需求告诉 AI 就行：**

> 帮我基于 template_mdocx 适配一个新模板。我的学校是 XX 大学，Word 模板文件在 `my_template.docx`，里面用到的样式名如下：正文用 `Normal`，标题用 `标题1`、`标题2`，题注用 `题注`，参考文献用 `参考文献`。请参考 `hust_thesis/` 目录下的实现方式，帮我创建一个 `my_univ/` 模块。

核心要改的只是 [generator.py](hust_thesis/generator.py) 里的样式映射：

```python
_STYLE_MAP = {
    "body": "Normal",        # ← 改成你模板里的正文样式名
    "heading_1": "Heading 1", # ← 改成你模板里的一级标题样式名
    "heading_2": "Heading 2",
    "caption": "Caption",
    "bibliography": "Bibliography",
}
```

### 项目结构速览

```
common/          ← 核心引擎，不碰它
hust_thesis/     ← 华科模板（你的模板也照这个写）
  ├── generator.py    ← 样式映射 + 生成逻辑
  ├── formatters.py   ← 参考文献格式
  └── template.docx   ← Word 模板文件
```

---

## 环境要求

- Python >= 3.12
- [uv](https://docs.astral.sh/uv/)（一行命令安装：`pip install uv`）

## 许可证

Apache-2.0
