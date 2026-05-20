# md2word — Markdown 转 Word 文档工具

> 中文 | [English](README-en.md)

将 Markdown 文档转换为格式精美的 Word (.docx) 文件，基于**模板引导段落**实现样式自定义——在 Word 中改格式，工具自动适配，无需写 CSS 或代码。

```bash
pip install -e ".[all]"
md2word input.md -o output.docx
```

## 亮点

- **模板即样式**：Word 中的「一级标题」「正文」等引导段落 = 样式定义。在 Word 里改字体、字号、颜色，工具自动读取并应用，**不需要写 CSS、不需要改代码**
- **四套内置主题**：学术论文、官方公文、技术文档、自媒体排版，开箱即用
- **LaTeX 公式 → Word OMML**：基于 matplotlib + latex2mathml，公式在 Word 中**原生可编辑**，非截图
- **Mermaid 图表**：流程图、时序图、甘特图自动渲染为 SVG 嵌入
- **代码语法高亮**：Pygments 驱动，200+ 语言着色
- **脚注**：Markdown 标准 `[^1]` 语法 → Word 原生脚注
- **三线表**：学术期刊风格的 `--three-line-table`
- **交叉引用**：Markdown 内链 → Word REF 域，Ctrl+F9 更新
- **目录 + 标题编号 + 分页**：一键生成 TOC、自动编号、H1 前分页
- **Watch 模式**：监听文件变化自动重新转换
- **配置文件**：`md2word.yaml` 保存常用参数，减少命令行输入
- **零模板依赖**：不装任何模板也能用——自动搜索内置主题

## 安装

需要 Python 3.10+。

```bash
# 从项目根目录安装
pip install -e .

# 安装全部可选增强
pip install -e ".[all]"
```

可选依赖按需安装：

| 组件 | 安装命令 | 功能 |
|------|---------|------|
| 代码高亮 | `pip install -e ".[highlight]"` | Pygments 语法着色 |
| 数学公式 | `pip install -e ".[math]"` | LaTeX → OMML/SVG |
| SVG 支持 | `pip install -e ".[svg]"` | 原生 SVG 嵌入 |
| 配置/监听 | `pip install -e ".[all]"` | YAML 配置 + Watch 模式 |

查看可用组件：

```bash
md2word --check-deps
```

## 快速开始

```bash
# 基本转换（自动使用内置学术论文模板）
md2word 文章.md -o 文章.docx

# 指定主题
md2word 文章.md --theme academic -o 文章.docx

# 使用自定义模板
md2word 文章.md -t 我的模板.docx -o 文章.docx

# 批量转换
md2word 第一章.md 第二章.md 第三章.md

# 标准输入
cat 文章.md | md2word -o 文章.docx
```

## 内置主题

| 主题 | 适用场景 | 风格特征 |
|------|---------|---------|
| `academic` | 学位论文、期刊投稿 | 宋体系列 + 黑体标题居中 + 首行缩进 + 标准学术版心 |
| `official` | 政府公文、红头文件 | 仿宋三号正文 + 黑体标题 + GB 标准页边距 |
| `tech` | 技术文档、API 手册 | 微软雅黑 + 深蓝层级标题 + 紧凑排版 + Consolas 代码 |
| `media` | 公众号、自媒体 | 大号标题(32pt) + 橙色品牌色 + 1.8 倍高行距 + 楷体引用 |

## 进阶用法

### 学术排版

```bash
md2word 论文.md --theme academic --toc --number-headings --three-line-table --page-break
```

一键生成：目录、层级编号 (1/1.1/1.1.1)、三线表、每章分页。

### 脚注

```markdown
Markdown 标准语法[^1]，自动转为 Word 原生脚注。

[^1]: 这是脚注内容，可跨多行。
```

### 交叉引用

```markdown
参考[数据说明](#数据说明)。      → 生成 Word REF 域
详细设置见[表1](#表1)。          → 引用表格
请参见[图1](#图1)。              → 引用图片
```

在 Word 中按 Ctrl+A → F9 更新域，自动填充标题/编号。

### 数学公式

```markdown
内联公式：$E = mc^2$
块级公式：$$\sum_{i=1}^{n} i = \frac{n(n+1)}{2}$$
```

支持 LaTeX → Word OMML（原生可编辑公式）并自动编号。

### 三线表

```bash
md2word 论文.md --three-line-table
```

学术期刊风格表格：顶线/底线加粗、表头下线、无竖线。

### Watch 模式

```bash
md2word 文章.md --watch
```

文件变化后自动重新转换，适合写作时实时预览。

### 配置文件

创建 `md2word.yaml` 于项目目录：

```yaml
theme: academic
toc: true
toc-depth: 1-3
number-headings: true
three-line-table: true
```

之后只需 `md2word input.md` 即可——参数自动读取。CLI 参数优先级高于配置文件。

## 自定义模板

**不需要写代码。** 在 Word 中新建 docx，插入以下关键词段落，调整格式，保存即可：

| 引导段落 | 对应元素 |
|---------|---------|
| 一级标题 | `# 标题` |
| 二级标题 | `## 标题` |
| 三级标题 | `### 标题` |
| 正文 | 普通段落 |
| 首行缩进 | 首行缩进的正文 |
| 图片 | 图片容器 |
| 图注 | 图片说明 |
| 引用 | `> 引用块` |
| 代码 | 代码块 |
| 无序列表 | `- 列表项` |
| 有序列表 | `1. 列表项` |
| 目录标题 | 目录上方标题 |

工具的检测规则：按段落**从头到尾**匹配关键词，**先匹配到的优先**。

快速生成模板：

```bash
# 基于学术主题生成模板
md2word --create-template 我的模板.docx --theme academic

# 验证模板完整性
md2word --validate-template 我的模板.docx

# 查看模板识别的样式
md2word --list-styles -t 我的模板.docx
```

## 全部 CLI 参数

| 参数 | 说明 |
|------|------|
| `inputs` | 输入 Markdown 文件（支持多个） |
| `-o, --output` | 输出 .docx 路径 |
| `-t, --template` | 模板 .docx 文件 |
| `--theme` | 内置主题名 |
| `--image-width` | 图片最大宽度（英寸，默认 5.5） |
| `--toc / --no-toc` | 启用/禁用目录 |
| `--toc-depth` | 目录深度（如 `1-3`） |
| `--number-headings` | 标题自动编号 |
| `--page-break` | H1 前插入分页符 |
| `--three-line-table` | 三线表样式 |
| `--no-footnotes` | 禁用脚注 |
| `--no-highlight` | 禁用代码高亮 |
| `--no-math` | 禁用数学公式 |
| `--no-mermaid` | 禁用 Mermaid 图表 |
| `--watch` | 监听文件变化自动转换 |
| `--config` | 指定配置文件路径 |
| `--create-template` | 生成模板文件 |
| `--validate-template` | 验证模板完整性 |
| `--list-styles` | 查看模板样式 |
| `--list-themes` | 列出可用主题 |
| `--check-deps` | 检查依赖完整性 |
| `--version` | 显示版本号 |

## 依赖关系

| 特性 | 依赖 | 可选 |
|------|------|------|
| 核心转换 | python-docx, markdown, Pillow | 否 |
| 代码高亮 | Pygments | 是 |
| 数学公式 | matplotlib, latex2mathml | 是 |
| SVG 嵌入 | resvg | 是 |
| Watch 模式 | watchdog | 是 |
| YAML 配置 | PyYAML（纯文本回退） | 否（无 PyYAML 时自动降级） |

## 工作原理

```
Markdown → HTML → 按块分派 → 模板样式提取 → python-docx 构建 → .docx
                           ↕
                   脚注/公式/Mermaid/代码高亮
```

工具的独特之处在于**样式提取机制**：不是用代码定义格式，而是从 Word 模板中读取已存在的段落样式。打开模板 docx → 修改引导段落格式 → 保存 → md2word 自动使用新样式。

## 项目结构

```
src/md2word/
├── cli.py              — CLI 入口、参数解析、Watch 模式
├── converter.py        — 核心转换管道
├── template.py         — 引导段落样式提取 + 模板验证
├── config.py           — 配置文件加载（YAML/JSON/TOML）
├── footnotes.py        — 脚注提取 + Word 原生脚注写入
├── themes.py           — 4 套内置主题定义 + 模板生成
├── image_utils.py      — 图片下载、缩放、SVG 处理
├── syntax.py           — Pygments 语法高亮
├── math_omml.py        — LaTeX → Word OMML 转换
├── math_renderer.py    — LaTeX → SVG 渲染（matplotlib 回退）
├── mermaid_renderer.py — Mermaid → SVG（API / mmdc）
template/
├── 学术论文.docx
├── 官方公文.docx
├── 技术文档.docx
└── 自媒体排版.docx
```

## 技术栈

Python 3.10+, python-docx, lxml, markdown, Pillow, requests, Pygments, matplotlib, latex2mathml, watchdog

## 与 Pandoc 对比

| 能力 | md2word | Pandoc |
|------|---------|--------|
| 安装复杂度 | 纯 Python, pip install | Haskell 生态, 较重型 |
| 模板机制 | Word 改格式即生效 | 需写 LaTeX/自定义 writer |
| 公式 | OMML 原生可编辑 | 需 --reference-docx |
| 三线表 | 一键开关 | 需自定义模板 |
| 中文排版 | 显式东亚字体, 防日文回退 | 依赖模板配置 |
| Mermaid | 内建 | 需预处理 |
| Watch 模式 | 内建 | 无 |
| 脚注 | Markdown 标准语法 | 支持 |
| 交叉引用 | Word REF 域 | 需 Pandoc 过滤器 |
| PDF 输出 | 不支持 | 内建 |

md2word 定位是**Markdown → Word 的专用工具**，专注于中文排版和 Word 原生特性，不做通用文档转换。
