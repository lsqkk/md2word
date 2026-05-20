# md2word — Markdown 转 Word 文档工具

> 中文 | [English](README-en.md)

将 Markdown 文档转换为符合中国文档标准的 Word (.docx) 文件，**原生支持党政机关公文格式（GB/T 9704-2012）、学术论文格式（GB/T 7713-2015）、红头文件生成**——无需 LaTeX、无需 CSS、无需手动排版。

```bash
pip install -e ".[all]"
md2word 文章.md -o 文章.docx    # 自动识别内置主题
md2word 通知.md --redhead "XX市人民政府"   # 一键生成红头文件
```

---

## 为什么选择 md2word？

### 直面中文排版的真实痛点

中文文档排版有独特的规范体系——公文有 GB/T 9704-2012（仿宋三号正文、37mm 上边距），学术论文有 GB/T 7713-2015（宋体小四、首行缩进两字符），这些格式在 Pandoc、LaTeX 等通用工具中配置复杂且容易出错。md2word 从设计之初就以中文规范为第一优先级：

| 场景 | 通用工具的做法 | md2word 的做法 |
|------|--------------|---------------|
| **红头文件** | 不支持或需手动拼接 | `--redhead "XX单位"` 一键生成（红头 + 分隔线 + 文号） |
| **公文格式** | 需手写 LaTeX 模板 | 内置 GB/T 9704-2012 标准主题，开箱即用 |
| **东亚字体** | 易回退为日文/韩文字体 | 显式指定 SimSun/SimHei/FangSong/KaiTi，杜绝字体回退 |
| **三线表** | 需手写表格样式 | `--three-line-table` 一键切换 |
| **公式** | 转图片或需 MathML 手写 | LaTeX → OMML 原生可编辑公式 |
| **交叉引用** | 需 Pandoc 过滤器 | Markdown 内链 → Word REF 域 |

### 模板系统：在 Word 里改格式，工具自动理解

大多数文档工具的模板需要写代码（LaTeX 模板、CSS 样式表、Pandoc filters）。md2word 的模板就是 **一个普通的 .docx 文件**——打开它，修改引导段落（如"一级标题""正文"）的字体、字号、颜色，保存后用新模板转换即可。

这在团队协作中非常有价值：非技术同事可以在 Word 中调整样式规范，开发者无需反复改代码。

### 与 Pandoc 的对比

| 能力 | md2word | Pandoc |
|------|---------|--------|
| **安装** | 纯 Python, `pip install` | Haskell 生态, 较重型 |
| **模板修改** | 打开 Word 改格式 | 需写 LaTeX/自定义 writer |
| **中文规范** | 内置 GB 标准 + 东亚字体策略 | 依赖外部模板配置 |
| **红头文件** | `--redhead` 一键生成 | 不支持 |
| **公式** | OMML 原生可编辑 | 需 `--reference-docx` |
| **Mermaid** | 内建渲染 | 需外部预处理 |
| **三线表** | `--three-line-table` | 需自定义模板 |
| **增量转换** | 内建（内容哈希缓存） | 无 |
| **Watch 模式** | 内建 | 无 |
| **PDF 输出** | 不支持（定位是 Word 专用） | 内建 |

---

## 安装

需要 Python 3.10+。

```bash
# 从项目根目录安装
pip install -e .

# 安装全部可选增强
pip install -e ".[all]"
```

可选依赖：

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

---

## 快速开始

```bash
# 基本转换（自动搜索内置主题）
md2word 文章.md -o 文章.docx

# 指定主题
md2word 文章.md --theme academic -o 文章.docx

# 使用自定义模板
md2word 文章.md -t 我的模板.docx -o 文章.docx

# 批量转换——自动为每个 .md 生成同名 .docx
md2word 第一章.md 第二章.md 第三章.md

# 标准输入
cat 文章.md | md2word -o 文章.docx
```

---

## 内置主题（6 套）

| 主题 | 适用场景 | 风格特征 |
|------|---------|---------|
| `academic` | 学位论文、期刊投稿 | 宋体系列 + 黑体标题居中 + 首行缩进 + 标准学术版心 |
| `academic-plus` | 学术论文（增强版） | 在 academic 基础上增加摘要/关键词/参考文献引导段落 |
| `official` | 政府公文 | 仿宋三号正文 + 黑体标题 + GB/T 9704-2012 标准页边距 |
| `tech` | 技术文档、API 手册 | 微软雅黑 + 深蓝层级标题 + 紧凑排版 + Consolas 代码 |
| `media` | 公众号、自媒体 | 大号标题(32pt) + 橙色品牌色 + 1.8 倍高行距 + 楷体引用 |
| `redhead` | **红头文件** | GB/T 9704-2012 标准 + 红头样式 + 黑体标题 + 仿宋正文 |

### 红头文件专属功能

```bash
md2word 通知.md --redhead "XX市人民政府" --theme redhead
```

自动生成：
- **红色发文机关名称**（28pt 宋体, #CC0000）+ "文件" 后缀
- **红色分隔线**（满宽）
- **公文文号占位符**（如"XX市人民政府文件"下方）
- GB/T 9704-2012 标准页边距（上 37mm、下 35mm、左 28mm、右 26mm）

---

## 新功能（v1.3.0）

### 项目级配置 `.md2word/` 目录

在项目根目录创建 `.md2word/` 文件夹，放置配置文件和模板：

```
项目目录/
├── .md2word/
│   ├── config.yaml       # 项目级配置（自动加载）
│   └── template.docx     # 项目级模板（自动识别）
├── src/
└── 文章.md
```

配置查找优先级：`.md2word/config.yaml` > `md2word.yaml` > `md2word.yml` > `md2word.json` > `pyproject.toml`

```yaml
# .md2word/config.yaml
theme: academic
toc: true
toc-depth: 1-3
number-headings: true
three-line-table: true
```

### 增量转换（`--incremental`）

基于 MD5 内容哈希的智能跳过机制——重复转换时，未变动的文件直接跳过：

```bash
md2word 第一章.md 第二章.md 第三章.md --incremental
```

首次运行完整转换，后续只处理内容有变化的文件。缓存文件 `.md2word_cache.json` 自动管理。

### 红头文件（`--redhead`）

参见上方红头文件专属功能说明。

### GB 标准合规检查（`--gb-check`）

自动检测文档格式是否符合 GB/T 9704-2012 党政机关公文格式标准：

```bash
md2word 通知.md --theme official --gb-check
```

输出示例：
```
  ⚠  GB/T 9704-2012 公文上边距应为 37mm，当前 30.0mm
  ⚠  GB/T 9704-2012 公文左边距应为 28mm，当前 25.0mm
```

### 页码格式（`--page-number`）

```bash
md2word 文档.md --page-number "-- %d --"
```

自动在页脚居中插入页码，支持任意格式（`%d` 为页码占位符）。同时自动插入总页数域。

### 项目目录（`--project-dir`）

显式指定项目根目录，便于从子目录使用项目配置：

```bash
md2word 文章.md --project-dir /path/to/project
```

---

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

文件变化后自动重新转换，适合写作时实时预览。支持 watchdog（高效）和 polling（回退）两种模式。

### 配置文件

```yaml
# md2word.yaml
theme: academic
toc: true
toc-depth: 1-3
number-headings: true
three-line-table: true
```

之后只需 `md2word input.md` 即可——参数自动读取。CLI 参数优先级高于配置文件。

---

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

---

## 全部 CLI 参数

| 参数 | 说明 |
|------|------|
| `inputs` | 输入 Markdown 文件（支持多个） |
| `-o, --output` | 输出 .docx 路径 |
| `-t, --template` | 模板 .docx 文件 |
| `--theme` | 内置主题名（academic / academic-plus / official / tech / media / redhead） |
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
| `--redhead` | 红头文件发文机关名称 |
| `--page-number` | 页码格式（如 `-- %d --`） |
| `--gb-check` | GB 标准合规检查 |
| `--incremental` | 增量转换（内容哈希缓存） |
| `--project-dir` | 项目根目录 |
| `--create-template` | 生成模板文件 |
| `--validate-template` | 验证模板完整性 |
| `--list-styles` | 查看模板样式 |
| `--list-themes` | 列出可用主题 |
| `--check-deps` | 检查依赖完整性 |
| `--version` | 显示版本号 |

---

## 依赖关系

| 特性 | 依赖 | 可选 |
|------|------|------|
| 核心转换 | python-docx, markdown, Pillow | 否 |
| 代码高亮 | Pygments | 是 |
| 数学公式 | matplotlib, latex2mathml | 是 |
| SVG 嵌入 | resvg | 是 |
| Watch 模式 | watchdog | 是（无 watchdog 时自动降级为 polling） |
| YAML 配置 | PyYAML | 是（无 PyYAML 时降级为纯文本解析） |

---

## 工作原理

```
Markdown → HTML → 按块分派 → 模板样式提取 → python-docx 构建 → .docx
                           ↕
                   脚注/公式/Mermaid/代码高亮
```

工具的独特之处在于**样式提取机制**：不是用代码定义格式，而是从 Word 模板中读取已存在的段落样式。打开模板 docx → 修改引导段落格式 → 保存 → md2word 自动使用新样式。

---

## 项目结构

```
src/md2word/
├── cli.py              — CLI 入口、参数解析、Watch 模式
├── converter.py        — 核心转换管道（红头文件、GB 检查、页码、增量缓存）
├── template.py         — 引导段落样式提取 + 模板验证
├── config.py           — 配置文件加载（YAML/JSON/TOML + .md2word/ 目录）
├── footnotes.py        — 脚注提取 + Word 原生脚注写入
├── themes.py           — 6 套内置主题定义 + 模板生成
├── image_utils.py      — 图片下载、缩放、SVG 处理
├── syntax.py           — Pygments 语法高亮
├── math_omml.py        — LaTeX → Word OMML 转换
├── math_renderer.py    — LaTeX → SVG 渲染（matplotlib 回退）
├── mermaid_renderer.py — Mermaid → SVG（API / mmdc）
template/
├── 学术论文.docx        academic theme
├── 官方公文.docx        official theme
├── 技术文档.docx        tech theme
├── 自媒体排版.docx      media theme
```

## 技术栈

Python 3.10+, python-docx, lxml, markdown, Pillow, requests, Pygments, matplotlib, latex2mathml, watchdog

## 许可证

MIT
