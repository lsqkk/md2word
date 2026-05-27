# md2word — CLI 命令

> Agent 技能文件。使用前请将路径修改为实际安装位置。

## 基本转换

```bash
# 单文件 → docx（自动搜索内置主题模板）
md2word input.md -o output.docx

# 批量转换（每个 .md → 同名 .docx）
md2word file1.md file2.md file3.md

# glob 批量匹配
md2word "docs/**/*.md"

# 指定输出目录
md2word file1.md file2.md --out-dir ./output

# 标准输入
cat doc.md | md2word -o output.docx
```

## 模板与主题

```bash
# 指定内置主题
md2word input.md --theme academic -o output.docx
md2word input.md --theme official -o output.docx
md2word input.md --theme redhead -o output.docx

# 使用自定义 .docx 模板
md2word input.md -t 我的模板.docx -o output.docx

# 生成主题模板（供自定义修改）
md2word --create-template 新模板.docx --theme academic

# 验证模板中必需的引导段落是否齐全
md2word --validate-template 模板.docx

# 列出模板中已识别的引导段落及其格式
md2word --list-styles -t 模板.docx

# 列出所有可用主题
md2word --list-themes
```

## 红头文件

```bash
# 一键生成红头文件（发文机关名 + "文件" + 红色分隔线 + 文号）
md2word 通知.md --redhead "XX市人民政府" --theme redhead

# 自定义文号年份和编号
md2word 通知.md --redhead "XX市人民政府" --redhead-year 2026 --redhead-number 12

# 自动行为：TOC 默认关闭、正文首行缩进、标题方正小标宋 26pt
# 如需强制开启目录，追加 --toc
```

## 格式控制

```bash
md2word input.md --toc                       # 生成目录（红头文件默认无目录）
md2word input.md --no-toc                    # 显式禁用目录
md2word input.md --toc-depth 1-4             # 目录深度（默认 1-3）
md2word input.md --number-headings           # 标题自动编号（1, 1.1, 1.1.1）
md2word input.md --page-break                # H1 前分页
md2word input.md --three-line-table          # 三线表（顶线+表头下线+底线，无竖线）
md2word input.md --image-width 4.5           # 图片最大宽度（英寸，默认 5.5）
md2word input.md --page-number "-- %d --"    # 页码格式（%d 为页码占位符）
md2word input.md --no-footnotes              # 禁用脚注处理
md2word input.md --no-highlight              # 禁用代码高亮
md2word input.md --no-math                   # 禁用公式渲染
md2word input.md --no-mermaid                # 禁用 Mermaid 图表
md2word input.md --no-update-check           # 禁用 GitHub 版本更新检查
```

## 配置文件

```bash
# 自动检测（按优先级）：
# .md2word/config.yaml > md2word.yaml > md2word.yml > md2word.json > pyproject.toml
md2word input.md

# 指定配置文件
md2word input.md --config 路径/config.yaml

# 项目级目录（自动加载 .md2word/config.yaml 和 .md2word/template.docx）
md2word input.md --project-dir /path/to/project
```

## 高级功能

```bash
# GB 标准合规检查（检查页边距/字体是否符合 GB/T 9704-2012）
md2word input.md --theme official --gb-check

# 增量转换（MD5 缓存，仅处理内容变化的文件）
md2word file1.md file2.md --incremental

# 详细输出模式（显示各阶段进度）
md2word input.md --verbose

# Watch 模式（文件变化自动重新转换）
md2word input.md --watch
md2word dir/ file.md --watch

# 查看版本
md2word --version
```

## 完整参数列表

| 参数 | 说明 |
|------|------|
| `inputs` | 输入 Markdown 文件（支持多个和 glob 模式） |
| `-o, --output` | 输出 .docx 路径 |
| `--out-dir` | 批量输出目录 |
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
| `--no-update-check` | 禁用版本更新检查 |
| `--redhead` | 红头文件发文机关名称 |
| `--redhead-year` | 红头文件文号年份 |
| `--redhead-number` | 红头文件文号编号 |
| `--page-number` | 页码格式（如 `-- %d --`） |
| `--gb-check` | GB 标准合规检查 |
| `--incremental` | 增量转换（内容哈希缓存） |
| `--project-dir` | 项目根目录 |
| `--config` | 指定配置文件路径 |
| `--watch` | 监听文件变化自动转换 |
| `--verbose` | 显示详细转换过程 |
| `--style-map` | 样式映射（配置文件中使用） |
| `--create-template` | 生成模板文件 |
| `--validate-template` | 验证模板完整性 |
| `--list-styles` | 查看模板样式 |
| `--list-themes` | 列出可用主题 |
| `--check-deps` | 检查依赖完整性 |
| `--version` | 显示版本号 |
