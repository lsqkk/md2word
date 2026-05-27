# md2word — 配置文件详解

> Agent 技能文件。

## 文件查找优先级

1. `.md2word/config.yaml`（项目级目录）
2. `md2word.yaml`（当前/父目录）
3. `md2word.yml`（当前/父目录）
4. `md2word.json`（当前/父目录）
5. `pyproject.toml` → `[tool.md2word]` 节

搜索从当前目录向上遍历直到找到配置文件或到达根目录。

## 项目级配置目录

```
project/
├── .md2word/
│   ├── config.yaml       # 项目专属配置
│   └── template.docx     # 项目专属模板（自动识别）
├── src/
└── docs/
    └── 文章.md
```

## 配置项全表

```yaml
# 主题（official / academic / academic-plus / tech / media / redhead）
theme: academic

# 自定义模板路径（优先级高于 theme）
template: path/to/template.docx

# 输出路径（仅单文件时有效）
output: output.docx

# 目录设置
toc: true
toc-depth: "1-3"

# 标题编号
number-headings: true

# 分页
page-break: true

# 三线表
three-line-table: true

# 图片宽度（英寸）
image-width: 5.5

# 样式映射（覆盖默认样式槽）
style_map:
  code: CustomCodeStyle

# 禁用功能
no-highlight: false
no-math: false
no-mermaid: false
no-footnotes: false

# 红头文件
redhead: XX市人民政府
redhead-year: 2026
redhead-number: 12

# 页码格式（%d 为占位符）
page-number: "-- %d --"

# GB 合规检查
gb-check: false

# 增量转换
incremental: false

# 项目目录
project-dir: /path/to/project

# Watch 模式
watch: false

# 详细输出模式
verbose: false

# 版本更新检查（默认开启）
update-check: true
```

**CLI 参数优先级 > 配置文件 > 默认值**
