# md2word — Markdown to Word Tool

**项目路径:** `D:\git\lsqkk\md2word\` (已注册为系统命令 `md2word`)

## 安装
```bash
pip install -e .
# 或带所有可选依赖
pip install -e ".[all]"
```

## 核心用法
```bash
md2word input.md -o output.docx                     # 默认模板
md2word input.md --theme academic                   # 指定主题
md2word input.md -t 模板文件.docx -o output.docx     # 自定义模板
md2word file1.md file2.md                           # 批量转换
md2word input.md --watch                            # 监听文件自动转换
cat doc.md | md2word -o out.docx                    # 标准输入
```

## 高级功能
```bash
# 目录与编号
md2word input.md --toc --toc-depth 1-3 --number-headings --page-break

# 学术排版
md2word input.md --theme academic --number-headings --three-line-table

# 模板操作
md2word --create-template my-template.docx --theme academic
md2word --validate-template my-template.docx
md2word --list-styles -t template.docx

# 配置与检查
md2word --check-deps
md2word --config md2word.yaml input.md
```

## 主题

| 参数 | 文件名 | 风格 |
|------|--------|------|
| `official` | 官方公文.docx | 仿宋+黑体，GB标准页边距 |
| `academic` | 学术论文.docx | 宋体系列，首行缩进，学术排版 |
| `tech` | 技术文档.docx | 微软雅黑+深蓝标题，紧凑 |
| `media` | 自媒体排版.docx | 大标题+橙色+高行距 |

## 引导段落模板系统

模板中的特定关键词段落定义样式。用户只需在 Word 中修改这些段落的格式即可自定义：

| 关键词 | 对应元素 |
|--------|----------|
| 一级标题 | `# 标题` |
| 二级标题 | `## 标题` |
| 三级标题 | `### 标题` |
| 正文 | 普通段落 |
| 首行缩进 | 首行缩进的正文 |
| 图片/图注 | 图片容器和说明 |
| 引用 | `> 引用块` |
| 代码 | 代码块 |
| 无序/有序列表 | 列表项 |
| 目录标题 | 目录标题 |

## 配置文件 (`md2word.yaml`)

支持 `md2word.yaml`、`md2word.yml`、`md2word.json`、`pyproject.toml ([tool.md2word])`：

```yaml
template: template/技术文档.docx
theme: tech
toc: true
toc-depth: "1-3"
number-headings: true
page-break: true
three-line-table: true
```

## 新功能 (v1.2.0)

- **配置文件** — 自动读取 md2word.yaml，CLI 参数优先
- **Watch 模式** (`--watch`) — 监听 .md 文件变化自动转换
- **脚注** — `[^1]` 语法 → Word 原生脚注
- **三线表** (`--three-line-table`) — 学术表格式
- **交叉引用** — `[文字](#标题名)` → Word REF 域
- **页眉页脚** — 自动保留模板中的页眉页脚
- **公式编号** — 块级公式自动 SEQ 编号
- **依赖检查** — `--check-deps` 查看功能完整性
- **更好的错误提示** — 模板缺失、文件未找到等友好提示
