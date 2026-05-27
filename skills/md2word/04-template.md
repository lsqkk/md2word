# md2word — 引导段落模板系统

> Agent 技能文件。

## 原理

模板 docx 中包含以特定关键词开头的"引导段落"。工具读取这些段落的格式（字体、字号、颜色、对齐、缩进、行距），应用到对应的 Markdown 元素上。

**用户只需在 Word 中修改引导段落的格式，即可自定义输出样式，无需写代码。**

## 关键词列表

| 关键词（中/英文） | 样式槽位 | 对应元素 |
|-------------------|---------|----------|
| 一级标题 / heading 1 | `h1` | `# 标题` |
| 二级标题 / heading 2 | `h2` | `## 标题` |
| 三级标题 / heading 3 | `h3` | `### 标题` |
| 四级标题 / heading 4 | `h4` | `#### 标题` |
| 五级标题 / heading 5 | `h5` | `##### 标题` |
| 正文 | `body` | 普通段落 |
| 首行缩进 | `body_indent` | 首行缩进的正文 |
| 图片 | `image` | `<img>` 容器 |
| 图注 | `figcaption` | 图片下方说明 |
| 引用 / quote | `quote` | `> 引用块` |
| 代码 / code block | `code` | 代码块 |
| 无序列表 / bullet list | `bullet_list` | `- 列表项` |
| 有序列表 / ordered list | `number_list` | `1. 列表项` |
| 目录标题 / table of contents | `toc_title` | 目录上方标题 |
| 摘要 / abstract | `abstract` | YAML frontmatter 摘要 |
| 关键词 / keywords | `keywords` | YAML frontmatter 关键词 |
| 参考文献 / references | `references` | 参考文献段落 |

## 检测规则

1. 按引导段落**在模板中的出现顺序**匹配
2. 先匹配到的优先（靠前的段落优先级高）
3. 匹配时去标点、去空格、忽略大小写
4. 一个槽位被匹配后，模板中后续的同关键词段落不再生效

## 模板验证

```bash
# 验证模板是否缺失必需段落
md2word --validate-template 模板.docx

# 必需的：一级标题、二级标题、三级标题、正文
# 推荐的：首行缩进、图片、图注、引用、代码、无序/有序列表、目录标题
```

## 快速生成模板

```bash
# 基于学术主题生成模板
md2word --create-template 我的模板.docx --theme academic

# 基于自媒体主题
md2word --create-template 我的模板.docx --theme media
```

## 样式映射

通过配置文件覆盖默认样式槽：

```yaml
# md2word.yaml
style_map:
  code: CustomCodeStyle
  quote: 自定义引用
```
