# md2word — 红头文件 Markdown 模板

> Agent 技能文件。

快速创建红头文件 markdown 的推荐结构：

```markdown
# 关于XXXXX的通知

各有关单位：

　　现将有关事项通知如下：

一、XXXXX

（一）XXXXX

　　详细内容……

（二）XXXXX

二、XXXXX

<p style="text-align: right">西安交通大学教务处</p>
<p style="text-align: right">2026年5月20日</p>
```

## 格式说明

| 元素 | 渲染效果 |
|------|---------|
| `# 标题` | 方正小标宋简体 26pt 居中（公文标题） |
| `## 标题` | 黑体 16pt（一级结构层次） |
| `### 标题` | 楷体 16pt（二级结构层次） |
| 正文段落 | 仿宋 16pt 首行缩进（三号字） |
| 落款和日期 | 用 `<p style="text-align: right">` 实现右对齐 |

## 转换命令

```bash
md2word 通知.md --redhead "XX单位" --theme redhead

# 自定义文号
md2word 通知.md --redhead "XX单位" --redhead-year 2026 --redhead-number 12
```
