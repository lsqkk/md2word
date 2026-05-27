# md2word — 安装

> Agent 技能文件。使用前请将如下路径修改为实际安装位置。

## 核心安装

```bash
pip install -e D:/git/lsqkk/md2word
```

## 全部可选依赖

```bash
pip install -e "D:/git/lsqkk/md2word[all]"
```

## 按需安装

| 组件 | 安装命令 | 功能 |
|------|---------|------|
| 代码高亮 | `pip install -e "D:/git/lsqkk/md2word[highlight]"` | Pygments 语法着色 |
| 数学公式 | `pip install -e "D:/git/lsqkk/md2word[math]"` | LaTeX → OMML 可编辑公式 |
| SVG 嵌入 | `pip install -e "D:/git/lsqkk/md2word[svg]"` | 原生 SVG 嵌入 docx |
| 文件监听 | `pip install -e "D:/git/lsqkk/md2word[watch]"` | Watch 模式（watchdog） |

## 检查依赖状态

```bash
md2word --check-deps
```
