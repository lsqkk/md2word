# md2word — 版本历史

> Agent 技能文件。

### v1.9.2（当前）
- **Mermaid 渲染重写**：SVG foreignObject 转原生 `<text>` 元素，解决 Word 不渲染 foreignObject 文字的问题
- **Mermaid 栅格图备用**：SVG 渲染失败时自动降级为 mermaid.ink JPEG API
- **Mermaid 文本自适应定位**：从 label transform 和 foreignObject 尺寸动态计算文本中心坐标
- **TOC 修复**：去除 `\h` 超链接开关，避免 Word 弹"域引用外部文件"提示
- **Heading 样式**：自动应用 Word Heading 1/2/3 样式，确保 TOC `\o` 开关可识别
- **URL-safe base64**：修复含中文的 diagram 因 `+` 号被 URL 截断导致 404 的问题

### v1.9.1
- **TOOL 文档拆分**：MD2WORD-TOOL.md 拆分为 `skills/md2word/` 多模块文件，按需读取
- **版本更新检查**：转换成功后自动检查 GitHub 新版本，24h 缓存免限速
- **CLI 新增 `--no-update-check`**：禁用版本检查
- **配置支持**：`md2word.yaml` 中添加 `update-check: false` 禁用它
- 新增模块 `update_check.py`，25 个新测试
- total: 251 tests

### v1.8.0
- **ConvertOptions 数据类**：`convert()` 接受单一配置对象
- **配置系统升级**：PyYAML 硬依赖，支持嵌套配置（style_map、列表）
- **模板自定义 XML 标记**：`<w:customXml>` 注入机器可读 slot 标识
- **Verbose 模式**：`--verbose` 显示全流程进度
- **错误上下文**：块处理错误包含 80 字符文本预览
- 39 个新测试，2450 行测试代码

### v1.6.0
- **abstract/keywords 渲染管道完成**：frontmatter 写入文档段落
- **书签冲突修复**：相同标题自动追加 `-1`、`-2` 后缀
- **`style_map` 实现**：覆盖默认样式槽名
- **红头文号可配置**：`--redhead-year` + `--redhead-number`
- **代码块空白修复**：`text.strip('\n')` 保留缩进
- **converter.py 清理**：删除死代码，抽取 cache.py
- 61 个新测试

### v1.5.0
- **架构重构**：`converter.py` 拆分为 6 个模块
- **ConversionContext**：全局状态封装为数据类
- **ConversionReport 增强**：增加 severity 级别
- **新语法**：`~~删除线~~`、`==高亮==`、`^上标^`、`~下标~`
- **任务列表**：`- [x]` / `- [ ]` 完整支持
- **YAML frontmatter** → docx 属性
- **交叉引用**：`[文字](#锚点)` → Word REF 域
- **`--out-dir`** + **glob 模式**

### v1.3.0
- `--redhead AUTHORITY`：红头文件一键生成
- `--gb-check`：GB/T 9704-2012 合规检查
- `--page-number FMT`：页码格式
- `--incremental`：MD5 增量缓存
- `.md2word/config.yaml` 自动识别
- academic-plus + redhead 主题
- 107 个测试
