# 更新日志

## v0.0.2 (2026-07-24)

### 项目重命名

- **仓库**: `Python-Excel-to-PPT-Vocabulary-Tool` → `xlsx-to-pptx`
- **二进制**: `单词卡片转换` → `英语助记卡片生成`
- **窗口标题**: `单词PPT生成器` → `英语助记卡片生成`

### 新增功能

#### PPTX → PNG 渲染管线

- 用 `resvg` + `usvg` + `tiny-skia` 替换 `ab_glyph` 逐像素渲染
- 修复自闭合 XML 元素（`<a:off/>`）被忽略导致排版异常的 bug
- 跨平台字体探测：Latin / CJK / Emoji 三层回退链
- `catch_unwind` 保护损坏系统字体导致的崩溃
- 每张 slide 独立渲染，单张失败不阻止其他
- `AtomicBool` 取消支持（GUI 取消按钮）
- 空白页检测：文字密度 < 0.01% 自动告警

#### 全栈诊断系统

- `DiagStore` 横切基础设施：`info` / `warn` / `error` 三级事件
- NDJSON 结构化日志，`serde_json` 序列化
- 11 个 `target` 命名空间覆盖全管线：`font_probe` / `parse` / `slide` / `render` / `verify` / `reader` / `generator` / `template` / `cli` / `gui` / `export`
- 所有模块注入 `&mut DiagStore`，零 `_diag` 未使用参数
- 修复 5 处 `let _ =` 静默吞错误

#### CLI `diag` 子命令

```bash
英语助记卡片生成 diag export.ndjson --summary         # 会话总览
英语助记卡片生成 diag export.ndjson --blank-slides     # 空白 slide 检查
英语助记卡片生成 diag export.ndjson --font-trace       # 字体探测链
英语助记卡片生成 diag export.ndjson --errors           # 所有 WARN/ERROR
英语助记卡片生成 diag export.ndjson --slide 2          # 单张详情
英语助记卡片生成 diag export.ndjson --json             # 原始 JSON（脚本消费）
```

#### 自定义 PPTX 模板

- 用户在 PowerPoint 中自由设计排版，`{{占位符}}` 标记数据位置
- 支持 6 个字段：`{{单词}}` `{{音标}}` `{{词根词缀}}` `{{例句}}` `{{例句释义}}` `{{单词释义}}`
- 占位符扫描/替换/校验，兼容 PowerPoint 拆分文本
- 条目数超过模板幻灯片数时自动复制最后一张

### 文档

- `AGENTS.md`: AI 代理规则（架构 / TDD / DiagStore API / 错误处理 / commit 规范）
- `.github/copilot-instructions.md`: Copilot 专属指令
- `docs/diagnostics.md`: NDJSON 格式规范 + Agent 诊断工作流 + jq / Python 示例
- `docs/superpowers/specs/`: 设计规格（架构 / 字体策略 / 审计缺口）
- `docs/superpowers/plans/`: 14 任务 TDD 实现计划

### CI/CD

- 版本标签（`v*.*.*`）推送自动触发 Release 工作流
- 三平台构建：Linux x86_64 / Windows x86_64 (mingw) / macOS x86_64
- 产物自动附加到 GitHub Release

### 清理

- 删除旧 Python 原型文件（`主程序.py` / `创建模板.py` / `打包配置.py`）
- 删除 `单词PPT生成器.exe`
- 删除 `HANDOVER.md`（内容合并到 `AGENTS.md` + `README.md`）
- 合并 `io_other_error` 和 `if_same_then_else` 的 clippy 修复

### 测试

- 全仓 74 个测试，零失败
- 新增 20 个 PNG 导出测试（20/20）
- 新增 4 个 DiagStore 测试（4/4）

---

## v0.0.1 (2026-02)

- 初始原型：Python 实现
- 基础 Excel → PPTX 转换
