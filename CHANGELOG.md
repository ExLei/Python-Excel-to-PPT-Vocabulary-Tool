# 更新日志

## v0.0.2 (2026-07-24)

### 语言重写：Python → Rust

- **零运行时依赖**：单二进制 9.7 MB，无需安装 Python / pip
- **编译时安全**：所有权模型消除内存 bug，`catch_unwind` 保护字体加载崩溃
- **跨平台**：Linux / Windows (mingw 交叉编译) / macOS 三平台一致行为

### 架构重写：单体脚本 → Workspace 三 Crate

```
crates/
├── core/  # 纯逻辑，零 UI 依赖
├── cli/   # 命令行 (clap)
└── gui/   # 图形界面 (egui/eframe)
```

- CLI + GUI 合并为单二进制：`argc > 1` → CLI，否则 → GUI
- 公共 API 向后兼容，调用方无需修改

### PNG 导出管线重写：ab_glyph → resvg + SVG

| 层 | 旧方案 | 新方案 |
|---|---|---|
| XML 解析 | `Event::Start` only，自闭合元素忽略 | `Event::Start` + `Event::Empty` |
| 中间格式 | 无（直接 pixel buffer） | SVG（resvg 原生消费） |
| 字体渲染 | `ab_glyph` 逐像素 rasterize | `resvg` + `rustybuzz`（HarfBuzz 文本整形） |
| 字体查找 | 手动扫描系统目录 | `fontdb` 自动发现 + 通用族映射 |

- 修复 `<a:off/>` 等自闭合 XML 元素被忽略的根因 bug
- 跨平台字体探测：Latin / CJK / Emoji 三层回退链
- 空白页检测：文字密度 < 0.01% 自动告警

### PPTX 生成重写：python-pptx → 手写 OOXML

- 内嵌固定模板文本替换（`__WORD__` / `__PHONETIC__` 等占位符）
- 用户自定义模板模式：`{{单词}}` `{{音标}}` 等 6 字段占位符
- PowerPoint 拆分文本兼容：`{{` `单词` `}}` 跨 `<a:r>` 元素处理
- 条目数超模板幻灯片数时自动复制最后一张
- OOXML 结构 100% 合规（模板来自 PowerPoint 认可版本）

### 字体系统重写：嵌入字体 → 运行时探测

- 移除嵌入字体文件，GUI 体积从 12 MB 降至 ~3 MB
- 系统字体自动扫描：IPA 音标 + CJK 中文正常显示
- Latin 回退链：`Segoe UI` → `Helvetica` → `Arial` → `DejaVu Sans`
- CJK 回退链：`Microsoft YaHei` → `PingFang SC` → `Noto Sans CJK SC` → `WenQuanYi Micro Hei`
- Emoji 回退链：`Segoe UI Emoji` → `Apple Color Emoji` → `Noto Color Emoji`

### 错误处理重写：eprintln → DiagStore + NDJSON

- `DiagStore` 横切基础设施：`info` / `warn` / `error` 三级事件
- 11 个 `target` 命名空间覆盖全管线
- 所有模块注入 `&mut DiagStore`，零 `_diag` 未使用参数
- 修复 5 处 `let _ =` 静默吞错误
- NDJSON 结构化日志，`serde_json` 序列化
- CLI `diag` 子命令：`--summary` / `--blank-slides` / `--font-trace` / `--errors` / `--slide`


### 测试

- 全仓 74 个测试，零失败
- TDD 强制执行：先写测试 → 验证失败 → 最少代码通过 → 重构

### CI/CD

- 版本标签（`v*.*.*`）推送自动触发 Release 工作流
- 三平台构建：Linux x86_64 / Windows x86_64 (mingw) / macOS x86_64

### 项目重命名

- **仓库**: `Python-Excel-to-PPT-Vocabulary-Tool` → `xlsx-to-pptx`
- **二进制**: `单词卡片转换` → `英语助记卡片生成`

---

## v0.0.1 (2026-02)

Python 原型 — 基础 Excel → PPTX 转换
