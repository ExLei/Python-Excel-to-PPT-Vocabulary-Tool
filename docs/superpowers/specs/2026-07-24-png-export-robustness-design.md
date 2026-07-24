# PNG 导出健壮性重写 — 设计规格

**日期**: 2026-07-24  
**状态**: 已批准  
**关联**: 根因修复 — XML 自闭合元素解析 + ab_glyph→resvg

---

## 1. 问题诊断

| # | 根因 | 影响 |
|---|------|------|
| 1 | `quick_xml` 未处理 `Event::Empty`（自闭合元素如 `<a:off/>` 被忽略） | 所有 PPTX 位置/样式属性丢失，排版异常 |
| 2 | `ab_glyph` 不支持 CJK 字形 | 中文渲染为方框 |
| 3 | 字体硬编码 `"DejaVu Sans"` | 跨平台字体缺失时静默失败（空白 PNG） |
| 4 | 零日志/零诊断 | 无法判断失败发生在哪个环节 |
| 5 | 无字体回退链 | 模板指定字体/emoji 缺失时无替代方案 |
| 6 | 错误静默吞没 | `render_svg_to_png` 成功返回但文字可能未渲染 |

---

## 2. 架构

```
┌──────────────────────────────────────────────────────────────┐
│                       export_png (span)                       │
│                                                              │
│  ┌──────────┐   ┌──────────┐   ┌──────────┐   ┌──────────┐  │
│  │  font    │──→│  pptx    │──→│  svg     │──→│  resvg   │  │
│  │  probe   │   │  parse   │   │  generate│   │  render  │  │
│  └────┬─────┘   └────┬─────┘   └────┬─────┘   └────┬─────┘  │
│       │              │              │              │         │
│       ▼              ▼              ▼              ▼         │
│  ┌──────────────────────────────────────────────────────┐    │
│  │                    DiagStore                          │    │
│  │  · 收集结构化事件（font_probe / placeholder / slide） │    │
│  │  · 持久化到 SQLite (.diag.db)                         │    │
│  │  · tracing span → 实时 stderr (进度/--verbose)       │    │
│  └──────────────────────────────────────────────────────┘    │
└──────────────────────────────────────────────────────────────┘
```

---

## 3. 字体发现与回退

### 3.1 探测策略

不猜字体名。用 `fontdb` 扫描全部系统字体，按**平台优先硬编码列表**探测：

| 链 | 优先级列表 |
|----|-----------|
| Latin | `"Segoe UI"` → `"Helvetica"` → `"Arial"` → `"DejaVu Sans"` |
| CJK | `"Microsoft YaHei"` → `"PingFang SC"` → `"Noto Sans CJK SC"` → `"WenQuanYi Micro Hei"` |
| Emoji | `"Segoe UI Emoji"` → `"Apple Color Emoji"` → `"Noto Color Emoji"` |

每条链找到第一个存在即停止。记录完整探测结果到 `font_probe` 表。

### 3.2 SVG font-family 拼装

```
"${模板中指定的字体}, ${cjk_exists}, ${latin_exists}, ${emoji_exists}, sans-serif"
```

- 模板字体优先（用户意图）
- 随后 CJK/Latin/Emoji 实际存在的字体
- `sans-serif` 兜底（fontdb 的通用族映射）
- **字形级回退由 rustybuzz 自动处理**，不需要我们做字形覆盖分析

### 3.3 fontdb 通用族映射

探测完成后，调用：
```rust
fontdb.set_sans_serif_family(best_latin);
fontdb.set_serif_family(best_latin);
fontdb.set_monospace_family(best_latin);
```

---

## 4. 诊断数据库（SQLite）

### 4.1 Schema

```sql
CREATE TABLE export_session (
    id INTEGER PRIMARY KEY,
    started_at TEXT NOT NULL,
    total_slides INTEGER,
    output_dir TEXT,
    status TEXT  -- 'running' | 'completed' | 'failed'
);

CREATE TABLE font_probe (
    id INTEGER PRIMARY KEY,
    session_id INTEGER REFERENCES export_session(id),
    chain TEXT NOT NULL,       -- 'latin' | 'cjk' | 'emoji'
    query_order INTEGER,       -- 查询序号
    font_name TEXT NOT NULL,   -- 被查询的字体名
    face_count INTEGER,        -- 找到的面数
    selected BOOLEAN,          -- 是否选中
    file_path TEXT             -- 字体文件路径
);

CREATE TABLE placeholder_parse (
    id INTEGER PRIMARY KEY,
    session_id INTEGER REFERENCES export_session(id),
    slide_index INTEGER,       -- 从 0 开始
    placeholder_name TEXT,
    emu_x INTEGER, emu_y INTEGER,
    emu_cx INTEGER, emu_cy INTEGER,
    px_x INTEGER, px_y INTEGER,
    px_w INTEGER,
    font_size_raw INTEGER,     -- hundredths of a point
    font_size_pt REAL,
    color_raw_hex TEXT,
    bold BOOLEAN,
    align_center BOOLEAN
);

CREATE TABLE slide_log (
    id INTEGER PRIMARY KEY,
    session_id INTEGER REFERENCES export_session(id),
    slide_index INTEGER,
    word TEXT,                 -- entry.word
    svg_bytes INTEGER,
    png_bytes INTEGER,
    text_pixels INTEGER,       -- 非白像素数
    text_density_pct REAL,     -- text_pixels / total_pixels * 100
    fields_rendered TEXT,      -- JSON: ["单词","音标"]
    fields_skipped TEXT,       -- JSON: [{"name":"例句","reason":"empty"}]
    sampled_pixels_json TEXT,  -- JSON: [{"x":960,"y":540,"rgba":[0,0,0,255]}]
    duration_ms INTEGER,
    error TEXT                 -- NULL if success
);

CREATE TABLE diagnostic_event (
    id INTEGER PRIMARY KEY,
    session_id INTEGER REFERENCES export_session(id),
    slide_index INTEGER,       -- NULL if session-level
    level TEXT NOT NULL,       -- 'WARN' | 'ERROR'
    category TEXT NOT NULL,    -- 'font' | 'parse' | 'render' | 'verify'
    message TEXT NOT NULL,
    fields_json TEXT,          -- 附加上下文
    resolved BOOLEAN DEFAULT 0
);
```

### 4.2 写入时机

- `export_session`: 导出开始时 INSERT，完成时 UPDATE status
- `font_probe`: 每条链每个字体探测立即 INSERT
- `placeholder_parse`: 每个占位符解析后 INSERT
- `slide_log`: 每张 slide 渲染完成后 INSERT
- `diagnostic_event`: 任何 WARN/ERROR 发生时 INSERT

### 4.3 Agent 诊断查询示例

```sql
-- 哪些 slide 可能空白？
SELECT slide_index, word, text_density_pct
FROM slide_log WHERE text_density_pct < 0.01;

-- 字体回退的完整决策过程
SELECT chain, font_name, face_count, selected, file_path
FROM font_probe ORDER BY chain, query_order;

-- EMU→PX 转换是否有异常值？
SELECT placeholder_name, emu_x, px_x, emu_y, px_y
FROM placeholder_parse WHERE px_x = 0 OR px_y = 0;
```

---

## 5. 终端输出

```
 字体探测: 拉丁→DejaVu Sans  CJK→Noto Sans CJK SC  emoji→缺失
 正在生成: [1/5] apple ........... ✓  (63KB, 0.61%)
 正在生成: [2/5] banana .......... ✓  (58KB, 0.55%)
 正在生成: [3/5] cherry .......... ⚠ 文字极少 → .diag.db
 正在生成: [4/5] dragon fruit .... ✓  (71KB, 0.72%)
 正在生成: [5/5] elderberry ...... ✓  (65KB, 0.58%)

 完成: 5 张 → /tmp/output/
 ⚠ 1 个警告 → 详情: /tmp/output/export_20260724_213001.diag.db
```

- `--verbose`: 输出完整 tracing spans 到 stderr（含所有字段）
- `--diag-db <path>`: 指定 .diag.db 路径（默认输出目录内）

---

## 6. 空白页检测

每张 slide 渲染后：
1. 对 `(w/2, h/2)`、`(w/4, h/4)`、`(3w/4, 3h/4)` 三点采样
2. 统计非白像素数 → `text_pixels`
3. 计算 `text_density_pct = text_pixels / (1920*1080) * 100`
4. < 0.01% → `WARN` event + `slide_log.text_density_pct` 记录
5. < 0.001% 且 SVG 中有 text 元素 → `ERROR` event（大概率字体问题）

---

## 7. 错误处理策略

```
每个阶段都可能失败，但不应阻止后续 slide：
┌──────────┐
│ 字体探测  │ → Err? → 记录 diagnostic_event + 使用 "sans-serif" 继续
└──────────┘
┌──────────┐
│ PPTX 解析 │ → Err? → 记录 + 该 slide 跳过 + 继续下一张
└──────────┘
┌──────────┐
│ SVG 生成  │ → Err? → 逻辑错误，记录 + 跳过
└──────────┘
┌──────────┐
│ PNG 渲染  │ → Err? → 记录 + 跳过
└──────────┘
┌──────────┐
│ 空白检测  │ → WARN/ERROR → 记录 diagnostic_event，不阻止输出
└──────────┘
```

**唯一硬错误**：所有 slide 都失败时才返回 `Err`。

---

## 8. 代码结构

```
crates/core/src/png_export/
├── mod.rs            # 公共 API + export_with_layout + 常量 (~80 行)
├── pipeline.rs       # font_probe → parse → SVG → render → verify (~290 行)
└── diag.rs           # DiagStore + NDJSON 事件定义 (~180 行)

crates/cli/src/
└── lib.rs            # + diag 子命令（--summary, --blank-slides, --font-trace, --errors, --slide）
```

**新增依赖：**
- `tracing` + `tracing-subscriber`（结构化日志，含 json feature）
- `ttf-parser`（直接依赖，复用已编译 — 用于字形覆盖检测）
- 删除 `ab_glyph`

**不引入的依赖：**
- ~~`rusqlite`~~ → 用 `tracing-subscriber` JSON layer 输出 NDJSON，由 `diag` 子命令查询
---

## 9. TDD 测试策略

每模块测试先于实现：

| 模块 | 测试内容 |
|------|---------|
| `font_probe` | 探测已知/未知字体；回退链构建正确；全缺失时不 panic |
| `layout` | Event::Empty 处理；EMU→PX 换算正确；颜色解析 |
| `svg_gen` | 有效 SVG XML；中文/IPA/emoji 嵌入；字段跳过逻辑 |
| `render` | SVG→PNG 魔术字节；尺寸正确；字体缺失时报错非崩溃 |
| `diag` | SQLite 创建/写入/查询；session 生命周期 |
| `verify` | 全白/正常/边缘密度检测 |
| 集成 | CLI `export-png` 端到端；多 slide；模板/无模板；字体缺失场景 |

---

## 10. 自检

- [x] 无 TODO/占位符
- [x] 架构与功能描述一致
- [x] 范围聚焦：PNG 导出管线 + 诊断，不涉及 GUI/CLI 重构
- [x] 无模糊需求

---

规格已写入 `docs/superpowers/specs/2026-07-24-png-export-robustness-design.md`。请审查，修改后我调用 writing-plans 创建实现计划。

---

## 11. 全面代码审计发现（LSP + grep + 调用链分析）

### 11.1 调用链全景

```
CLI: run() → Command::ExportPng                          GUI: start_export_png()
  │                                                         │
  ├─ export_with_template(-t)  ─┐                           ├─ export_entries_to_png()
  └─ export_entries_to_png()  ─┤                           │
                                ▼                            │
                         export_with_layout()  ◄─────────────┘
                                │
                    ┌───────────┼───────────┐
                    ▼                       ▼
            render_slide_to_svg()    render_svg_to_png()
                    │                       │
                    ▼                       ▼
              xml_escape()           resvg + usvg + fontdb
```

**生产入口：** CLI 2 处（`lib.rs:310,312`），GUI 1 处（`app.rs:285`）  
**公共 API：** 7 个符号（5 fn + 1 struct + 1 enum）  
**测试覆盖：** 11 个测试（`png_export_test.rs`）

### 11.2 审计发现的缺口

| ID | 严重度 | 问题 | 状态 |
|----|-------|------|------|
| G1 | Low | `PngExportError::NoFontFound` 从未构造（死变体） | 本次重写修复：字体探测后使用 |
| G2 | Low | `SpState.cy` 设置但从未写入 `PlaceholderLayout` | 添加 `h` 字段或移除 |
| G3 | **High** | `collect_attrs` 静默吞没所有 XML 解析失败（x/y/cx/cy/sz/color） | 本次重写：每项失败 emit `warn!` 到 diag |
| G4 | **High** | `render_svg_to_png` 无渲染后验证（字体缺失也不报错） | 本次重写：空白页检测 + 字体覆盖率检查 |
| G5 | Medium | GUI PNG 导出无取消机制（对比 PPT 生成有 `cancel_flag`） | 本次重写：添加 `AtomicBool` + 进度回调 |
| G6 | Low | `parse_template` 硬编码 `slide1.xml` | 保持（设计意图） |
| G7 | Medium | 部分布局静默丢弃未列出字段 | 本次重写：`slide_log.fields_skipped` 记录 |
| G8 | Low | EMU 值无范围校验（恶意 PPTX 可能产生垃圾坐标） | 本次重写：边界检查 + warn |
| G9 | Low | GUI `create_dir_all` 错误被 `let _ =` 丢弃 | 本次重写：修复为 `?` 传播 |
| G10 | Low | CLI 无 Ctrl+C 信号处理 | 暂缓（工具性质决定） |

### 11.3 公共 API 影响评估

**保持不变的接口**（CLI/GUI 无需修改）：
- `export_entries_to_png(entries, output_dir)` — 签名不变
- `export_with_template(entries, tpl_path, output_dir)` — 签名不变
- `parse_template(tpl_path)` — 签名不变
- `PlaceholderLayout` — 字段不变
- `PngExportError` — 变体不变

**新增公共接口**：
- `render_slide_to_svg(entry, layout)` → `String` — 已在 
- `render_svg_to_png(svg)` → `Result<Vec<u8>, PngExportError>` — 已在
- `export_with_layout(...)` 内部签名增加 `cancel_flag: &AtomicBool` 和 `diag: &mut DiagStore`（私有不影响外部）

**调用方无需修改** — CLI 和 GUI 的 import 和调用点完全兼容。
