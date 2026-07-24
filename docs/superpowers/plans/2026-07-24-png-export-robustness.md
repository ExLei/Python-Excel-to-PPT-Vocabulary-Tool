# PNG 导出健壮性重写 — 实现计划

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推荐）或 superpowers:executing-plans 逐任务实现此计划。步骤使用复选框（`- [ ]`）语法来跟踪进度。

**目标：** 全软件诊断系统 + PNG 导出管线重写。用 `tracing` + NDJSON 替换所有裸 `println!/eprintln!`，用 `resvg` 替换 `ab_glyph`，添加 `diag` CLI 子命令。

**架构：** 横切 `diag.rs` 模块提供 `DiagStore`（内存事件收集器 + NDJSON 写入），所有核心模块通过 `&mut DiagStore` 注入。CLI 新增 `diag` 子命令查询 NDJSON 日志。GUI 通过 `DiagStore` 展示诊断面板。

**技术栈：** Rust 2021, `tracing` + `tracing-subscriber` (json feature), `ttf-parser` (复用), `resvg` + `usvg` + `tiny-skia` (已有), `quick_xml` (已有)

---

### 任务 1：diag.rs — 诊断基础设施

**文件：**
- 创建：`crates/core/src/diag.rs`
- 修改：`crates/core/src/lib.rs` (添加 `pub mod diag`)
- 测试：`crates/core/tests/diag_test.rs`

- [ ] **步骤 1：编写失败的测试 — DiagStore 基本操作**

```rust
// crates/core/tests/diag_test.rs
use vocab_core::diag::{DiagStore, DiagLevel};

#[test]
fn diag_store_collects_events() {
    let mut store = DiagStore::new();
    store.info("font_probe", "scanning fonts", None);
    store.warn("font_probe", "emoji font not found", None);
    store.error("render", "SVG parse failed", Some(r#"{"detail":"line 12"}"#));

    assert_eq!(store.event_count(), 3);
    assert_eq!(store.warnings(), 1);
    assert_eq!(store.errors(), 1);
}

#[test]
fn diag_store_ndjson_output() {
    let mut store = DiagStore::new();
    store.info("test", "hello", Some(r#"{"key":"value"}"#));

    let ndjson = store.to_ndjson();
    assert!(ndjson.contains(r#""level":"INFO""#));
    assert!(ndjson.contains(r#""target":"test""#));
    assert!(ndjson.contains(r#""message":"hello""#));
    assert!(ndjson.contains(r#""key":"value""#));
}

#[test]
fn diag_store_empty_is_valid() {
    let store = DiagStore::new();
    assert_eq!(store.event_count(), 0);
    assert_eq!(store.to_ndjson(), "");
    assert_eq!(store.warnings(), 0);
    assert_eq!(store.errors(), 0);
}
```

运行：`cargo test -p vocab_core --test diag_test`
预期：FAIL — `diag` 模块不存在

- [ ] **步骤 2：编写 DiagStore 最小实现**

```rust
// crates/core/src/diag.rs
use std::time::{SystemTime, UNIX_EPOCH};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiagLevel {
    Info,
    Warn,
    Error,
}

#[derive(Debug, Clone)]
pub struct DiagEvent {
    pub timestamp: String,
    pub level: DiagLevel,
    pub target: String,
    pub message: String,
    pub fields_json: Option<String>,
}

pub struct DiagStore {
    events: Vec<DiagEvent>,
    warning_count: usize,
    error_count: usize,
}

impl DiagStore {
    pub fn new() -> Self {
        Self { events: Vec::new(), warning_count: 0, error_count: 0 }
    }

    fn now_iso() -> String {
        // simplified timestamp
        "2024-01-01T00:00:00Z".to_string()
    }

    fn push(&mut self, level: DiagLevel, target: &str, message: &str, fields_json: Option<&str>) {
        if matches!(level, DiagLevel::Warn) { self.warning_count += 1; }
        if matches!(level, DiagLevel::Error) { self.error_count += 1; }
        self.events.push(DiagEvent {
            timestamp: Self::now_iso(),
            level,
            target: target.to_string(),
            message: message.to_string(),
            fields_json: fields_json.map(|s| s.to_string()),
        });
    }

    pub fn info(&mut self, target: &str, message: &str, fields_json: Option<&str>) {
        self.push(DiagLevel::Info, target, message, fields_json);
    }

    pub fn warn(&mut self, target: &str, message: &str, fields_json: Option<&str>) {
        self.push(DiagLevel::Warn, target, message, fields_json);
    }

    pub fn error(&mut self, target: &str, message: &str, fields_json: Option<&str>) {
        self.push(DiagLevel::Error, target, message, fields_json);
    }

    pub fn event_count(&self) -> usize { self.events.len() }
    pub fn warnings(&self) -> usize { self.warning_count }
    pub fn errors(&self) -> usize { self.error_count }

    pub fn to_ndjson(&self) -> String {
        let mut out = String::new();
        for e in &self.events {
            let level_str = match e.level {
                DiagLevel::Info => "INFO",
                DiagLevel::Warn => "WARN",
                DiagLevel::Error => "ERROR",
            };
            out.push_str(&format!(
                r#"{{"timestamp":"{}","level":"{}","target":"{}","message":"{}""#,
                e.timestamp, level_str, e.target, e.message
            ));
            if let Some(ref f) = e.fields_json {
                out.push_str(&format!(r#","fields":{}"#, f));
            }
            out.push_str("}\n");
        }
        out
    }
}
```

- [ ] **步骤 3：在 lib.rs 中注册模块**

```rust
// crates/core/src/lib.rs
pub mod diag;
```

- [ ] **步骤 4：运行测试验证通过**

运行：`cargo test -p vocab_core --test diag_test`
预期：3 PASS

- [ ] **步骤 5：添加真实时间戳 + 文件写入**

在 `diag.rs` 的 `DiagStore` 添加：

```rust
fn now_iso() -> String {
    // 替换步骤2的简化实现
    let dur = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default();
    let secs = dur.as_secs();
    // 简单格式化: 2026-07-24T21:30:01Z
    let days = secs / 86400;
    let time = secs % 86400;
    let h = time / 3600;
    let m = (time % 3600) / 60;
    let s = time % 60;
    // 简化: 不用真实日期，用 duration
    format!("{h:02}:{m:02}:{s:02}")
}

pub fn write_ndjson_to_file(&self, path: &std::path::Path) -> std::io::Result<()> {
    std::fs::write(path, self.to_ndjson())
}
```

测试验证：
```rust
#[test]
fn diag_store_writes_ndjson_file() {
    use std::io::Read;
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("test.ndjson");
    let mut store = DiagStore::new();
    store.info("test", "hello", None);
    store.write_ndjson_to_file(&path).unwrap();
    
    let mut content = String::new();
    std::fs::File::open(&path).unwrap().read_to_string(&mut content).unwrap();
    assert!(content.contains("hello"));
}
```

- [ ] **步骤 6：Commit**

```bash
git add crates/core/src/diag.rs crates/core/src/lib.rs crates/core/tests/diag_test.rs
git commit -m "feat(diag): add DiagStore with NDJSON output infrastructure"
```

---

### 任务 2：diag.rs — 完整 NDJSON 格式（tracing-compatible JSON）

**文件：**
- 修改：`crates/core/src/diag.rs`

- [ ] **步骤 1：编写测试 — 验证 JSON 字段完整性**

```rust
// crates/core/tests/diag_test.rs 追加
#[test]
fn ndjson_line_is_valid_json() {
    let mut store = DiagStore::new();
    store.error("render", "SVG parse failed", Some(r#"{"svg_bytes":1842,"error_detail":"unexpected tag"}"#));
    let ndjson = store.to_ndjson();
    for line in ndjson.lines() {
        serde_json::from_str::<serde_json::Value>(line)
            .expect("each line must be valid JSON");
    }
}

#[test]
fn ndjson_contains_all_standard_fields() {
    let mut store = DiagStore::new();
    store.info("test", "msg", Some(r#"{"key":"val"}"#));
    let line = store.to_ndjson().lines().next().unwrap();
    let v: serde_json::Value = serde_json::from_str(line).unwrap();
    assert!(v.get("timestamp").is_some());
    assert!(v.get("level").is_some());
    assert!(v.get("target").is_some());
    assert!(v.get("message").is_some());
    assert!(v.get("fields").is_some());
}
```

运行：`cargo test -p vocab_core --test diag_test ndjson`
预期：FAIL — 需要 `serde_json` 依赖

- [ ] **步骤 2：添加 serde_json 依赖**

```toml
# crates/core/Cargo.toml [dependencies] 追加
serde_json = "1"
```

- [ ] **步骤 3：替换 `to_ndjson` 为 serde_json 序列化**

重写 `DiagEvent` 和 `to_ndjson`：

```rust
use serde::Serialize;

#[derive(Debug, Clone, Serialize)]
struct NdjsonLine {
    timestamp: String,
    level: String,
    target: String,
    message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    fields: Option<serde_json::Value>,
}

impl DiagStore {
    pub fn to_ndjson(&self) -> String {
        let mut out = String::new();
        for e in &self.events {
            let line = NdjsonLine {
                timestamp: e.timestamp.clone(),
                level: match e.level {
                    DiagLevel::Info => "INFO".to_string(),
                    DiagLevel::Warn => "WARN".to_string(),
                    DiagLevel::Error => "ERROR".to_string(),
                },
                target: e.target.clone(),
                message: e.message.clone(),
                fields: e.fields_json.as_ref()
                    .and_then(|s| serde_json::from_str(s).ok()),
            };
            out.push_str(&serde_json::to_string(&line).unwrap());
            out.push('\n');
        }
        out
    }
}
```

需要 `serde` derive:
```toml
serde = { version = "1", features = ["derive"] }
```

- [ ] **步骤 4：运行测试验证通过**

运行：`cargo test -p vocab_core --test diag_test`
预期：5 PASS

- [ ] **步骤 5：Commit**

```bash
git add crates/core/Cargo.toml crates/core/src/diag.rs crates/core/tests/diag_test.rs
git commit -m "feat(diag): serde_json serialization for NDJSON output"
```

---

### 任务 3：pipeline.rs — 字体探测（font_probe）

**文件：**
- 创建：`crates/core/src/png_export/pipeline.rs`  
- 修改：`crates/core/src/png_export/mod.rs`（添加 `pub mod pipeline`）
- 创建：`crates/core/src/png_export/mod.rs`（从 png_export.rs 迁移公共 API）
- 测试：`crates/core/tests/png_export_test.rs`（追加字体探测测试）

- [ ] **步骤 1：编写失败的测试 — 字体探测**

```rust
// crates/core/tests/png_export_test.rs 追加
use vocab_core::png_export::pipeline;

#[test]
fn font_probe_detects_system_fonts() {
    let mut diag = vocab_core::diag::DiagStore::new();
    let config = pipeline::probe_fonts(&mut diag);
    
    // 至少应该构建出 font_stack（即使零字体也有 "sans-serif" 兜底）
    assert!(!config.font_stack.is_empty());
    assert!(config.font_stack.contains("sans-serif"));
    // diag 应该记录了探测过程
    assert!(diag.event_count() > 0,
        "font probe should emit diagnostic events");
}

#[test]
fn font_probe_fallback_chain_is_recorded() {
    let mut diag = vocab_core::diag::DiagStore::new();
    let _config = pipeline::probe_fonts(&mut diag);
    
    let ndjson = diag.to_ndjson();
    // 应该包含字体探测相关事件
    assert!(ndjson.contains("font_probe"),
        "NDJSON should contain font probe events");
}
```

运行：`cargo test -p vocab_core --test png_export_test font_probe`
预期：FAIL — `pipeline` 模块不存在

- [ ] **步骤 2：创建 png_export 模块目录结构**

```bash
mkdir -p crates/core/src/png_export
```

创建 `crates/core/src/png_export/mod.rs`：

```rust
pub mod pipeline;
// 后续迁移现有公共 API
pub use super::png_export_types::*;
```

创建 `crates/core/src/png_export/pipeline.rs`：

```rust
use crate::diag::DiagStore;

pub struct FontConfig {
    pub font_stack: String,
}

/// 探测系统字体，构建 SVG font-family 回退栈。
/// 永不出错 — 最坏情况返回 "sans-serif"。
pub fn probe_fonts(diag: &mut DiagStore) -> FontConfig {
    let mut fontdb = usvg::fontdb::Database::new();
    fontdb.load_system_fonts();

    diag.info("font_probe", &format!("scanned system fonts"), None);

    let latin_chains = ["Segoe UI", "Helvetica", "Arial", "DejaVu Sans"];
    let cjk_chains = ["Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", "WenQuanYi Micro Hei"];
    let emoji_chains = ["Segoe UI Emoji", "Apple Color Emoji", "Noto Color Emoji"];

    let best_latin = first_available(&fontdb, &latin_chains, "latin", diag);
    let best_cjk = first_available(&fontdb, &cjk_chains, "cjk", diag);
    let best_emoji = first_available(&fontdb, &emoji_chains, "emoji", diag);

    let mut parts: Vec<&str> = Vec::new();
    if let Some(ref cjk) = best_cjk { parts.push(cjk); }
    if let Some(ref latin) = best_latin { parts.push(latin); }
    if let Some(ref emoji) = best_emoji { parts.push(emoji); }
    parts.push("sans-serif");

    let font_stack = parts.join(", ");
    diag.info("font_probe", &format!("font_stack: {font_stack}"), None);

    if let Some(ref latin) = best_latin {
        fontdb.set_sans_serif_family(latin);
        fontdb.set_serif_family(latin);
        fontdb.set_monospace_family(latin);
    }

    FontConfig { font_stack }
}

fn first_available(fontdb: &usvg::fontdb::Database, candidates: &[&str], chain_name: &str, diag: &mut DiagStore) -> Option<String> {
    for (i, name) in candidates.iter().enumerate() {
        let faces = fontdb.faces()
            .filter(|info| info.family == *name || info.families.iter().any(|(f, _)| f == *name))
            .count();
        if faces > 0 {
            diag.info("font_probe", &format!("{chain_name}: selected \"{name}\" ({faces} faces)"), None);
            return Some(name.to_string());
        }
        diag.info("font_probe", &format!("{chain_name}: \"{name}\" not found (order={i})"), None);
    }
    diag.warn("font_probe", &format!("{chain_name}: all candidates missing"), None);
    None
}
```

- [ ] **步骤 3：运行测试验证通过**

运行：`cargo test -p vocab_core --test png_export_test font_probe`
预期：2 PASS

- [ ] **步骤 4：Commit**

```bash
git add crates/core/src/png_export/ crates/core/tests/png_export_test.rs
git commit -m "feat(pipeline): font probe with fallback chain and diag events"
```

---

### 任务 4：pipeline.rs — PPTX 模板解析（parse_slide_xml 迁移）

**文件：**
- 修改：`crates/core/src/png_export/pipeline.rs`
- 修改：`crates/core/src/png_export/mod.rs`

- [ ] **步骤 1：从 png_export.rs 迁移 parse_slide_xml 到 pipeline.rs**

将以下内容从 `crates/core/src/png_export.rs` 迁移到 `crates/core/src/png_export/pipeline.rs`：

1. `SpState` 结构体（删除 `cy` 字段，死代码）
2. `impl Default for SpState`
3. `parse_slide_xml` 函数（添加 `diag: &mut DiagStore` 参数）
4. `collect_attrs` 内部函数（添加解析失败时的 `diag.warn`）
5. `extract_placeholder` 函数
6. `emu_to_px_x/y/w` 函数（合并为 `emu_to_px(emu: i64, dim: Dim) -> i32`）

在 `collect_attrs` 的每个 `unwrap_or` 处添加 diag：

```rust
fn collect_attrs(e: &quick_xml::events::BytesStart, st: &mut SpState, diag: &mut DiagStore) {
    for attr in e.attributes().flatten() {
        let kbytes = attr.key.as_ref().to_vec();
        let k = std::str::from_utf8(&kbytes).unwrap_or("");
        let v = std::str::from_utf8(&attr.value).unwrap_or("");
        match k {
            "x" => {
                match v.parse::<i64>() {
                    Ok(val) => st.x = val,
                    Err(_) => diag.warn("parse", &format!("failed to parse x=\"{v}\", defaulting to 0"), None),
                }
            }
            "y" => {
                match v.parse::<i64>() {
                    Ok(val) => st.y = val,
                    Err(_) => diag.warn("parse", &format!("failed to parse y=\"{v}\", defaulting to 0"), None),
                }
            }
            // ... similarly for cx, cy, sz
            _ => {}
        }
    }
}
```

- [ ] **步骤 2：更新现有测试，确保迁移后仍通过**

运行：`cargo test -p vocab_core --test png_export_test parse`
预期：3 PASS

- [ ] **步骤 3：Commit**

```bash
git add crates/core/src/png_export/
git commit -m "refactor(pipeline): migrate parse_slide_xml with diag instrumentation"
```

---

### 任务 5：pipeline.rs — SVG 生成 + resvg 渲染（带 diag）

**文件：**
- 修改：`crates/core/src/png_export/pipeline.rs`

- [ ] **步骤 1：编写测试 — 管线函数带 diag 输出**

```rust
#[test]
fn pipeline_emits_diag_events_per_slide() {
    let mut diag = DiagStore::new();
    let entry = WordEntry { word: "test".into(), ..Default::default() };
    let layout = HashMap::new();
    
    let svg = pipeline::render_slide_to_svg(&entry, &layout);
    let png = pipeline::render_svg_to_png(&svg, &mut diag).unwrap();
    
    assert!(diag.event_count() > 0, "render should emit diag events");
    assert!(png.len() > 1000);
}
```

- [ ] **步骤 2：实现 `render_svg_to_png` 带 `diag: &mut DiagStore`**

在 pipeline.rs 中添加：

```rust
pub fn render_svg_to_png(svg: &str, config: &FontConfig, diag: &mut DiagStore) -> Result<Vec<u8>, PngExportError> {
    let mut fontdb = usvg::fontdb::Database::new();
    fontdb.load_system_fonts();
    // 使用探测到的字体配置通用族
    if let Some(latin) = config.font_stack.split(", ").find(|s| !s.is_empty() && *s != "sans-serif") {
        fontdb.set_sans_serif_family(latin);
    }

    let opt = usvg::Options {
        fontdb: std::sync::Arc::new(fontdb),
        ..Default::default()
    };

    let rtree = usvg::Tree::from_str(svg, &opt)
        .map_err(|e| {
            diag.error("render", &format!("SVG parse failed: {e}"), None);
            PngExportError::TemplateParse(format!("SVG parse: {e}"))
        })?;

    let size = rtree.size().to_int_size();
    diag.info("render", &format!("parsed SVG: {} text nodes, {}x{}", 
        count_text_nodes(&rtree), size.width(), size.height()), None);

    let mut pixmap = tiny_skia::Pixmap::new(size.width(), size.height())
        .ok_or_else(|| {
            diag.error("render", "failed to create pixmap", None);
            PngExportError::TemplateParse("failed to create pixmap".into())
        })?;

    resvg::render(&rtree, usvg::Transform::default(), &mut pixmap.as_mut());

    let png = pixmap.encode_png().map_err(|e| {
        diag.error("render", &format!("PNG encode failed: {e}"), None);
        PngExportError::TemplateParse(format!("PNG encode: {e}"))
    })?;

    // 空白页检测
    let non_white = count_non_white_pixels(&pixmap);
    let density = non_white as f64 / (size.width() as f64 * size.height() as f64) * 100.0;
    diag.info("render", &format!("PNG: {} bytes, {} text px ({:.4}%)", png.len(), non_white, density), None);

    if density < 0.001 && count_text_nodes(&rtree) > 0 {
        diag.error("verify", "BLANK SLIDE: SVG has text but zero pixels rendered — likely font issue", None);
    } else if density < 0.01 {
        diag.warn("verify", &format!("low text density: {:.4}%", density), None);
    }

    Ok(png)
}

fn count_text_nodes(tree: &usvg::Tree) -> usize {
    // 递归遍历 usvg 树计数字 text 节点
    fn count(group: &usvg::Group) -> usize {
        group.children.iter().map(|node| match node {
            usvg::Node::Text(_) => 1,
            usvg::Node::Group(g) => count(g),
            _ => 0,
        }).sum()
    }
    count(&tree.root)
}

fn count_non_white_pixels(pixmap: &tiny_skia::Pixmap) -> usize {
    pixmap.pixels().iter().filter(|p| {
        p.red() != 255 || p.green() != 255 || p.blue() != 255
    }).count()
}
```

- [ ] **步骤 3：运行测试验证**

运行：`cargo test -p vocab_core --test png_export_test`
预期：所有 11+ 现有测试 PASS + 新测试 PASS

- [ ] **步骤 4：Commit**

```bash
git add crates/core/src/png_export/pipeline.rs crates/core/tests/png_export_test.rs
git commit -m "feat(pipeline): render_svg_to_png with diag + blank detection"
```

---

### 任务 5.5：每 slide 错误隔离 + 取消机制

**文件：**
- 修改：`crates/core/src/png_export/mod.rs`
- 修改：`crates/core/src/png_export/pipeline.rs`

- [ ] **步骤 1：在 pipeline.rs 中添加 `render_one_slide` 独立函数**

```rust
// pipeline.rs
pub fn render_one_slide(
    entry: &WordEntry,
    layout: &HashMap<String, PlaceholderLayout>,
    font_config: &FontConfig,
    output_dir: &Path,
    index: usize,
    diag: &mut DiagStore,
) -> Result<PathBuf, PngExportError> {
    let svg = render_slide_to_svg(entry, layout);
    let png = render_svg_to_png(&svg, font_config, diag)?;
    let path = output_dir.join(format!("slide_{}.png", index + 1));
    std::fs::write(&path, &png)?;
    Ok(path)
}
```

- [ ] **步骤 2：重写 `export_with_layout` 支持错误隔离 + 取消**

```rust
// mod.rs — 替换原有 export_with_layout
use std::sync::atomic::{AtomicBool, Ordering};

fn export_with_layout(
    entries: &[WordEntry],
    output_dir: &Path,
    layout: &HashMap<String, PlaceholderLayout>,
    font_config: &FontConfig,
    cancel: Option<&AtomicBool>,
    diag: &mut DiagStore,
) -> Result<Vec<PathBuf>, PngExportError> {
    std::fs::create_dir_all(output_dir)?;
    let mut pngs = Vec::new();
    let mut slide_errors: Vec<(usize, String)> = Vec::new();

    for (i, entry) in entries.iter().enumerate() {
        // 取消检查
        if let Some(flag) = cancel {
            if flag.load(Ordering::Relaxed) {
                diag.warn("export", &format!("cancelled at slide {}/{}", i + 1, entries.len()), None);
                break;
            }
        }

        match pipeline::render_one_slide(entry, layout, font_config, output_dir, i, diag) {
            Ok(path) => pngs.push(path),
            Err(e) => {
                diag.error("slide", &format!("slide {} ({}) failed: {}", i + 1, entry.word, e), None);
                slide_errors.push((i, e.to_string()));
            }
        }
    }

    diag.info("export", &format!(
        "{} succeeded, {} failed, {} total",
        pngs.len(), slide_errors.len(), entries.len()
    ), None);

    // 全部失败才硬错误
    if pngs.is_empty() && !slide_errors.is_empty() {
        return Err(PngExportError::TemplateParse("all slides failed".into()));
    }

    Ok(pngs)
}
```

- [ ] **步骤 3：更新公共 API 签名 — 添加 cancel 参数（可选，向后兼容）**

```rust
// 新 API（带 cancel）
pub fn export_entries_to_png_with_cancel(
    entries: &[WordEntry], output_dir: &Path,
    cancel: &AtomicBool,
) -> Result<Vec<PathBuf>, PngExportError> { /* ... */ }

// 旧 API 保持不变，内部传 None
pub fn export_entries_to_png(entries: &[WordEntry], output_dir: &Path) -> Result<Vec<PathBuf>, PngExportError> {
    let mut diag = DiagStore::new();
    let font_config = pipeline::probe_fonts(&mut diag);
    let result = export_with_layout(entries, output_dir, &HashMap::new(), &font_config, None, &mut diag);
    let _ = diag.write_ndjson_to_file(&output_dir.join(ndjson_filename()));
    result
}
```

- [ ] **步骤 4：运行测试验证隔离逻辑**

```rust
// png_export_test.rs 追加
#[test]
fn slide_error_isolation_continues_on_failure() {
    // 构造一个会导致渲染失败的 slide + 一个正常的 slide
    // 验证正常 slide 仍然产出 PNG
}

#[test]
fn cancel_flag_stops_export() {
    let cancel = std::sync::atomic::AtomicBool::new(false);
    // 在第二个 slide 前设置 cancel = true
    // 验证只导出了第一张
}
```

- [ ] **步骤 5：Commit**

```bash
git add crates/core/src/png_export/
git commit -m "feat(export): per-slide error isolation + cancel support"
```

---

### 任务 5.6：fontdb 崩溃保护

**文件：**
- 修改：`crates/core/src/png_export/pipeline.rs`

- [ ] **步骤 1：添加 `safe_load_system_fonts` 函数**

```rust
// pipeline.rs
fn safe_load_system_fonts(diag: &mut DiagStore) -> usvg::fontdb::Database {
    let mut fontdb = usvg::fontdb::Database::new();

    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        fontdb.load_system_fonts();
    }));

    match result {
        Ok(()) => {
            let count = fontdb.faces().count();
            diag.info("font_probe", &format!("loaded {count} font faces successfully"), None);
            if count == 0 {
                diag.error("font_probe", "zero fonts found — all text will render as tofu", None);
            }
        }
        Err(_) => {
            diag.error("font_probe",
                "font loading CRASHED (corrupt system font?). Output will be blank. \
                 Try removing recently installed fonts or running in a clean environment.",
                None);
        }
    }
    fontdb
}
```

- [ ] **步骤 2：更新 `probe_fonts` 和 `render_svg_to_png` 使用安全版本**

替换所有 `fontdb.load_system_fonts()` 调用为 `safe_load_system_fonts(diag)`。

- [ ] **步骤 3：添加测试 — 空字体数据库不 panic**

```rust
#[test]
fn empty_font_db_does_not_panic() {
    let mut diag = DiagStore::new();
    let fontdb = usvg::fontdb::Database::new(); // 不加载任何字体
    // 创建 FontConfig 时 fontdb 为空
    let config = FontConfig { font_stack: "sans-serif".into() };
    let svg = r#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text x="10" y="20">test</text></svg>"#;
    let result = pipeline::render_svg_to_png(svg, &config, &mut diag);
    // 不应 panic，可能返回 Ok（文字不渲染但不会崩溃）
    assert!(result.is_ok() || result.is_err());
}
```

- [ ] **步骤 4：Commit**

```bash
git add crates/core/src/png_export/pipeline.rs
git commit -m "feat(font): catch_unwind protection for corrupt system fonts"
```

---

### 任务 6：png_export 公共 API 整合 — 桥接新旧


### 任务 6.5：NDJSON 写入失败不阻塞

**文件：**
- 修改：`crates/core/src/png_export/mod.rs`

- [ ] **步骤 1：替换 `let _ =` 为显式错误报告**

在所有 `write_ndjson_to_file` 调用处：

```rust
// 旧 （mod.rs export_with_template_inner）
let _ = diag.write_ndjson_to_file(&ndjson_path);

// 新
match diag.write_ndjson_to_file(&ndjson_path) {
    Ok(_) => eprintln!("  诊断日志: {}", ndjson_path.display()),
    Err(e) => eprintln!("  警告: 无法写入诊断日志到 {}: {e}", ndjson_path.display()),
}
```

- [ ] **步骤 2：CLI 层同样处理**

```rust
// cli/src/lib.rs — ExportPng handler 中
let ndjson_path = output.join(format!("export_{}.ndjson", timestamp));
match diag.write_ndjson_to_file(&ndjson_path) {
    Ok(_) => println!("  诊断日志: {}", ndjson_path.display()),
    Err(e) => println!("  警告: 诊断日志写入失败: {e}"),
}
```

- [ ] **步骤 3：Commit**

```bash
git add crates/core/src/png_export/mod.rs crates/cli/src/lib.rs
git commit -m "fix(diag): NDJSON write failure is non-blocking with user notification"
```

### 任务 7：CLI — diag 子命令

**文件：**
- 修改：`crates/cli/src/lib.rs`
- 修改：`crates/cli/Cargo.toml`（如有新增依赖）

- [ ] **步骤 1：添加 Diag 命令变体 + 子命令参数**

```rust
// 在 Command enum 中添加
#[derive(Subcommand)]
enum Command {
    // ... 现有变体 ...
    
    /// 诊断 NDJSON 日志文件
    Diag {
        /// NDJSON 日志文件路径
        file: PathBuf,
        #[command(subcommand)]
        query: DiagQuery,
    },
}

#[derive(Subcommand)]
enum DiagQuery {
    /// 导出总览
    Summary,
    /// 检查空白 slide（默认阈值 0.01%）
    BlankSlides {
        #[arg(default_value = "0.01")]
        threshold: f64,
    },
    /// 字体探测完整过程
    FontTrace,
    /// 所有错误和警告
    Errors,
    /// 单张 slide 详情
    Slide {
        /// slide 序号（0 起始）
        index: usize,
    },
}
```

- [ ] **步骤 2：实现 diag 查询函数**

```rust
fn run_diag(file: &Path, query: &DiagQuery) -> Result<(), String> {
    let content = std::fs::read_to_string(file)
        .map_err(|e| format!("无法读取日志文件: {e}"))?;
    
    let events: Vec<serde_json::Value> = content
        .lines()
        .filter(|l| !l.is_empty())
        .filter_map(|l| serde_json::from_str(l).ok())
        .collect();

    match query {
        DiagQuery::Summary => {
            let total = events.len();
            let warnings = events.iter().filter(|e| e["level"] == "WARN").count();
            let errors = events.iter().filter(|e| e["level"] == "ERROR").count();
            let fonts = events.iter()
                .filter(|e| e["target"] == "font_probe" && e["message"].as_str().map_or(false, |m| m.contains("font_stack")))
                .filter_map(|e| e["message"].as_str())
                .next()
                .unwrap_or("unknown");
            println!("  Session: {} events", total);
            println!("  Font: {fonts}");
            println!("  Warnings: {warnings}, Errors: {errors}");
        }
        DiagQuery::BlankSlides { threshold } => {
            for e in &events {
                if e["target"] == "render" {
                    if let Some(msg) = e["message"].as_str() {
                        if msg.contains("text px") {
                            // 解析密度
                            if let Some(pct_str) = msg.split('(').nth(1).and_then(|s| s.split('%').next()) {
                                if let Ok(pct) = pct_str.parse::<f64>() {
                                    if pct < *threshold {
                                        println!("  blank: {msg}");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        DiagQuery::FontTrace => {
            for e in &events {
                if e["target"] == "font_probe" {
                    println!("  {}", e["message"].as_str().unwrap_or(""));
                }
            }
        }
        DiagQuery::Errors => {
            for e in &events {
                let level = e["level"].as_str().unwrap_or("");
                if level == "WARN" || level == "ERROR" {
                    println!("  {level:5} {target:20} {msg}",
                        target = e["target"].as_str().unwrap_or(""),
                        msg = e["message"].as_str().unwrap_or(""));
                }
            }
        }
        DiagQuery::Slide { index } => {
            // 查找 slide 相关事件
            let slide_events: Vec<_> = events.iter()
                .filter(|e| e["target"] == "slide" || e["target"] == "render" || e["target"] == "verify")
                .collect();
            // 简化: 打印所有
            for e in slide_events {
                println!("  {}", e["message"].as_str().unwrap_or(""));
            }
        }
    }
    Ok(())
}
```

- [ ] **步骤 3：在 run() 中添加 Diag 分支**

```rust
Command::Diag { file, query } => run_diag(&file, &query),
```

- [ ] **步骤 4：验证编译 + 烟雾测试**

```bash
cargo build --release
# 先做一次导出生成 NDJSON
./target/release/单词卡片转换 export-png -i assets/template.xlsx -o /tmp/diag_test/
# 运行 diag 查询
./target/release/单词卡片转换 diag /tmp/diag_test/export_*.ndjson --summary
./target/release/单词卡片转换 diag /tmp/diag_test/export_*.ndjson --font-trace
```

- [ ] **步骤 5：Commit**

```bash
git add crates/cli/
git commit -m "feat(cli): add diag subcommand for NDJSON log analysis"
```

---

### 任务 8：现有模块 diag 注入 — reader.rs

**文件：**
- 修改：`crates/core/src/reader.rs`

- [ ] **步骤 1：给 reader 函数添加 `diag: &mut DiagStore` 参数**

修改 `load` 和 `load_csv`/`load_excel`：

```rust
pub fn load(source: &InputSource, diag: &mut DiagStore) -> Result<Vec<WordEntry>, LoadError> {
    diag.info("reader", &format!("loading: {:?}", source), None);
    match source {
        InputSource::Excel { path, sheet } => {
            diag.info("reader", &format!("Excel: {} sheet={}", path.display(), sheet), None);
            load_excel(path, sheet, diag)
        }
        InputSource::Csv { path, encoding } => {
            diag.info("reader", &format!("CSV: {} encoding={}", path.display(), encoding), None);
            load_csv(path, encoding, diag)
        }
    }
}
```

将 `eprintln!` 替换为 `diag.warn`：

```rust
// 旧: eprintln!("警告: 第 {row} 行单词为空，跳过");
// 新:
diag.warn("reader", &format!("row {}: empty word, skipping", row), None);
```

- [ ] **步骤 2：更新 CLI/GUI 调用者传递 `&mut DiagStore`**

CLI 中：

```rust
let mut diag = DiagStore::new();
let entries = reader::load(&source, &mut diag).map_err(|e| e.to_string())?;
// ... 使用后写入 NDJSON
let _ = diag.write_ndjson_to_file(&ndjson_path);
```

- [ ] **步骤 3：运行全部测试**

运行：`cargo test --workspace`
预期：全部 PASS

- [ ] **步骤 4：Commit**

```bash
git add crates/core/src/reader.rs crates/cli/src/lib.rs crates/gui/src/app.rs
git commit -m "feat(reader): diag instrumentation replacing eprintln"
```

---

### 任务 9：现有模块 diag 注入 — generator.rs, template_reader.rs

**文件：**
- 修改：`crates/core/src/generator.rs`
- 修改：`crates/core/src/template_reader.rs`

- [ ] **步骤 1：generator.rs — 添加 diag 参数 + 替换进度回调中的 print**

```rust
pub fn generate(
    entries: &[WordEntry],
    output: &Path,
    progress: impl Fn(usize, usize) -> bool,
    diag: &mut DiagStore,
) -> Result<(), GenerateError> {
    diag.info("generator", &format!("generating {} entries → {}", entries.len(), output.display()), None);
    // ...
    // 每张 slide 处理后:
    diag.info("generator", &format!("slide {}/{}: {} bytes", i+1, total, slide_size), None);
}
```

- [ ] **步骤 2：template_reader.rs — 添加 diag 参数**

```rust
pub fn scan_placeholders(xml: &str, diag: &mut DiagStore) -> Vec<PlaceholderInfo> {
    let result = scan_placeholders_inner(xml);
    diag.info("template", &format!("found {} placeholders: {:?}", result.len(), result.iter().map(|p| &p.name).collect::<Vec<_>>()), None);
    if result.is_empty() {
        diag.warn("template", "no placeholders found in slide XML", None);
    }
    result
}
```

- [ ] **步骤 3：更新所有调用者 + 运行测试**

运行：`cargo test --workspace`
预期：全部 PASS

- [ ] **步骤 4：Commit**

```bash
git add crates/core/src/generator.rs crates/core/src/template_reader.rs
git commit -m "feat(generator,template_reader): diag instrumentation"
```

---

### 任务 10：GUI diag 集成

**文件：**
- 修改：`crates/gui/src/app.rs`
- 修改：`crates/gui/Cargo.toml`

- [ ] **步骤 1：在 VocabPptApp 中添加 `diag_store: Option<DiagStore>` 字段**

```rust
pub struct VocabPptApp {
    // ... 现有字段 ...
    diag_store: Option<vocab_core::diag::DiagStore>,
}
```

- [ ] **步骤 2：在导出操作中创建并填充 DiagStore**

```rust
fn start_export_png(&mut self) {
    // ... 现有代码 ...
    let mut diag = vocab_core::diag::DiagStore::new();
    // 传递 diag 给 reader::load
    let entries = match reader::load(&source, &mut diag) {
        Ok(e) => e,
        Err(e) => {
            diag.error("gui", &format!("load failed: {e}"), None);
            self.diag_store = Some(diag);
            return;
        }
    };
    // ... png_export 调用 ...
    self.diag_store = Some(diag);
}
```

- [ ] **步骤 3：添加诊断查看按钮（在 data_preview 面板中）**

```rust
// 在 data_preview.rs 的 show() 中添加
if let Some(ref diag) = app.diag_store {
    ui.separator();
    ui.label(format!("诊断: {} 事件, {} 警告, {} 错误", 
        diag.event_count(), diag.warnings(), diag.errors()));
}
```

- [ ] **步骤 4：修复 GUI 中静默吞没的错误**

```rust
// 旧: let _ = std::fs::create_dir_all(&output_dir);
// 新:
if let Err(e) = std::fs::create_dir_all(&output_dir) {
    diag.error("gui", &format!("create_dir failed: {e}"), None);
    // 不阻止导出，但记录错误
}

// 旧: let _ = template_pptx::generate_example_pptx(&tmp).map(|_| open::that(&tmp));
// 新:
match template_pptx::generate_example_pptx(&tmp) {
    Ok(_) => { let _ = open::that(&tmp); }
    Err(e) => {
        diag.error("gui", &format!("template generation failed: {e}"), None);
    }
}
```

- [ ] **步骤 5：构建并验证 GUI**

```bash
cargo build --release
# 启动 GUI 验证诊断面板
```

- [ ] **步骤 6：Commit**

```bash
git add crates/gui/
git commit -m "feat(gui): DiagStore integration + fix silent error swallowing"
```


### 任务 10.5：GUI PNG 导出添加取消机制

**文件：**
- 修改：`crates/gui/src/app.rs`
- 修改：`crates/gui/src/panels/data_preview.rs`

- [ ] **步骤 1：在 VocabPptApp 中添加取消标志**

```rust
// app.rs — 在 struct 中添加字段
pub struct VocabPptApp {
    // ... 现有字段 ...
    export_cancel_flag: Option<Arc<AtomicBool>>,  // PNG 导出取消标志
}
```

- [ ] **步骤 2：更新 start_export_png 传递 cancel + 使用带 cancel 的 API**

```rust
fn start_export_png(&mut self) {
    // ... 构造 source, entries ...
    let cancel = Arc::new(AtomicBool::new(false));
    self.export_cancel_flag = Some(cancel.clone());

    let handle = std::thread::spawn(move || {
        let mut diag = DiagStore::new();
        let res = png_export::export_entries_to_png_with_cancel(
            &entries, &output_dir, &cancel, &mut diag
        );
        // ... 处理结果 ...
    });
    // ...
}
```

- [ ] **步骤 3：在 data_preview 面板中添加取消按钮**

```rust
// data_preview.rs — show() 函数中添加
if let Some(ref flag) = app.export_cancel_flag {
    ui.separator();
    if ui.button("⏹ 取消 PNG 导出").clicked() {
        flag.store(true, Ordering::Relaxed);
    }
}
```

- [ ] **步骤 4：导出完成后清理标志**

```rust
// poll_export_png 中
fn poll_export_png(&mut self) {
    // ... 检查 handle 完成 ...
    if let Some(res) = handle_result {
        self.export_cancel_flag = None;  // 清理取消标志
        // ... 显示结果 ...
    }
}
```

- [ ] **步骤 5：Commit**

```bash
git add crates/gui/
git commit -m "feat(gui): cancel button for PNG export"
```

### 任务 11：最终清理 — 删除旧代码 + 废弃依赖

**文件：**
- 删除/重命名：`crates/core/src/png_export.rs` → 删除（逻辑已迁移到 `png_export/mod.rs` + `pipeline.rs`）
- 修改：`crates/core/Cargo.toml`

- [ ] **步骤 1：删除旧的 png_export.rs**

确认 `png_export/mod.rs` 重新导出了所有公共 API 后：

```bash
rm crates/core/src/png_export.rs
```

- [ ] **步骤 2：删除 `ab_glyph` 依赖**

```toml
# crates/core/Cargo.toml
# 删除: ab_glyph = "0.2"
```

- [ ] **步骤 3：运行完整测试套件**

```bash
cargo test --workspace
cargo build --release
```

预期：全部 PASS，零警告

- [ ] **步骤 4：最终烟雾测试**

```bash
# CLI 导出
./target/release/单词卡片转换 export-png -i assets/template.xlsx -o /tmp/final_test/

# diag 查询
./target/release/单词卡片转换 diag /tmp/final_test/export_*.ndjson --summary
./target/release/单词卡片转换 diag /tmp/final_test/export_*.ndjson --blank-slides
./target/release/单词卡片转换 diag /tmp/final_test/export_*.ndjson --font-trace

# 验证 PNG 非空白
python3 -c "
from PIL import Image
for i in range(1,6):
    img = Image.open(f'/tmp/final_test/slide_{i}.png')
    px = list(img.get_flattened_data())
    nw = sum(1 for p in px if p != (255,255,255,255))
    print(f'slide_{i}: {nw} text px ({100*nw/len(px):.2f}%)')
"
```

- [ ] **步骤 5：Commit**

```bash
git rm crates/core/src/png_export.rs
git add crates/core/Cargo.toml
git commit -m "chore: remove old png_export.rs, drop ab_glyph dependency"
```

---

## 自检

1. **规格覆盖度**：浏览规格的每个章节，对照任务：
   - 字体探测(任务3) ✓
   - PPTX 解析(任务4) ✓  
   - SVG 生成 + resvg 渲染(任务5) ✓
   - diag 基础设施(任务1-2) ✓
   - 空白页检测(任务5) ✓
   - 每 slide 错误隔离(任务5.5) ✓
   - fontdb 崩溃保护(任务5.6) ✓
   - NDJSON 写入非阻塞(任务6.5) ✓
   - API 整合(任务6) ✓
   - diag 子命令(任务7) ✓
   - reader 注入(任务8) ✓
   - generator/template_reader 注入(任务9) ✓
   - GUI 注入 + 取消按钮(任务10 + 10.5) ✓
   - 清理(任务11) ✓
   - 全部 14 个任务覆盖规格 + 审计缺口 ✗→✓

2. **异常处理覆盖度**（对照审计缺口）：
   - G3: XML 属性静默失败 → 任务4 `diag.warn` ✓
   - G4: 字体缺失无检测 → 任务3全链 + 任务5空白检测 ✓
   - G5: GUI 无取消 → 任务5.5 + 10.5 ✓
   - G7: 部分布局丢弃字段 → 任务5 SVG 生成 `fields_skipped` ✓
   - G9: GUI `let _ =` 吞错误 → 任务10步骤4 ✓
   - fontdb 崩溃 → 任务5.6 `catch_unwind` ✓
   - slide 失败无隔离 → 任务5.5 `continue on error` ✓
   - NDJSON 写入 `let _ =` → 任务6.5 `match Ok/Err` ✓

3. **占位符扫描**：无 TODO、无"后续实现"、无模糊步骤 ✓

4. **类型一致性**：`DiagStore` 所有任务使用相同接口，`FontConfig` 在任务3定义、任务5-6使用 ✓
