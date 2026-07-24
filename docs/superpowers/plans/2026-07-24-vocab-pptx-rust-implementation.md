# 单词PPT生成器 Rust 重构 — 实现计划

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推荐）或 superpowers:executing-plans 逐任务实现此计划。步骤使用复选框（`- [ ]`）语法来跟踪进度。

**目标：** 将 Python-Excel-to-PPT-Vocabulary-Tool 重构为基于 Rust（egui + ppt-rs + calamine）的跨平台 CLI+GUI 双模式应用。

**架构：** Workspace 三层：`core`（纯逻辑，零 UI 依赖）→ `cli`（clap 薄壳）/ `gui`（egui 状态机）。PPTX 生成用 `ppt-rs`（已验证活跃），Excel 用 `calamine`，CSV 用 `csv` + `encoding_rs`。

**技术栈：** Rust 2024 edition, egui/eframe, ppt-rs, calamine, clap, csv, encoding_rs, rfd, thiserror, open

---

## 文件结构

| 文件 | 职责 |
|------|------|
| `crates/core/src/types.rs` | `WordEntry`, `SlideConfig`, `InputSource`, `SlideTheme`, 错误类型 |
| `crates/core/src/reader.rs` | `load()`, `list_sheets()` — Excel/CSV 读取与校验 |
| `crates/core/src/generator.rs` | `generate()` — PPTX 生成（ppt-rs），含进度回调与取消 |
| `crates/core/src/template.rs` | `export_template()` — 内置模板嵌入与导出 |
| `crates/core/src/lib.rs` | 公开 API re-export |
| `crates/core/tests/*.rs` | 单元测试（每模块一个） |
| `crates/cli/src/main.rs` | clap 命令解析，调 core API |
| `crates/gui/src/main.rs` | eframe 启动 |
| `crates/gui/src/app.rs` | 状态机 (`AppState`) |
| `crates/gui/src/panels/file_picker.rs` | 文件选择 + sheet 下拉 |
| `crates/gui/src/panels/data_preview.rs` | 数据预览表格 |
| `crates/gui/src/panels/output_config.rs` | 输出路径 + 主题 |
| `assets/template.xlsx` | 内置 Excel 模板 |

---

## Phase 0: Workspace 脚手架

### 任务 0.1: 创建 workspace 和三个 crate

**文件：**
- 创建：`Cargo.toml`
- 创建：`crates/core/Cargo.toml`
- 创建：`crates/core/src/lib.rs`
- 创建：`crates/cli/Cargo.toml`
- 创建：`crates/cli/src/main.rs`
- 创建：`crates/gui/Cargo.toml`
- 创建：`crates/gui/src/main.rs`

- [ ] **步骤 1: 创建 workspace Cargo.toml**

```toml
[workspace]
members = ["crates/core", "crates/cli", "crates/gui"]
resolver = "2"

[workspace.package]
version = "0.1.0"
edition = "2021"
license = "MIT"
```

- [ ] **步骤 2: 创建 core/Cargo.toml**

```toml
[package]
name = "core"
version.workspace = true
edition.workspace = true
license.workspace = true

[dependencies]
calamine = "0.26"
csv = "1.3"
encoding_rs = "0.8"
thiserror = "2"
ppt-rs = "0.2"
zip = "2"
```

- [ ] **步骤 3: 创建 core/src/lib.rs（空壳）**

```rust
pub mod types;
pub mod reader;
pub mod generator;
pub mod template;
```

- [ ] **步骤 4: 创建 cli/Cargo.toml + main.rs**

```toml
[package]
name = "cli"
version.workspace = true
edition.workspace = true
license.workspace = true

[[bin]]
name = "单词ppt"

[dependencies]
core = { path = "../core" }
clap = { version = "4", features = ["derive"] }
```

```rust
// cli/src/main.rs
fn main() {
    println!("单词PPT生成器 CLI (WIP)");
}
```

- [ ] **步骤 5: 创建 gui/Cargo.toml + main.rs**

```toml
[package]
name = "gui"
version.workspace = true
edition.workspace = true
license.workspace = true

[dependencies]
core = { path = "../core" }
eframe = "0.31"
egui = "0.31"
egui_extras = { version = "0.31", features = ["table"] }
rfd = "0.15"
```

```rust
// gui/src/main.rs
fn main() {
    println!("单词PPT生成器 GUI (WIP)");
}
```

- [ ] **步骤 6: 创建 assets/ 目录，复制模板**

从 Python 项目复制 `单词表模板.xlsx` 到 `assets/template.xlsx`。

- [ ] **步骤 7: 验证编译**

```bash
cd 单词PPT生成器 && cargo build
```

预期：三个 crate 均编译通过。

- [ ] **步骤 8: Commit**

```bash
git add -A && git commit -m "chore: workspace 脚手架 (core/cli/gui + assets)"
```

---

## Phase 1: Core — 数据模型

### 任务 1.1: 定义 types.rs 和错误类型

**文件：**
- 创建：`crates/core/src/types.rs`
- 创建：`crates/core/tests/types_test.rs`

- [ ] **步骤 1: 编写失败的测试**

```rust
// crates/core/tests/types_test.rs
use core::types::*;

#[test]
fn word_entry_all_fields_present() {
    let entry = WordEntry {
        word: "apple".into(),
        phonetic: "/ˈæpl/".into(),
        morphology: "".into(),
        example: "I eat an apple.".into(),
        example_definition: "我吃苹果。".into(),
        definition: "苹果".into(),
    };
    assert_eq!(entry.word, "apple");
    assert_eq!(entry.morphology, ""); // 允许空
}

#[test]
fn slide_config_default_values() {
    let config = SlideConfig::default();
    assert_eq!(config.width, 16.0);
    assert_eq!(config.height, 9.0);
    assert_eq!(config.word_font_size, 72.0);
    assert_eq!(config.text_wrap_threshold, 40);
}

#[test]
fn input_source_excel_variant() {
    let src = InputSource::Excel {
        path: "test.xlsx".into(),
        sheet: "Sheet1".into(),
    };
    match src {
        InputSource::Excel { sheet, .. } => assert_eq!(sheet, "Sheet1"),
        _ => panic!("expected Excel variant"),
    }
}

#[test]
fn slide_theme_default_is_variant() {
    let theme = SlideTheme::default();
    // Default 应该存在
    assert!(matches!(theme, SlideTheme::Default));
}
```

- [ ] **步骤 2: 运行测试验证失败**

```bash
cargo test -p core
```

预期：FAIL — `WordEntry` not defined, `SlideConfig` not defined

- [ ] **步骤 3: 实现 types.rs**

```rust
// crates/core/src/types.rs
use std::path::PathBuf;

#[derive(Debug, Clone)]
pub struct WordEntry {
    pub word: String,
    pub phonetic: String,
    pub morphology: String,
    pub example: String,
    pub example_definition: String,
    pub definition: String,
}

impl WordEntry {
    /// 返回 true 如果此条应该跳过（word 为空或全空）
    pub fn should_skip(&self) -> bool {
        self.word.trim().is_empty()
            || (self.word.is_empty()
                && self.phonetic.is_empty()
                && self.morphology.is_empty()
                && self.example.is_empty()
                && self.example_definition.is_empty()
                && self.definition.is_empty())
    }
}

#[derive(Debug, Clone)]
pub struct SlideConfig {
    pub width: f32,
    pub height: f32,
    pub word_font_size: f32,
    pub phonetic_font_size: f32,
    pub content_font_size: f32,
    pub font_family_cjk: String,
    pub font_family_latin: String,
    pub text_wrap_threshold: usize,
}

impl Default for SlideConfig {
    fn default() -> Self {
        Self {
            width: 16.0,
            height: 9.0,
            word_font_size: 72.0,
            phonetic_font_size: 32.0,
            content_font_size: 32.0,
            font_family_cjk: "Noto Sans CJK SC".into(),
            font_family_latin: "Calibri".into(),
            text_wrap_threshold: 40,
        }
    }
}

#[derive(Debug, Clone)]
pub enum InputSource {
    Excel { path: PathBuf, sheet: String },
    Csv { path: PathBuf, encoding: String },
}

#[derive(Debug, Clone, Default)]
pub enum SlideTheme {
    #[default]
    Default,
    Blue,
    Green,
    Gray,
}

#[derive(Debug, thiserror::Error)]
pub enum LoadError {
    #[error("缺少列: {0:?}")]
    MissingColumns(Vec<String>),
    #[error("第 {row} 行必填字段为空: {field}")]
    EmptyRequiredField { row: usize, field: String },
    #[error("文件格式无效: {0}")]
    InvalidFormat(String),
    #[error("无法打开文件: {0}")]
    IoError(#[from] std::io::Error),
    #[error("编码错误: {0}")]
    EncodingError(String),
    #[error("Excel 读取错误: {0}")]
    ExcelError(String),
}

#[derive(Debug, thiserror::Error)]
pub enum GenerateError {
    #[error("没有可生成的条目")]
    NoEntries,
    #[error("文件已存在: {0}")]
    FileExists(String),
    #[error("已取消")]
    Cancelled,
    #[error("PPTX 生成错误: {0}")]
    PptxError(String),
    #[error("IO 错误: {0}")]
    IoError(#[from] std::io::Error),
}

#[derive(Debug, thiserror::Error)]
pub enum TemplateError {
    #[error("无法写入模板: {0}")]
    IoError(#[from] std::io::Error),
    #[error("内置模板损坏")]
    CorruptEmbedded,
}
```

- [ ] **步骤 4: 运行测试验证通过**

```bash
cargo test -p core
```

预期：PASS（4 tests）

- [ ] **步骤 5: Commit**

```bash
git add -A && git commit -m "feat(core): 数据模型 — WordEntry, SlideConfig, InputSource, SlideTheme, 错误类型"
```

---

## Phase 2: Core — Reader（Excel + CSV）

### 任务 2.1: Excel 列名校验 + list_sheets

**文件：**
- 创建：`crates/core/src/reader.rs`
- 创建：`crates/core/tests/reader_test.rs`

- [ ] **步骤 1: 编写失败的测试**

```rust
// crates/core/tests/reader_test.rs
use core::reader::*;
use core::types::*;
use std::path::Path;

#[test]
fn list_sheets_returns_sheet_names() {
    // 用 Python 模板作为测试输入
    let sheets = list_sheets(Path::new("../../assets/template.xlsx")).unwrap();
    assert!(!sheets.is_empty());
    assert!(sheets.contains(&"单词表".to_string()));
}

#[test]
fn load_excel_missing_columns() {
    // 需要先创建一个缺列的测试 Excel 文件
    // 此测试依赖外部文件，先用已知的模板验证正常路径
    let source = InputSource::Excel {
        path: "../../assets/template.xlsx".into(),
        sheet: "单词表".into(),
    };
    let entries = load(&source).unwrap();
    assert!(!entries.is_empty());
    assert_eq!(entries[0].word, "apple");
}
```

- [ ] **步骤 2: 运行测试验证失败**

```bash
cargo test -p core
```

预期：FAIL — `list_sheets` not found in module `reader`

- [ ] **步骤 3: 实现 reader.rs（第一版：仅 Excel）**

```rust
// crates/core/src/reader.rs
use crate::types::*;
use calamine::{open_workbook, Reader, Xlsx};
use std::path::Path;

const REQUIRED_COLUMNS: [(&str, &str); 6] = [
    ("word", "英文单词"),
    ("phonetic", "英文音标"),
    ("morphology", "词根词缀"),
    ("example", "例句"),
    ("example_definition", "例句释义"),
    ("definition", "单词释义"),
];

pub fn list_sheets(path: &Path) -> Result<Vec<String>, LoadError> {
    let workbook: Xlsx<_> = open_workbook(path)
        .map_err(|e| LoadError::ExcelError(e.to_string()))?;
    Ok(workbook.sheet_names().to_vec())
}

pub fn load(source: &InputSource) -> Result<Vec<WordEntry>, LoadError> {
    match source {
        InputSource::Excel { path, sheet } => load_excel(path, sheet),
        InputSource::Csv { path, encoding } => load_csv(path, encoding),
    }
}

fn load_excel(path: &Path, sheet: &str) -> Result<Vec<WordEntry>, LoadError> {
    let mut workbook: Xlsx<_> = open_workbook(path)
        .map_err(|e| LoadError::ExcelError(e.to_string()))?;

    let range = workbook
        .worksheet_range(sheet)
        .map_err(|e| LoadError::ExcelError(format!("sheet '{}' not found: {}", sheet, e)))?;

    let mut rows = range.rows();
    let header_row = rows.next()
        .ok_or(LoadError::InvalidFormat("文件为空".into()))?;

    // 建立列名 → 列索引映射
    let header_map: Vec<Option<&str>> = header_row.iter()
        .map(|cell| Some(cell.get_string()?.as_str()))
        .collect();

    // 校验所有必填列
    let mut missing = Vec::new();
    for (key, display) in REQUIRED_COLUMNS {
        if !header_map.iter().any(|h| h == &Some(display)) {
            missing.push(key.to_string());
        }
    }
    if !missing.is_empty() {
        return Err(LoadError::MissingColumns(missing));
    }

    // 建立 key → 列索引
    let col_index: Vec<usize> = REQUIRED_COLUMNS.iter()
        .map(|(_, display)| {
            header_map.iter()
                .position(|h| h == &Some(display))
                .expect("column should exist after validation")
        })
        .collect();

    // 解析数据行
    let mut entries = Vec::new();
    for (row_idx, row) in rows.enumerate() {
        let get = |idx: usize| -> String {
            row.get(idx)
                .and_then(|c| c.get_string())
                .unwrap_or_default()
                .to_string()
        };
        let entry = WordEntry {
            word: get(col_index[0]),
            phonetic: get(col_index[1]),
            morphology: get(col_index[2]),
            example: get(col_index[3]),
            example_definition: get(col_index[4]),
            definition: get(col_index[5]),
        };
        if entry.should_skip() {
            continue;
        }
        if entry.word.trim().is_empty() {
            // word 为空，跳过并警告（stderr）
            eprintln!("警告: 第 {} 行缺少英文单词，已跳过", row_idx + 2);
            continue;
        }
        entries.push(entry);
    }
    Ok(entries)
}

fn load_csv(_path: &Path, _encoding: &str) -> Result<Vec<WordEntry>, LoadError> {
    // 任务 2.2 实现
    Err(LoadError::InvalidFormat("CSV 暂未实现".into()))
}
```

- [ ] **步骤 4: 运行测试验证通过**

```bash
cargo test -p core
```

预期：types_test 4 tests PASS; reader_test 2 tests PASS

- [ ] **步骤 5: Commit**

```bash
git add -A && git commit -m "feat(core): Excel 读取 — list_sheets + load (列名校验, 数据解析)"
```

### 任务 2.2: CSV 读取 + 多编码

**文件：**
- 修改：`crates/core/src/reader.rs`（实现 `load_csv`）

- [ ] **步骤 1: 编写失败的测试**

```rust
#[test]
fn load_csv_utf8_returns_entries() {
    // 创建临时 CSV 文件
    let tmp = std::env::temp_dir().join("test_words.csv");
    std::fs::write(&tmp, "英文单词,英文音标,词根词缀,例句,例句释义,单词释义\napple,/æpl/,ap-,I eat an apple.,我吃苹果。,苹果\n").unwrap();
    let source = InputSource::Csv {
        path: tmp.clone(),
        encoding: "UTF-8".into(),
    };
    let entries = load(&source).unwrap();
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].word, "apple");
    let _ = std::fs::remove_file(&tmp);
}

#[test]
fn load_csv_missing_columns_reports_all() {
    let tmp = std::env::temp_dir().join("bad.csv");
    std::fs::write(&tmp, "英文单词,例句\napple,hello\n").unwrap();
    let source = InputSource::Csv {
        path: tmp.clone(),
        encoding: "UTF-8".into(),
    };
    let err = load(&source).unwrap_err();
    assert!(matches!(err, LoadError::MissingColumns(_)));
    let _ = std::fs::remove_file(&tmp);
}

#[test]
fn load_csv_skips_empty_word_rows() {
    let tmp = std::env::temp_dir().join("empty_word.csv");
    std::fs::write(&tmp, "英文单词,英文音标,词根词缀,例句,例句释义,单词释义\n,//,,,,\napple,/æpl/,,I eat.,我吃。,苹果\n").unwrap();
    let source = InputSource::Csv {
        path: tmp.clone(),
        encoding: "UTF-8".into(),
    };
    let entries = load(&source).unwrap();
    assert_eq!(entries.len(), 1); // 空 word 行被跳过
    assert_eq!(entries[0].word, "apple");
    let _ = std::fs::remove_file(&tmp);
}
```

- [ ] **步骤 2: 运行测试验证失败**

```bash
cargo test -p core
```

预期：新 CSV 测试 FAIL — "CSV 暂未实现"

- [ ] **步骤 3: 实现 load_csv**

```rust
// reader.rs 中替换 load_csv 存根
fn load_csv(path: &Path, encoding: &str) -> Result<Vec<WordEntry>, LoadError> {
    let file = std::fs::File::open(path)?;
    let mut reader = if encoding.to_uppercase() == "UTF-8" || encoding.to_uppercase() == "UTF8" {
        csv::ReaderBuilder::new()
            .has_headers(true)
            .from_reader(file)
    } else {
        // 使用 encoding_rs 转码
        let bytes = std::fs::read(path)?;
        let (decoded, _, had_errors) = encoding_rs::Encoding::for_label(encoding.as_bytes())
            .ok_or_else(|| LoadError::EncodingError(format!("不支持的编码: {}", encoding)))?
            .decode(&bytes);
        if had_errors {
            eprintln!("警告: 编码转换中有无法映射的字符");
        }
        csv::ReaderBuilder::new()
            .has_headers(true)
            .from_reader(decoded.as_bytes())
    };

    let headers = reader.headers()
        .map_err(|e| LoadError::InvalidFormat(format!("CSV 表头解析失败: {}", e)))?;

    // 同 Excel 校验逻辑
    let header_names: Vec<&str> = headers.iter().collect();

    let mut missing = Vec::new();
    for (key, display) in REQUIRED_COLUMNS {
        if !header_names.contains(&display) {
            missing.push(key.to_string());
        }
    }
    if !missing.is_empty() {
        return Err(LoadError::MissingColumns(missing));
    }

    let col_index: Vec<usize> = REQUIRED_COLUMNS.iter()
        .map(|(_, display)| {
            header_names.iter()
                .position(|&h| h == *display)
                .expect("column should exist after validation")
        })
        .collect();

    let mut entries = Vec::new();
    for (row_idx, record) in reader.records().enumerate() {
        let record = record
            .map_err(|e| LoadError::InvalidFormat(format!("第 {} 行解析失败: {}", row_idx + 2, e)))?;

        let get = |idx: usize| -> String {
            record.get(idx).unwrap_or_default().to_string()
        };

        let entry = WordEntry {
            word: get(col_index[0]),
            phonetic: get(col_index[1]),
            morphology: get(col_index[2]),
            example: get(col_index[3]),
            example_definition: get(col_index[4]),
            definition: get(col_index[5]),
        };

        if entry.should_skip() {
            continue;
        }
        if entry.word.trim().is_empty() {
            eprintln!("警告: 第 {} 行缺少英文单词，已跳过", row_idx + 2);
            continue;
        }
        entries.push(entry);
    }
    Ok(entries)
}
```

- [ ] **步骤 4: 运行测试验证通过**

```bash
cargo test -p core
```

预期：全部 PASS

- [ ] **步骤 5: Commit**

```bash
git add -A && git commit -m "feat(core): CSV 读取 — 多编码支持, 列名校验, 空行跳过"
```

---

## Phase 3: Spike — ppt-rs 验证

### 任务 3.1: 验证 ppt-rs 基础 + CJK 字体

**文件：**
- 创建（临时）：`crates/core/examples/spike_pptrs.rs`

- [ ] **步骤 1: 编写 spike 示例**

```rust
// crates/core/examples/spike_pptrs.rs
use ppt_rs::prelude::*;

fn main() {
    let slides = vec![
        SlideContent::new("测试单词")
            .add_text("apple")
            .add_text("/ˈæpl/")
            .add_text("词根词缀：ap-ple")
            .add_text("例句：I eat an apple every day.")
            .add_text("例句释义：我每天吃一个苹果。")
            .add_text("单词释义：苹果"),
    ];
    let data = ppt_rs::create_pptx_with_content("Spike Test", slides).unwrap();
    std::fs::write("spike_output.pptx", data).unwrap();
    println!("✓ spike_output.pptx 已生成");
}
```

- [ ] **步骤 2: 运行 spike**

```bash
cargo run -p core --example spike_pptrs
```

预期：生成 `spike_output.pptx`，不崩溃。

- [ ] **步骤 3: 手动验证**

打开 `spike_output.pptx` 在 PowerPoint / WPS / LibreOffice 中检查：
- 中文是否正常渲染？
- 文本框位置是否可控？

**Go/No-Go 决策：**
- ✅ 通过 → 继续用 ppt-rs → 任务 3.2
- ❌ 失败 → 降级手工 OOXML（见备选方案，需重新评估工时）

- [ ] **步骤 4: Commit 或记录决策**

```bash
# 通过：
git add -A && git commit -m "spike: ppt-rs 验证通过 — 中文 OK, 基础 API 可用"

# 不通过：
# 记录到 docs/superpowers/specs/ 中的决策日志
```

### 任务 3.2: 实现 generator.rs

**文件：**
- 创建：`crates/core/src/generator.rs`
- 创建：`crates/core/tests/generator_test.rs`

- [ ] **步骤 1: 编写失败的测试**

```rust
// crates/core/tests/generator_test.rs
use core::generator::*;
use core::types::*;
use std::path::Path;

fn make_entries(count: usize) -> Vec<WordEntry> {
    (0..count).map(|i| WordEntry {
        word: format!("word{}", i),
        phonetic: "/test/".into(),
        morphology: "test-".into(),
        example: "This is a test.".into(),
        example_definition: "测试。".into(),
        definition: "测试".into(),
    }).collect()
}

#[test]
fn generate_empty_entries_errors() {
    let err = generate(&[], Path::new("test.pptx"), &SlideConfig::default(), SlideTheme::Default, |_, _| true);
    assert!(matches!(err, Err(GenerateError::NoEntries)));
}

#[test]
fn generate_one_slide_creates_pptx() {
    let entries = make_entries(1);
    let tmp = std::env::temp_dir().join("test_one.pptx");
    let result = generate(&entries, &tmp, &SlideConfig::default(), SlideTheme::Default, |_, _| true);
    assert!(result.is_ok());
    assert!(tmp.exists());
    // 验证是有效 zip（PPTX 本质是 zip）
    let file = std::fs::File::open(&tmp).unwrap();
    assert!(zip::ZipArchive::new(file).is_ok());
    let _ = std::fs::remove_file(&tmp);
}

#[test]
fn generate_calls_progress_correctly() {
    use std::sync::atomic::{AtomicUsize, Ordering};
    let entries = make_entries(3);
    let count = AtomicUsize::new(0);
    let tmp = std::env::temp_dir().join("test_progress.pptx");
    generate(&entries, &tmp, &SlideConfig::default(), SlideTheme::Default, |current, total| {
        count.fetch_add(1, Ordering::SeqCst);
        assert_eq!(total, 3);
        assert!(current <= total);
        true
    }).unwrap();
    assert_eq!(count.load(Ordering::SeqCst), 3);
    let _ = std::fs::remove_file(&tmp);
}

#[test]
fn generate_cancel_stops_early() {
    let entries = make_entries(5);
    let tmp = std::env::temp_dir().join("test_cancel.pptx");
    let err = generate(&entries, &tmp, &SlideConfig::default(), SlideTheme::Default, |current, _total| {
        current < 2 // 第 3 张时取消
    }).unwrap_err();
    assert!(matches!(err, GenerateError::Cancelled));
    let _ = std::fs::remove_file(&tmp);
}

#[test]
fn generate_file_exists_error() {
    let entries = make_entries(1);
    let tmp = std::env::temp_dir().join("test_exists.pptx");
    std::fs::write(&tmp, "existing").unwrap(); // 先创建占位文件
    let err = generate(&entries, &tmp, &SlideConfig::default(), SlideTheme::Default, |_, _| true).unwrap_err();
    assert!(matches!(err, GenerateError::FileExists(_)));
    let _ = std::fs::remove_file(&tmp);
}
```

- [ ] **步骤 2: 运行测试验证失败**

```bash
cargo test -p core
```

预期：generator_test 全部 FAIL

- [ ] **步骤 3: 实现 generator.rs**

```rust
// crates/core/src/generator.rs
use crate::types::*;
use ppt_rs::prelude::*;
use std::path::Path;

pub fn generate(
    entries: &[WordEntry],
    output: &Path,
    config: &SlideConfig,
    theme: SlideTheme,
    progress: impl Fn(usize, usize) -> bool,
) -> Result<(), GenerateError> {
    if entries.is_empty() {
        return Err(GenerateError::NoEntries);
    }
    if output.exists() {
        return Err(GenerateError::FileExists(output.display().to_string()));
    }

    let total = entries.len();
    let mut slides: Vec<SlideContent> = Vec::with_capacity(total);

    for (i, entry) in entries.iter().enumerate() {
        let slide = SlideContent::new(&entry.word)
            .add_text(&format!("音标：{}", entry.phonetic))
            .add_text(&format!("词根词缀：{}", entry.morphology))
            .add_text(&format!("例句：{}", entry.example))
            .add_text(&format!("例句释义：{}", entry.example_definition))
            .add_text(&format!("单词释义：{}", entry.definition));
        slides.push(slide);

        if !progress(i + 1, total) {
            return Err(GenerateError::Cancelled);
        }
    }

    let data = ppt_rs::create_pptx_with_content("单词PPT", slides)
        .map_err(|e| GenerateError::PptxError(e.to_string()))?;

    std::fs::write(output, data)
        .map_err(|e| GenerateError::IoError(e))?;

    Ok(())
}
```

- [ ] **步骤 4: 运行测试验证通过**

```bash
cargo test -p core
```

预期：全部 PASS

- [ ] **步骤 5: Commit**

```bash
git add -A && git commit -m "feat(core): PPTX 生成 — ppt-rs, 进度回调, 取消支持, 文件冲突检测"
```

---

## Phase 4: Core — Template

### 任务 4.1: 模板嵌入与导出

**文件：**
- 创建：`crates/core/assets/template.xlsx`（从 assets/ 复制）
- 创建：`crates/core/src/template.rs`
- 创建：`crates/core/tests/template_test.rs`

- [ ] **步骤 1: 编写失败的测试**

```rust
// crates/core/tests/template_test.rs
use core::template::*;
use std::path::Path;

#[test]
fn export_template_creates_file() {
    let tmp = std::env::temp_dir().join("test_template.xlsx");
    export_template(&tmp).unwrap();
    assert!(tmp.exists());
    assert!(tmp.metadata().unwrap().len() > 0);
    let _ = std::fs::remove_file(&tmp);
}

#[test]
fn exported_template_is_valid_xlsx() {
    use calamine::{open_workbook, Reader, Xlsx};
    let tmp = std::env::temp_dir().join("test_valid.xlsx");
    export_template(&tmp).unwrap();
    let wb: Xlsx<_> = open_workbook(&tmp).unwrap();
    assert!(!wb.sheet_names().is_empty());
    let _ = std::fs::remove_file(&tmp);
}
```

- [ ] **步骤 2: 实现 template.rs**

```rust
// crates/core/src/template.rs
use crate::types::TemplateError;
use std::path::Path;

const TEMPLATE_BYTES: &[u8] = include_bytes!("../assets/template.xlsx");

pub fn export_template(path: &Path) -> Result<(), TemplateError> {
    std::fs::write(path, TEMPLATE_BYTES)
        .map_err(TemplateError::IoError)
}
```

- [ ] **步骤 3: 复制模板到 core/assets/**

```bash
cp assets/template.xlsx crates/core/assets/template.xlsx
```

- [ ] **步骤 4: 运行测试验证通过**

```bash
cargo test -p core
```

预期：全部 PASS

- [ ] **步骤 5: Commit**

```bash
git add -A && git commit -m "feat(core): 模板嵌入 — include_bytes! + export_template"
```

---

## Phase 5: CLI

### 任务 5.1: 实现 CLI（clap 命令）

**文件：**
- 修改：`crates/cli/src/main.rs`

- [ ] **步骤 1: 编写失败的测试**

CLI 不写纯单元测试——用集成测试验证参数解析。创建临时测试：

```bash
# 手动测试命令
cargo run -p cli -- generate --help
```

- [ ] **步骤 2: 实现 main.rs**

```rust
// crates/cli/src/main.rs
use clap::{Parser, Subcommand};
use core::types::*;
use core::{reader, generator, template};
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "单词ppt", about = "从 Excel/CSV 生成单词 PPT 幻灯片")]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// 生成 PPT（单文件）
    Generate {
        #[arg(short, long)]
        input: PathBuf,
        #[arg(short, long, default_value = "output.pptx")]
        output: PathBuf,
        #[arg(short, long)]
        sheet: Option<String>,
        #[arg(short = 'e', long, default_value = "utf-8")]
        encoding: String,
        #[arg(short, long, default_value = "default")]
        theme: String,
        #[arg(short = 'f', long)]
        force: bool,
    },
    /// 批量生成（遍历目录）
    Batch {
        #[arg(short, long)]
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
        #[arg(short, long)]
        sheet: Option<String>,
        #[arg(short = 'e', long, default_value = "utf-8")]
        encoding: String,
        #[arg(short, long, default_value = "default")]
        theme: String,
        #[arg(short = 'f', long)]
        force: bool,
    },
    /// 创建 Excel 模板
    Template {
        #[arg(short, long, default_value = "单词表模板.xlsx")]
        output: PathBuf,
    },
}

fn parse_theme(s: &str) -> SlideTheme {
    match s.to_lowercase().as_str() {
        "blue" => SlideTheme::Blue,
        "green" => SlideTheme::Green,
        "gray" => SlideTheme::Gray,
        _ => SlideTheme::Default,
    }
}

fn guess_source(path: &PathBuf, sheet: Option<String>, encoding: String) -> InputSource {
    match path.extension().and_then(|e| e.to_str()).unwrap_or("") {
        "csv" => InputSource::Csv { path: path.clone(), encoding },
        _ => InputSource::Excel {
            path: path.clone(),
            sheet: sheet.unwrap_or_else(|| {
                reader::list_sheets(path)
                    .ok()
                    .and_then(|s| s.into_iter().next())
                    .unwrap_or_else(|| "Sheet1".into())
            }),
        },
    }
}

fn main() {
    let cli = Cli::parse();

    match cli.command {
        Command::Generate { input, output, sheet, encoding, theme, force } => {
            if output.exists() && !force {
                eprintln!("错误: 输出文件已存在: {}。使用 --force 覆盖。", output.display());
                std::process::exit(1);
            }
            let source = guess_source(&input, sheet, encoding);
            match reader::load(&source) {
                Ok(entries) => {
                    println!("加载了 {} 条单词", entries.len());
                    let theme = parse_theme(&theme);
                    let progress = |current: usize, total: usize| -> bool {
                        print!("\r进度: {}/{}", current, total);
                        true
                    };
                    match generator::generate(&entries, &output, &SlideConfig::default(), theme, progress) {
                        Ok(()) => println!("\n✓ 已生成: {}", output.display()),
                        Err(e) => { eprintln!("\n错误: {}", e); std::process::exit(1); }
                    }
                }
                Err(e) => { eprintln!("读取失败: {}", e); std::process::exit(1); }
            }
        }

        Command::Batch { input, output, sheet, encoding, theme, force } => {
            let entries = std::fs::read_dir(&input)
                .expect("无法读取输入目录")
                .filter_map(|e| e.ok())
                .filter(|e| {
                    let ext = e.path().extension().and_then(|x| x.to_str()).unwrap_or("");
                    matches!(ext, "xlsx" | "xls" | "csv")
                })
                .collect::<Vec<_>>();

            let mut success = 0;
            let mut failed = 0;

            for entry in &entries {
                let out = output.join(entry.file_name()).with_extension("pptx");
                let source = guess_source(&entry.path(), sheet.clone(), encoding.clone());
                match reader::load(&source) {
                    Ok(data) => {
                        let theme = parse_theme(&theme);
                        match generator::generate(&data, &out, &SlideConfig::default(), theme, |_, _| true) {
                            Ok(()) => success += 1,
                            Err(e) => { eprintln!("  ✗ {}: {}", entry.file_name().to_string_lossy(), e); failed += 1; }
                        }
                    }
                    Err(e) => { eprintln!("  ✗ {}: {}", entry.file_name().to_string_lossy(), e); failed += 1; }
                }
            }
            println!("\n处理 {} 个文件，成功 {} 个，失败 {} 个", entries.len(), success, failed);
            if failed > 0 { std::process::exit(1); }
        }

        Command::Template { output } => {
            match template::export_template(&output) {
                Ok(()) => println!("✓ 模板已创建: {}", output.display()),
                Err(e) => { eprintln!("错误: {}", e); std::process::exit(1); }
            }
        }
    }
}
```

- [ ] **步骤 3: 冒烟测试**

```bash
# 测试 help
cargo run -p cli -- --help
cargo run -p cli -- generate --help

# 测试模板创建
cargo run -p cli -- template -o /tmp/test_template.xlsx
ls -la /tmp/test_template.xlsx

# 测试生成（用模板中的示例数据）
cargo run -p cli -- generate -i assets/template.xlsx -o /tmp/test_output.pptx -f
```

预期：三个命令均正常执行。

- [ ] **步骤 4: Commit**

```bash
git add -A && git commit -m "feat(cli): 完整 CLI — generate, batch, template 子命令"
```

---

## Phase 6: GUI

### 任务 6.1: 基础 GUI 框架（状态机 + 布局骨架）

**文件：**
- 修改：`crates/gui/src/main.rs`
- 创建：`crates/gui/src/app.rs`
- 创建：`crates/gui/src/panels/file_picker.rs`
- 创建：`crates/gui/src/panels/data_preview.rs`
- 创建：`crates/gui/src/panels/output_config.rs`

- [ ] **步骤 1: 实现 main.rs**

```rust
// crates/gui/src/main.rs
mod app;
mod panels;

fn main() -> Result<(), eframe::Error> {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([960.0, 720.0]),
        ..Default::default()
    };
    eframe::run_native(
        "单词PPT生成器",
        options,
        Box::new(|_cc| Ok(Box::new(app::VocabPptApp::default()))),
    )
}
```

- [ ] **步骤 2: 实现 app.rs（状态机 + 顶部布局）**

```rust
// crates/gui/src/app.rs
use core::types::*;
use std::path::PathBuf;

pub enum AppState {
    Idle,
    Loading { path: PathBuf },
    Preview { entries: Vec<WordEntry>, source: InputSource },
    Generating { current: usize, total: usize },
    Done { count: usize },
    Error { message: String },
}

#[derive(Default)]
pub struct VocabPptApp {
    pub state: AppState,
    // 文件选择
    pub input_path: String,
    pub selected_format: FormatType,
    pub sheets: Vec<String>,
    pub selected_sheet: String,
    pub csv_encoding: String,
    // 输出
    pub output_path: String,
    pub selected_theme: SlideTheme,
    // 生成
    pub should_cancel: bool,
}

#[derive(Default, PartialEq)]
pub enum FormatType {
    #[default]
    Excel,
    Csv,
}

impl Default for AppState {
    fn default() -> Self { Self::Idle }
}

impl eframe::App for VocabPptApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::TopBottomPanel::top("top_bar").show(ctx, |ui| {
            ui.heading("单词PPT生成器");
        });

        egui::CentralPanel::default().show(ctx, |ui| {
            // 输入面板
            self.panels::file_picker::show(ui, self);

            // 输出面板
            self.panels::output_config::show(ui, self);

            // 状态面板
            self.panels::data_preview::show(ui, self);

            // 底部按钮
            ui.horizontal(|ui| {
                if ui.button("打开模板").clicked() {
                    self.handle_open_template();
                }
                let can_generate = matches!(self.state, AppState::Preview { ref entries, .. } if !entries.is_empty());
                let btn_text = match self.state {
                    AppState::Generating { .. } => "取消",
                    _ => "生成PPT",
                };
                if ui.add_enabled(can_generate || matches!(self.state, AppState::Generating { .. }),
                    egui::Button::new(btn_text)).clicked() {
                    self.handle_generate();
                }
            });
        });
    }
}

impl VocabPptApp {
    fn handle_open_template(&self) {
        let tmp = std::env::temp_dir().join("单词表模板.xlsx");
        if core::template::export_template(&tmp).is_ok() {
            let _ = open::that(&tmp);
        }
    }

    fn handle_generate(&mut self) {
        // 将在数据预览面板中实现
    }
}
```

- [ ] **步骤 3: 实现面板骨架**

```rust
// crates/gui/src/panels/file_picker.rs
use crate::app::*;

pub fn show(ui: &mut egui::Ui, app: &mut VocabPptApp) {
    egui::CollapsingHeader::new("输入")
        .default_open(true)
        .show(ui, |ui| {
            ui.horizontal(|ui| {
                ui.label("文件:");
                ui.text_edit_singleline(&mut app.input_path);
                if ui.button("浏览...").clicked() {
                    if let Some(path) = rfd::FileDialog::new()
                        .add_filter("Excel/CSV", &["xlsx", "xls", "csv"])
                        .pick_file() {
                        app.input_path = path.display().to_string();
                    }
                }
            });

            ui.horizontal(|ui| {
                ui.radio_value(&mut app.selected_format, FormatType::Excel, "Excel (.xlsx/.xls)");
                ui.radio_value(&mut app.selected_format, FormatType::Csv, "CSV");
            });

            if app.selected_format == FormatType::Excel && !app.sheets.is_empty() {
                egui::ComboBox::from_label("表格")
                    .selected_text(&app.selected_sheet)
                    .show_ui(ui, |ui| {
                        for sheet in &app.sheets.clone() {
                            ui.selectable_value(&mut app.selected_sheet, sheet.clone(), sheet);
                        }
                    });
            }
        });
}
```

```rust
// crates/gui/src/panels/output_config.rs
use crate::app::*;

pub fn show(ui: &mut egui::Ui, app: &mut VocabPptApp) {
    egui::CollapsingHeader::new("输出")
        .default_open(true)
        .show(ui, |ui| {
            ui.horizontal(|ui| {
                ui.label("路径:");
                ui.text_edit_singleline(&mut app.output_path);
                if ui.button("浏览...").clicked() {
                    if let Some(path) = rfd::FileDialog::new()
                        .add_filter("PowerPoint", &["pptx"])
                        .save_file() {
                        app.output_path = path.display().to_string();
                    }
                }
            });

            ui.horizontal(|ui| {
                ui.label("主题:");
                egui::ComboBox::from_id_salt("theme")
                    .selected_text(match app.selected_theme {
                        SlideTheme::Default => "默认",
                        SlideTheme::Blue => "蓝",
                        SlideTheme::Green => "绿",
                        SlideTheme::Gray => "灰",
                    })
                    .show_ui(ui, |ui| {
                        ui.selectable_value(&mut app.selected_theme, SlideTheme::Default, "默认");
                        ui.selectable_value(&mut app.selected_theme, SlideTheme::Blue, "蓝");
                        ui.selectable_value(&mut app.selected_theme, SlideTheme::Green, "绿");
                        ui.selectable_value(&mut app.selected_theme, SlideTheme::Gray, "灰");
                    });
            });
        });
}
```

```rust
// crates/gui/src/panels/data_preview.rs
use crate::app::*;
use egui_extras::{Column, TableBuilder};
use core::reader;

pub fn show(ui: &mut egui::Ui, app: &mut VocabPptApp) {
    match &app.state {
        AppState::Idle => {
            ui.label("请选择文件");
        }
        AppState::Loading { .. } => {
            ui.spinner();
            ui.label("加载中...");
        }
        AppState::Preview { entries, .. } => {
            if entries.is_empty() {
                ui.label("该表格无有效数据");
                return;
            }
            ui.label(format!("共 {} 条", entries.len()));

            let available_height = ui.available_height().min(300.0);
            egui::ScrollArea::vertical()
                .max_height(available_height)
                .show(ui, |ui| {
                    let preview_count = entries.len().min(100);
                    TableBuilder::new(ui)
                        .columns(Column::auto(), 6)
                        .header(20.0, |mut header| {
                            header.col(|ui| { ui.label("单词"); });
                            header.col(|ui| { ui.label("音标"); });
                            header.col(|ui| { ui.label("词根"); });
                            header.col(|ui| { ui.label("例句"); });
                        })
                        .body(|body| {
                            body.rows(18.0, preview_count, |mut row| {
                                let i = row.index();
                                if i < entries.len() {
                                    row.col(|ui| { ui.label(&entries[i].word); });
                                    row.col(|ui| { ui.label(&entries[i].phonetic); });
                                    row.col(|ui| { ui.label(&entries[i].morphology); });
                                    row.col(|ui| { ui.label(&entries[i].example); });
                                }
                            });
                        });
                });
        }
        AppState::Generating { current, total } => {
            ui.label(format!("生成中: {}/{}", current, total));
            ui.add(egui::ProgressBar::new(*current as f32 / *total as f32));
        }
        AppState::Done { count } => {
            ui.colored_label(egui::Color32::GREEN, format!("✓ 成功生成 {} 张幻灯片", count));
        }
        AppState::Error { message } => {
            ui.colored_label(egui::Color32::RED, format!("⚠ {}", message));
        }
    }
}
```

```rust
// crates/gui/src/panels/mod.rs
pub mod file_picker;
pub mod data_preview;
pub mod output_config;
```

- [ ] **步骤 4: 编译验证**

```bash
cargo build -p gui
```

预期：编译通过。

- [ ] **步骤 5: 冒烟测试**

```bash
cargo run -p gui
```

预期：窗口打开，显示布局骨架。手动测试文件选择、预览、生成流程。

- [ ] **步骤 6: Commit**

```bash
git add -A && git commit -m "feat(gui): 基础 GUI — 状态机, 文件选择, 数据预览, 生成流程"
```

---

## Phase 7: 集成收尾

### 任务 7.1: 总体验证 + 文档

- [ ] **步骤 1: 全量测试**

```bash
cargo test --workspace
cargo clippy --workspace
cargo fmt --check --all
```

- [ ] **步骤 2: 端到端冒烟**

```bash
# CLI Excel
cargo run -p cli -- generate -i assets/template.xlsx -o /tmp/e2e_excel.pptx -f
# CLI CSV
cargo run -p cli -- template -o /tmp/e2e_template.xlsx
# GUI 手动
cargo run -p gui
```

- [ ] **步骤 3: Commit**

```bash
git add -A && git commit -m "chore: 全量测试通过, 端到端验证"
```

---

## 附录：备选方案（ppt-rs spike 失败时）

如果 ppt-rs 无法满足需求（CJK 字体、文本框定位），降级为手工 OOXML：

1. 使用 `zip` crate 创建 ZIP 归档
2. 手写 `[Content_Types].xml`、`ppt/presentation.xml`、`ppt/slides/slide1.xml` 等
3. `ppt/slides/slide1.xml` 中包含 `<a:p>` 段落节点，设置 `a:ea` 字体
4. 参考 ECMA-376 标准和 Python 版生成的 pptx 结构

此备选方案将显著增加 Phase 3 工时（估计 +3-5 天）。
