# 单词 PPT 生成器 — Rust 重构设计规格

**日期**: 2026-07-24（修订 v2，基于独立审查）  
**状态**: 已批准（修订中）  
**来源**: Python-Excel-to-PPT-Vocabulary-Tool 重构

---

## 1. 概述

将现有的 Python + Tkinter 单词 PPT 生成器重构为基于 Rust 的跨平台应用程序。

### 1.1 目标

- 保持原有核心功能（Excel → PPT 单词幻灯片）
- 新增 CSV 输入支持、进度条（可取消）、PPT 主题选择、批量处理
- CLI + GUI 双模式
- 跨平台：Windows / macOS / Linux

### 1.2 非目标

- 不支持 PPT 以外的输出格式
- 不支持 Excel 以外的输入格式（CSV 为唯一新增）
- 不做插件系统
- 不做在线/云端版本

---

## 2. 架构

### 2.1 Workspace 结构

```
单词PPT生成器/                       # workspace root（用户可见名称）
├── Cargo.toml                       # [workspace] members
├── crates/
│   ├── core/                        # 纯逻辑，无 UI 依赖
│   │   ├── Cargo.toml
│   │   ├── src/
│   │   │   ├── lib.rs               # 公开 API
│   │   │   ├── types.rs             # WordEntry、SlideConfig、InputSource、SlideTheme
│   │   │   ├── reader.rs            # Excel/CSV 读取与列名校验
│   │   │   ├── generator.rs         # PPTX 生成
│   │   │   └── template.rs          # 模板创建与导出
│   │   └── tests/
│   │       ├── reader_test.rs
│   │       ├── generator_test.rs
│   │       ├── types_test.rs
│   │       └── template_test.rs
│   ├── cli/                         # CLI 入口（clap）
│   │   ├── Cargo.toml
│   │   └── src/main.rs
│   └── gui/                         # GUI 入口（egui + eframe）
│       ├── Cargo.toml
│       └── src/
│           ├── main.rs              # eframe 启动
│           ├── app.rs               # eframe::App trait 实现，状态机
│           └── panels/
│               ├── file_picker.rs   # 文件选择、格式切换、sheet 下拉
│               ├── data_preview.rs  # 数据预览表格
│               └── output_config.rs # 输出路径、主题下拉
└── assets/
    └── template.xlsx                # 内置 Excel 模板
```

### 2.2 依赖链

```
cli ──→ core
gui ──→ core
core（零 UI 依赖）
```

### 2.3 核心依赖

| 功能 | Crate | 原因 |
|------|-------|------|
| PPTX 生成 | `ppt-rs` | 已验证活跃（yingkitw/ppt-rs, 2026-07），完整 OOXML 合规 |
| GUI | `egui` + `eframe` | 即时模式，轻量，原生跨平台 |
| CLI 参数 | `clap` (derive) | 行业标准 |
| CSV 读取 | `csv` | 标准级；编码检测用 `encoding_rs` |
| 文件对话框 | `rfd` | 原生对话框，跨平台 |
| 错误处理 | `thiserror` | 结构化错误类型 |
| 系统文件打开 | `open` | 跨平台调用默认程序打开模板 |

### 2.4 命名约定

- 项目文件夹、GUI 标签、CLI 命令、CLI 二进制名：**中文**（面向用户）
  - CLI 二进制名：`单词ppt`（Windows）/ 建议 ASCII 别名 `wordppt`（Linux/macOS 终端兼容）
- Crate 名、.rs 文件名、代码符号：**英文**（面向开发者，行业惯例）
- 中文列名作为内部 key 时，错误消息携带中文显示名；测试断言的错误变体使用内部 key（如 `MissingColumn("word")` 而非 `MissingColumn("英文单词")`）

---

## 3. 数据模型

### 3.1 WordEntry

```rust
#[derive(Debug, Clone)]
struct WordEntry {
    word: String,                // 英文单词
    phonetic: String,            // 英文音标
    morphology: String,          // 词根词缀
    example: String,             // 例句
    example_definition: String,  // 例句释义
    definition: String,          // 单词释义
}
```

**校验规则**（与 Python 版行为一致）：
- `word`（英文单词）为空字符串 → 整行跳过，记录警告
- 其他字段为空 → 保留空字符串，不拒绝（原 Python 版 `str(value) if value is not None else ''`）
- 整行所有字段全空 → 跳过

### 3.2 SlideConfig

```rust
struct SlideConfig {
    // 幻灯片尺寸（英寸）
    width: f32,               // 默认 16
    height: f32,              // 默认 9
    // 字号（pt）
    word_font_size: f32,      // 默认 72
    phonetic_font_size: f32,  // 默认 32
    content_font_size: f32,   // 默认 32
    // 字体族（跨平台中文渲染）
    font_family_cjk: String,  // 默认 "Noto Sans CJK SC"
    font_family_latin: String,// 默认 "Calibri"
    text_wrap_threshold: usize, // 默认 40 字符
}
```

**字体回退策略**：
- Windows：优先 `Microsoft YaHei`，回退 `SimHei`
- macOS：优先 `PingFang SC`，回退 `Heiti SC`
- Linux：优先 `Noto Sans CJK SC`，回退 `WenQuanYi Micro Hei`
- 生成器在写入 PPTX 时设定 `latin` 和 `ea`（East Asian）两种字体族
- 文本框布局位置（left/top/width/height）作为编译期常量定义在 `generator.rs` 中，不在 `SlideConfig` 暴露——布局属于内部实现细节，YAGNI 暂不外部化

### 3.3 InputSource

```rust
enum InputSource {
    Excel { path: PathBuf, sheet: String },
    Csv { path: PathBuf, encoding: String },  // encoding 默认 "UTF-8"
}
```

### 3.4 SlideTheme（新增）

```rust
enum SlideTheme {
    Default,   // 白底黑字
    Blue,      // 蓝底白字
    Green,     // 绿底白字
    Gray,      // 灰底黑字
}
```

GUI 显示标签（中文）由 `gui` crate 的 `Display` impl 负责，不在 `core` 中硬编码翻译。

---

## 4. 核心流水线

### 4.1 公开 API（core/src/lib.rs）

```rust
// 列出 Excel 文件中的所有 sheet 名称
pub fn list_sheets(path: &Path) -> Result<Vec<String>, LoadError>;

// 读取 Excel 或 CSV，返回 WordEntry 列表
pub fn load(source: &InputSource) -> Result<Vec<WordEntry>, LoadError>;

// 生成 PPTX 文件。
// progress 回调: (current, total) → 返回 false 则取消生成
pub fn generate(
    entries: &[WordEntry],
    output: &Path,
    config: &SlideConfig,
    theme: SlideTheme,
    progress: impl Fn(usize, usize) -> bool,
) -> Result<(), GenerateError>;

// 导出内置模板到指定路径
pub fn export_template(path: &Path) -> Result<(), TemplateError>;
```

**输出文件冲突行为**：
- 文件已存在时 `generate` 返回 `GenerateError::FileExists(path)`，由调用方决定处理
- CLI: `--force` / `-f` 标志覆盖；无此标志则报错退出
- GUI: 弹出确认对话框 "文件已存在，是否覆盖？"
### 4.3 generator 模块

- 使用 `ppt-rs` crate 创建和操作 PPTX 文件
- 每张 16:9 幻灯片，字段布局与 Python 版一致（单词居中放大、音标次行、词根词缀、例句+释义、单词释义）
- 通过 `SlideContent` API 添加文本框，`Position`/`Size` 控制布局
- 文本超 `text_wrap_threshold` 自动换行
- 每生成一张幻灯片调用 `progress(idx, total)`：
  - 返回 `true` → 继续
  - 返回 `false` → 取消，`generate` 返回 `GenerateError::Cancelled`
- 主题通过 `SchemeColor` 修改幻灯片背景色和字体色
- 中文字体：通过 `SlideConfig.font_family_cjk` 设置，需 spike 验证 `ppt-rs` 的东亚字体支持
### 4.4 template 模块

- 模板路径：用 `concat!(env!("CARGO_MANIFEST_DIR"), "/../../assets/template.xlsx")` 或 `include_bytes!` + 运行时写出到临时文件
  - 或更简单：将 `assets/template.xlsx` 放在 workspace root，从 `core/Cargo.toml` 不可见。改为放在 `core/assets/template.xlsx`，用 `include_bytes!("../assets/template.xlsx")` 嵌入
- `export_template` 按需写出到用户指定路径

---

## 5. GUI 设计

### 5.1 布局

```
┌─────────────────────────────────────────────────┐
│  单词PPT生成器                          ─ □ ×   │
├─────────────────────────────────────────────────┤
│  ┌─ 输入 ─────────────────────────────────────┐ │
│  │  文件: [________________] [浏览...]        │ │
│  │  格式: ○ Excel (.xlsx/.xls)  ○ CSV        │ │
│  │  表格: [▼ Sheet1           ]  (Excel 模式) │ │
│  └────────────────────────────────────────────┘ │
│  ┌─ 输出 ─────────────────────────────────────┐ │
│  │  路径: [________________] [浏览...]        │ │
│  │  主题: [▼ 默认 | 蓝 | 绿 | 灰 ]           │ │
│  └────────────────────────────────────────────┘ │
│  ┌─ 状态 ─────────────────────────────────────┐ │
│  │  [数据预览 / 错误信息 / 加载中...]         │ │
│  │  ⚠ 无法打开文件：xxx.xlsx 被占用           │ │
│  └────────────────────────────────────────────┘ │
│  ┌─ 进度 ─────────────────────────────────────┐ │
│  │  ████████████████░░░░░░  32/50   [取消]    │ │
│  └────────────────────────────────────────────┘ │
│        [打开模板]          [生成PPT]            │
└─────────────────────────────────────────────────┘
```

### 5.2 面板职责

| 面板（文件） | 职责 |
|-------------|------|
| `file_picker.rs` | 文件选择对话框、格式切换（Excel/CSV radio）、sheet 下拉列表（调用 `core::list_sheets()`） |
| `data_preview.rs` | 读取后的数据表格预览，前 100 行截断，显示总条数，空数据提示 |
| `output_config.rs` | 输出路径选择、主题下拉、encoding 输入（仅 CSV 模式） |

### 5.3 状态机（app.rs）

```rust
enum AppState {
    Idle,                          // 初始状态
    Loading { path: PathBuf },     // 正在读取文件
    Preview { entries: Vec<WordEntry> }, // 数据已加载，可预览
    Generating { current: usize, total: usize }, // 正在生成 PPT
    Done { count: usize },         // 生成完成
    Error { message: String },     // 错误状态
}
```

### 5.4 各状态下的 GUI 行为

| 状态 | 文件选择 | 预览区 | "生成PPT"按钮 | 进度条 |
|------|----------|--------|-------------|--------|
| Idle | 可用 | 空，提示"请选择文件" | 禁用（灰色） | 隐藏 |
| Loading | 可用 | "加载中..." spinner | 禁用 | 隐藏 |
| Preview（有数据） | 可用 | 表格预览 | 可用 | 隐藏 |
| Preview（空数据） | 可用 | "该表格无有效数据" | 禁用 | 隐藏 |
| Generating | 禁用 | 保持上次预览 | 变为"取消" | 显示+更新 |
| Done | 可用 | 保持预览 | 可用 | 100% 绿色 |
| Error | 可用 | 错误信息（红色） | 禁用 | 隐藏 |

### 5.5 交互逻辑

1. 用户选择文件 → 自动检测格式（扩展名判断）
2. Excel 文件 → 调用 `core::list_sheets()` 填充下拉框 → 用户选择 sheet → 调用 `core::load()` 加载预览
3. CSV 文件 → 直接调用 `core::load()` 加载预览
4. 预览显示前 100 行（截断保护）；0 行时显示"无有效数据"
5. 点击"生成PPT" → 状态转为 Generating，按钮变为"取消"，异步生成 + 进度回调
6. 生成中点击"取消" → progress 回调返回 false → `core::generate()` 返回 `Cancelled`
7. 生成完成 → 状态栏显示 "成功生成 N 张幻灯片"
8. 如果输出文件已存在 → 弹确认对话框
9. "打开模板" → 导出内置模板到系统临时目录 → `open::that()` 打开

### 5.6 关键 Crate

- `egui` + `eframe`：GUI 框架
- `rfd`：原生文件对话框
- `egui_extras`：Table 组件用于数据预览

---

## 6. CLI 设计

### 6.1 命令

```bash
# Excel 单文件
单词ppt generate --input words.xlsx --sheet "单词表" --output words.pptx

# CSV（指定编码）
单词ppt generate --input words.csv --output words.pptx --encoding GBK

# 强制覆盖已存在的输出文件
单词ppt generate --input words.xlsx --output words.pptx --force

# 批量（遍历目录，continue-on-error）
单词ppt batch --input ./data/ --output ./output/ --sheet "单词表"

# 创建模板
单词ppt template --output 模板.xlsx
```

### 6.2 参数

| 参数 | 适用命令 | 说明 | 默认值 |
|------|----------|------|--------|
| `--input` / `-i` | generate, batch | 输入文件/目录路径 | 必填（除 template） |
| `--output` / `-o` | generate, batch, template | 输出文件/目录路径 | `output.pptx` |
| `--sheet` / `-s` | generate, batch | Excel sheet 名称 | 第一个 sheet |
| `--encoding` / `-e` | generate, batch | CSV 编码：utf-8/gbk/gb2312/gb18030 | `utf-8` |
| `--theme` / `-t` | generate, batch | 主题：default/blue/green/gray | `default` |
| `--force` / `-f` | generate, batch | 覆盖已存在的输出文件 | false |

### 6.3 批量错误策略

`batch` 子命令：**continue-on-error**。遍历目录下所有 `.xlsx`/`.xls`/`.csv` 文件，逐个处理：
- 单个文件失败 → 打印错误到 stderr，继续下一个
- 全部完成后，stdout 打印汇总：`处理 N 个文件，成功 M 个，失败 K 个`
- 退出码：全部成功 → 0，有失败 → 1

---

## 7. 测试策略

## 8. 技术风险

### 8.1 PPTX 生成 — 已确认方案

**方案**：使用 `ppt-rs` crate（crates.io, yingkitw/ppt-rs, GitHub 46★, 2026-07-11 活跃）。

经 AnySearch 验证，`ppt-rs` 提供：
- `Presentation` builder / `SlideContent`（bullet、text）
- `Color`、`Position`、`Size`、`Transform` 元素类型
- `SchemeColor` 主题色支持
- `create_pptx_with_content()` 高级 API
- 完整 OOXML 标准合规

**风险等级**：低。crate 活跃、API 成熟、满足项目需求（文本框定位、中文字体、背景色）。

**仍需 spike 验证**：中文 CJK 字体族 `a:ea` 属性设置是否被正确渲染。

### 8.2 模板打开（低）

调用系统默认程序打开 `.xlsx`：使用 `open` crate，跨平台一致。

### 8.3 中文编码路径（低）

Windows 上中文路径可能触发编码问题。使用 `camino::Utf8Path` 或严格 UTF-8 路径处理。


---

## 9. 新增功能清单

| 功能 | 说明 | Python 版 |
|------|------|-----------|
| CSV 输入 | `csv` crate + `encoding_rs` 多编码支持 | 无 |
| 进度条（可取消） | `generate` 接受 `Fn(usize, usize) -> bool` 回调 | 无 |
| PPT 主题 | 4 种配色（默认/蓝/绿/灰） | 无 |
| 批量处理 | `batch` 子命令，continue-on-error，汇总报告 | 无 |
| 跨平台 | Windows/macOS/Linux | 仅 Windows exe |
| 数据预览 | GUI 内预览表格 | 无，直接生成 |
| GUI 状态管理 | 显式状态机：Idle/Loading/Preview/Generating/Done/Error | 无 |
| 文件覆盖保护 | CLI `--force` / GUI 确认对话框 | 静默覆盖 |
