# 单词卡片转换

Excel/CSV 词汇表 → 16:9 PPTX 单词幻灯片，一键生成。

纯 Rust 实现，零运行时依赖。**一个二进制通吃 CLI + GUI**：双击启动图形界面，命令行传参即走 CLI。

- **双模式一个文件**：双击 → GUI；传参 → CLI
- **多格式输入**：Excel (.xlsx) / CSV（UTF-8 / GBK / GB2312 / GB18030）
- **自定义模板**：用户在 PowerPoint 中设计排版，软件填入数据
- **PPTX 渲染 PNG**：解析幻灯片 XML 精确还原排版，零外部依赖
- **批量处理**：遍历目录，continue-on-error，汇总报告
- **跨平台**：Windows / macOS / Linux
- **系统字体**：自动探测 + 多层回退，IPA 音标和中文正常显示
- **可诊断**：NDJSON 结构化日志 + `diag` 子命令，Agent 可自动定位问题
- **文件监听**：外部修改自动刷新预览

## 快速开始

### 图形界面

双击 `单词卡片转换.exe`：

1. **选择输入文件** — 支持 .xlsx / .csv，自动检测格式
2. **选择工作表** — 下拉菜单列出所有 sheet（Excel 模式）
3. **（可选）选择 PPTX 模板** — 含 `{{占位符}}` 的自定义版式
4. **数据预览** — 自动加载前 100 行
5. **生成PPT** — 显示进度，可随时取消
6. **导出PNG** — 有预览数据即可导出，内部自动生成 PPTX 再渲染

### 命令行

```bash
# 基本用法（-o 可选，默认与输入同级目录）
单词卡片转换 generate -i words.xlsx

# 使用自定义 PPTX 模板
单词卡片转换 generate -i words.xlsx -t 我的版式.pptx

# CSV 输入 + 指定编码
单词卡片转换 generate -i words.csv -e GBK

# 批量转换
单词卡片转换 batch -i ./data/ -o ./output/

# 强制覆盖
单词卡片转换 generate -i words.xlsx -f

# 导出 PPTX 为 PNG
单词卡片转换 export-png -i output.pptx -o ./png/

# Excel/CSV 直接导出 PNG（内部生成 PPTX → 渲染）
单词卡片转换 export-png -i words.xlsx -t 模板.pptx -o ./png/

# 生成 PPTX 示例模板（含 {{占位符}}）
单词卡片转换 template-pptx -o 示例模板.pptx

# 生成 Excel 词汇表模板
单词卡片转换 template -o 单词表模板.xlsx

# 诊断导出问题
单词卡片转换 diag export.ndjson --summary
单词卡片转换 diag export.ndjson --blank-slides
单词卡片转换 diag export.ndjson --font-trace
## 数据格式

Excel 需包含以下列（CSV 同理）：

| 列名 | 说明 | 必填 |
|------|------|------|
| 英文单词 | 单词拼写 | **是** |
| 英文音标 | IPA 音标 | 否 |
| 词根词缀 | 词根词缀分析 | 否 |
| 例句 | 包含该单词的例句 | 否 |
| 例句释义 | 例句的中文释义 | 否 |
| 单词释义 | 单词的中文释义 | 否 |

- `英文单词` 为空 → 整行跳过
- 其他字段为空 → 保留空位

## 模板使用

### PPTX 自定义模板

在 PowerPoint 中自由设计排版，软件读取后批量填入数据。

**工作流程：**

```text
1. 生成示例模板     →  单词卡片转换 template-pptx -o 我的版式.pptx
2. 在 PowerPoint 中编辑  →  拖动文本框、改字体、调大小、换颜色
3. 保存为模板       →  这就是你的模板文件
4. 填入数据生成 PPT  →  单词卡片转换 generate -i words.xlsx -t 我的版式.pptx
```

**占位符规范：**

| 占位符 | 对应字段 |
|--------|---------|
| `{{单词}}` | 英文单词（必填） |
| `{{音标}}` | IPA 音标 |
| `{{词根词缀}}` | 词根词缀 |
| `{{例句}}` | 例句 |
| `{{例句释义}}` | 例句的中文释义 |
| `{{单词释义}}` | 单词的中文释义 |

**规则：**
- `{{单词}}` 必须存在，否则报错
- 不需要的字段不放占位符即可
- 前后文字保留：`词根词缀：{{词根词缀}}` → `词根词缀：ap-ple`
- 条目数超过模板幻灯片数时自动复制最后一张
- PNG 导出会解析模板 XML，精确还原字号、颜色、对齐、位置

### Excel 数据模板

```bash
单词卡片转换 template -o 单词表.xlsx
```

## 构建

```bash
# 开发
cargo build && cargo test --workspace

# Windows 交叉编译（需 mingw-w64）
cargo build --release --target x86_64-pc-windows-gnu
```

产物：`target/x86_64-pc-windows-gnu/release/单词卡片转换.exe` (~4.6 MB)

## 架构

```
crates/
├── core/              # 纯逻辑，零 UI 依赖
│   ├── reader.rs           # Excel (calamine) / CSV 读取
│   ├── generator.rs        # PPTX 生成（模板文本替换）
│   ├── template_reader.rs  # 扫描/替换 {{占位符}}
│   ├── template_pptx.rs    # 生成示例 PPTX 模板
│   ├── template.rs         # Excel 模板导出
│   ├── png_export.rs       # PPTX → PNG（解析 XML → SVG → resvg 渲染）
│   └── types.rs            # WordEntry, 错误类型
├── cli/               # CLI 逻辑 (clap)，含 diag 诊断子命令
│   └── lib.rs
└── gui/               # GUI (egui/eframe) + 统一入口
    ├── main.rs             # args>1 → CLI；否则 → GUI + 系统字体
    ├── app.rs              # 状态机
    └── panels/
        ├── file_picker.rs
        ├── data_preview.rs
        └── output_config.rs
```

## 诊断

每次 PNG 导出自动生成 NDJSON 诊断日志。Agent 或人类可用 `diag` 子命令查询——无需写 SQL，无需 jq。

```bash
# 两步诊断任何问题
单词卡片转换 export-png -i words.xlsx -o ./output/     # 导出（自动生成日志）
单词卡片转换 diag ./output/export_*.ndjson --summary     # 查看诊断

# 常用查询
单词卡片转换 diag export.ndjson --blank-slides           # 哪些 slide 可能空白？
单词卡片转换 diag export.ndjson --font-trace             # 用了什么字体？哪些缺失？
单词卡片转换 diag export.ndjson --errors                 # 所有错误和警告
单词卡片转换 diag export.ndjson --slide 2                # 单张 slide 完整详情
```

**完整参考：** [诊断指南](docs/diagnostics.md) — NDJSON 格式规范、Agent 诊断工作流、jq/Python 脚本示例、日志文件管理。

## 故障排除

| 问题 | 诊断命令 | 解读 |
|------|---------|------|
| PNG 空白/文字缺失 | `diag export.ndjson --blank-slides` | density=0% + 有 text 元素 → 字体；density=0% + 无 text 元素 → 数据为空 |
| 中文显示为方框 | `diag export.ndjson --font-trace` | cjk chain 全部 [✗] → 安装中文字体 |
| 音标显示为方框 | `diag export.ndjson --font-trace` | latin chain 全部 [✗] → 安装 DejaVu / Segoe UI |
| emoji 显示为方框 | `diag export.ndjson --font-trace` | emoji chain 全部 [✗] → 安装 emoji 字体 |
| 部分字段未显示 | `diag export.ndjson --slide N` | 查看 fields_skipped → reason 列 |
| 导出卡住 | `tail -5 export.ndjson` | 最后一条事件定位卡住位置 |
| CSV 中文乱码 | — | 用 `-e GBK` 指定编码 |
| 模板缺少 `{{单词}}` | — | 至少一张幻灯片须含 `{{单词}}` |

## 许可证

GNU Affero General Public License v3.0，详见 [LICENSE](LICENSE)。
