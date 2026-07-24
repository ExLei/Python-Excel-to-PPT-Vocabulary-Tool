# 诊断系统参考

> **受众：** AI Agent + 人类用户  
> **版本：** v2.0  
> **最后更新：** 2026-07-24

---

## 目录

1. [快速开始](#快速开始)
2. [NDJSON 日志格式](#ndjson-日志格式)
3. [`diag` 子命令参考](#diag-子命令参考)
4. [Agent 诊断工作流](#agent-诊断工作流)
5. [人类故障排除指南](#人类故障排除指南)
6. [程序化查询（jq / Python）](#程序化查询)
7. [日志文件生命周期](#日志文件生命周期)

---

## 快速开始

每次 PNG 导出自动生成诊断日志。两步诊断任何问题：

```bash
# 步骤 1：导出
单词卡片转换 export-png -i words.xlsx -o ./output/

# 步骤 2：查看诊断
单词卡片转换 diag ./output/export_*.ndjson --summary
```

---

## NDJSON 日志格式

每行一个 JSON 对象。所有事件共享以下字段：

```json
{
  "timestamp": "21:30:01.234",
  "level": "INFO",
  "target": "png_export::render",
  "message": "PNG: 63196 bytes, 12740 text px (0.61%)",
  "fields": {}
}
```

### 字段说明

| 字段 | 类型 | 说明 |
|------|------|------|
| `timestamp` | string | 事件时间（自导出开始的相对时间，格式 `MM:SS.mmm`） |
| `level` | `"INFO"` \| `"WARN"` \| `"ERROR"` | 严重级别 |
| `target` | string | 事件来源模块（见下方 target 表） |
| `message` | string | 人类可读描述 |
| `fields` | object \| null | 结构化附加上下文（键值对） |

### `target` 命名空间

| target | 发出者 | 触发时机 |
|--------|-------|---------|
| `export` | png_export/mod.rs | 导出开始/结束、slide 计数汇总 |
| `font_probe` | png_export/pipeline.rs | 字体探测链、选中/缺失字体 |
| `parse` | png_export/pipeline.rs | PPTX XML 解析、属性转换、解析失败 |
| `slide` | png_export/pipeline.rs | 每张 slide 的字段渲染/跳过、SVG 大小 |
| `render` | png_export/pipeline.rs | PNG 字节数、文字像素数、密度 |
| `verify` | png_export/pipeline.rs | 空白页检测、低密度警告 |
| `reader` | reader.rs | 文件打开、编码检测、行解析/跳过 |
| `generator` | generator.rs | PPTX 生成进度、slide 复制、ZIP 写入 |
| `template` | template_reader.rs | 占位符扫描/替换结果 |
| `cli` | cli/lib.rs | 命令调用、参数、结果 |
| `gui` | gui/app.rs | 文件选择、导出启动、错误 |

### `fields` 常见键

**export 事件：**
```json
{"entry_count": 5, "output_dir": "/tmp/output"}
{"succeeded": 5, "failed": 0, "total": 5}
```

**font_probe 事件：**
```json
{"chain": "latin", "candidate": "Segoe UI", "faces_found": 0, "selected": false}
{"chain": "latin", "candidate": "DejaVu Sans", "faces_found": 1, "selected": true, "path": "/usr/share/fonts/TTF/DejaVuSans.ttf"}
{"font_stack": "Noto Sans CJK SC, DejaVu Sans, sans-serif"}
```

**slide 事件：**
```json
{"index": 0, "word": "apple"}
{"fields_rendered": ["单词", "音标", "单词释义"], "fields_skipped": [{"name": "例句", "reason": "empty"}]}
{"svg_bytes": 1842, "text_elements": 3}
```

**render 事件：**
```json
{"png_bytes": 63196, "text_pixels": 12740, "text_density_pct": 0.61}
{"sampled": [{"x": 960, "y": 540, "rgba": [30, 30, 30, 255]}, {"x": 480, "y": 270, "rgba": [255, 255, 255, 255]}]}
```

**verify 事件（WARN/ERROR）：**
```json
{"reason": "low_text_density", "density_pct": 0.006, "threshold": 0.01}
{"reason": "blank_slide", "text_elements_in_svg": 3, "text_pixels": 0, "likely_cause": "font issue"}
```

**parse 事件（WARN）：**
```json
{"attribute": "x", "raw_value": "invalid", "fallback": 0}
{"attribute": "sz", "raw_value": "abc", "fallback": 24.0}
```

**reader 事件：**
```json
{"file": "words.xlsx", "format": "xlsx", "sheet": "Sheet1"}
{"encoding": "GBK", "had_decode_errors": false}
{"total_parsed": 50, "total_skipped": 2}
{"row": 3, "reason": "empty_word"}
```

**generator 事件：**
```json
{"slide_index": 0, "placeholders_found": ["单词", "音标"]}
{"template_source": "embedded"}
{"entries_processed": 5, "total_entries": 5}
```

---

## `diag` 子命令参考

### `--summary`

导出会话总览。**诊断第一步必用。**

```bash
单词卡片转换 diag export.ndjson --summary
```

输出示例：

```
  Session: 00:02.145 (2.1s)
  Slides: 5 total, 5 succeeded, 0 failed
  Font: cjk=Noto Sans CJK SC, latin=DejaVu Sans, emoji=未安装
  Warnings: 1
    → slide_2: low text density (0.006%)
  Errors: 0
  Log: /tmp/output/export_20260724_213001.ndjson
```

**Agent 解读：** 如果 `failed > 0`，查看 `--errors`。如果 `warnings > 0`，查看 `--blank-slides`。如果 emoji=未安装，通知用户安装 emoji 字体。

### `--blank-slides [阈值]`

列出文字密度低于阈值的 slide。默认阈值 0.01%。

```bash
单词卡片转换 diag export.ndjson --blank-slides
单词卡片转换 diag export.ndjson --blank-slides 0.05   # 自定义阈值
```

输出示例：

```
  slide_2: cherry    density=0.006% (127px / 2073600px)
    ├─ sampled: (960,540)=#1e1e1e  (480,270)=#ffffff  (1440,810)=#ffffff
    ├─ rendered: 单词(cherry), 音标(/ˈtʃeri/), 单词释义(樱桃)
    ├─ skipped: 例句(empty), 词根词缀(empty)
    └─ cause: text rendered but only in center region → tiny font or off-center position
                → check PPTX template x/y values for fields other than "单词"

  slide_5: dragon fruit    density=0.000% ← BLANK
    ├─ SVG: 342 bytes, 0 text elements
    ├─ all fields empty in source data → no content to render
    └─ cause: source data row has no non-empty fields → skip this row in Excel
```

**Agent 解读：**
- `density=0.000%` + `0 text elements` → 数据为空，不是 bug
- `density=0.000%` + `N text elements` → 字体问题，查 `--font-trace`
- `density<0.01%` + `sampled` 显示中心有文字 → 位置问题，查 PPTX 模板坐标

### `--font-trace`

字体探测的完整决策过程。

```bash
单词卡片转换 diag export.ndjson --font-trace
```

输出示例：

```
  latin chain:
    [✗] Segoe UI            — 0 faces
    [✗] Helvetica           — 0 faces
    [✗] Arial               — 0 faces
    [✓] DejaVu Sans         — 1 face  (/usr/share/fonts/TTF/DejaVuSans.ttf)
    → selected: DejaVu Sans

  cjk chain:
    [✗] Microsoft YaHei     — 0 faces (Windows 字体)
    [✗] PingFang SC         — 0 faces (macOS 字体)
    [✓] Noto Sans CJK SC    — 1 face  (/usr/share/fonts/opentype/NotoSansCJK-Regular.ttc)
    → selected: Noto Sans CJK SC

  emoji chain:
    [✗] Segoe UI Emoji      — 0 faces
    [✗] Apple Color Emoji   — 0 faces
    [✗] Noto Color Emoji    — 0 faces
    ⚠ ALL EMOJI FONTS MISSING — ✓✗😀 等字符将显示为方框

  font-stack: "Noto Sans CJK SC, DejaVu Sans, sans-serif"
```

**Agent 解读：**
- 所有链都至少有一个 `[✓]` → 字体正常
- 某链全部 `[✗]` → 该语言字符将无法显示
- `emoji chain` 全 `[✗]` → 不是严重问题，emoji 在词汇表中少见
- `latin chain` 全 `[✗]` → 严重，英文和 IPA 音标都会缺失
- `cjk chain` 全 `[✗]` → 严重，中文无法显示

### `--errors`

列出所有 WARN 和 ERROR 事件。

```bash
单词卡片转换 diag export.ndjson --errors
```

输出示例：

```
  WARN  font_probe         emoji chain: all 3 candidates missing
  WARN  verify             slide_2: low text density 0.006% (< 0.01% threshold)
  ERROR render             slide_4: SVG parse failed at line 12: unexpected closing tag </text>
  ERROR font_probe         font loading CRASHED — corrupt system font detected
```

**Agent 解读：**
- `ERROR font_probe ... CRASHED` → 系统有损坏字体，建议用户检查 `~/.fonts` 或 `C:\Windows\Fonts`
- `ERROR render ... SVG parse failed` → 内部 bug，SVG 生成器产生了无效 XML
- `WARN verify ... low text density` → 非致命，但需人工确认

### `--slide <N>`

单张 slide 完整诊断。N 从 0 开始。

```bash
单词卡片转换 diag export.ndjson --slide 2
```

输出示例：

```
  slide_2: cherry
    ├─ SVG: 2104 bytes, 3 text elements
    ├─ PNG: 58234 bytes, 1920×1080
    ├─ text pixels: 9408 (0.45%)
    ├─ sampled: (960,540)=#1e1e1e  (480,270)=#1e1e1e  (1440,810)=#ffffff
    ├─ fields rendered:
    │   单词(cherry)    x=144  y=100  sz=72pt  bold  color=#1e1e1e
    │   音标(/ˈtʃeri/)  x=144  y=180  sz=24pt        color=#646464
    │   单词释义(樱桃)   x=144  y=280  sz=26pt        color=#1e1e1e
    └─ fields skipped:
        例句(empty) — 源数据中例句列为空
        词根词缀(empty) — 源数据中词根词缀列为空
```

**Agent 解读：**
- `fields rendered` 列出实际渲染的字段及其布局 → 验证 PPTX 模板解析是否正确
- `fields skipped` 列出原因 → 判断是数据问题还是模板问题
- `sampled` 三点像素 → 判断文字是否在预期位置

### `--json`

输出原始 JSON，供脚本消费。

```bash
# 获取所有空白 slide 的 JSON
单词卡片转换 diag export.ndjson --json --blank-slides

# 获取字体探测结果
单词卡片转换 diag export.ndjson --json --font-trace
```

---

## Agent 诊断工作流

### 场景 1：用户报告 "PNG 一片空白"

```
诊断路径：
  ┌─────────────────────────────────────────────────────────┐
  │ 1. diag --summary                                       │
  │    → failed > 0?  → --errors                            │
  │    → warnings > 0? → --blank-slides                     │
  ├─────────────────────────────────────────────────────────┤
  │ 2. diag --blank-slides                                  │
  │    → density=0% + text_elements=0?                      │
  │       → 源数据为空，不是 bug。告诉用户检查 Excel。        │
  │    → density=0% + text_elements>0?                       │
  │       → 字体问题。继续步骤 3。                            │
  │    → density>0% but <0.01%?                              │
  │       → 文字渲染了但极小/偏位。继续步骤 4。                │
  ├─────────────────────────────────────────────────────────┤
  │ 3. diag --font-trace                                    │
  │    → latin chain 全 [✗]? → 无拉丁字体，英文/IPA 缺失     │
  │    → cjk chain 全 [✗]?   → 无 CJK 字体，中文缺失        │
  │    → font-stack 只有 "sans-serif"?                       │
  │       → 零字体可用。建议安装 DejaVu Sans + Noto Sans CJK │
  │    → ERROR "font loading CRASHED"?                       │
  │       → 损坏的系统字体。建议运行 fc-cache -f 或清理字体   │
  ├─────────────────────────────────────────────────────────┤
  │ 4. diag --slide <N>                                     │
  │    → sampled 三点都是 #ffffff? → 文字完全没渲染           │
  │    → sampled 中心有文字但密度低? → 字体太小或位置偏了     │
  │    → fields rendered 为空? → 所有字段都被跳过             │
  └─────────────────────────────────────────────────────────┘
```

### 场景 2：用户报告 "某些字段没显示"

```
诊断路径：
  1. diag --slide <问题slide的序号>
  2. 查看 "fields skipped" 段落
     → reason="empty"           → 源数据该列为空
     → reason="no_placeholder"  → 模板缺少该字段的占位符
     → reason="parse_error"     → 模板中该占位符的 XML 解析失败
  3. 如果 reason="no_placeholder":
     → 用 PowerPoint 打开模板，确认是否放了 {{字段名}}
```

### 场景 3：用户报告 "中文显示为方框" 或 "音标显示为方框"

```
诊断路径：
  1. diag --font-trace
  2. 查看 "cjk chain"（中文）或 "latin chain"（音标）
  3. 全部 [✗] → 指导用户安装对应字体：
     - Windows: 微软雅黑 / Segoe UI（通常已自带）
     - macOS:   PingFang SC / Helvetica（通常已自带）  
     - Linux:   sudo apt install fonts-noto-cjk fonts-dejavu
```

### 场景 4：批量导出后自动质量检查（CI/CD）

```bash
#!/bin/bash
# 自动检查：任何 slide 文字密度 < 0.01% 则失败

BLANK_COUNT=$(单词卡片转换 diag export.ndjson --json --blank-slides 0.01 | jq 'length')

if [ "$BLANK_COUNT" -gt 0 ]; then
    echo "::error:: $BLANK_COUNT slides have low text density"
    单词卡片转换 diag export.ndjson --blank-slides 0.01
    exit 1
fi
echo "::notice:: All slides passed blank check"
```

### 场景 5：用户报告 "导出到一半卡住了"

```
诊断路径：
  1. 查看 NDJSON 最后几行：
     tail -5 export.ndjson
  2. 最后一条 slide 事件是哪个？→ 定位卡住位置
  3. 如果是 "cancelled" → 用户点了取消按钮（正常）
  4. 如果是某 slide 之后无后续事件 → 该 slide 渲染卡住
     → diag --slide <最后成功的slide序号+1> 查看
```

---

## 人类故障排除指南

### "导出成功了但 PNG 是白的"

```bash
# 1. 看总览
单词卡片转换 diag export.ndjson --summary

# 2. 如果显示 "Warnings: 1 → slide_X: low text density" 
单词卡片转换 diag export.ndjson --slide X

# 3. 根据输出判断：
#    - "0 text elements" → Excel 数据为空，检查源文件
#    - "3 text elements, 0 pixels" → 字体问题
#    - "sampled 显示中心有文字" → PPTX 模板位置不对
```

### "中文/音标显示为方框"

```bash
单词卡片转换 diag export.ndjson --font-trace
# 看 cjk chain（中文）或 latin chain（音标）是否全部 [✗]
# 如果是 → 安装对应字体
# Linux:   sudo apt install fonts-noto-cjk fonts-dejavu
# Windows: 确保系统有 微软雅黑 或 Segoe UI
# macOS:   通常自带 PingFang SC / Helvetica
```

### "生成了 PPTX 但有些字段是空的"

这个诊断目前针对 PNG 导出。PPTX 生成的诊断在后续版本中完善。

### "不想用命令行，GUI 怎么看诊断？"

导出完成后，GUI 底部状态栏显示诊断摘要。点击"查看诊断"按钮展开详情面板。

---

## 程序化查询

### jq 查询

```bash
# 所有 ERROR 事件
grep '"ERROR"' export.ndjson | jq '{target, message, fields}'

# 每张 slide 的文字密度
grep '"text_density"' export.ndjson | jq '{index: .fields.index, word: .fields.word, density: .fields.text_density_pct}'

# 字体探测摘要
grep '"font_probe"' export.ndjson | grep '"selected":true' | jq '{chain: .fields.chain, font: .fields.candidate}'

# 统计：多少 slide 被跳过、多少字段被跳过
grep '"fields_skipped"' export.ndjson | jq '[.fields.fields_skipped[]?.name] | unique'

# 找出密度最低的 slide
grep '"text_density"' export.ndjson | jq -s 'sort_by(.fields.text_density_pct) | .[0]'
```

### Python 批量分析

```python
import json, sys

def analyze(ndjson_path: str):
    events = []
    with open(ndjson_path) as f:
        for line in f:
            if line.strip():
                events.append(json.loads(line))

    # 统计
    errors = [e for e in events if e["level"] == "ERROR"]
    warns  = [e for e in events if e["level"] == "WARN"]
    slides = [e for e in events if "text_density_pct" in (e.get("fields") or {})]

    print(f"Total events: {len(events)}")
    print(f"Errors: {len(errors)}, Warnings: {len(warns)}")

    # 空白 slide
    for s in slides:
        density = s["fields"]["text_density_pct"]
        if density < 0.01:
            print(f"  ⚠ slide {s['fields']['index']} ({s['fields']['word']}): {density:.4f}%")

    # 字体问题
    font_events = [e for e in events if e["target"] == "font_probe"]
    missing_chains = set()
    for e in font_events:
        if "all candidates missing" in e["message"]:
            missing_chains.add(e["fields"].get("chain", "unknown"))
    if missing_chains:
        print(f"Missing font chains: {missing_chains}")

if __name__ == "__main__":
    analyze(sys.argv[1])
```

---

## 日志文件生命周期

### 命名

```
输出目录/export_YYYYMMDD_HHmmss.ndjson

示例：
  ./output/export_20260724_213001.ndjson
```

### 大小估算

| Slide 数量 | 日志大小 |
|-----------|---------|
| 5 | ~3 KB |
| 50 | ~25 KB |
| 500 | ~250 KB |
| 5000 | ~2.5 MB |

### 清理

日志文件不自动删除。建议：

```bash
# 保留最近 10 个日志
ls -t ./output/export_*.ndjson | tail -n +11 | xargs rm -f
```
