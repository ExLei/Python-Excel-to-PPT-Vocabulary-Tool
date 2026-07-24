# Copilot Instructions — 英语助记卡片生成

> 此文件为 GitHub Copilot 提供项目专属上下文。AGENTS.md 中的完整规则同样适用。

## 关键规则（Copilot 必须遵守）

1. **TDD 强制** — 在实现 `fn` 之前先写 `#[test]`
2. **每个可失败操作注入 `diag: &mut DiagStore`** — 零静默失败
3. **禁止 `let _ =` 吞没错误** — 用 `match` + `diag.error/warn`
4. **公共 API 签名不可变** — 新功能用新函数，旧函数保持兼容
5. **不引入 `rusqlite` 或新重量依赖** — 用 NDJSON
6. **`r#"..."#` 包含 `"#` 时用 `r##"..."##`**

## 常用模式

```rust
// 读取数据 + 诊断
let entries = reader::load(&source, &mut diag)
    .map_err(|e| { diag.error("reader", &format!("{e}"), None); e })?;

// 处理每项 + 取消支持
for (i, item) in items.iter().enumerate() {
    if cancel.load(Ordering::Relaxed) { diag.warn("task", "cancelled", None); break; }
    match process_one(item, &mut diag) {
        Ok(r) => results.push(r),
        Err(e) => diag.error("task", &format!("item {i}: {e}"), None),
    }
}

// 渲染后验证
let density = count_pixels(&output);
diag.info("render", &format!("{density:.2}%"), None);
if density < 0.01 { diag.warn("verify", "low density", None); }
```

## 诊断查询（Agent 自检用）

任务完成后运行：
```bash
cargo test --workspace                          # 全绿
cargo run --release -- diag <log> --summary     # 无 ERROR
```

完整规则见 `AGENTS.md` 和 `docs/diagnostics.md`。
