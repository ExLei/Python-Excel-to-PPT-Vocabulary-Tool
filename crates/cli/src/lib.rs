use std::fs;
use std::path::{Path, PathBuf};

use clap::{Parser, Subcommand};
use vocab_core::generator::{generate, generate_from_template};
use vocab_core::diag::DiagStore;
use vocab_core::png_export;
use vocab_core::reader::{list_sheets, load};
use vocab_core::template::export_template;
use vocab_core::template_pptx::generate_example_pptx;
use vocab_core::types::InputSource;

/// 单词PPT生成器 — 从 Excel/CSV 词汇表一键生成 PPTX 单词课件
#[derive(Parser)]
#[command(name = "单词ppt", version, about)]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// 单文件生成：将词汇表转换为 PPTX
    Generate {
        /// 输入文件路径 (.xlsx / .xls / .csv)
        #[arg(short, long, value_name = "FILE")]
        input: PathBuf,

        /// 输出 PPTX 路径 (默认与输入同目录)
        #[arg(short, long, value_name = "FILE")]
        output: Option<PathBuf>,

        /// PPTX 模板文件 (含 {{占位符}} 的用户自定义模板)
        #[arg(short = 't', long, value_name = "FILE")]
        template: Option<PathBuf>,

        /// Excel sheet 名称 (不指定则自动取第一个)
        #[arg(short = 'S', long, value_name = "NAME")]
        sheet: Option<String>,

        /// CSV 文件编码 (默认 utf-8)
        #[arg(short, long, value_name = "ENCODING", default_value = "utf-8")]
        encoding: String,

        /// 强制覆盖已存在的输出文件
        #[arg(short, long)]
        force: bool,
    },

    /// 批量生成：遍历目录中所有词汇表
    Batch {
        /// 输入目录路径
        #[arg(short, long, value_name = "DIR")]
        input: PathBuf,

        /// 输出目录路径
        #[arg(short, long, value_name = "DIR", default_value = ".")]
        output: PathBuf,

        /// Excel sheet 名称 (不指定则自动取第一个)
        #[arg(short = 'S', long, value_name = "NAME")]
        sheet: Option<String>,

        /// CSV 文件编码 (默认 utf-8)
        #[arg(short, long, value_name = "ENCODING", default_value = "utf-8")]
        encoding: String,

        /// 强制覆盖已存在的输出文件
        #[arg(short, long)]
        force: bool,

        /// PPTX 模板文件 (含 {{占位符}} 的用户自定义模板)
        #[arg(short = 't', long, value_name = "FILE")]
        template: Option<PathBuf>,
    },

    /// 生成 PPTX 示例模板 (含 {{占位符}} 供用户修改)
    TemplatePptx {
        /// 输出示例模板文件路径
        #[arg(short, long, value_name = "FILE", default_value = "示例模板.pptx")]
        output: PathBuf,
    },

    ExportPng {
        /// 输入文件路径 (.xlsx / .csv)
        #[arg(short, long, value_name = "FILE")]
        input: PathBuf,

        /// 输出目录 (默认与输入同目录的 png_output)
        #[arg(short, long, value_name = "DIR")]
        output: Option<PathBuf>,

        /// PPTX 模板文件 (提取排版)
        #[arg(short = 't', long, value_name = "FILE")]
        template: Option<PathBuf>,

        /// Excel sheet 名称
        #[arg(short = 'S', long, value_name = "NAME")]
        sheet: Option<String>,

        /// CSV 编码 (默认 utf-8)
        #[arg(short, long, value_name = "ENCODING", default_value = "utf-8")]
        encoding: String,
    },
    /// 导出词汇表 Excel 模板
    Template {
        /// 输出模板文件路径 (默认 单词表模板.xlsx)
        #[arg(short, long, value_name = "FILE", default_value = "单词表模板.xlsx")]
        output: PathBuf,
    },

    /// 查看导出诊断日志
    Diag {
        /// NDJSON 诊断日志文件路径
        #[arg(value_name = "FILE")]
        file: PathBuf,

        /// 显示汇总统计
        #[arg(short, long)]
        summary: bool,

        /// 显示空白 slide 检测事件
        #[arg(long)]
        blank_slides: bool,

        /// 显示字体探测追踪
        #[arg(long)]
        font_trace: bool,

        /// 显示所有错误和警告
        #[arg(long)]
        errors: bool,

        /// 查看特定 slide 的详细信息
        #[arg(long, value_name = "N")]
        slide: Option<usize>,

        /// 以 JSON 格式输出 (脚本消费)
        #[arg(long)]
        json: bool,
    },
}

/// 根据扩展名和参数构建 InputSource
fn guess_source(path: &Path, sheet: Option<&str>, encoding: &str, diag: &mut DiagStore) -> Result<InputSource, String> {
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();

    match ext.as_str() {
        "csv" => Ok(InputSource::Csv {
            path: path.to_path_buf(),
            encoding: encoding.to_string(),
        }),
        _ => {
            let resolved_sheet = if let Some(s) = sheet {
                s.to_string()
            } else {
                list_sheets(path, diag)
                    .map_err(|e| format!("无法列出 sheet: {e}"))?
                    .into_iter()
                    .next()
                    .ok_or_else(|| format!("工作簿无 sheet: {}", path.display()))?
            };
            Ok(InputSource::Excel {
                path: path.to_path_buf(),
                sheet: resolved_sheet,
            })
        }
    }
}

fn generate_one(
    source: &InputSource,
    output: &Path,
    force: bool,
    template: Option<&Path>,
    diag: &mut DiagStore,
) -> Result<usize, String> {
    if force && output.exists() {
        fs::remove_file(output).map_err(|e| format!("无法删除 {output:?}: {e}"))?;
    }

    let entries = load(source, diag).map_err(|e| e.to_string())?;
    let total = entries.len();

    if let Some(tpl) = template {
        generate_from_template(&entries, tpl, output, |current, _total| {
            print!("\r生成进度: {current}/{_total}");
            true
        }, diag)
        .map_err(|e| e.to_string())?;
    } else {
        generate(&entries, output, |current, _total| {
            print!("\r生成进度: {current}/{_total}");
            true
        }, diag)
        .map_err(|e| e.to_string())?;
    }

    println!("\r完成: {} — {total} 页", output.display());
    Ok(total)
}

/// CLI 模式入口，由 main.rs 或 GUI 入口调用
pub fn run() -> Result<(), String> {
    let cli = Cli::parse();
    let mut diag = DiagStore::new();
    match cli.command {
        Command::Generate {
            input,
            output,
            template,
            sheet,
            encoding,
            force,
        } => {
            let source = guess_source(&input, sheet.as_deref(), &encoding, &mut diag)?;
            let output = output.unwrap_or_else(|| {
                let dir = input.parent().unwrap_or(Path::new("."));
                let stem = input
                    .file_stem()
                    .and_then(|s| s.to_str())
                    .unwrap_or("output");
                dir.join(format!("{stem}.pptx"))
            });
            generate_one(&source, &output, force, template.as_deref(), &mut diag).map(|_| ())
        }

        Command::Batch {
            input,
            output,
            sheet,
            encoding,
            force,
            template,
        } => {
            if !output.exists() {
                fs::create_dir_all(&output).map_err(|e| format!("无法创建输出目录: {e}"))?;
            }

            let entries = fs::read_dir(&input)
                .map_err(|e| format!("无法读取目录 {}: {e}", input.display()))?;

            let exts = ["xlsx", "xls", "csv"];
            let mut success: Vec<String> = Vec::new();
            let mut failures: Vec<(String, String)> = Vec::new();

            for entry in entries {
                let entry = match entry {
                    Ok(e) => e,
                    Err(_) => continue,
                };
                let path = entry.path();
                if !path.is_file() {
                    continue;
                }
                let ext = path
                    .extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("")
                    .to_lowercase();
                if !exts.contains(&ext.as_str()) {
                    continue;
                }

                let stem = path
                    .file_stem()
                    .and_then(|s| s.to_str())
                    .unwrap_or("unknown");
                let out_path = output.join(format!("{stem}.pptx"));
                let fname = path
                    .file_name()
                    .and_then(|s| s.to_str())
                    .unwrap_or("unknown");

                print!("{fname} → {} ... ", out_path.display());

                let source = match guess_source(&path, sheet.as_deref(), &encoding, &mut diag) {
                    Ok(s) => s,
                    Err(e) => {
                        failures.push((fname.to_string(), e));
                        eprintln!("失败");
                        continue;
                    }
                };

                match generate_one(&source, &out_path, force, template.as_deref(), &mut diag) {
                    Ok(_) => {
                        success.push(fname.to_string());
                    }
                    Err(e) => {
                        failures.push((fname.to_string(), e));
                    }
                }
            }

            println!();
            println!("===== 批量生成汇总 =====");
            println!("成功: {} 个", success.len());
            for s in &success {
                println!("  [OK] {s}");
            }
            println!("失败: {} 个", failures.len());
            for (f, e) in &failures {
                eprintln!("  [FAIL] {f}: {e}");
            }

            if failures.is_empty() {
                Ok(())
            } else {
                Err("部分文件生成失败".to_string())
            }
        }

        Command::TemplatePptx { output } => {
            println!("生成 PPTX 示例模板到 {} ...", output.display());
            generate_example_pptx(&output).map_err(|e| format!("生成失败: {e}"))?;
            println!("完成: {}", output.display());
            Ok(())
        }
        Command::ExportPng {
            input,
            output,
            template,
            sheet,
            encoding,
        } => {
            let output = output.unwrap_or_else(|| {
                let dir = input.parent().unwrap_or(Path::new("."));
                dir.join("png_output")
            });
            let source = guess_source(&input, sheet.as_deref(), &encoding, &mut diag)?;
            let entries = load(&source, &mut diag).map_err(|e| e.to_string())?;
            println!(
                "导出 {} 条词汇为 PNG 到 {} ...",
                entries.len(),
                output.display()
            );
            let pngs = if let Some(tpl) = &template {
                png_export::export_with_template(&entries, tpl, &output)
            } else {
                png_export::export_entries_to_png(&entries, &output)
            }
                .map_err(|e| format!("PNG 导出失败: {e}"))?;
            println!("完成: {} 张图片", pngs.len());
            for p in &pngs {
                println!("  {}", p.display());
            }
            Ok(())
        }

        Command::Template { output } => {
            println!("导出模板到 {} ...", output.display());
            export_template(&output).map_err(|e| format!("导出失败: {e}"))?;
            println!("完成: {}", output.display());
            Ok(())
        }

        Command::Diag {
            file,
            summary,
            blank_slides,
            font_trace,
            errors,
            slide,
            json,
        } => {
            let content = fs::read_to_string(&file)
                .map_err(|e| format!("无法读取诊断文件 {}: {e}", file.display()))?;

            let events: Vec<serde_json::Value> = content
                .lines()
                .filter(|l| !l.trim().is_empty())
                .filter_map(|l| serde_json::from_str::<serde_json::Value>(l).ok())
                .collect();

            if summary {
                let total = events.len();
                let info_count = events
                    .iter()
                    .filter(|e| e.get("level").and_then(|v| v.as_str()) == Some("INFO"))
                    .count();
                let warn_count = events
                    .iter()
                    .filter(|e| e.get("level").and_then(|v| v.as_str()) == Some("WARN"))
                    .count();
                let error_count = events
                    .iter()
                    .filter(|e| e.get("level").and_then(|v| v.as_str()) == Some("ERROR"))
                    .count();

                println!("诊断文件: {}", file.display());
                println!("── 汇总 ──");
                println!("  总事件: {total}");
                println!("  INFO:   {info_count}");
                println!("  WARN:   {warn_count}");
                println!("  ERROR:  {error_count}");

                if total > 0 {
                    // Count by target
                    let mut targets: std::collections::HashMap<&str, usize> =
                        std::collections::HashMap::new();
                    for e in &events {
                        if let Some(t) = e.get("target").and_then(|v| v.as_str()) {
                            *targets.entry(t).or_insert(0) += 1;
                        }
                    }
                    if !targets.is_empty() {
                        println!("── 按模块 ──");
                        let mut sorted: Vec<_> = targets.iter().collect();
                        sorted.sort_by(|a, b| b.1.cmp(a.1));
                        for (target, count) in sorted {
                            println!("  {target}: {count}");
                        }
                    }

                    // Show last 3 errors/warnings
                    let issues: Vec<_> = events
                        .iter()
                        .filter(|e| {
                            let lvl = e.get("level").and_then(|v| v.as_str()).unwrap_or("");
                            lvl == "ERROR" || lvl == "WARN"
                        })
                        .collect();
                    if !issues.is_empty() {
                        println!("── 最近 {} 条问题 ──", issues.len().min(5));
                        for e in issues.iter().rev().take(5) {
                            let lvl = e.get("level").and_then(|v| v.as_str()).unwrap_or("?");
                            let tgt = e.get("target").and_then(|v| v.as_str()).unwrap_or("?");
                            let msg = e.get("message").and_then(|v| v.as_str()).unwrap_or("?");
                            println!("  [{lvl}] {tgt}: {msg}");
                        }
                    }
                }
            } else {
                // Filter events based on flags
                let filtered: Vec<&serde_json::Value> = if let Some(n) = slide {
                    events
                        .iter()
                        .filter(|e| {
                            let msg = e.get("message").and_then(|v| v.as_str()).unwrap_or("");
                            let in_msg =
                                msg.contains(&format!("slide {} ", n))
                                    || msg.contains(&format!("slide_{}", n))
                                    || msg.contains(&format!("slide {}\n", n));
                            let in_fields = e
                                .get("fields")
                                .and_then(|v| v.get("slide"))
                                .and_then(|v| v.as_u64())
                                .map(|s| s == n as u64)
                                .unwrap_or(false);
                            in_msg || in_fields
                        })
                        .collect()
                } else if blank_slides {
                    events
                        .iter()
                        .filter(|e| {
                            let msg = e.get("message").and_then(|v| v.as_str()).unwrap_or("");
                            msg.contains("blank") || msg.contains("very low content density")
                        })
                        .collect()
                } else if font_trace {
                    events
                        .iter()
                        .filter(|e| {
                            e.get("target").and_then(|v| v.as_str()) == Some("font_probe")
                        })
                        .collect()
                } else if errors {
                    events
                        .iter()
                        .filter(|e| {
                            let lvl = e.get("level").and_then(|v| v.as_str()).unwrap_or("");
                            lvl == "ERROR" || lvl == "WARN"
                        })
                        .collect()
                } else {
                    events.iter().collect()
                };

                if json {
                    println!(
                        "{}",
                        serde_json::to_string_pretty(&filtered)
                            .map_err(|e| format!("JSON 序列化失败: {e}"))?
                    );
                } else {
                    for e in &filtered {
                        let lvl = e.get("level").and_then(|v| v.as_str()).unwrap_or("?");
                        let ts = e.get("timestamp").and_then(|v| v.as_str()).unwrap_or("?");
                        let tgt = e.get("target").and_then(|v| v.as_str()).unwrap_or("?");
                        let msg = e.get("message").and_then(|v| v.as_str()).unwrap_or("?");
                        println!("[{lvl}] {ts} {tgt}: {msg}");
                        if let Some(fields) = e.get("fields").and_then(|v| v.as_object()) {
                            for (k, v) in fields {
                                println!("  {k}: {v}");
                            }
                        }
                    }
                }

                if filtered.is_empty() && !json {
                    println!("(无匹配事件)");
                }
            }

            Ok(())
        }
    }
}
