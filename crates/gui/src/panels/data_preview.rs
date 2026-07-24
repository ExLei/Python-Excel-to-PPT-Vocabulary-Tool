use std::sync::atomic::Ordering;

use egui::{Color32, ProgressBar, RichText, Ui};
use egui_extras::{Column, TableBuilder};

use crate::app::{AppState, VocabPptApp};

const PREVIEW_LIMIT: usize = 100;
const COLUMNS: [&str; 6] = ["单词", "音标", "词根", "例句", "释义", "单词释义"];

pub fn show(ui: &mut Ui, app: &mut VocabPptApp) {
    ui.heading("状态");

    match &app.state {
        AppState::Idle => {
            ui.label(RichText::new("请选择文件").color(Color32::GRAY));
        }
        AppState::Loading { .. } => {
            ui.horizontal(|ui| {
                ui.spinner();
                ui.label("加载中...");
            });
        }
        AppState::Preview { entries } => {
            let count = entries.len();
            let display_count = count.min(PREVIEW_LIMIT);
            ui.label(format!("共 {count} 条记录", count = count));
            if app.is_exporting_png() {
                ui.horizontal(|ui| {
                    ui.label(RichText::new("正在导出PNG...").color(Color32::GRAY));
                    if ui.button("取消").clicked() {
                        if let Some(ref f) = app.cancel_png_flag {
                            f.store(true, Ordering::Relaxed);
                        }
                    }
                });
                ui.separator();
            }
            if count > PREVIEW_LIMIT {
                ui.label(
                    RichText::new(format!("(仅显示前 {PREVIEW_LIMIT} 行)")).color(Color32::GRAY),
                );
            }

            let max_h = ui.available_height().max(100.0);
            egui::ScrollArea::vertical()
                .max_height(max_h)
                .show(ui, |ui| {
                    let available_width = ui.available_width();
                    let col_width = available_width / 6.0 - 4.0;

                    TableBuilder::new(ui)
                        .striped(true)
                        .cell_layout(egui::Layout::left_to_right(egui::Align::Center))
                        .column(Column::initial(col_width).resizable(true))
                        .column(Column::initial(col_width).resizable(true))
                        .column(Column::initial(col_width).resizable(true))
                        .column(Column::initial(col_width).resizable(true))
                        .column(Column::initial(col_width).resizable(true))
                        .column(Column::initial(col_width).resizable(true))
                        .header(20.0, |mut header| {
                            for col_name in &COLUMNS {
                                header.col(|ui| {
                                    ui.strong(*col_name);
                                });
                            }
                        })
                        .body(|body| {
                            body.rows(20.0, display_count, |mut row| {
                                let idx = row.index();
                                let entry = &entries[idx];
                                row.col(|ui| {
                                    ui.label(truncate(&entry.word, 20));
                                });
                                row.col(|ui| {
                                    ui.label(truncate(&entry.phonetic, 20));
                                });
                                row.col(|ui| {
                                    ui.label(truncate(&entry.morphology, 20));
                                });
                                row.col(|ui| {
                                    ui.label(truncate(&entry.example, 30));
                                });
                                row.col(|ui| {
                                    ui.label(truncate(&entry.example_definition, 20));
                                });
                                row.col(|ui| {
                                    ui.label(truncate(&entry.definition, 30));
                                });
                            });
                        });
                });
        }
        AppState::Generating => {
            let (current, total) = app.gen_progress().unwrap_or((0, 0));
            let pct = if total > 0 {
                current as f32 / total as f32
            } else {
                0.0
            };
            ui.horizontal(|ui| {
                ui.add(ProgressBar::new(pct).desired_width(300.0));
                ui.label(format!("{current}/{total}"));
                if ui.button("取消").clicked() {
                    if let Some(ref flag) = app.cancel_flag {
                        flag.store(true, Ordering::Relaxed);
                    }
                }
            });
        }
        AppState::Done { count } => {
            ui.label(
                RichText::new(format!("成功生成 {count} 张")).color(Color32::from_rgb(0, 160, 0)),
            );
        }
        AppState::Error { message } => {
            ui.label(RichText::new(message).color(Color32::RED));
        }
    }

    // ── 诊断摘要 ──
    let total = app.diag.event_count();
    if total > 0 {
        ui.separator();
        ui.heading("诊断");
        let warns = app.diag.warnings();
        let errs = app.diag.errors();
        let infos = total - warns - errs;
        ui.label(format!(
            "事件: {} 条 (信息 {}, 警告 {}, 错误 {})",
            total, infos, warns, errs
        ));
    }
}

fn truncate(s: &str, max: usize) -> String {
    let cleaned: String = s.chars().take(max).collect();
    if s.chars().count() > max {
        format!("{}…", cleaned)
    } else {
        cleaned
    }
}
