use std::path::Path;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::SystemTime;

use eframe::egui;
use vocab_core::diag::DiagStore;
use vocab_core::types::{GenerateError, InputSource, WordEntry};
use vocab_core::{generator, png_export, reader, template, template_pptx};

use crate::panels::{data_preview, file_picker, output_config};

#[derive(Default, Clone, Copy, PartialEq)]
pub enum InputMode {
    #[default]
    Excel,
    Csv,
}

#[derive(Default)]
pub enum AppState {
    #[default]
    Idle,
    Loading {
        _path: PathBuf,
    },
    Preview {
        entries: Vec<WordEntry>,
    },
    Generating,
    Done {
        count: usize,
    },
    Error {
        message: String,
    },
}
impl AppState {
    pub(crate) fn is_generating(&self) -> bool {
        matches!(self, AppState::Generating)
    }
}

struct GenShared {
    current: usize,
    total: usize,
    result: Option<Result<(), GenerateError>>,
    diag: Option<DiagStore>,
}
pub struct VocabPptApp {
    pub state: AppState,
    pub input_path: String,
    pub input_mode: InputMode,
    pub sheets: Vec<String>,
    pub selected_sheet: String,
    pub csv_encoding: String,
    pub output_path: String,
    pub template_path: String,
    pub(crate) cancel_flag: Option<Arc<AtomicBool>>,
    pub(crate) cancel_png_flag: Option<Arc<AtomicBool>>,
    gen_shared: Option<Arc<Mutex<GenShared>>>,
    pub(crate) needs_load: bool,
    pending_load: bool,
    pub(crate) needs_generate: bool,
    pub(crate) pick_input_file: bool,
    pub(crate) pick_output_file: bool,
    pub(crate) pick_template_file: bool,
    pub(crate) needs_template_pptx: bool,
    pub(crate) needs_export_png: bool,
    export_png_handle: Option<Arc<Mutex<Option<Result<Vec<PathBuf>, String>>>>>,
    pub(crate) export_msg: Option<String>,
    // 自动刷新：跟踪文件变化
    last_input_path: String,
    pub diag: DiagStore,
    last_modified: Option<SystemTime>,
}

impl Default for VocabPptApp {
    fn default() -> Self {
        Self {
            state: AppState::Idle,
            input_path: String::new(),
            input_mode: InputMode::Excel,
            sheets: Vec::new(),
            selected_sheet: String::new(),
            csv_encoding: "UTF-8".into(),
            output_path: "output.pptx".into(),
            template_path: String::new(),
            cancel_flag: None,
            cancel_png_flag: None,
            gen_shared: None,
            needs_load: false,
            pending_load: false,
            needs_generate: false,
            pick_input_file: false,
            pick_output_file: false,
            pick_template_file: false,
            export_png_handle: None,
            export_msg: None,
            diag: DiagStore::new(),
            needs_template_pptx: false,
            needs_export_png: false,
            last_input_path: String::new(),
            last_modified: None,
        }
    }
}

/// 自动重命名：若文件已存在，追加 (1) (2) ... 直到找到可用名。
fn find_available_path(path: &Path) -> PathBuf {
    if !path.exists() {
        return path.to_path_buf();
    }
    let stem = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("output");
    let ext = path.extension().and_then(|s| s.to_str()).unwrap_or("pptx");
    let parent = path.parent().unwrap_or(Path::new("."));
    for n in 1..1000u32 {
        let candidate = parent.join(format!("{stem} ({n}).{ext}"));
        if !candidate.exists() {
            return candidate;
        }
    }
    // 极端情况：回退到带时间戳的文件名
    parent.join(format!("{stem}_{}.{ext}", std::process::id()))
}
impl VocabPptApp {
    fn detect_format(&mut self) {
        let p = std::path::Path::new(&self.input_path);
        self.input_mode = match p.extension().and_then(|e| e.to_str()) {
            Some("csv") => InputMode::Csv,
            _ => InputMode::Excel,
        };
        if self.input_mode == InputMode::Excel && !self.input_path.is_empty() {
            self.fetch_sheets();
        }
        // 自动推导输出路径（与输入同目录）
        if !self.input_path.is_empty() {
            let parent_dir = {
                let p = Path::new(&self.input_path);
                p.parent().unwrap_or(Path::new(".")).to_path_buf()
            };
            let stem = {
                let p = Path::new(&self.input_path);
                p.file_stem()
                    .and_then(|s| s.to_str())
                    .unwrap_or("output")
                    .to_string()
            };
            self.output_path = parent_dir
                .join(format!("{stem}.pptx"))
                .display()
                .to_string();
        }
        // Auto-trigger load after file selection
        self.needs_load = true;
    }
    fn fetch_sheets(&mut self) {
        self.diag.clear();
        self.sheets = reader::list_sheets(std::path::Path::new(&self.input_path), &mut self.diag)
            .unwrap_or_default();
        self.selected_sheet = self.sheets.first().cloned().unwrap_or_default();
    }
    fn try_load(&mut self, ctx: &egui::Context) {
        let source = match self.input_mode {
            InputMode::Excel => InputSource::Excel {
                path: PathBuf::from(&self.input_path),
                sheet: if self.selected_sheet.is_empty() {
                    "Sheet1".into()
                } else {
                    self.selected_sheet.clone()
                },
            },
            InputMode::Csv => InputSource::Csv {
                path: PathBuf::from(&self.input_path),
                encoding: self.csv_encoding.clone(),
            },
        };
        self.diag.clear();
        match reader::load(&source, &mut self.diag) {
            Ok(entries) => self.state = AppState::Preview { entries },
            Err(e) => {
                self.state = AppState::Error {
                    message: format!("读取失败: {e}"),
                }
            }
        }
        ctx.request_repaint();
    }
    fn start_generate(&mut self, ctx: &egui::Context) {
        let entries = match &self.state {
            AppState::Preview { entries } => entries.clone(),
            _ => return,
        };
        let total = entries.len();

        // 自动重命名：若输出文件已存在，追加 (1) (2) ...
        let output_path = PathBuf::from(&self.output_path);
        let output = find_available_path(&output_path);
        self.output_path = output.display().to_string();

        let use_template = !self.template_path.is_empty();
        let template = if use_template {
            Some(PathBuf::from(&self.template_path))
        } else {
            None
        };
        let shared = Arc::new(Mutex::new(GenShared {
            current: 0,
            total,
            result: None,
            diag: None,
        }));
        let cancel = Arc::new(AtomicBool::new(false));
        self.gen_shared = Some(shared.clone());
        self.cancel_flag = Some(cancel.clone());
        self.state = AppState::Generating;
        ctx.request_repaint();
        thread::spawn(move || {
            let mut diag = DiagStore::new();
            let r = if let Some(tpl) = template {
                generator::generate_from_template(
                    &entries,
                    &tpl,
                    &output,
                    |cur, _| {
                        if let Ok(mut s) = shared.lock() {
                            s.current = cur;
                        }
                        !cancel.load(Ordering::Relaxed)
                    },
                    &mut diag,
                )
            } else {
                generator::generate(
                    &entries,
                    &output,
                    |cur, _| {
                        if let Ok(mut s) = shared.lock() {
                            s.current = cur;
                        }
                        !cancel.load(Ordering::Relaxed)
                    },
                    &mut diag,
                )
            };
            if let Ok(mut s) = shared.lock() {
                s.current = total;
                s.result = Some(r);
                s.diag = Some(diag);
            }
        });
    }
    fn poll_generation(&mut self, ctx: &egui::Context) {
        let shared = match &self.gen_shared {
            Some(s) => s.clone(),
            None => return,
        };
        // Take result + diag in a scoped block to drop the MutexGuard before modifying self
        let done: Option<(Result<(), GenerateError>, usize, Option<DiagStore>)> = {
            shared
                .lock()
                .ok()
                .and_then(|mut s| s.result.take().map(|r| (r, s.total, s.diag.take())))
        };
        if let Some((result, total, gen_diag)) = done {
            self.gen_shared = None;
            self.cancel_flag = None;
            if let Some(gd) = gen_diag {
                self.diag = gd;
            }
            self.state = match result {
                Ok(()) => AppState::Done { count: total },
                Err(e) => AppState::Error {
                    message: e.to_string(),
                },
            };
            ctx.request_repaint();
        }
    }
    pub(crate) fn gen_progress(&self) -> Option<(usize, usize)> {
        self.gen_shared
            .as_ref()?
            .lock()
            .ok()
            .map(|s| (s.current, s.total))
    }

    pub(crate) fn is_exporting_png(&self) -> bool {
        self.export_png_handle.is_some()
    }

    fn open_example_pptx(&mut self) {
        let tmp = std::env::temp_dir().join("示例模板.pptx");
        match template_pptx::generate_example_pptx(&tmp) {
            Ok(()) => {
                if let Err(e) = open::that(&tmp) {
                    self.export_msg = Some(format!("无法打开文件: {e}"));
                }
            }
            Err(e) => {
                self.export_msg = Some(format!("生成示例PPT失败: {e}"));
            }
        }
    }
    fn open_template(&mut self) {
        let tmp = std::env::temp_dir().join("vocab_template.xlsx");
        match template::export_template(&tmp) {
            Ok(()) => {
                if let Err(e) = open::that(&tmp) {
                    self.export_msg = Some(format!("无法打开文件: {e}"));
                }
            }
            Err(e) => {
                self.export_msg = Some(format!("导出模板失败: {e}"));
            }
        }
    }
    fn start_export_png(&mut self, ctx: &egui::Context) {
        let entries = match &self.state {
            AppState::Preview { entries } => entries.clone(),
            _ => return,
        };
        let output_dir = Path::new(&self.output_path)
            .parent()
            .unwrap_or(Path::new("."))
            .join("png_output");
        let template_path = if self.template_path.is_empty() {
            None
        } else {
            Some(PathBuf::from(&self.template_path))
        };
        let handle = Arc::new(Mutex::new(None));
        let cancel = Arc::new(AtomicBool::new(false));
        self.cancel_png_flag = Some(cancel.clone());
        self.export_png_handle = Some(handle.clone());
        ctx.request_repaint();
        thread::spawn(move || {
            if let Err(e) = std::fs::create_dir_all(&output_dir) {
                *handle.lock().unwrap() = Some(Err(format!("创建输出目录失败: {e}")));
                return;
            }
            let res: Result<Vec<PathBuf>, String> = if let Some(tpl) = &template_path {
                let pptx_path = output_dir.join("_generated.pptx");
                if pptx_path.exists() {
                    let _ = std::fs::remove_file(&pptx_path);
                }
                match vocab_core::generator::generate_from_template(
                    &entries,
                    tpl,
                    &pptx_path,
                    |_, _| true,
                    &mut vocab_core::diag::DiagStore::new(),
                ) {
                    Err(e) => Err(format!("生成 PPTX 失败: {e}")),
                    Ok(()) => {
                        let mut rd = vocab_core::diag::DiagStore::new();
                        match vocab_core::renderer::render_pptx(&pptx_path, &output_dir, &mut rd) {
                            Ok(pngs) => {
                                let _ = std::fs::remove_file(&pptx_path);
                                Ok(pngs)
                            }
                            Err(vocab_core::renderer::RenderError::NoRenderer) => {
                                let _ = std::fs::remove_file(&pptx_path);
                                Err(vocab_core::renderer::RENDERER_HELP.to_string())
                            }
                            Err(e) => {
                                let _ = std::fs::remove_file(&pptx_path);
                                Err(format!("外部渲染器失败: {e}"))
                            }
                        }
                    }
                }
            } else {
                png_export::export_entries_to_png_with_cancel(&entries, &output_dir, Some(&cancel))
                    .map_err(|e| e.to_string())
            };
            *handle.lock().unwrap() = Some(res);
        });
    }
    fn poll_export_png(&mut self, ctx: &egui::Context) {
        let handle = match self.export_png_handle.as_ref() {
            Some(h) => Arc::clone(h),
            None => return,
        };
        let mut guard = handle.lock().unwrap();
        if let Some(res) = guard.take() {
            drop(guard);
            self.export_png_handle = None;
            self.cancel_png_flag = None;
            match &res {
                Ok(paths) => {
                    let dir = paths
                        .first()
                        .and_then(|p| p.parent())
                        .map(|d| d.display().to_string())
                        .unwrap_or_default();
                    self.export_msg = Some(format!("PNG 导出完成: {} 张 → {dir}", paths.len()));
                }
                Err(e) => {
                    self.export_msg = Some(format!("PNG 导出失败: {e}"));
                }
            }
            ctx.request_repaint();
        }
    }
}

impl eframe::App for VocabPptApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        if self.pick_input_file {
            self.pick_input_file = false;
            if let Some(p) = rfd::FileDialog::new()
                .add_filter("Excel/CSV", &["xlsx", "xls", "csv"])
                .pick_file()
            {
                self.input_path = p.display().to_string();
                self.detect_format();
            }
        }
        if self.pick_output_file {
            self.pick_output_file = false;
            if let Some(p) = rfd::FileDialog::new()
                .add_filter("PPTX", &["pptx"])
                .save_file()
            {
                self.output_path = p.display().to_string();
            }
        }
        if self.pick_template_file {
            self.pick_template_file = false;
            if let Some(p) = rfd::FileDialog::new()
                .add_filter("PPTX 模板", &["pptx"])
                .pick_file()
            {
                self.template_path = p.display().to_string();
            }
        }
        if self.needs_template_pptx {
            self.needs_template_pptx = false;
            self.open_example_pptx();
        }
        if self.needs_load {
            self.needs_load = false;
            self.state = AppState::Loading {
                _path: PathBuf::from(&self.input_path),
            };
            self.pending_load = true;
            ctx.request_repaint();
        }
        if self.needs_generate {
            self.needs_generate = false;
            self.start_generate(ctx);
        }
        if self.needs_export_png {
            self.needs_export_png = false;
            self.start_export_png(ctx);
        }
        self.poll_export_png(ctx);
        self.poll_generation(ctx);

        // ── 自动刷新：监测文件变化 ──
        if !self.state.is_generating() && !self.input_path.is_empty() {
            // 1) 输入路径变化 → 自动重新加载
            if self.input_path != self.last_input_path {
                self.last_input_path = self.input_path.clone();
                let path = Path::new(&self.input_path);
                self.last_modified = path.metadata().ok().and_then(|m| m.modified().ok());
                if path.exists() {
                    self.needs_load = true;
                }
            }
            // 2) 文件被外部修改 → 自动刷新预览
            else if let AppState::Preview { .. } = &self.state {
                if let Ok(meta) = std::fs::metadata(&self.input_path) {
                    if let Ok(mtime) = meta.modified() {
                        if self.last_modified != Some(mtime) {
                            self.last_modified = Some(mtime);
                            self.needs_load = true;
                        }
                    }
                }
            }
        }

        // ── Button bar (fixed at bottom) ──
        egui::TopBottomPanel::bottom("bottom_bar").show(ctx, |ui| {
            ui.horizontal(|ui| {
                if ui.button("打开Excel模板").clicked() {
                    self.open_template();
                }
                if ui.button("示例PPT模板").clicked() {
                    self.needs_template_pptx = true;
                }
                let exporting = self.export_png_handle.is_some();
                let has_data = matches!(&self.state, AppState::Preview { entries, .. } if !entries.is_empty());
                ui.add_enabled_ui(has_data && !exporting, |ui| {
                    if ui.button("导出PNG").clicked() {
                        self.needs_export_png = true;
                    }
                });
                if exporting {
                    ui.label("正在导出PNG...");
                    if ui.button("取消").clicked() {
                        if let Some(ref f) = self.cancel_png_flag {
                            f.store(true, Ordering::Relaxed);
                        }
                    }
                }
                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                    if self.state.is_generating() {
                        if ui.button("取消").clicked() {
                            if let Some(ref f) = self.cancel_flag {
                                f.store(true, Ordering::Relaxed);
                            }
                        }
                    } else {
                        let en = matches!(&self.state, AppState::Preview { entries, .. } if !entries.is_empty());
                        ui.add_enabled_ui(en, |ui| {
                            if ui.button("生成PPT").clicked() {
                                self.needs_generate = true;
                            }
                        });
                    }
                });
            });
        });

        // ── 导出消息通知 ──
        if let Some(msg) = &self.export_msg.clone() {
            egui::TopBottomPanel::top("export_notify").show(ctx, |ui| {
                ui.horizontal(|ui| {
                    if msg.contains("失败") {
                        ui.colored_label(egui::Color32::RED, msg);
                    } else {
                        ui.colored_label(egui::Color32::GREEN, msg);
                    }
                    if ui.button("✕").clicked() {
                        self.export_msg = None;
                    }
                });
            });
        }

        // ── Main content ──
        egui::CentralPanel::default().show(ctx, |ui| {
            file_picker::show(ui, self);
            ui.separator();
            output_config::show(ui, self);
            ui.separator();
            data_preview::show(ui, self);
        });

        if self.pending_load {
            self.pending_load = false;
            self.try_load(ctx);
        }
    }
}
