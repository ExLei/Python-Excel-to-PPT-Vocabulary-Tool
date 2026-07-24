#![windows_subsystem = "windows"]

#[cfg(target_os = "windows")]
mod win_console {
    const ATTACH_PARENT_PROCESS: u32 = 0xFFFFFFFF;

    extern "system" {
        fn AttachConsole(dwProcessId: u32) -> i32;
    }

    /// CLI 模式下附加到父进程控制台，确保 println/eprintln 有输出。
    pub fn ensure_console() {
        unsafe {
            // AttachConsole 失败也无妨（比如从非控制台启动），静默忽略
            AttachConsole(ATTACH_PARENT_PROCESS);
        }
    }
}

use eframe::egui;
use std::path::Path;
use std::sync::Arc;

mod app;
mod panels;

use app::VocabPptApp;

/// 扫描系统字体目录，按优先级加载拉丁字体（IPA 覆盖）和 CJK 字体。
fn load_system_fonts(fonts: &mut egui::FontDefinitions) {
    // ── 平台字体目录 ──
    let font_dirs: Vec<std::path::PathBuf> = if cfg!(target_os = "windows") {
        vec![Path::new("C:\\Windows\\Fonts").to_path_buf()]
    } else if cfg!(target_os = "macos") {
        vec![
            Path::new("/System/Library/Fonts").to_path_buf(),
            Path::new("/Library/Fonts").to_path_buf(),
        ]
    } else {
        vec![
            Path::new("/usr/share/fonts").to_path_buf(),
            Path::new("/usr/local/share/fonts").to_path_buf(),
        ]
    };

    #[rustfmt::skip]
    let latin_fonts: &[&str] = if cfg!(target_os = "windows") {
        &["segoeui.ttf", "segoeuib.ttf", "calibri.ttf", "arial.ttf"]
    } else if cfg!(target_os = "macos") {
        &["helvetica.ttc", "helveticaneue.ttc"]
    } else {
        &["dejavusans.ttf", "liberationsans-regular.ttf", "notosans-regular.ttf"]
    };

    #[rustfmt::skip]
    let cjk_fonts: &[&str] = if cfg!(target_os = "windows") {
        &["msyh.ttc", "msyhbd.ttc", "simsun.ttc", "simsunb.ttf"]
    } else if cfg!(target_os = "macos") {
        &["pingfang.ttc", "heiti.ttc", "stheitisc-light.ttc"]
    } else {
        &["notosanscjk-regular.ttc", "notosanscjk-light.ttc",
          "notoserifcjk-regular.ttc", "wqy-microhei.ttc"]
    };

    #[rustfmt::skip]
    let mono_fonts: &[&str] = if cfg!(target_os = "windows") {
        &["consola.ttf", "cour.ttf"]
    } else if cfg!(target_os = "macos") {
        &["menlo.ttc", "monaco.ttf"]
    } else {
        &["dejavusansmono.ttf", "liberationmono-regular.ttf"]
    };

    for dir in &font_dirs {
        if !dir.exists() {
            continue;
        }
        let Ok(entries) = std::fs::read_dir(dir) else {
            continue;
        };

        let mut paths: Vec<std::path::PathBuf> = entries.flatten().map(|e| e.path()).collect();

        for p in &paths.clone() {
            if p.is_dir() {
                if let Ok(sub) = std::fs::read_dir(p) {
                    paths.extend(sub.flatten().map(|e| e.path()));
                }
            }
        }

        for path in &paths {
            let fname = path
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("")
                .to_lowercase();

            if !(fname.ends_with(".ttf") || fname.ends_with(".otf") || fname.ends_with(".ttc")) {
                continue;
            }

            let is_latin = latin_fonts.iter().any(|p| fname == *p);
            let is_cjk = cjk_fonts.iter().any(|p| fname == *p);
            let is_mono = mono_fonts.iter().any(|p| fname == *p);

            if !is_latin && !is_cjk && !is_mono {
                continue;
            }

            if let Ok(data) = std::fs::read(path) {
                let name = path
                    .file_stem()
                    .and_then(|n| n.to_str())
                    .unwrap_or("System")
                    .to_owned();

                fonts
                    .font_data
                    .insert(name.clone(), Arc::new(egui::FontData::from_owned(data)));

                fonts
                    .families
                    .entry(egui::FontFamily::Proportional)
                    .or_default()
                    .push(name.clone());

                if is_mono {
                    fonts
                        .families
                        .entry(egui::FontFamily::Monospace)
                        .or_default()
                        .push(name);
                }
            }
        }
    }
}

fn run_gui() -> Result<(), eframe::Error> {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default().with_inner_size([960.0, 700.0]),
        ..Default::default()
    };
    eframe::run_native(
        "英语助记卡片生成",
        options,
        Box::new(|cc: &eframe::CreationContext| {
            let mut fonts = egui::FontDefinitions::default();
            load_system_fonts(&mut fonts);
            cc.egui_ctx.set_fonts(fonts);
            Ok(Box::<VocabPptApp>::default())
        }),
    )
}

fn main() {
    // 有命令行参数 → CLI 模式；无参数 → GUI 模式
    let args: Vec<String> = std::env::args().collect();
    if args.len() > 1 {
        // CLI 模式 — 附加到父进程控制台以输出信息
        #[cfg(target_os = "windows")]
        win_console::ensure_console();

        if let Err(e) = cli::run() {
            eprintln!("{e}");
            std::process::exit(1);
        }
    } else {
        if let Err(e) = run_gui() {
            eprintln!("GUI 启动失败: {e}");
            std::process::exit(1);
        }
    }
}
