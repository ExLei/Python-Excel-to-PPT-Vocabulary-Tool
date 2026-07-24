use egui::{TextEdit, Ui};

use crate::app::VocabPptApp;

/// Returns true if output path is non-empty (enables generate button).
pub fn show(ui: &mut Ui, app: &mut VocabPptApp) -> bool {
    ui.heading("输出");

    ui.horizontal(|ui| {
        ui.label("输出路径:");
        let w = ui.available_width() - 70.0;
        ui.add_sized(
            [w, 20.0],
            TextEdit::singleline(&mut app.output_path).hint_text("output.pptx"),
        );
        if ui.button("浏览...").clicked() {
            app.pick_output_file = true;
        }
    });

    ui.horizontal(|ui| {
        ui.label("PPTX模板:");
        let w = ui.available_width() - 70.0;
        ui.add_sized(
            [w, 20.0],
            TextEdit::singleline(&mut app.template_path)
                .hint_text("(可选) 含 {{占位符}} 的自定义模板"),
        );
        if ui.button("浏览...").clicked() {
            app.pick_template_file = true;
        }
    });

    !app.output_path.is_empty()
}
