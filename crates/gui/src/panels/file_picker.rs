use egui::{ComboBox, TextEdit, Ui};

use crate::app::{InputMode, VocabPptApp};

pub fn show(ui: &mut Ui, app: &mut VocabPptApp) {
    ui.heading("输入");

    // File path row
    ui.horizontal(|ui| {
        ui.label("文件:");
        let w = ui.available_width() - 70.0;
        ui.add_sized(
            [w, 20.0],
            TextEdit::singleline(&mut app.input_path).hint_text("选择 Excel 或 CSV 文件"),
        );
        if ui.button("浏览...").clicked() {
            app.pick_input_file = true;
        }
    });

    // Format radio + sheet/encoding
    ui.horizontal(|ui| {
        ui.label("格式:");
        ui.radio_value(&mut app.input_mode, InputMode::Excel, "Excel");
        ui.radio_value(&mut app.input_mode, InputMode::Csv, "CSV");

        match app.input_mode {
            InputMode::Excel => {
                ui.label("工作表:");
                if app.sheets.is_empty() {
                    ui.label("(无)");
                } else {
                    ComboBox::from_id_salt("sheet")
                        .selected_text(&app.selected_sheet)
                        .width(120.0)
                        .show_ui(ui, |ui| {
                            for s in &app.sheets.clone() {
                                ui.selectable_value(&mut app.selected_sheet, s.clone(), s);
                            }
                        });
                }
            }
            InputMode::Csv => {
                ui.label("编码:");
                ui.add_sized([80.0, 20.0], TextEdit::singleline(&mut app.csv_encoding));
            }
        }
    });
    if !app.input_path.is_empty()
        && ui.button("刷新").clicked() {
            app.needs_load = true;
        }
}
