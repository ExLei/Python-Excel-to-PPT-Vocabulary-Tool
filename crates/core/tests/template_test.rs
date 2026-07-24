use std::fs;
use std::path::Path;

use calamine::{Reader, Xlsx};
use tempfile::tempdir;
use vocab_core::template;
use vocab_core::types::TemplateError;

#[test]
fn export_template_creates_file() {
    let dir = tempdir().unwrap();
    let path = dir.path().join("template.xlsx");
    let result = template::export_template(&path);
    assert!(result.is_ok(), "export should succeed: {result:?}");
    assert!(path.exists(), "file should exist");
    let metadata = fs::metadata(&path).unwrap();
    assert!(metadata.len() > 0, "file should not be empty");
}

#[test]
fn exported_template_is_valid_xlsx() {
    let dir = tempdir().unwrap();
    let path = dir.path().join("template.xlsx");
    template::export_template(&path).unwrap();

    let workbook: Xlsx<_> = calamine::open_workbook(&path).unwrap();
    let sheet_names = workbook.sheet_names().to_owned();
    assert!(!sheet_names.is_empty(), "should have at least one sheet");
}

#[test]
fn export_to_invalid_path_errors() {
    let result = template::export_template(Path::new("/nonexistent/deep/template.xlsx"));
    assert!(matches!(result, Err(TemplateError::IoError(_))));
}
