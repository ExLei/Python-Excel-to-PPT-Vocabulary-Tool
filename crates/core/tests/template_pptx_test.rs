use std::fs;
use std::io::Read;

use tempfile::tempdir;
use vocab_core::template_pptx;
use zip::ZipArchive;

#[test]
fn generate_example_creates_valid_pptx() {
    let dir = tempdir().unwrap();
    let output = dir.path().join("example.pptx");

    template_pptx::generate_example_pptx(&output).unwrap();
    assert!(output.exists(), "output file should exist");

    let data = fs::read(&output).unwrap();
    assert!(!data.is_empty());
    assert_eq!(&data[0..2], b"PK", "output should be a valid ZIP archive");
}

#[test]
fn generate_example_contains_all_six_placeholders() {
    let dir = tempdir().unwrap();
    let output = dir.path().join("example.pptx");

    template_pptx::generate_example_pptx(&output).unwrap();

    // Extract slide1.xml and verify placeholders
    let data = fs::read(&output).unwrap();
    let reader = std::io::Cursor::new(&data);
    let mut archive = ZipArchive::new(reader).unwrap();
    let slide = archive.by_name("ppt/slides/slide1.xml").unwrap();
    let slide_text =
        String::from_utf8(slide.bytes().collect::<Result<Vec<u8>, _>>().unwrap()).unwrap();

    // All 6 standard placeholders should be present
    assert!(slide_text.contains("{{单词}}"), "missing {{单词}}");
    assert!(slide_text.contains("{{音标}}"), "missing {{音标}}");
    assert!(slide_text.contains("{{词根词缀}}"), "missing {{词根词缀}}");
    assert!(slide_text.contains("{{例句}}"), "missing {{例句}}");
    assert!(slide_text.contains("{{例句释义}}"), "missing {{例句释义}}");
    assert!(slide_text.contains("{{单词释义}}"), "missing {{单词释义}}");
}

#[test]
fn generate_example_overwrites_existing() {
    let dir = tempdir().unwrap();
    let output = dir.path().join("example.pptx");

    // First generation
    template_pptx::generate_example_pptx(&output).unwrap();
    // Second generation should succeed (overwrite)
    let result = template_pptx::generate_example_pptx(&output);
    assert!(result.is_ok(), "should overwrite existing file");
}
