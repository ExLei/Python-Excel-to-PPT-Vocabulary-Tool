use std::fs;
use std::io::Write;
use std::path::Path;

use tempfile::tempdir;
use vocab_core::png_export;
use vocab_core::types::WordEntry;
use zip::{write::SimpleFileOptions, ZipWriter};

/// Helper: create a minimal PPTX with one slide containing {{单词}} placeholder
fn create_single_placeholder_pptx(path: &Path) {
    let file = fs::File::create(path).unwrap();
    let mut zip = ZipWriter::new(file);
    let opts = SimpleFileOptions::default();

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"#,
    ).unwrap();

    // _rels/.rels
    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#,
    ).unwrap();

    // ppt/presentation.xml
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>"#,
    ).unwrap();

    // ppt/_rels/presentation.xml.rels
    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#,
    ).unwrap();

    // Minimal slideMaster and slideLayout and theme
    zip.start_file("ppt/slideMasters/slideMaster1.xml", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>"#,
    ).unwrap();

    zip.start_file("ppt/slideLayouts/slideLayout1.xml", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
</p:sldLayout>"#,
    ).unwrap();

    zip.start_file("ppt/theme/theme1.xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme>
  </a:themeElements>
</a:theme>"#,
    ).unwrap();

    // slide1.xml with {{单词}} placeholder at known position
    zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
    let slide1 = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="word"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="10972800" cy="1371600"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:r><a:rPr sz="3600" b="1"><a:solidFill><a:srgbClr val="1A1A1A"/></a:solidFill></a:rPr><a:t>{{单词}}</a:t></a:r>
        </a:p>
      </p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>"#;
    zip.write_all(slide1.as_bytes()).unwrap();

    zip.finish().unwrap();
}

// ── Red: PPTX template parse tests ──

#[test]
fn parse_template_extracts_word_placeholder_position() {
    let dir = tempdir().unwrap();
    let tpl = dir.path().join("template.pptx");
    create_single_placeholder_pptx(&tpl);

    let layout = png_export::parse_template(&tpl).expect("should parse template successfully");

    let word_pl = layout
        .get("单词")
        .expect("should find {{单词}} placeholder");

    // Position: x=914400, y=914400 EMU → roughly x=144px, y=144px at 1920x1080
    assert!(
        word_pl.x > 0,
        "word x should be positive, got {}",
        word_pl.x
    );
    assert!(
        word_pl.y > 0,
        "word y should be positive, got {}",
        word_pl.y
    );
    assert!(
        word_pl.w > 0,
        "word width should be positive, got {}",
        word_pl.w
    );
    // Font size: sz=3600 hundredths-of-a-point → 36pt
    assert!(
        (word_pl.font_size_pt - 36.0).abs() < 0.5,
        "expected font size 36pt, got {}",
        word_pl.font_size_pt
    );
    assert!(word_pl.bold, "word should be bold");
}

#[test]
fn parse_template_rejects_invalid_pptx() {
    let dir = tempdir().unwrap();
    let path = dir.path().join("not_a_pptx.pptx");
    fs::write(&path, b"not a zip file").unwrap();

    let result = png_export::parse_template(&path);
    assert!(result.is_err(), "should error on non-ZIP file");
}

#[test]
fn parse_template_rejects_no_placeholder() {
    let dir = tempdir().unwrap();
    let tpl = dir.path().join("no_placeholder.pptx");
    let file = fs::File::create(&tpl).unwrap();
    let mut zip = ZipWriter::new(file);
    let opts = SimpleFileOptions::default();

    // minimal PPTX skeleton with slide that has NO {{...}} markers
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"#,
    ).unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#,
    ).unwrap();

    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>"#,
    ).unwrap();
    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#,
    ).unwrap();
    zip.start_file("ppt/slideMasters/slideMaster1.xml", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>"#,
    ).unwrap();
    zip.start_file("ppt/slideLayouts/slideLayout1.xml", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
</p:sldLayout>"#,
    ).unwrap();
    zip.start_file("ppt/theme/theme1.xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements><a:clrScheme name="Office"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme></a:themeElements>
</a:theme>"#,
    ).unwrap();

    // slide with text but NO placeholders
    zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:r><a:t>plain text, no placeholder</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>"#,
    ).unwrap();

    zip.finish().unwrap();

    let result = png_export::parse_template(&tpl);
    assert!(result.is_err(), "should error when no placeholders found");
}

// ── Red: SVG generation tests ──

fn make_layout_for_test(
) -> std::collections::HashMap<String, vocab_core::png_export::PlaceholderLayout> {
    use vocab_core::png_export::PlaceholderLayout;
    let mut m = std::collections::HashMap::new();
    m.insert(
        "单词".to_string(),
        PlaceholderLayout {
            name: "单词".into(),
            x: 144,
            y: 144,
            w: 1728,
            font_size_pt: 36.0,
            bold: true,
            color: image::Rgba([30, 30, 30, 255]),
            align_center: false,
        },
    );
    m.insert(
        "音标".to_string(),
        PlaceholderLayout {
            name: "音标".into(),
            x: 144,
            y: 220,
            w: 1728,
            font_size_pt: 24.0,
            bold: false,
            color: image::Rgba([100, 100, 100, 255]),
            align_center: false,
        },
    );
    m
}

fn make_full_layout_for_test(
) -> std::collections::HashMap<String, vocab_core::png_export::PlaceholderLayout> {
    use vocab_core::png_export::PlaceholderLayout;
    let mut m = make_layout_for_test();
    m.insert(
        "单词释义".to_string(),
        PlaceholderLayout {
            name: "单词释义".into(),
            x: 144,
            y: 280,
            w: 1728,
            font_size_pt: 26.0,
            bold: false,
            color: image::Rgba([30, 30, 30, 255]),
            align_center: false,
        },
    );
    m.insert(
        "例句".to_string(),
        PlaceholderLayout {
            name: "例句".into(),
            x: 144,
            y: 360,
            w: 1728,
            font_size_pt: 24.0,
            bold: false,
            color: image::Rgba([30, 30, 30, 255]),
            align_center: false,
        },
    );
    m
}

#[test]
fn svg_fields_not_in_layout_are_skipped() {
    let entry = WordEntry {
        word: "test".into(),
        phonetic: String::new(),
        morphology: "test-morph".into(),
        example: String::new(),
        example_definition: String::new(),
        definition: String::new(),
    };
    let layout = make_layout_for_test();
    let svg = png_export::render_slide_to_svg(&entry, &layout);
    // "单词" in layout → present
    assert!(svg.contains("test"), "SVG should contain word");
    // "词根词缀" NOT in layout → skipped
    assert!(
        !svg.contains("test-morph"),
        "morphology should be skipped when not in layout"
    );
}

#[test]
fn svg_generation_produces_valid_xml() {
    let entry = WordEntry {
        word: "apple".into(),
        phonetic: "/ˈæpl/".into(),
        morphology: String::new(),
        example: String::new(),
        example_definition: String::new(),
        definition: String::new(),
    };
    let layout = make_layout_for_test();

    let svg = png_export::render_slide_to_svg(&entry, &layout);

    // Must be valid SVG XML
    assert!(
        svg.starts_with("<svg"),
        "SVG should start with <svg> tag, got: {svg:.100}"
    );
    assert!(svg.contains("</svg>"), "SVG should close with </svg>");
    assert!(svg.contains("apple"), "SVG should contain the word text");
    assert!(svg.contains("/ˈæpl/"), "SVG should contain phonetic text");
    // Must have xmlns and viewBox
    assert!(svg.contains("xmlns="), "SVG should have xmlns attribute");
    assert!(
        svg.contains("viewBox="),
        "SVG should have viewBox attribute"
    );
}

#[test]
fn svg_positions_text_at_layout_coordinates() {
    let entry = WordEntry {
        word: "test".into(),
        phonetic: String::new(),
        morphology: String::new(),
        example: String::new(),
        example_definition: String::new(),
        definition: String::new(),
    };
    let layout = make_layout_for_test();

    let svg = png_export::render_slide_to_svg(&entry, &layout);

    // The "单词" placeholder has x=144, y=144, font_size_pt=36
    // SVG text element should have x/y attributes near these values
    assert!(
        svg.contains("x=\"144\""),
        "SVG text for 单词 should be at x=144"
    );
    // y in SVG usually includes the font ascent, so it should be near 144 + font_size
    assert!(
        svg.contains("font-size=\"36\""),
        "SVG text should use 36pt font for 单词"
    );
}

#[test]
fn svg_handles_chinese_characters() {
    let entry = WordEntry {
        word: "苹果".into(),
        phonetic: String::new(),
        morphology: String::new(),
        example: "我每天吃一个苹果。".into(),
        example_definition: String::new(),
        definition: "一种水果".into(),
    };
    let layout = make_full_layout_for_test();

    let svg = png_export::render_slide_to_svg(&entry, &layout);

    // All Chinese characters should be present in the SVG
    assert!(svg.contains("苹果"), "SVG should contain Chinese word");
    assert!(
        svg.contains("我每天吃一个苹果"),
        "SVG should contain Chinese example"
    );
    assert!(
        svg.contains("一种水果"),
        "SVG should contain Chinese definition"
    );
}

#[test]
fn svg_handles_ipa_characters() {
    let entry = WordEntry {
        word: "pronunciation".into(),
        phonetic: "/prəˌnʌnsiˈeɪʃən/".into(),
        morphology: String::new(),
        example: String::new(),
        example_definition: String::new(),
        definition: String::new(),
    };
    let layout = make_layout_for_test();

    let svg = png_export::render_slide_to_svg(&entry, &layout);

    // IPA characters (ə, ʌ, ʃ) should be in SVG
    assert!(svg.contains("ə"), "SVG should contain IPA schwa character");
    assert!(svg.contains("ʌ"), "SVG should contain IPA character");
    assert!(svg.contains("ʃ"), "SVG should contain IPA esh character");
}

// ── Red: PNG rendering tests ──

fn make_test_svg() -> String {
    r##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1920 1080" width="1920" height="1080">
  <rect width="1920" height="1080" fill="#ffffff"/>
  <text x="80" y="180" font-family="sans-serif" font-size="72" fill="#1e1e1e">测试文字</text>
</svg>"##
        .to_string()
}

#[test]
fn end_to_end_template_to_png() {
    let dir = tempdir().unwrap();
    let tpl = dir.path().join("template.pptx");
    create_single_placeholder_pptx(&tpl);

    let entries = vec![WordEntry {
        word: "hello".into(),
        phonetic: String::new(),
        morphology: String::new(),
        example: String::new(),
        example_definition: String::new(),
        definition: String::new(),
    }];

    let output_dir = dir.path().join("png_output");
    let pngs = png_export::export_with_template(&entries, &tpl, &output_dir)
        .expect("should export PNGs from template");

    assert_eq!(pngs.len(), 1, "should produce 1 PNG per entry");
    assert!(pngs[0].exists(), "PNG file should exist");

    let png_data = std::fs::read(&pngs[0]).unwrap();
    assert!(png_data.len() > 1000, "PNG should be non-trivial");
    let img = image::load_from_memory(&png_data).unwrap();
    assert_eq!(img.width(), 1920);
    assert_eq!(img.height(), 1080);
}

// ── Font probe tests ──

#[test]
fn probe_fonts_returns_non_empty() {
    let mut diag = vocab_core::diag::DiagStore::new();
    let cfg = vocab_core::png_export::pipeline::probe_fonts(&mut diag);
    assert!(
        !cfg.font_stack.is_empty(),
        "probe_fonts must return a font_stack"
    );
}

#[test]
fn probe_fonts_emits_diag_events() {
    let mut diag = vocab_core::diag::DiagStore::new();
    let _cfg = vocab_core::png_export::pipeline::probe_fonts(&mut diag);
    assert!(diag.event_count() > 0, "probe_fonts must emit diag events");
}

// probe_fonts now returns a single FontConfig directly (no first_available needed)

#[test]
fn probe_fonts_returns_valid_config() {
    let mut diag = vocab_core::diag::DiagStore::new();
    let cfg = vocab_core::png_export::pipeline::probe_fonts(&mut diag);
    assert!(!cfg.font_stack.is_empty());
    assert!(cfg.font_stack.contains("sans-serif"));
}

// ── count_text_nodes tests ──

#[test]
fn count_text_nodes_counts_text_elements() {
    let svg = r##"<svg xmlns="http://www.w3.org/2000/svg">
  <text x="10" y="20">hello</text>
  <text x="10" y="40">world</text>
  <rect width="100" height="100"/>
  <text x="30" y="60">third</text>
</svg>"##;
    let count = vocab_core::png_export::count_text_nodes(svg);
    assert_eq!(count, 3, "should count 3 text elements");
}

#[test]
fn count_text_nodes_returns_zero_for_rect_only() {
    let svg = r##"<svg xmlns="http://www.w3.org/2000/svg">
  <rect width="1920" height="1080" fill="#ffffff"/>
</svg>"##;
    let count = vocab_core::png_export::count_text_nodes(svg);
    assert_eq!(count, 0, "should return 0 when no text elements");
}

// ── count_non_white_pixels tests ──

fn make_all_white_png() -> Vec<u8> {
    use image::{ImageBuffer, Rgb};
    let img: ImageBuffer<Rgb<u8>, Vec<u8>> = ImageBuffer::from_pixel(100, 50, Rgb([255, 255, 255]));
    let mut buf = std::io::Cursor::new(Vec::new());
    img.write_to(&mut buf, image::ImageFormat::Png).unwrap();
    buf.into_inner()
}

#[test]
fn count_non_white_pixels_all_white_returns_zero() {
    let png = make_all_white_png();
    let count = vocab_core::png_export::count_non_white_pixels(&png).expect("should decode PNG");
    assert_eq!(count, 0, "all-white PNG should have 0 non-white pixels");
}

#[test]
fn count_non_white_pixels_counts_colored_pixels() {
    use image::{ImageBuffer, Rgb};
    let mut img: ImageBuffer<Rgb<u8>, Vec<u8>> =
        ImageBuffer::from_pixel(10, 10, Rgb([255, 255, 255]));
    // Set top-left corner to black
    img.put_pixel(0, 0, Rgb([0, 0, 0]));
    img.put_pixel(1, 0, Rgb([128, 128, 128]));
    // Set rightmost pixel to near-white (not exactly white)
    img.put_pixel(9, 9, Rgb([254, 255, 255]));
    let mut buf = std::io::Cursor::new(Vec::new());
    img.write_to(&mut buf, image::ImageFormat::Png).unwrap();
    let png = buf.into_inner();
    let count = vocab_core::png_export::count_non_white_pixels(&png).expect("should decode PNG");
    assert_eq!(
        count, 3,
        "should count 3 non-white pixels (including near-white)"
    );
}

// ── Updated render_svg_to_png tests with FontConfig ──

fn make_test_font_config() -> vocab_core::png_export::pipeline::FontConfig {
    let mut diag = vocab_core::diag::DiagStore::new();
    vocab_core::png_export::pipeline::probe_fonts(&mut diag)
}

#[test]
fn render_svg_to_png_produces_valid_png_bytes() {
    let svg = make_test_svg();
    let font_cfg = make_test_font_config();
    let mut diag = vocab_core::diag::DiagStore::new();
    let png_bytes = vocab_core::png_export::render_svg_to_png(&svg, &font_cfg, &mut diag)
        .expect("should render SVG to PNG");

    assert_eq!(
        &png_bytes[0..8],
        &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A],
        "output should start with PNG magic bytes"
    );
    assert!(
        png_bytes.len() > 1000,
        "PNG should be non-trivial size, got {}",
        png_bytes.len()
    );
}

#[test]
fn render_svg_to_png_has_correct_dimensions() {
    let svg = make_test_svg();
    let font_cfg = make_test_font_config();
    let mut diag = vocab_core::diag::DiagStore::new();
    let png_bytes = vocab_core::png_export::render_svg_to_png(&svg, &font_cfg, &mut diag).unwrap();

    let img = image::load_from_memory(&png_bytes).expect("should decode PNG");
    assert_eq!(img.width(), 1920, "PNG width should be 1920");
    assert_eq!(img.height(), 1080, "PNG height should be 1080");
}

// ── Blank slide detection tests ──

#[test]
fn blank_slide_detection_density_below_001_pct_with_text_emits_error() {
    // An SVG with text nodes but the text renders as white-on-white (density ≈ 0)
    let svg = r##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1920 1080" width="1920" height="1080">
  <rect width="1920" height="1080" fill="#ffffff"/>
  <text x="80" y="180" font-family="sans-serif" font-size="72" fill="#ffffff">invisible text</text>
</svg>"##;
    let font_cfg = make_test_font_config();
    let mut diag = vocab_core::diag::DiagStore::new();
    let _png = vocab_core::png_export::render_svg_to_png(&svg, &font_cfg, &mut diag)
        .expect("should render");

    let text_nodes = vocab_core::png_export::count_text_nodes(&svg);
    let non_white = vocab_core::png_export::count_non_white_pixels(&_png).unwrap();
    let total = 1920u64 * 1080u64;
    let density = non_white as f64 / total as f64;

    assert!(text_nodes > 0, "SVG has text nodes");
    assert!(
        density < 0.00001,
        "white text on white background yields near-zero density: {}",
        density
    );
    assert!(
        diag.errors() > 0,
        "diag should have at least one error for near-blank slide, got {}",
        diag.errors()
    );
}

#[test]
fn render_svg_to_png_tolerates_unknown_font() {
    // Unknown font names should not cause errors — resvg uses internal fallback
    let svg = make_test_svg();
    let font_cfg = vocab_core::png_export::pipeline::FontConfig {
        font_stack: "__nonexistent_font_xyzzy_12345__".to_string(),
        best_latin: None,
        best_cjk: None,
    };
    let mut diag = vocab_core::diag::DiagStore::new();
    let result = vocab_core::png_export::render_svg_to_png(&svg, &font_cfg, &mut diag);
    assert!(
        result.is_ok(),
        "should not fail for unknown font — resvg handles fallback internally"
    );
}
