use std::collections::HashMap;

use image::Rgba;
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::diag::DiagStore;

/// Font configuration: identifies an available system font.
/// CSS font-family stack string with fallback chains resolved.
#[derive(Debug, Clone)]
pub struct FontConfig {
    pub font_stack: String,
    pub best_latin: Option<String>,
    pub best_cjk: Option<String>,
}

/// Probe system fonts and build a cross-platform font-family stack.
///
/// Checks prioritized candidate lists for Latin, CJK, and Emoji fonts.
/// Emits diagnostic events for every candidate checked and the final selection.
pub fn probe_fonts(diag: &mut DiagStore) -> FontConfig {
    let mut db = usvg::fontdb::Database::new();
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        db.load_system_fonts();
    }));
    if result.is_err() {
        diag.error(
            "font_probe",
            "fontdb.load_system_fonts() panicked — possible corrupt system font",
            None,
        );
        return FontConfig {
            font_stack: "sans-serif".into(),
            best_latin: None,
            best_cjk: None,
        };
    }

    let face_count = db.faces().count();
    diag.info(
        "font_probe",
        &format!("scanned {face_count} font faces"),
        None,
    );

    let latin_chain = ["Segoe UI", "Helvetica", "Arial", "DejaVu Sans"];
    let cjk_chain = [
        "Microsoft YaHei",
        "PingFang SC",
        "Noto Sans CJK SC",
        "WenQuanYi Micro Hei",
    ];
    let emoji_chain = ["Segoe UI Emoji", "Apple Color Emoji", "Noto Color Emoji"];

    let best_latin = find_first(&db, &latin_chain, "latin", diag);
    let best_cjk = find_first(&db, &cjk_chain, "cjk", diag);
    let best_emoji = find_first(&db, &emoji_chain, "emoji", diag);

    let mut parts: Vec<&str> = Vec::new();
    if let Some(ref cjk) = best_cjk {
        parts.push(cjk);
    }
    if let Some(ref latin) = best_latin {
        parts.push(latin);
    }
    if let Some(ref emoji) = best_emoji {
        parts.push(emoji);
    }
    parts.push("sans-serif");

    let font_stack = parts.join(", ");
    diag.info("font_probe", &format!("font_stack: {font_stack}"), None);

    if let Some(ref latin) = best_latin {
        db.set_sans_serif_family(latin);
        db.set_serif_family(latin);
    }

    FontConfig {
        font_stack,
        best_latin,
        best_cjk,
    }
}

fn find_first(
    db: &usvg::fontdb::Database,
    candidates: &[&str],
    chain: &str,
    diag: &mut DiagStore,
) -> Option<String> {
    for (i, name) in candidates.iter().enumerate() {
        let count = db
            .faces()
            .filter(|f| f.families.iter().any(|(fam, _)| fam == name))
            .count();
        if count > 0 {
            diag.info(
                "font_probe",
                &format!("{chain}: selected \"{name}\" ({count} faces, order={i})"),
                None,
            );
            return Some(name.to_string());
        }
        diag.info(
            "font_probe",
            &format!("{chain}: \"{name}\" not found (order={i})"),
            None,
        );
    }
    diag.warn(
        "font_probe",
        &format!(
            "{chain}: all {count} candidates missing",
            count = candidates.len()
        ),
        None,
    );
    None
}

// ── PPTX slide XML parsing ──

#[derive(Debug)]
struct SpState {
    x: i64,
    y: i64,
    cx: i64,
    font_sz: f32,
    bold: bool,
    color: Rgba<u8>,
    align_ctr: bool,
    text: String,
}

impl Default for SpState {
    fn default() -> Self {
        SpState {
            x: 0,
            y: 0,
            cx: 0,
            font_sz: 24.0,
            bold: false,
            color: Rgba([30, 30, 30, 255]),
            align_ctr: false,
            text: String::new(),
        }
    }
}

fn collect_attrs(e: &quick_xml::events::BytesStart, st: &mut SpState, diag: &mut DiagStore) {
    for attr in e.attributes().flatten() {
        let kbytes = attr.key.as_ref().to_vec();
        let k = match std::str::from_utf8(&kbytes) {
            Ok(s) => s,
            Err(_) => {
                diag.warn("parse_slide_xml", "invalid UTF-8 in attribute key", None);
                continue;
            }
        };
        let v = match std::str::from_utf8(&attr.value) {
            Ok(s) => s,
            Err(_) => {
                diag.warn("parse_slide_xml", "invalid UTF-8 in attribute value", None);
                continue;
            }
        };
        match k {
            "x" => {
                st.x = match v.parse() {
                    Ok(val) => val,
                    Err(_) => {
                        diag.warn("parse_slide_xml", "failed to parse x attribute", None);
                        0
                    }
                }
            }
            "y" => {
                st.y = match v.parse() {
                    Ok(val) => val,
                    Err(_) => {
                        diag.warn("parse_slide_xml", "failed to parse y attribute", None);
                        0
                    }
                }
            }
            "cx" => {
                st.cx = match v.parse() {
                    Ok(val) => val,
                    Err(_) => {
                        diag.warn("parse_slide_xml", "failed to parse cx attribute", None);
                        0
                    }
                }
            }
            "sz" => {
                st.font_sz = match v.parse::<f32>() {
                    Ok(val) => val / 100.0,
                    Err(_) => {
                        diag.warn("parse_slide_xml", "failed to parse sz attribute", None);
                        2400.0 / 100.0
                    }
                }
            }
            "b" => st.bold = v == "1",
            "val" => {
                if let Ok(rgb) = u32::from_str_radix(v, 16) {
                    st.color = Rgba([
                        ((rgb >> 16) & 0xFF) as u8,
                        ((rgb >> 8) & 0xFF) as u8,
                        (rgb & 0xFF) as u8,
                        255,
                    ]);
                }
            }
            "algn" => st.align_ctr = v == "ctr",
            _ => {}
        }
    }
}

fn extract_placeholder(text: &str) -> Option<String> {
    if let (Some(s), Some(e)) = (text.find("{{"), text.find("}}")) {
        if s < e {
            return Some(text[s + 2..e].to_string());
        }
    }
    None
}

pub(super) fn parse_slide_xml(
    xml: &str,
    diag: &mut DiagStore,
) -> Result<HashMap<String, super::PlaceholderLayout>, super::PngExportError> {
    let mut reader = Reader::from_str(xml);
    let mut layouts = HashMap::new();
    let mut buf = Vec::new();
    let mut in_sp = false;
    let mut sp_depth = 0u32;
    let mut st = SpState::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let tag = std::str::from_utf8(name.as_ref()).unwrap_or("");
                if tag == "p:sp" {
                    in_sp = true;
                    sp_depth += 1;
                    st = SpState::default();
                }
                collect_attrs(&e, &mut st, diag);
            }
            Ok(Event::Empty(e)) => {
                collect_attrs(&e, &mut st, diag);
            }
            Ok(Event::Text(e)) => {
                if in_sp {
                    st.text.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let tag = std::str::from_utf8(name.as_ref()).unwrap_or("");
                if tag == "p:sp" {
                    sp_depth -= 1;
                    if sp_depth == 0 {
                        in_sp = false;
                        if !st.text.is_empty() {
                            if let Some(ph_name) = extract_placeholder(&st.text) {
                                layouts.insert(
                                    ph_name.clone(),
                                    super::PlaceholderLayout {
                                        name: ph_name,
                                        x: super::emu_to_px_x(st.x),
                                        y: super::emu_to_px_y(st.y),
                                        w: super::emu_to_px_w(st.cx),
                                        font_size_pt: st.font_sz.max(12.0),
                                        bold: st.bold,
                                        color: st.color,
                                        align_center: st.align_ctr,
                                    },
                                );
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    if layouts.is_empty() {
        Err(super::PngExportError::TemplateParse(
            "no placeholders found".into(),
        ))
    } else {
        Ok(layouts)
    }
}
