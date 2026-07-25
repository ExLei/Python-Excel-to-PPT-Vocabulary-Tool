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
