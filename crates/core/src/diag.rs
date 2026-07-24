use serde::Serialize;
use std::time::{SystemTime, UNIX_EPOCH};
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "UPPERCASE")]
pub enum DiagLevel {
    Info,
    Warn,
    Error,
}

#[derive(Debug, Clone)]
pub struct DiagEvent {
    pub timestamp: String,
    pub level: DiagLevel,
    pub target: String,
    pub message: String,
    pub fields_json: Option<String>,
}

pub struct DiagStore {
    events: Vec<DiagEvent>,
    warning_count: usize,
    error_count: usize,
}

#[derive(Serialize)]
struct NdjsonLine<'a> {
    timestamp: &'a str,
    level: DiagLevel,
    target: &'a str,
    message: &'a str,
    #[serde(skip_serializing_if = "Option::is_none")]
    fields: Option<serde_json::Value>,
}

impl DiagStore {
    pub fn new() -> Self {
        Self {
            events: Vec::new(),
            warning_count: 0,
            error_count: 0,
        }
    }

    pub fn clear(&mut self) {
        self.events.clear();
        self.warning_count = 0;
        self.error_count = 0;
    }
}

impl Default for DiagStore {
    fn default() -> Self {
        Self::new()
    }
}

impl DiagStore {

    fn now_iso() -> String {
        let dur = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default();
        let secs = dur.as_secs();
        let h = (secs / 3600) % 24;
        let m = (secs % 3600) / 60;
        let s = secs % 60;
        format!("{h:02}:{m:02}:{s:02}")
    }

    fn push(&mut self, level: DiagLevel, target: &str, message: &str, fields_json: Option<&str>) {
        if matches!(level, DiagLevel::Warn) {
            self.warning_count += 1;
        }
        if matches!(level, DiagLevel::Error) {
            self.error_count += 1;
        }
        self.events.push(DiagEvent {
            timestamp: Self::now_iso(),
            level,
            target: target.to_string(),
            message: message.to_string(),
            fields_json: fields_json.map(|s| s.to_string()),
        });
    }

    pub fn info(&mut self, target: &str, message: &str, fields_json: Option<&str>) {
        self.push(DiagLevel::Info, target, message, fields_json);
    }

    pub fn warn(&mut self, target: &str, message: &str, fields_json: Option<&str>) {
        self.push(DiagLevel::Warn, target, message, fields_json);
    }

    pub fn error(&mut self, target: &str, message: &str, fields_json: Option<&str>) {
        self.push(DiagLevel::Error, target, message, fields_json);
    }

    pub fn event_count(&self) -> usize {
        self.events.len()
    }
    pub fn warnings(&self) -> usize {
        self.warning_count
    }
    pub fn errors(&self) -> usize {
        self.error_count
    }

    pub fn to_ndjson(&self) -> String {
        let mut out = String::new();
        for e in &self.events {
            let fields = e
                .fields_json
                .as_deref()
                .and_then(|s| serde_json::from_str::<serde_json::Value>(s).ok());
            let line = NdjsonLine {
                timestamp: &e.timestamp,
                level: e.level,
                target: &e.target,
                message: &e.message,
                fields,
            };
            // serde_json::to_string on NdjsonLine never fails (all fields are infallible)
            out.push_str(&serde_json::to_string(&line).unwrap());
            out.push('\n');
        }
        out
    }

    pub fn write_ndjson_to_file(&self, path: &std::path::Path) -> std::io::Result<()> {
        std::fs::write(path, self.to_ndjson())
    }
}
