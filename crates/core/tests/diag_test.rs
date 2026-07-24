use vocab_core::diag::DiagStore;

#[test]
fn diag_store_collects_events() {
    let mut store = DiagStore::new();
    store.info("font_probe", "scanning fonts", None);
    store.warn("font_probe", "emoji font not found", None);
    store.error("render", "SVG parse failed", Some(r#"{"detail":"line 12"}"#));

    assert_eq!(store.event_count(), 3);
    assert_eq!(store.warnings(), 1);
    assert_eq!(store.errors(), 1);
}

#[test]
fn diag_store_ndjson_output() {
    let mut store = DiagStore::new();
    store.info("test", "hello", Some(r#"{"key":"value"}"#));

    let ndjson = store.to_ndjson();
    assert!(ndjson.contains(r#""level":"INFO""#));
    assert!(ndjson.contains(r#""target":"test""#));
    assert!(ndjson.contains(r#""message":"hello""#));
    assert!(ndjson.contains(r#""key":"value""#));
}

#[test]
fn diag_store_empty_is_valid() {
    let store = DiagStore::new();
    assert_eq!(store.event_count(), 0);
    assert_eq!(store.to_ndjson(), "");
    assert_eq!(store.warnings(), 0);
    assert_eq!(store.errors(), 0);
}

#[test]
fn diag_store_writes_ndjson_file() {
    use std::io::Read;
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("test.ndjson");
    let mut store = DiagStore::new();
    store.info("test", "hello", None);
    store.write_ndjson_to_file(&path).unwrap();

    let mut content = String::new();
    std::fs::File::open(&path).unwrap().read_to_string(&mut content).unwrap();
    assert!(content.contains("hello"));
}
