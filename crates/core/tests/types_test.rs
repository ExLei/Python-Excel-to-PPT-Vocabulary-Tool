use vocab_core::types::*;

#[test]
fn word_entry_all_fields_present() {
    let entry = WordEntry {
        word: "apple".into(),
        phonetic: "/ˈæpl/".into(),
        morphology: "".into(),
        example: "I eat an apple.".into(),
        example_definition: "我吃苹果。".into(),
        definition: "苹果".into(),
    };
    assert_eq!(entry.word, "apple");
    assert_eq!(entry.morphology, "");
}

#[test]
fn input_source_excel_variant() {
    let src = InputSource::Excel {
        path: "test.xlsx".into(),
        sheet: "Sheet1".into(),
    };
    match src {
        InputSource::Excel { sheet, .. } => assert_eq!(sheet, "Sheet1"),
        _ => panic!("expected Excel variant"),
    }
}
