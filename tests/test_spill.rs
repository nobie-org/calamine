use calamine::{open_workbook, Reader, Xlsx};

#[test]
fn test_dynamic_array_spill_detection() {
    let path = format!("{}/tests/spill.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open spill.xlsx");

    let range = workbook
        .worksheet_range("Sheet1")
        .expect("Cannot get Sheet1 range");

    // A1 — source formula cell; should not be marked as spilled
    let a1 = range
        .get_value((0, 0))
        .expect("A1 should be within the produced range");
    assert!(!a1.is_spilled, "A1 must not be marked as spilled");

    // A2 — inside spill range; should be marked as spilled
    let a2 = range
        .get_value((1, 0))
        .expect("A2 should be within the produced range");
    assert!(a2.is_spilled, "A2 must be marked as spilled");
}
