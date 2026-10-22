use pptx_edit::{DeckSession, EditCtx, EditError};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/table-basic.pptx");

fn ids(session: &DeckSession) -> (String, String) {
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let table = slide
        .shapes
        .iter()
        .find(|shape| shape.graphic.is_some())
        .unwrap();
    (slide.id.clone(), table.id.clone())
}

fn sizes(session: &DeckSession) -> (Vec<i64>, Vec<i64>, i64, i64) {
    let snapshot = session.snapshot().unwrap();
    let shape = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.graphic.is_some())
        .unwrap();
    let Some(pptx_parse::GraphicFrameData::Table(table)) = &shape.graphic else {
        panic!("expected a table")
    };
    (
        table.grid.clone(),
        table.rows.iter().map(|row| row.height).collect(),
        shape.width,
        shape.height,
    )
}

fn part(bytes: &[u8], path: &str) -> String {
    ooxml_opc::unzip_parts(bytes)
        .unwrap()
        .into_iter()
        .find(|(name, _)| name == path)
        .map(|(_, bytes)| String::from_utf8(bytes).unwrap())
        .unwrap()
}

#[test]
fn table_sizes_change_live_undo_and_redo_as_one_step() {
    let session = DeckSession::open(DECK, 1).unwrap();
    let (slide, shape) = ids(&session);
    let before = sizes(&session);
    session.add_undo_barrier();
    session
        .set_table_sizes(
            &EditCtx::local("test"),
            &slide,
            &shape,
            vec![2000000, 1000000, 2500000],
            vec![400000, 500000, 600000],
        )
        .unwrap();
    session.add_undo_barrier();
    let after = (
        vec![2000000, 1000000, 2500000],
        vec![400000, 500000, 600000],
        5500000,
        1500000,
    );
    assert_eq!(sizes(&session), after);
    assert!(session.undo());
    assert_eq!(sizes(&session), before);
    assert!(session.redo());
    assert_eq!(sizes(&session), after);
}

#[test]
fn saved_table_sizes_patch_the_grid_rows_and_frame_and_keep_merges() {
    let session = DeckSession::open(DECK, 2).unwrap();
    let (slide, shape) = ids(&session);
    session
        .set_table_sizes(
            &EditCtx::local("test"),
            &slide,
            &shape,
            vec![2000000, 1000000, 2500000],
            vec![400000, 500000, 600000],
        )
        .unwrap();
    let saved = session.save().unwrap();
    let xml = part(&saved, "ppt/slides/slide1.xml");
    assert!(xml.contains(
        r#"<a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="1000000"/><a:gridCol w="2500000"/></a:tblGrid>"#
    ));
    for height in ["400000", "500000", "600000"] {
        assert!(xml.contains(&format!(r#"<a:tr h="{height}">"#)), "{height}");
    }
    assert!(xml.contains(r#"<a:ext cx="5500000" cy="1500000"/>"#));
    for merge in [
        r#"<a:tc gridSpan="2">"#,
        r#"<a:tc hMerge="1">"#,
        r#"<a:tc rowSpan="2">"#,
        r#"<a:tc vMerge="1">"#,
    ] {
        assert!(xml.contains(merge), "{merge}");
    }
    let reopened = DeckSession::open(&saved, 3).unwrap();
    assert_eq!(sizes(&reopened), sizes(&session));
    for path in ["ppt/presentation.xml", "ppt/slideLayouts/slideLayout1.xml"] {
        assert_eq!(part(&saved, path), part(DECK, path), "{path}");
    }
}

#[test]
fn unchanged_table_sizes_leave_the_slide_untouched() {
    let session = DeckSession::open(DECK, 4).unwrap();
    let (slide, shape) = ids(&session);
    let (widths, heights, _, _) = sizes(&session);
    session
        .set_table_sizes(&EditCtx::local("test"), &slide, &shape, widths, heights)
        .unwrap();
    let saved = session.save().unwrap();
    assert_eq!(
        part(&saved, "ppt/slides/slide1.xml"),
        part(DECK, "ppt/slides/slide1.xml")
    );
}

#[test]
fn invalid_table_sizes_are_refused_without_changes() {
    let session = DeckSession::open(DECK, 5).unwrap();
    let (slide, shape) = ids(&session);
    let before = sizes(&session);
    let context = EditCtx::local("test");
    for (widths, heights) in [
        (vec![1, 1], vec![1, 1, 1]),
        (vec![1, 1, 1], vec![1, 1]),
        (vec![0, 0, 0], vec![1, 1, 1]),
        (vec![1, -1, 1], vec![1, 1, 1]),
        (vec![i64::MAX, i64::MAX, 1], vec![1, 1, 1]),
    ] {
        let error = session
            .set_table_sizes(&context, &slide, &shape, widths.clone(), heights)
            .unwrap_err();
        assert!(
            matches!(error, EditError::InvalidGeometry(_)),
            "{widths:?}: {error:?}"
        );
    }
    assert!(
        session
            .set_table_sizes(&context, &slide, "missing", vec![1, 1, 1], vec![1, 1, 1])
            .is_err()
    );
    assert_eq!(sizes(&session), before);
    assert!(!session.can_undo());
}

fn remote_table_edit(edit: impl Fn(&mut serde_json::Value)) -> (DeckSession, Vec<u8>) {
    use yrs::updates::decoder::Decode;
    use yrs::{Doc, Map, MapRef, Out, ReadTxn, Transact, Update};
    let origin = DeckSession::open(DECK, 10).unwrap();
    let replica =
        DeckSession::open_from_update_with_source(&origin.encode_state_as_update_v1(), DECK, 11)
            .unwrap();
    let doc = Doc::with_client_id(12);
    doc.transact_mut()
        .apply_update(Update::decode_v1(&replica.encode_state_as_update_v1()).unwrap())
        .unwrap();
    let before = doc.transact().state_vector();
    {
        let mut txn = doc.transact_mut();
        let shapes = txn.get_map("pptx:shapes").unwrap();
        let maps: Vec<MapRef> = shapes
            .iter(&txn)
            .filter_map(|(_, value)| value.cast::<MapRef>().ok())
            .collect();
        for shape in maps {
            if let Some(Out::Any(yrs::Any::String(json))) = shape.get(&txn, "graphicJson") {
                let mut value: serde_json::Value = serde_json::from_str(&json).unwrap();
                if value["type"] == "table" {
                    edit(&mut value);
                    shape.insert(&mut txn, "graphicJson", value.to_string());
                }
            }
        }
    }
    let update = doc.transact().encode_state_as_update_v1(&before);
    (replica, update)
}

#[test]
fn a_remote_update_with_negative_table_sizes_is_refused() {
    let (replica, update) = remote_table_edit(|table| table["rows"][1]["height"] = (-5).into());
    let before = sizes(&replica);
    assert!(replica.apply_update_v1(&update).is_err());
    assert_eq!(sizes(&replica), before);
    replica.save().unwrap();
}

#[test]
fn a_remote_update_that_changes_the_table_grid_count_still_saves() {
    let (replica, update) =
        remote_table_edit(|table| table["grid"].as_array_mut().unwrap().push(1000000.into()));
    replica.apply_update_v1(&update).unwrap();
    let saved = replica.save().unwrap();
    assert_eq!(
        part(&saved, "ppt/slides/slide1.xml"),
        part(DECK, "ppt/slides/slide1.xml")
    );
}

#[test]
fn a_frame_taller_than_its_rows_keeps_its_extra_height() {
    let session = DeckSession::open(DECK, 6).unwrap();
    let (slide, shape) = ids(&session);
    let context = EditCtx::local("test");
    session
        .resize_shape(&context, &slide, &shape, 5715000, 2000000)
        .unwrap();
    let (mut widths, heights, _, _) = sizes(&session);
    widths[0] += 100000;
    session
        .set_table_sizes(&context, &slide, &shape, widths, heights)
        .unwrap();
    let (_, _, width, height) = sizes(&session);
    assert_eq!((width, height), (5815000, 2000000));
}

#[test]
fn rows_with_no_minimum_height_keep_it_through_a_column_resize() {
    let parts: Vec<_> = ooxml_opc::unzip_parts(DECK)
        .unwrap()
        .into_iter()
        .map(|(name, bytes)| {
            if name == "ppt/slides/slide1.xml" {
                let xml = String::from_utf8(bytes).unwrap();
                (
                    name,
                    xml.replace(r#"<a:tr h="370840">"#, r#"<a:tr h="0">"#)
                        .into_bytes(),
                )
            } else {
                (name, bytes)
            }
        })
        .collect();
    let deck = ooxml_opc::rezip_parts(&parts).unwrap();
    let session = DeckSession::open(&deck, 7).unwrap();
    let (slide, shape) = ids(&session);
    let (mut widths, heights, _, frame_height) = sizes(&session);
    assert_eq!(heights, vec![0, 0, 0]);
    widths[2] += 100000;
    session
        .set_table_sizes(&EditCtx::local("test"), &slide, &shape, widths, heights)
        .unwrap();
    assert_eq!(sizes(&session).3, frame_height);
}
