//! Public read/write contracts, using PHP fixtures and real spreadsheet engines.
//! Fixtures originate from spatie/simple-excel at b6cea2b77b609770642a1cc2d91df5a1eb868fae
//! under its MIT license: <https://github.com/spatie/simple-excel/tree/b6cea2b77b609770642a1cc2d91df5a1eb868fae/tests/stubs>.

use std::{
    fs,
    io::{self, Write},
    path::PathBuf,
    sync::{Arc, Mutex},
};

use spatie_simple_excel::{
    Error, FileType, Format, Row, SimpleExcelReader as Reader, SimpleExcelWriter as Writer, Value,
    WriterEngine, WriterOptions,
};

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

fn rows(reader: Reader) -> Vec<Row> {
    reader
        .get_rows()
        .unwrap()
        .collect::<Result<_, _>>()
        .unwrap()
}

#[test]
fn reads_php_csv_fixtures_and_ignores_surplus_columns() {
    let data = rows(Reader::create(fixture("header-and-rows.csv")).unwrap());
    assert_eq!(data.len(), 2);
    assert_eq!(data[0]["email"], Value::from("john@example.com"));
    assert_eq!(data[1]["first_name"], Value::from("mary jane"));
    let short = rows(Reader::create(fixture("header-too-short.csv")).unwrap());
    assert_eq!(
        short[0].values(),
        &[Value::from("john@example.com"), Value::from("john")]
    );
    let padded = rows(Reader::create(fixture("row-too-short.csv")).unwrap());
    assert_eq!(padded[0]["first_name"], Value::from(""));
}

#[test]
fn handles_empty_and_header_only_files() {
    let reader = Reader::create(fixture("empty.csv")).unwrap();
    assert_eq!(reader.get_headers().unwrap(), None);
    assert!(rows(reader).is_empty());
    assert!(rows(Reader::create(fixture("only-header.csv")).unwrap()).is_empty());
}

#[test]
fn custom_headers_replace_keys_but_preserve_original_headers() {
    let reader = Reader::create(fixture("header-and-rows.csv"))
        .unwrap()
        .use_headers(["address", "given", "surname"]);
    assert_eq!(
        reader.get_headers().unwrap().unwrap(),
        ["address", "given", "surname"]
    );
    assert_eq!(
        reader.get_original_headers().unwrap().unwrap(),
        ["email", "first_name", "last_name"]
    );
    assert_eq!(rows(reader)[0]["given"], Value::from("john"));
    let raw = Reader::create(fixture("rows-without-header.csv"))
        .unwrap()
        .no_header_row();
    assert_eq!(raw.get_headers().unwrap(), None);
    assert_eq!(rows(raw.clone())[0][0], Value::from("john@example.com"));
    assert_eq!(
        rows(raw.use_headers(["email"]))[0].values(),
        &[Value::from("john@example.com")]
    );
}

#[test]
fn applies_header_transforms_in_php_order_and_preserves_empty_keys() {
    let reader = Reader::create(fixture("header-with-spaces.csv"))
        .unwrap()
        .trim_header_row()
        .headers_to_snake_case()
        .format_headers_using(|s| format!("{s}_suffix"));
    assert_eq!(rows(reader)[0]["first_name_suffix"], Value::from("john"));
    let headers = Reader::create(fixture("empty-header-title.csv"))
        .unwrap()
        .trim_header_row_with("il ")
        .get_headers()
        .unwrap()
        .unwrap();
    assert_eq!(headers, ["ema", "", "ast"]);
    let headers = Reader::create(fixture("headers-not-snake-case.csv"))
        .unwrap()
        .headers_to_snake_case()
        .get_headers()
        .unwrap()
        .unwrap();
    assert_eq!(headers, ["email", "first_name", "last_name", "job_title"]);
}

#[test]
fn trims_values_before_key_aware_formatting_without_changing_other_columns() {
    let reader = Reader::create(fixture("values-with-spaces.csv"))
        .unwrap()
        .trim_values()
        .format_values_using(|value, key| {
            if key == "email" {
                Value::from(value.to_string().to_uppercase())
            } else {
                value
            }
        });
    let data = rows(reader);
    assert_eq!(data[0]["email"], Value::from("JOHN@EXAMPLE.COM"));
    assert_eq!(data[0]["first_name"], Value::from("john"));
    let padded = rows(
        Reader::create(fixture("padded-values.csv"))
            .unwrap()
            .trim_values_with("*"),
    );
    assert_eq!(padded[0]["email"], Value::from("john@example.com"));
    let dates = rows(
        Reader::create(fixture("formatted_dates.xlsx"))
            .unwrap()
            .trim_values(),
    );
    assert!(matches!(dates[0]["created_at"], Value::DateTime(_)));
}

#[test]
fn php_character_ranges_digit_boundaries_and_duplicate_headers_are_preserved() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("data.csv");
    fs::write(
        &path,
        "A..zEmail2FA,A..zEmail2FA,UPPERCase\n..*Tom*..,..*Ada*.., x \n",
    )
    .unwrap();
    let reader = Reader::create(&path)
        .unwrap()
        .trim_header_row_with("A..z.")
        .headers_to_snake_case()
        .trim_values_with(".*");
    // PHP overwrites duplicate keys without moving the first key's position.
    let data = rows(reader);
    assert_eq!(data[0].headers().unwrap(), ["2", ""]);
    assert_eq!(data[0].values(), &[Value::from("Ada"), Value::from(" x ")]);
}

#[test]
fn csv_encoding_delimiter_and_enclosure_can_be_combined() {
    let data = rows(
        Reader::create(fixture("shift-jis-encoding.csv"))
            .unwrap()
            .use_encoding("SHIFT_JIS")
            .unwrap(),
    );
    assert_eq!(data[0]["お名前"], Value::from("太郎"));
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("input");
    fs::write(&path, b"'name';'note'\n'Andr\xe9';'a;b'\n").unwrap();
    let data = rows(
        Reader::create_with_type(path, FileType::Csv)
            .unwrap()
            .use_delimiter(b';')
            .unwrap()
            .use_encoding("Windows-1252")
            .unwrap()
            .use_field_enclosure(b'\'')
            .unwrap(),
    );
    assert_eq!(
        data[0].values(),
        &[Value::from("André"), Value::from("a;b")]
    );
}

#[test]
fn limits_are_repeatable_and_zero_does_not_yield_a_row() {
    let reader = Reader::create(fixture("header-and-rows.csv"))
        .unwrap()
        .skip(1)
        .take(1);
    for _ in 0..2 {
        assert_eq!(
            rows(reader.clone())[0]["first_name"],
            Value::from("mary jane")
        );
    }
    assert!(rows(reader.take(0)).is_empty());
}

#[test]
fn reads_php_xlsx_sheets_sparse_cells_and_offset_headers() {
    let reader = Reader::create(fixture("multiple_sheets.xlsx")).unwrap();
    assert_eq!(reader.get_sheet_names().unwrap(), ["sheet1", "sheet2"]);
    assert!(reader.has_sheet("sheet2").unwrap());
    assert!(!reader.has_sheet("Sheet2").unwrap());
    assert_eq!(
        reader.clone().from_sheet(2).get_headers().unwrap().unwrap(),
        ["contact", "email"]
    );
    assert_eq!(
        reader
            .clone()
            .from_sheet_name("sheet2")
            .get_headers()
            .unwrap(),
        reader.clone().from_sheet(2).get_headers().unwrap()
    );
    assert!(matches!(
        reader.clone().from_sheet(0).get_rows(),
        Err(Error::SheetNotFound(_))
    ));
    assert!(matches!(
        reader.from_sheet_name("absent").get_rows(),
        Err(Error::SheetNotFound(_))
    ));
    let data = rows(Reader::create(fixture("columns-returned-in-wrong-order.xlsx")).unwrap());
    assert_eq!(data[0]["id"].to_string(), "11223344");
    assert_eq!(data[0]["place"], Value::from(""));
    assert_eq!(data[0]["status"], Value::from("yes"));
    let data = rows(
        Reader::create(fixture("header-not-on-first-row.xlsx"))
            .unwrap()
            .header_on_row(2)
            .use_headers(["first_name", "last_name"]),
    );
    assert_eq!(data[0]["first_name"], Value::from("Taylor"));
    assert_eq!(data[1]["last_name"], Value::from("Wathan"));
}

#[test]
fn preserves_php_date_formats_and_optional_empty_rows() {
    let data = rows(
        Reader::create(fixture("formatted_dates.xlsx"))
            .unwrap()
            .preserve_date_time_formatting(),
    );
    assert_eq!(data[0]["created_at"], Value::from("9/20/2024"));
    assert_eq!(data[1]["created_at"], Value::from("9/19/2024"));
    let reader = Reader::create(fixture("empty_rows.xlsx")).unwrap();
    assert_eq!(rows(reader.clone()).len(), 2);
    assert_eq!(rows(reader.preserve_empty_rows()).len(), 3);
}

#[test]
fn writes_exact_csv_with_default_bom_and_counts_headers() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("data.csv");
    let mut writer = Writer::create(&path).unwrap();
    writer
        .add_rows([
            [("first_name", "John"), ("last_name", "Doe")],
            [("first_name", "Jane"), ("last_name", "Doe")],
        ])
        .unwrap();
    assert_eq!(writer.get_number_of_rows(), 3);
    writer.close().unwrap();
    writer.close().unwrap();
    assert_eq!(
        fs::read(&path).unwrap(),
        b"\xef\xbb\xbffirst_name,last_name\nJohn,Doe\nJane,Doe\n"
    );
    assert!(matches!(
        writer.add_row(Row::new(["late"])),
        Err(Error::Closed)
    ));
}

#[test]
fn explicit_headers_raw_rows_no_headers_and_csv_quoting() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("data.csv");
    let mut writer = Writer::create_with_options(
        &path,
        WriterOptions {
            delimiter: b';',
            add_bom: false,
            ..WriterOptions::default()
        },
    )
    .unwrap();
    writer
        .add_header(["name", "note"])
        .unwrap()
        .add_row(Row::new(["Jane", "a;b\n\"c\""]))
        .unwrap();
    writer.close().unwrap();
    assert_eq!(
        fs::read_to_string(&path).unwrap(),
        "name;note\nJane;\"a;b\n\"\"c\"\"\"\n"
    );
    let mut writer = Writer::create_without_bom(&path).unwrap();
    writer
        .no_header_row()
        .add_row([("name", "Jane"), ("note", "Doe")])
        .unwrap();
    writer.close().unwrap();
    assert_eq!(fs::read_to_string(&path).unwrap(), "Jane,Doe\n");
}

#[test]
fn xlsx_and_ods_roundtrip_types_dates_formulas_and_per_sheet_headers() {
    for extension in ["xlsx", "ods"] {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join(format!("data.{extension}"));
        let date = chrono::NaiveDate::from_ymd_opt(2024, 9, 20)
            .unwrap()
            .and_hms_opt(13, 14, 15)
            .unwrap();
        let mut writer = Writer::create(&path).unwrap();
        writer.name_current_sheet("People").unwrap();
        writer
            .add_row([
                ("name", Value::from("Jane")),
                ("age", Value::from(42)),
                ("active", Value::from(true)),
                ("date", date.into()),
                ("formula", Value::Formula("=1+2".into())),
                ("literal", Value::from("=1+2")),
                ("elapsed", chrono::Duration::seconds(178_245).into()),
            ])
            .unwrap();
        writer
            .add_new_sheet_and_make_it_current(Some("Other"))
            .unwrap()
            .add_row([("note", "second sheet")])
            .unwrap();
        assert_eq!(writer.get_number_of_rows(), 4);
        writer.close().unwrap();
        let reader = Reader::create(&path).unwrap();
        assert_eq!(reader.get_sheet_names().unwrap(), ["People", "Other"]);
        let data = rows(reader.clone());
        assert_eq!(data[0]["age"].to_string(), "42");
        assert_eq!(data[0]["active"], Value::Bool(true));
        assert_eq!(data[0]["date"], Value::DateTime(date));
        assert_eq!(
            data[0]["elapsed"],
            Value::Duration(chrono::Duration::seconds(178_245))
        );
        assert_eq!(data[0]["literal"], Value::from("=1+2"));
        assert_eq!(data[0]["formula"].to_string(), "0");
        assert_eq!(
            rows(reader.clone().keep_formulas())[0]["formula"],
            Value::Formula("=1+2".into())
        );
        assert_eq!(
            rows(reader.from_sheet(2))[0]["note"],
            Value::from("second sheet")
        );
    }
}

#[test]
fn xlsx_header_row_and_cell_styles_and_layout_survive_constant_memory_writes() {
    for explicit in [false, true] {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("data.xlsx");
        let mut writer = Writer::create(&path).unwrap();
        writer.set_header_style(Format::new().set_bold()).unwrap();
        writer
            .get_current_sheet()
            .unwrap()
            .set_column_width(0, 25)
            .unwrap()
            .set_row_height(0, 20)
            .unwrap();
        if explicit {
            writer.add_header(["name", "age"]).unwrap();
        }
        let mut row = Row::from_pairs([("name", Value::from("Jane")), ("age", Value::from(42))]);
        row.set_cell_format(
            1,
            Format::new()
                .set_underline(spatie_simple_excel::rust_xlsxwriter::FormatUnderline::Single),
        )
        .unwrap();
        writer
            .add_row_with_style(row, &Format::new().set_italic())
            .unwrap();
        writer.close().unwrap();
        let mut book: calamine::Xlsx<_> = calamine::open_workbook(&path).unwrap();
        let styles = book.worksheet_style("Sheet1").unwrap();
        assert!(styles.get((0, 0)).unwrap().font.as_ref().unwrap().is_bold());
        assert!(
            styles
                .get((1, 0))
                .unwrap()
                .font
                .as_ref()
                .unwrap()
                .is_italic()
        );
        assert!(
            styles
                .get((1, 1))
                .unwrap()
                .font
                .as_ref()
                .unwrap()
                .has_underline()
        );
        assert!(
            !styles
                .get((1, 1))
                .unwrap()
                .font
                .as_ref()
                .unwrap()
                .is_italic()
        );
        assert_eq!(
            rows(Reader::create(&path).unwrap())[0]["name"],
            Value::from("Jane")
        );
    }
}

#[derive(Clone, Default)]
struct Sink(Arc<Mutex<Vec<u8>>>);
impl Write for Sink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.0.lock().unwrap().extend_from_slice(bytes);
        Ok(bytes.len())
    }
    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn downloads_stream_csv_and_finalize_readable_spreadsheets() {
    for extension in ["csv", "xlsx", "ods"] {
        let sink = Sink::default();
        let mut writer =
            Writer::stream_download(format!("export.{extension}"), sink.clone()).unwrap();
        writer.add_row([("name", "Jane")]).unwrap();
        writer.flush().unwrap();
        assert_eq!(sink.0.lock().unwrap().is_empty(), extension != "csv");
        writer.to_browser().unwrap();
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join(format!("data.{extension}"));
        fs::write(&path, &*sink.0.lock().unwrap()).unwrap();
        assert_eq!(
            rows(Reader::create(&path).unwrap())[0]["name"],
            Value::from("Jane")
        );
    }
}

#[test]
fn errors_are_observable_and_invalid_options_do_not_truncate_existing_csv() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("data.csv");
    fs::write(&path, "keep me").unwrap();
    assert!(
        Writer::create_with_options(
            &path,
            WriterOptions {
                delimiter: b'"',
                ..WriterOptions::default()
            }
        )
        .is_err()
    );
    assert_eq!(fs::read_to_string(&path).unwrap(), "keep me");
    assert!(matches!(
        Reader::create("file.bad"),
        Err(Error::UnsupportedType(_))
    ));
    assert!(Reader::create("missing.csv").unwrap().get_rows().is_err());
    assert!(
        Reader::create(&path)
            .unwrap()
            .use_encoding("not-an-encoding")
            .is_err()
    );
    let mut writer = Writer::create(&path).unwrap();
    assert!(writer.add_new_sheet_and_make_it_current(None).is_err());
    assert!(writer.add_row(Row::new([f64::NAN])).is_err());
    assert_eq!(writer.get_number_of_rows(), 0);
    assert!(matches!(writer.get_writer().unwrap(), WriterEngine::Csv(_)));
}

#[test]
fn csv_is_read_lazily_and_both_iterators_remain_independent() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("data.csv");
    let mut data = b"name\nJane\n".to_vec();
    data.extend(std::iter::repeat_n(b'x', 100_000));
    data.push(b'\n');
    fs::write(&path, data).unwrap();
    let reader = Reader::create(&path).unwrap();
    let mut first = reader.get_rows().unwrap();
    let mut second = reader.get_rows().unwrap();
    assert_eq!(first.next().unwrap().unwrap()["name"], Value::from("Jane"));
    assert_eq!(second.next().unwrap().unwrap()["name"], Value::from("Jane"));
    fs::OpenOptions::new()
        .append(true)
        .open(&path)
        .unwrap()
        .write_all(b"Late\n")
        .unwrap();
    for iterator in [&mut first, &mut second] {
        assert_eq!(
            iterator.next().unwrap().unwrap()["name"]
                .as_str()
                .unwrap()
                .len(),
            100_000
        );
        assert_eq!(
            iterator.next().unwrap().unwrap()["name"],
            Value::from("Late")
        );
        assert!(iterator.next().is_none());
    }
}

#[test]
fn drop_finalizes_files_but_close_reports_output_failures() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("data.xlsx");
    {
        let mut writer = Writer::create(&path).unwrap();
        writer.add_row([("name", "Jane")]).unwrap();
    }
    assert_eq!(
        rows(Reader::create(&path).unwrap())[0]["name"],
        Value::from("Jane")
    );
    struct Broken;
    impl Write for Broken {
        fn write(&mut self, _: &[u8]) -> io::Result<usize> {
            Err(io::Error::other("broken output"))
        }
        fn flush(&mut self) -> io::Result<()> {
            Err(io::Error::other("broken output"))
        }
    }
    let mut writer = Writer::stream_download("data.xlsx", Broken).unwrap();
    writer.add_row([("name", "Jane")]).unwrap();
    assert!(
        writer
            .close()
            .unwrap_err()
            .to_string()
            .contains("broken output")
    );
    assert!(matches!(
        writer.add_row([("name", "late")]),
        Err(Error::Closed)
    ));
}
