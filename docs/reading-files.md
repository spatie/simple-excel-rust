---
title: Reading files
weight: 4
---

## Reading a CSV

Imagine you have a CSV with this content.

```csv
email,first_name
john@example.com,john
jane@example.com,jane
```

```rust
# use spatie_simple_excel::SimpleExcelReader;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let directory = tempfile::tempdir()?;
# let path_to_csv = directory.path().join("people.csv");
# std::fs::write(&path_to_csv, "email,first_name\njohn@example.com,john\njane@example.com,jane\n")?;
let rows = SimpleExcelReader::create(path_to_csv)?.get_rows()?;

for row in rows {
    let row_properties = row?;
    println!("{}", row_properties["email"]);
#   assert!(row_properties["email"].as_str().unwrap().ends_with("@example.com"));
}
# Ok(())
# }
```

Rows keep column order. Use `row["email"]` for a known header, `row.get("email")` for an optional header, or `row[0]` for a numeric column. Indexing a missing key panics; `get` returns `None`. Values are typed `Value` variants; CSV values are strings. With headers, short rows are padded with empty strings, and extra columns are discarded. Duplicate header names keep the last value at the first key's position.

### Reading an Excel file

Reading an Excel file is identical to reading a CSV file. Just make sure that the path given to the `create` method of `SimpleExcelReader` ends with `xlsx`. ODS files use the same interface with an `ods` extension.

Use `SimpleExcelReader::create_with_type(path, FileType::Csv)` when the path has no extension or a different one. `FileType` is exported by this crate.

Reader configuration consumes and returns the reader. Files open when you request rows, headers or sheet metadata, not when you call `create`.

### Working with iterators

`get_rows` returns an iterator of `Result<Row, Error>`. Process each row with a `for` loop or Rust's iterator adapters. Read errors terminate the iterator.

Here's a quick, silly example where we only want to process rows that have a `first_name` that contains more than 5 characters.

```rust
# use spatie_simple_excel::SimpleExcelReader;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let path_to_csv = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/header-and-rows.csv");
let rows = SimpleExcelReader::create(path_to_csv)?.get_rows()?;
for row in rows {
    let row_properties = row?;
    if row_properties["first_name"].as_str().is_some_and(|name| name.len() > 5) {
        println!("{}", row_properties["first_name"]);
#       assert_eq!(row_properties["first_name"].as_str(), Some("mary jane"));
    }
}
# Ok(())
# }
```

Use `collect::<Result<Vec<_>, _>>()?` when you deliberately want all rows in memory. Do not use `filter_map(Result::ok)` unless discarding read errors is intentional. Each `get_rows()` call opens a fresh independent iterator, with the same skip and take settings. Drop the iterator to close the file; no reader `close()` call is needed.

<a id="memory-usage"></a>

### Memory usage

CSV reads stream through `csv` and `encoding_rs_io`. Memory depends on the current record and fixed-size buffers, not the total row count. Encoding conversion also happens incrementally.

XLSX reads use `calamine`, which loads the selected sheet and shared strings into memory. Keeping formulas or preserving date formatting also loads formula or style ranges. Iterating avoids a second full copy of the sheet, but does not make XLSX reading constant-memory.

ODS reads load the entire workbook into memory. For inputs too large to fit in memory, use CSV. Calling `collect` retains the returned rows in memory for every format.

### Reading a file without headers

If the file you are reading does not contain a header row, use `no_header_row()`.

```rust
# use spatie_simple_excel::SimpleExcelReader;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let path_to_csv = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/rows-without-header.csv");
let rows = SimpleExcelReader::create(path_to_csv)?
    .no_header_row()
    .get_rows()?;
# let row = rows.into_iter().next().unwrap()?;
# assert_eq!(row[0].as_str(), Some("john@example.com"));
# assert!(row.headers().is_none());
# Ok(())
# }
```

### Manually setting the headers

If you would like to use a specific array of values for the headers, use `use_headers()`.

```rust
# use spatie_simple_excel::SimpleExcelReader;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let path_to_csv = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/header-and-rows.csv");
let rows = SimpleExcelReader::create(path_to_csv)?
    .use_headers(["email_address", "given_name"])
    .get_rows()?;
# let row = rows.into_iter().next().unwrap()?;
# assert_eq!(row["given_name"].as_str(), Some("john"));
# Ok(())
# }
```

If your file already contains a header row, it will be ignored and replaced with your custom headers. If it does not contain a header row, also use `no_header_row()`.

### Working with multiple sheet documents

Excel files can include multiple spreadsheets. Select the sheet with `from_sheet()` using a one-based index, or `from_sheet_name()` using its exact name. `get_sheet_names()` retrieves all sheet names and `has_sheet()` checks for a name.

```rust
# use spatie_simple_excel::SimpleExcelReader;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let path_to_xlsx = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/multiple_sheets.xlsx");
let reader = SimpleExcelReader::create(path_to_xlsx)?;
let sheets = reader.get_sheet_names()?;
let has_sheet = reader.has_sheet("sheet1")?;
let rows = reader.clone().from_sheet(2).get_rows()?;
let named_rows = reader.from_sheet_name("sheet2").get_rows()?;
# assert_eq!(sheets, ["sheet1", "sheet2"]);
# assert!(has_sheet);
# assert_eq!(rows.count(), named_rows.count());
# Ok(())
# }
```

An absent sheet returns `Error::SheetNotFound`. CSV exposes one sheet named `Sheet1`.

### Retrieving header row values

Use `get_headers()` to retrieve the header row as an optional vector. If you set custom headers, these are returned instead of the file's headers. `get_original_headers()` returns the file's headers after trimming and formatting but before custom replacement.

```rust
# use spatie_simple_excel::SimpleExcelReader;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let path_to_csv = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/header-and-rows.csv");
let headers = SimpleExcelReader::create(path_to_csv)?.get_headers()?;
# assert_eq!(headers.unwrap(), ["email", "first_name", "last_name"]);
# Ok(())
# }
```

### Dealing with headers that are not on the first line

Use `header_on_row()` to indicate the row at which the headers are present. Any data above this row is discarded. The argument starts at 0. Blank rows do not count unless spreadsheet empty rows are explicitly preserved.

```rust
# use spatie_simple_excel::SimpleExcelReader;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let path_to_xlsx = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/header-not-on-first-row.xlsx");
let rows = SimpleExcelReader::create(path_to_xlsx)?
    .trim_header_row()
    .header_on_row(2)
    .get_rows()?;
# assert_eq!(rows.into_iter().next().unwrap()?["firstname"].as_str(), Some("Taylor"));
# Ok(())
# }
```

### Trimming headers

If the header row contains additional characters, use `trim_header_row()`. It removes space, tab, newline, carriage return, NUL and vertical tab. `trim_header_row_with("a..z")` accepts a character set with ASCII ranges. Unicode characters are treated as characters rather than individual UTF-8 bytes.

### Convert headers to snake_case

Use `headers_to_snake_case()` to convert headers such as `First Name`, `jobTitle` and `Column2` into `first_name`, `job_title` and `column_2`.

### Manually formatting headers

Use `format_headers_using()` with a closure. Formatting runs after trimming and snake_case conversion.

```rust
# use spatie_simple_excel::SimpleExcelReader;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let path_to_csv = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/header-and-rows.csv");
let rows = SimpleExcelReader::create(path_to_csv)?
    .format_headers_using(|header| format!("{header}_simple_excel"))
    .get_rows()?;
# assert_eq!(rows.into_iter().next().unwrap()?["email_simple_excel"].as_str(), Some("john@example.com"));
# Ok(())
# }
```

### Trimming and formatting values

Just like you can clean up header names, you can clean up the data values themselves. Use `trim_values()` to strip whitespace from every string value. `trim_values_with("*")` strips a set of characters from both ends, not a suffix. Non-string values, such as dates, are untouched.

For full control, use `format_values_using()`. The closure receives the value and its header key. For unkeyed rows, the key is the numeric index rendered as a string.

```rust
# use spatie_simple_excel::{SimpleExcelReader, Value};
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let path_to_csv = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/values-with-spaces.csv");
let rows = SimpleExcelReader::create(path_to_csv)?
    .trim_values()
    .format_values_using(|value, key| match value {
        Value::String(text) if key == "email" => text.to_lowercase().into(),
        other => other,
    })
    .get_rows()?;
# assert_eq!(rows.into_iter().next().unwrap()?["email"].as_str(), Some("john@example.com"));
# Ok(())
# }
```

### Manually working with the reader object

`get_reader()` opens the underlying `calamine::Sheets` reader. `get_csv_reader()` opens the configured `csv::Reader`. These are independent engines: changing them does not change later `get_rows()` calls. All engine crates are re-exported where needed for advanced use.

### Limiting the result set

The `take` method limits the number of data rows. The `skip` method defines where to start after the header. For example, `.skip(10).take(5)` returns data rows 11 through 15. Both settings are reapplied on every `get_rows()` call. `take(0)` returns no rows.

### Reading cells that contain formulas

Normally, cells containing formulas return the cached computed value stored in the file. Use `keep_formulas()` to return `Value::Formula` strings instead. This package does not evaluate formulas. Files without cached results may return an empty value or the engine's stored default.

Formula strings start with `=` in both XLSX and ODS; the ODS namespace prefix is removed. Spreadsheet cell errors, such as `#DIV/0!`, are returned as `Value::Error`.

### Preserve date formatting

By default spreadsheet dates and times are returned as `Value::DateTime(chrono::NaiveDateTime)`. They have no timezone. Durations use `Value::Duration`. Use `preserve_date_time_formatting()` to return XLSX date strings using the stored number format.

```rust
# use spatie_simple_excel::SimpleExcelReader;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let path_to_xlsx = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/formatted_dates.xlsx");
let rows = SimpleExcelReader::create(path_to_xlsx)?
    .preserve_date_time_formatting()
    .get_rows()?;
# assert_eq!(rows.into_iter().next().unwrap()?["created_at"].as_str(), Some("9/20/2024"));
# Ok(())
# }
```

Date formatting uses `ssfmt`'s Excel number-format support. It does not reproduce every locale-specific Excel display rule. Built-in format IDs 14, 22 and 47 display as `m/d/yyyy`, `m/d/yyyy h:mm` and `mm:ss.0`. Invalid date values or unparseable date formats return errors. The date-format option only applies to XLSX.

### Preserve empty rows

Use `preserve_empty_rows()` to preserve spreadsheet empty rows within the used range, including leading empty rows. Trailing rows beyond that range cannot be recovered. CSV blank lines are ignored.

Leading and interior empty cells keep their column positions. Unkeyed spreadsheet rows omit trailing blank cells; keyed rows pad missing values with empty strings.

### Using an alternative delimiter, enclosure or encoding

The CSV delimiter defaults to a comma, and the enclosure defaults to a double quote. Use `use_delimiter(b';')`, `use_field_enclosure(b'\'')` and `use_encoding("SHIFT_JIS")` to change them. Delimiters and enclosures must be distinct, non-NUL ASCII bytes other than CR or LF. Encoding labels follow the WHATWG registry. Malformed encoded sequences become U+FFFD, and BOMs are removed during decoding.
