---
title: Writing files
weight: 5
---

## Writing files

Here's how you can write a CSV file:

```rust
# use spatie_simple_excel::SimpleExcelWriter;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let directory = tempfile::tempdir()?;
# let path_to_csv = directory.path().join("people.csv");
let mut writer = SimpleExcelWriter::create(&path_to_csv)?;
writer
    .add_row([("first_name", "John"), ("last_name", "Doe")])?
    .add_row([("first_name", "Jane"), ("last_name", "Doe")])?;
writer.close()?;
# assert_eq!(std::fs::read_to_string(path_to_csv)?, "\u{feff}first_name,last_name\nJohn,Doe\nJane,Doe\n");
# Ok(())
# }
```

The file contains:

```csv
first_name,last_name
John,Doe
Jane,Doe
```

Pair arrays become keyed rows. The first keyed row on each sheet supplies the automatic header. Later rows are written in input order, not reordered by key. Use `Value` for mixed-type rows. Text is always literal, including text beginning with `=`; use `Value::Formula` explicitly when you want a spreadsheet formula. CSV has no formula cell type, so spreadsheet applications may still interpret exported text as formulas when opening a CSV. Sanitize untrusted CSV exports for your destination application.

Spreadsheet formulas are stored with a cached result of zero until a spreadsheet application recalculates them. `Value::Error` writes its displayed text, not an error cell.

### Manually set the header from array

Instead of automatically detecting a header row, you can set it manually. `Row::new()` creates an unkeyed row and does not generate headers.

```rust
# use spatie_simple_excel::{SimpleExcelWriter, Row};
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let directory = tempfile::tempdir()?;
# let path_to_csv = directory.path().join("people.csv");
let mut writer = SimpleExcelWriter::create(&path_to_csv)?;
writer
    .add_header(["first_name", "last_name"])?
    .add_row(Row::new(["John", "Doe"]))?
    .add_row(Row::new(["Jane", "Doe"]))?;
writer.close()?;
# assert_eq!(writer.get_number_of_rows(), 3);
# Ok(())
# }
```

### Writing an Excel file

Writing an Excel file is identical to writing a CSV. Just make sure that the path given to `create` ends with `xlsx`. ODS uses the same interface with an `ods` extension. `create_with_type` overrides extension detection.

The `close` method finalizes the archive. Call it before reading the file and to receive finalization errors. Drop attempts finalization too, but cannot report errors. Repeated `close()` calls succeed; adding rows after closing returns `Error::Closed`. Failed writes may leave partial output. Writing is not transactional.

The Excel limits are 1,048,576 rows and 16,384 columns per sheet; create a new sheet before reaching the row limit. Spreadsheet numbers must be finite. Integers beyond the exact f64 range may lose precision; write identifiers as text.

### Memory usage

CSV writes stream through `csv`, with memory proportional to the current record and output buffers. `add_rows` consumes its input incrementally rather than collecting it.

XLSX uses `rust_xlsxwriter` constant-memory worksheets. Previous rows spool to temporary files rather than accumulating in RAM. Formats, workbook metadata and ZIP finalization still allocate memory, and temporary disk usage grows with the export. Cells must be written in increasing row order; the wrapper handles this automatically. Native engine access cannot edit rows that have already been flushed.

ODS writes use `spatie-spreadsheet-ods` and hold the workbook in memory. Choose CSV or XLSX for large exports that need bounded row-data memory.

### Streaming an Excel file to the browser

Instead of writing to a named file, provide an output implementing `std::io::Write`. Your HTTP application supplies response headers such as `Content-Type` and `Content-Disposition`. `to_browser()` closes the writer; it never exits the process.

```rust
# use spatie_simple_excel::SimpleExcelWriter;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let directory = tempfile::tempdir()?;
# let path = directory.path().join("response.xlsx");
# let response_body = std::fs::File::create(&path)?;
let mut writer = SimpleExcelWriter::stream_download("your-export.xlsx", response_body)?;
writer
    .add_row([("first_name", "John"), ("last_name", "Doe")])?
    .add_row([("first_name", "Jane"), ("last_name", "Doe")])?;
writer.to_browser()?;
# let rows = spatie_simple_excel::SimpleExcelReader::create(path)?.get_rows()?;
# assert_eq!(rows.count(), 2);
# Ok(())
# }
```

CSV output streams directly. Use `flush()` periodically for large CSV streams. XLSX and ODS must finish their ZIP archive first; downloads spool to a temporary file and copy to your output at close. Calling `flush()` does not stream a partial spreadsheet archive. The sink must be owned (`'static`); use a file, stdout handle or an owned adapter for your web framework.

Configure the native engine through `get_writer()` or `get_current_sheet()` before adding rows.

### Writing multiple rows at once

Use `add_rows` instead of `add_row` to add multiple rows. Any iterator is accepted and consumed incrementally.

```rust
# use spatie_simple_excel::SimpleExcelWriter;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let directory = tempfile::tempdir()?;
# let path_to_csv = directory.path().join("people.csv");
let mut writer = SimpleExcelWriter::create(path_to_csv)?;
writer.add_rows([
    [("first_name", "John"), ("last_name", "Doe")],
    [("first_name", "Jane"), ("last_name", "Doe")],
])?;
writer.close()?;
# assert_eq!(writer.get_number_of_rows(), 3);
# Ok(())
# }
```

`add_rows_with_style` applies one XLSX format to all supplied rows. An input iterator yielding `Result` can instead be handled with a `for` loop and `writer.add_row(row?)?`.

### Writing a file without titles

Use `no_header_row()` if a title row should not be added automatically.

```rust
# use spatie_simple_excel::SimpleExcelWriter;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let directory = tempfile::tempdir()?;
# let path_to_csv = directory.path().join("people.csv");
let mut writer = SimpleExcelWriter::create(&path_to_csv)?;
writer.no_header_row().add_row([("first_name", "Jane"), ("last_name", "Doe")])?;
writer.close()?;
# assert_eq!(std::fs::read_to_string(path_to_csv)?, "\u{feff}Jane,Doe\n");
# Ok(())
# }
```

### Adding layout

The `rust_xlsxwriter::Format` builder formats XLSX rows. It is exported as `Format`. CSV ignores formatting. ODS uses native `spreadsheet_ods` styles through `get_writer()`; supplying an XLSX format to ODS returns an error.

```rust
# use spatie_simple_excel::{SimpleExcelWriter, Row};
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let directory = tempfile::tempdir()?;
# let path_to_xlsx = directory.path().join("people.xlsx");
use spatie_simple_excel::{Format, rust_xlsxwriter::{Color, FormatBorder}};

let style = Format::new()
    .set_bold()
    .set_font_size(15)
    .set_font_color(Color::Blue)
    .set_text_wrap()
    .set_background_color(Color::Yellow)
    .set_border(FormatBorder::Thin);

let mut writer = SimpleExcelWriter::create(path_to_xlsx)?;
writer.set_header_style(style.clone())?;
writer.add_row_with_style(Row::new(["values", "of", "the", "row"]), &style)?;
writer.close()?;
# assert_eq!(writer.get_number_of_rows(), 1);
# Ok(())
# }
```

To style headers, call `set_header_style`. It applies to both automatic and explicit headers. To style a cell, call `row.set_cell_format(column, format)?`; the cell format replaces the row format for that cell. Date cells have a default `yyyy-mm-dd hh:mm:ss` format and durations use `[hh]:mm:ss`; explicit styles replace these defaults too.

### Setting column widths and row heights

Access the current worksheet to set column widths and row heights. Do this before writing rows in constant-memory mode.

```rust
# use spatie_simple_excel::SimpleExcelWriter;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let directory = tempfile::tempdir()?;
# let path_to_xlsx = directory.path().join("document.xlsx");
let mut writer = SimpleExcelWriter::create(path_to_xlsx)?;
let sheet = writer.get_current_sheet()?;
sheet.set_column_range_width(0, 7, 25)?;
sheet.set_row_height(0, 15)?;
for column in [0, 2, 7] {
    sheet.set_column_width(column, 40)?;
}
sheet.set_column_range_width(8, 11, 10)?;
writer.close()?;
# Ok(())
# }
```

Native engine column and row coordinates start at zero.

### Creating additional sheets

By default the writer writes to the first sheet. Use `name_current_sheet()` to name it, and `add_new_sheet_and_make_it_current()` for another sheet. Each sheet starts with automatic header detection enabled unless you called `no_header_row()` globally.

```rust
# use spatie_simple_excel::SimpleExcelWriter;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let directory = tempfile::tempdir()?;
# let path_to_xlsx = directory.path().join("document.xlsx");
let mut writer = SimpleExcelWriter::create(&path_to_xlsx)?;
writer.name_current_sheet("Posts")?.add_row([("title", "My first post")])?;
writer.add_new_sheet_and_make_it_current(Some("Comments"))?
    .add_row([("comment", "Hello"), ("author", "Jane")])?;
writer.close()?;
# assert_eq!(spatie_simple_excel::SimpleExcelReader::create(path_to_xlsx)?.get_sheet_names()?, ["Posts", "Comments"]);
# Ok(())
# }
```

Pass `None` for names such as `Sheet2`. CSV does not support sheet naming or additional sheets and returns an error.

### Using an alternative delimiter

Configure the delimiter at creation, before the file is opened.

```rust
# use spatie_simple_excel::SimpleExcelWriter;
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let directory = tempfile::tempdir()?;
# let path_to_csv = directory.path().join("people.csv");
use spatie_simple_excel::WriterOptions;

let mut writer = SimpleExcelWriter::create_with_options(path_to_csv, WriterOptions {
    delimiter: b';',
    add_bom: false,
    ..WriterOptions::default()
})?;
writer.add_row([("name", "Jane"), ("note", "a;b")])?;
writer.close()?;
# Ok(())
# }
```

### Getting the number of rows written

`get_number_of_rows()` returns the total across all sheets, including automatically added and explicit header rows. One keyed data row with an automatic header counts as two rows.

### Disable BOM

Use `SimpleExcelWriter::create_without_bom(path)?` or `WriterOptions { add_bom: false, ..WriterOptions::default() }` to disable the UTF-8 BOM. This must be configured at creation. CSV defaults to a BOM, comma delimiter, double quote enclosure and LF record endings.

### Manually working with the writer object

`get_writer()` returns `WriterEngine::Csv`, `WriterEngine::Xlsx` or `WriterEngine::Ods` with a mutable native engine. `get_current_sheet()` is a convenient shortcut for XLSX layout. Native writes bypass row counts and header tracking, so use wrapper methods for rows and sheet creation. Closing releases the native engine; subsequent access returns `Error::Closed`.
