---
title: PHP parity and memory usage
weight: 6
---

This port follows [spatie/simple-excel at b6cea2b](https://github.com/spatie/simple-excel/tree/b6cea2b77b609770642a1cc2d91df5a1eb868fae). The PHP package has a README rather than a separate guide tree; the reading and writing guides follow its headings, examples and wording where the behavior matches.

## Memory usage

- **CSV reading and writing** stream through `csv` and `encoding_rs_io`. Memory depends on the current record and fixed-size buffers, not the total row count. Calling `collect` explicitly retains the rows.
- **XLSX writing** uses `rust_xlsxwriter` constant-memory worksheets. Previous rows spool to temporary files. Row data does not accumulate in RAM; formats, workbook metadata and ZIP finalization have their own allocations. Temporary disk usage grows with the export.
- **XLSX reading** uses `calamine`. The selected sheet and shared strings are loaded into memory. Formula preservation and date formatting also load formula/style ranges. The row iterator avoids a second full copy of the sheet, but it is not a streaming XLSX parser.
- **ODS reading and writing** load the workbook into memory through `calamine` and `spreadsheet-ods`. ODS is included for format parity, not a low-memory guarantee.
- **Downloads** stream CSV directly to an owned `Write` sink. XLSX and ODS finalize into a temporary file and copy to the sink at close, rather than creating another full archive in a `Vec<u8>`. HTTP headers and response lifecycle belong to the caller.

## Behavior mapping

The test names below refer to [tests/spreadsheets.rs](https://github.com/spatie/simple-excel-rust/blob/main/tests/spreadsheets.rs). The shipped CSV/XLSX fixtures come from the PHP package at the revision above, under its MIT license. All listed tests execute real engines; no spreadsheet engine is mocked.

| PHP behavior | Rust equivalent | Verification |
| --- | --- | --- |
| `create`, file extension or explicit type | `create`, `create_with_type`, `FileType` | CSV fixtures, XLSX/ODS roundtrip and encoding tests |
| `getRows`, lazy collection | `get_rows` yielding `Result<Row, Error>` | CSV, spreadsheet and independent-iterator tests |
| Empty/header-only inputs, short/long rows | No data, empty-string padding, extra columns discarded | `handles_empty_and_header_only_files`, `reads_php_csv_fixtures_and_ignores_surplus_columns` |
| Header lookup, replacement and no headers | `get_headers`, `get_original_headers`, `use_headers`, `no_header_row` | `custom_headers_replace_keys_but_preserve_original_headers` |
| `headerOnRow`, sparse spreadsheet columns | `header_on_row`, absolute column positions | `reads_php_xlsx_sheets_sparse_cells_and_offset_headers` |
| Trim headers, snake_case, custom callback | `trim_header_row`, `_with`, `headers_to_snake_case`, `format_headers_using` | Header transformation and character-range tests |
| Trim values, key-aware callback | `trim_values`, `_with`, `format_values_using` | `trims_values_before_key_aware_formatting_without_changing_other_columns` |
| Delimiter, enclosure, encoding | Byte setters and encoding labels | `csv_encoding_delimiter_and_enclosure_can_be_combined` |
| Sheet names, existence, name/index selection | Same names in snake_case; one-based sheet indices | PHP multiple-sheet fixtures and XLSX/ODS roundtrip |
| `skip`, `take` | Same names, repeatable settings | `limits_are_repeatable_and_zero_does_not_yield_a_row` |
| Cached formulas / `keepFormulas` | Cached values / `Value::Formula` | XLSX/ODS roundtrip |
| Dates and preserved formatting | `chrono` values / `preserve_date_time_formatting` | Original PHP date fixture with exact `9/20/2024` output |
| `preserveEmptyRows` | `preserve_empty_rows` within spreadsheet used range | Original PHP empty-row fixture |
| Automatic/explicit headers, raw entity rows | Pair arrays, `add_header`, `Row::new` | Exact CSV output and row-count tests |
| `addRow`, `addRows`, styles | `add_row`, `add_rows`, `_with_style` methods | CSV bytes, XLSX/ODS roundtrip and actual XLSX styles |
| Header and per-cell styles | `set_header_style`, `Row::set_cell_format` for XLSX | `xlsx_header_row_and_cell_styles_and_layout_survive_constant_memory_writes` |
| Column widths and row heights | Native worksheet through `get_current_sheet` | XLSX style/layout test and executed guide |
| Add/name sheets and reset header detection | `add_new_sheet_and_make_it_current`, `name_current_sheet` | XLSX/ODS roundtrip |
| BOM and alternate delimiter | `WriterOptions`, `create_without_bom` | Exact byte-level CSV tests |
| `getPath`, `getWriter`, `getReader` | Path access, exposed native engines | Public API and integration tests |
| `streamDownload`, `toBrowser`, flushing | Owned `Write` sink, no process exit | `downloads_stream_csv_and_finalize_readable_spreadsheets` |
| Close/destructor behavior | `close` reports errors, Drop is best effort | `drop_finalizes_files_but_close_reports_output_failures` |

## Deliberate Rust differences

- Reader configuration consumes and returns `Self`; writers remain mutable because rows are added over time. Recoverable failures return typed `Error` values. A reader opens files when rows or metadata are requested, not at construction. Dropping an iterator closes its file resources.
- Rust iterators are single-pass, not rewindable collections. Call `get_rows()` again for a new traversal. Skip and take do not mutate the reusable reader's settings. Zero means no rows for `take(0)`.
- Ordered rows replace PHP arrays. Unkeyed rows use numeric indexing. Callbacks receive numeric keys as strings. PHP numeric-string key coercion is not reproduced. Missing-key indexing panics like Rust collection indexing; `get` returns `None` instead.
- Dates are timezone-free `NaiveDateTime` values, not PHP objects with an implicit default timezone. Excel's built-in date display IDs 14, 22 and 47 use OpenSpout's documented display choices. Other formatting follows `ssfmt`; locale quirks and unsupported number formats are not guaranteed to be byte-identical to OpenSpout. Malformed date/format values return errors.
- Formulas are not evaluated by either reader. XLSX/ODS formula writes default to a cached zero until a spreadsheet application recalculates. ODS formula strings have their `of:` namespace prefix removed for the common `Value::Formula("=...")` interface. Literal strings never become spreadsheet formula cells automatically.
- Styles and advanced configuration use native Rust engines rather than emulating OpenSpout classes. XLSX `Format` values are not ODS styles. ODS layout and styling are available through the exposed `WorkBook`; supplying an XLSX format to ODS returns an error rather than pretending it worked. CSV ignores styles.
- Native engine coordinates start at zero. Native writes bypass wrapper row counters and header tracking. Constant-memory XLSX disallows editing previous rows. Create a new sheet before the Excel row limit instead of relying on an automatic rollover.
- String trimming handles Unicode characters; ASCII `a..z` ranges are supported. Invalid encoding labels and invalid CSV delimiter/enclosure settings fail explicitly. Decoding uses WHATWG label rules and replaces malformed sequences with U+FFFD.
- Spreadsheet readers use the engine's used range. Empty rows beyond that range and explicit trailing blank cells in unkeyed rows are not preserved. Leading/interior empty spreadsheet cells keep their absolute positions; keyed missing cells become empty strings.
- Finite spreadsheet numbers are required. Integer cells in exported spreadsheets use f64 storage and can lose precision beyond the exact f64 integer range. Write large IDs as text. Cell errors read as `Value::Error`; writing those values exports their displayed text, not an executable error cell.
- Browser response headers, PHP stream wrappers, process exit and engine-specific OpenSpout configuration objects have no standalone Rust equivalent. Use an owned Rust output sink and native engine access. File writes are not atomic, and Drop cannot communicate errors.

README and guide examples run as rustdoc tests, including hidden fixture setup and output assertions. CI executes the suite on Linux, macOS and Windows, plus Rust 1.89 and a security audit. See [testing](testing.md) for commands.
