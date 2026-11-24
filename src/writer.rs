use std::{
    fmt,
    fs::File,
    io::{self, Seek, Write},
    path::{Path, PathBuf},
};

use rust_xlsxwriter::{Format, Formula, Workbook, Worksheet};
use spreadsheet_ods::{Sheet, WorkBook};
use tempfile::NamedTempFile;

use crate::{Error, FileType, Row, Value, reader::validate_csv_pair};

/// Writer creation settings. CSV is UTF-8 with a BOM, comma delimiter and double quotes by default.
#[derive(Debug, Clone)]
pub struct WriterOptions {
    /// Overrides extension-based file type detection.
    pub file_type: Option<FileType>,
    /// CSV delimiter, as a non-NUL ASCII byte other than CR or LF.
    pub delimiter: u8,
    /// CSV enclosure, distinct from the delimiter.
    pub enclosure: u8,
    /// Writes the UTF-8 BOM at creation time for CSV only.
    pub add_bom: bool,
}

impl Default for WriterOptions {
    fn default() -> Self {
        Self {
            file_type: None,
            delimiter: b',',
            enclosure: b'"',
            add_bom: true,
        }
    }
}

/// The underlying engines, exposed for advanced configuration.
///
/// Writing rows or adding sheets directly bypasses the wrapper's row counters
/// and header handling. Prefer `get_current_sheet` for XLSX layout changes.
pub enum WriterEngine {
    /// A streaming CSV writer.
    Csv(Box<csv::Writer<Box<dyn Write>>>),
    /// An XLSX workbook whose sheets use constant-memory mode.
    Xlsx(Box<Workbook>),
    /// An in-memory ODS workbook.
    Ods(Box<WorkBook>),
}

impl fmt::Debug for WriterEngine {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::Csv(_) => "WriterEngine::Csv",
            Self::Xlsx(_) => "WriterEngine::Xlsx",
            Self::Ods(_) => "WriterEngine::Ods",
        })
    }
}

/// A row-by-row CSV, XLSX or ODS writer.
///
/// Call `close` to observe finalization errors. Drop attempts best-effort cleanup
/// and finalization but cannot report failures. XLSX cells spool to temporary
/// files; ODS holds the workbook in memory. Neither spreadsheet format is a live
/// HTTP byte stream: downloads finalize to a temporary file before copying out.
pub struct SimpleExcelWriter {
    path: PathBuf,
    engine: Option<WriterEngine>,
    output: Option<Box<dyn Write>>,
    temporary: Option<NamedTempFile>,
    current_sheet: usize,
    next_row: u32,
    number_of_rows: u64,
    process_header: bool,
    first_row: bool,
    header_style: Option<Format>,
}

impl fmt::Debug for SimpleExcelWriter {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SimpleExcelWriter")
            .field("path", &self.path)
            .field("engine", &self.engine)
            .field("number_of_rows", &self.number_of_rows)
            .finish_non_exhaustive()
    }
}

impl SimpleExcelWriter {
    /// Creates a file writer with default options, detecting type from the extension.
    pub fn create(path: impl AsRef<Path>) -> Result<Self, Error> {
        Self::create_with_options(path, WriterOptions::default())
    }

    /// Creates a file writer without a CSV BOM.
    pub fn create_without_bom(path: impl AsRef<Path>) -> Result<Self, Error> {
        Self::create_with_options(
            path,
            WriterOptions {
                add_bom: false,
                ..WriterOptions::default()
            },
        )
    }

    /// Creates a file writer with an explicit type.
    pub fn create_with_type(path: impl AsRef<Path>, file_type: FileType) -> Result<Self, Error> {
        Self::create_with_options(
            path,
            WriterOptions {
                file_type: Some(file_type),
                ..WriterOptions::default()
            },
        )
    }

    /// Configures delimiter, enclosure, BOM and/or type before the output is opened.
    pub fn create_with_options(
        path: impl AsRef<Path>,
        options: WriterOptions,
    ) -> Result<Self, Error> {
        Self::open(path.as_ref(), options, None)
    }

    /// Writes a download to a caller-provided output, such as an HTTP body sink or stdout.
    /// The caller sets HTTP headers. CSV streams directly; XLSX/ODS spool until `close`.
    pub fn stream_download(
        name: impl AsRef<Path>,
        output: impl Write + 'static,
    ) -> Result<Self, Error> {
        Self::stream_download_with_options(name, output, WriterOptions::default())
    }

    /// Creates a download with explicit CSV and file type settings.
    pub fn stream_download_with_options(
        name: impl AsRef<Path>,
        output: impl Write + 'static,
        options: WriterOptions,
    ) -> Result<Self, Error> {
        Self::open(name.as_ref(), options, Some(Box::new(output)))
    }

    fn open(
        path: &Path,
        options: WriterOptions,
        mut output: Option<Box<dyn Write>>,
    ) -> Result<Self, Error> {
        let file_type = options
            .file_type
            .map_or_else(|| FileType::from_path(path), Ok)?;
        let mut temporary = None;
        let engine = match file_type {
            FileType::Csv => {
                validate_csv_pair(options.delimiter, options.enclosure)?;
                let mut sink: Box<dyn Write> = match output.take() {
                    Some(sink) => sink,
                    None => Box::new(File::create(path)?),
                };
                if options.add_bom {
                    sink.write_all(b"\xef\xbb\xbf")?;
                }
                WriterEngine::Csv(Box::new(
                    csv::WriterBuilder::new()
                        .delimiter(options.delimiter)
                        .quote(options.enclosure)
                        .flexible(true)
                        .from_writer(sink),
                ))
            }
            FileType::Xlsx => {
                if output.is_some() {
                    temporary = Some(NamedTempFile::new()?);
                }
                let mut workbook = Workbook::new();
                workbook.add_worksheet_with_constant_memory()?;
                WriterEngine::Xlsx(Box::new(workbook))
            }
            FileType::Ods => {
                if output.is_some() {
                    temporary = Some(NamedTempFile::new()?);
                }
                let mut workbook = WorkBook::new_empty();
                workbook.push_sheet(Sheet::new("Sheet1"));
                WriterEngine::Ods(Box::new(workbook))
            }
        };
        Ok(Self {
            path: path.into(),
            engine: Some(engine),
            output,
            temporary,
            current_sheet: 0,
            next_row: 0,
            number_of_rows: 0,
            process_header: true,
            first_row: true,
            header_style: None,
        })
    }

    /// Returns the output path or download name.
    pub fn get_path(&self) -> &Path {
        &self.path
    }
    /// Returns rows written, including automatic and explicit headers across all sheets.
    pub fn get_number_of_rows(&self) -> u64 {
        self.number_of_rows
    }
    /// Disables automatic header generation for subsequent rows.
    pub fn no_header_row(&mut self) -> &mut Self {
        self.process_header = false;
        self
    }
    /// Sets an XLSX header format. For ODS, use its native cell styles through `get_writer`.
    pub fn set_header_style(&mut self, style: Format) -> Result<&mut Self, Error> {
        if matches!(self.engine, Some(WriterEngine::Ods(_))) {
            return Err(Error::UnsupportedOperation(
                "Use native ODS styles through get_writer",
            ));
        }
        self.header_style = Some(style);
        Ok(self)
    }

    /// Returns the native engine for advanced configuration, or an error after closing.
    pub fn get_writer(&mut self) -> Result<&mut WriterEngine, Error> {
        self.engine.as_mut().ok_or(Error::Closed)
    }

    /// Returns the current XLSX worksheet for layout and default formatting.
    /// Configure widths and row heights before writing rows in constant-memory mode.
    pub fn get_current_sheet(&mut self) -> Result<&mut Worksheet, Error> {
        match self.engine.as_mut().ok_or(Error::Closed)? {
            WriterEngine::Xlsx(book) => Ok(book.worksheet_from_index(self.current_sheet)?),
            _ => Err(Error::UnsupportedOperation(
                "get_current_sheet requires XLSX; use get_writer for other formats",
            )),
        }
    }

    /// Adds a row. The first keyed row on each sheet generates an automatic header.
    /// Input is written in iteration order; later keys do not reorder columns.
    pub fn add_row(&mut self, row: impl Into<Row>) -> Result<&mut Self, Error> {
        self.add_row_internal(row.into(), None)
    }

    /// Adds a row with an XLSX format. Per-cell formats override the row format.
    /// CSV ignores formatting. ODS uses its native styles through `get_writer`.
    pub fn add_row_with_style(
        &mut self,
        row: impl Into<Row>,
        style: &Format,
    ) -> Result<&mut Self, Error> {
        self.add_row_internal(row.into(), Some(style))
    }

    fn add_row_internal(&mut self, row: Row, style: Option<&Format>) -> Result<&mut Self, Error> {
        self.validate_row(&row, style)?;
        if self.first_row
            && self.process_header
            && let Some(headers) = &row.headers
        {
            self.add_header(headers)?;
        }
        self.write_row(&row, style)?;
        self.first_row = false;
        Ok(self)
    }

    /// Writes rows incrementally from any iterator, without collecting it.
    pub fn add_rows<R: Into<Row>>(
        &mut self,
        rows: impl IntoIterator<Item = R>,
    ) -> Result<&mut Self, Error> {
        for row in rows {
            self.add_row(row)?;
        }
        Ok(self)
    }

    /// Writes multiple rows using one XLSX style.
    pub fn add_rows_with_style<R: Into<Row>>(
        &mut self,
        rows: impl IntoIterator<Item = R>,
        style: &Format,
    ) -> Result<&mut Self, Error> {
        for row in rows {
            self.add_row_with_style(row, style)?;
        }
        Ok(self)
    }

    /// Writes an explicit header and disables automatic detection for the current sheet.
    pub fn add_header(
        &mut self,
        headers: impl IntoIterator<Item = impl AsRef<str>>,
    ) -> Result<&mut Self, Error> {
        let row = Row::new(headers.into_iter().map(|h| Value::from(h.as_ref())));
        let style = self.header_style.clone();
        self.validate_row(&row, style.as_ref())?;
        self.write_row(&row, style.as_ref())?;
        self.first_row = false;
        Ok(self)
    }

    fn validate_row(&self, row: &Row, style: Option<&Format>) -> Result<(), Error> {
        let engine = self.engine.as_ref().ok_or(Error::Closed)?;
        if matches!(engine, WriterEngine::Xlsx(_))
            && (row.values.len() > 16384 || self.next_row >= 1048576)
        {
            return Err(Error::InvalidInput(
                "XLSX row or column limit exceeded; create another sheet".into(),
            ));
        }
        if matches!(engine, WriterEngine::Ods(_))
            && (style.is_some() || row.formats.iter().any(Option::is_some))
        {
            return Err(Error::UnsupportedOperation(
                "Use native ODS styles through get_writer",
            ));
        }
        if row
            .values
            .iter()
            .any(|v| matches!(v, Value::Float(n) if !n.is_finite()))
        {
            return Err(Error::InvalidInput(
                "Spreadsheet numbers must be finite".into(),
            ));
        }
        Ok(())
    }

    fn write_row(&mut self, row: &Row, style: Option<&Format>) -> Result<(), Error> {
        match self.engine.as_mut().ok_or(Error::Closed)? {
            WriterEngine::Csv(csv) => {
                csv.write_record(row.values.iter().map(ToString::to_string))?
            }
            WriterEngine::Xlsx(book) => {
                let sheet = book.worksheet_from_index(self.current_sheet)?;
                let default = Format::new();
                let date = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");
                let duration = Format::new().set_num_format("[hh]:mm:ss");
                for (col, value) in row.values.iter().enumerate() {
                    let format = row
                        .formats
                        .get(col)
                        .and_then(Option::as_ref)
                        .or(style)
                        .unwrap_or(match value {
                            Value::DateTime(_) => &date,
                            Value::Duration(_) => &duration,
                            _ => &default,
                        });
                    let col = col as u16;
                    match value {
                        Value::Empty => {
                            sheet.write_blank(self.next_row, col, format)?;
                        }
                        Value::String(s) | Value::Error(s) => {
                            sheet.write_string_with_format(self.next_row, col, s, format)?;
                        }
                        Value::Int(n) => {
                            sheet.write_number_with_format(
                                self.next_row,
                                col,
                                *n as f64,
                                format,
                            )?;
                        }
                        Value::Float(n) => {
                            sheet.write_number_with_format(self.next_row, col, *n, format)?;
                        }
                        Value::Bool(b) => {
                            sheet.write_boolean_with_format(self.next_row, col, *b, format)?;
                        }
                        Value::DateTime(d) => {
                            sheet.write_datetime_with_format(self.next_row, col, d, format)?;
                        }
                        Value::Duration(d) => {
                            sheet.write_number_with_format(
                                self.next_row,
                                col,
                                d.num_milliseconds() as f64 / 86_400_000.0,
                                format,
                            )?;
                        }
                        Value::Formula(s) => {
                            sheet.write_formula_with_format(
                                self.next_row,
                                col,
                                Formula::new(s),
                                format,
                            )?;
                        }
                    }
                }
            }
            WriterEngine::Ods(book) => {
                let sheet = book.sheet_mut(self.current_sheet);
                for (col, value) in row.values.iter().enumerate() {
                    let col = u32::try_from(col)
                        .map_err(|_| Error::InvalidInput("Too many ODS columns".into()))?;
                    let value = match value {
                        Value::Empty => spreadsheet_ods::Value::Empty,
                        Value::String(s) | Value::Error(s) => s.clone().into(),
                        Value::Int(n) => (*n as f64).into(),
                        Value::Float(n) => (*n).into(),
                        Value::Bool(b) => (*b).into(),
                        Value::DateTime(d) => (*d).into(),
                        Value::Duration(d) => (*d).into(),
                        Value::Formula(s) => {
                            sheet.set_formula(
                                self.next_row,
                                col,
                                format!("of:={}", s.trim_start_matches('=')),
                            );
                            0.0.into()
                        }
                    };
                    sheet.set_value(self.next_row, col, value);
                }
            }
        }
        self.next_row = self
            .next_row
            .checked_add(1)
            .ok_or_else(|| Error::InvalidInput("Too many rows".into()))?;
        self.number_of_rows += 1;
        Ok(())
    }

    /// Names the current spreadsheet sheet. CSV does not support sheets.
    pub fn name_current_sheet(&mut self, name: &str) -> Result<&mut Self, Error> {
        match self.engine.as_mut().ok_or(Error::Closed)? {
            WriterEngine::Csv(_) => {
                return Err(Error::UnsupportedOperation(
                    "CSV does not support sheet naming",
                ));
            }
            WriterEngine::Xlsx(book) => {
                book.worksheet_from_index(self.current_sheet)?
                    .set_name(name)?;
            }
            WriterEngine::Ods(book) => {
                if name.is_empty()
                    || book
                        .iter_sheets()
                        .enumerate()
                        .any(|(i, s)| i != self.current_sheet && s.name() == name)
                {
                    return Err(Error::InvalidInput(
                        "Sheet names must be nonempty and unique".into(),
                    ));
                }
                book.sheet_mut(self.current_sheet).set_name(name);
            }
        }
        Ok(self)
    }

    /// Adds a sheet, optionally named, and resets automatic header detection.
    pub fn add_new_sheet_and_make_it_current(
        &mut self,
        name: Option<&str>,
    ) -> Result<&mut Self, Error> {
        match self.engine.as_mut().ok_or(Error::Closed)? {
            WriterEngine::Csv(_) => {
                return Err(Error::UnsupportedOperation(
                    "CSV does not support multiple sheets",
                ));
            }
            WriterEngine::Xlsx(book) => {
                let mut sheet = book.new_worksheet_with_constant_memory()?;
                if let Some(name) = name {
                    sheet.set_name(name)?;
                }
                book.push_worksheet(sheet);
                self.current_sheet = book.worksheets().len() - 1;
            }
            WriterEngine::Ods(book) => {
                let name =
                    name.map_or_else(|| format!("Sheet{}", book.num_sheets() + 1), str::to_owned);
                if name.is_empty() || book.sheet_idx(&name).is_some() {
                    return Err(Error::InvalidInput(
                        "Sheet names must be nonempty and unique".into(),
                    ));
                }
                book.push_sheet(Sheet::new(name));
                self.current_sheet = book.num_sheets() - 1;
            }
        }
        self.next_row = 0;
        self.first_row = true;
        Ok(self)
    }

    /// Flushes CSV bytes to the output. Spreadsheet data is finalized only by `close`.
    pub fn flush(&mut self) -> Result<(), Error> {
        if let WriterEngine::Csv(writer) = self.engine.as_mut().ok_or(Error::Closed)? {
            writer.flush()?;
        }
        Ok(())
    }

    /// Finalizes output and releases resources. Repeated calls succeed.
    /// A failure closes the writer too; partially written output may remain.
    pub fn close(&mut self) -> Result<(), Error> {
        let Some(mut engine) = self.engine.take() else {
            return Ok(());
        };
        let path = self
            .temporary
            .as_ref()
            .map_or(self.path.as_path(), |f| f.path());
        match &mut engine {
            WriterEngine::Csv(writer) => writer.flush()?,
            WriterEngine::Xlsx(book) => book.save(path)?,
            WriterEngine::Ods(book) => spreadsheet_ods::write_ods(book, path)?,
        }
        if let Some(output) = &mut self.output {
            if let Some(file) = &mut self.temporary {
                file.rewind()?;
                io::copy(file, output)?;
            }
            output.flush()?;
        }
        self.temporary = None;
        self.output = None;
        Ok(())
    }

    /// Finalizes a download without exiting the process. The caller controls HTTP headers.
    pub fn to_browser(&mut self) -> Result<(), Error> {
        self.close()
    }
}

impl Drop for SimpleExcelWriter {
    fn drop(&mut self) {
        let _ = self.close();
    }
}
