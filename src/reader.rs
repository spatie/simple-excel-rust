use std::{
    fmt,
    fs::File,
    io::{BufReader, Read},
    path::{Path, PathBuf},
    sync::Arc,
};

use calamine::{Data, DataType, Range, Reader, Sheets, StyleRange};
use encoding_rs::Encoding;
use encoding_rs_io::DecodeReaderBytesBuilder;

use crate::{Error, FileType, Row, Value};

type HeaderFormatter = Arc<dyn Fn(&str) -> String + Send + Sync>;
type ValueFormatter = Arc<dyn Fn(Value, &str) -> Value + Send + Sync>;

/// A reusable file reader. Each `get_rows` call opens an independent iterator.
///
/// CSV is streamed. Calamine loads the selected XLSX sheet (and shared strings)
/// or the ODS workbook into memory. Iterator adapters do not change that cost.
#[derive(Clone)]
#[must_use]
pub struct SimpleExcelReader {
    path: PathBuf,
    file_type: FileType,
    delimiter: u8,
    enclosure: u8,
    encoding: &'static Encoding,
    sheet: usize,
    sheet_name: Option<String>,
    header: bool,
    header_row: usize,
    headers: Option<Vec<String>>,
    trim_headers: Option<String>,
    snake_case: bool,
    header_formatter: Option<HeaderFormatter>,
    trim_values: Option<String>,
    value_formatter: Option<ValueFormatter>,
    formulas: bool,
    format_dates: bool,
    empty_rows: bool,
    skip: usize,
    take: Option<usize>,
}

impl fmt::Debug for SimpleExcelReader {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SimpleExcelReader")
            .field("path", &self.path)
            .field("file_type", &self.file_type)
            .finish_non_exhaustive()
    }
}

const TRIM: &str = " \t\n\r\0\x0b";

impl SimpleExcelReader {
    /// Creates a reader, detecting CSV, XLSX or ODS from the extension.
    /// The file is opened when rows, headers or sheet metadata are requested.
    pub fn create(path: impl AsRef<Path>) -> Result<Self, Error> {
        Self::create_with_type(&path, FileType::from_path(&path)?)
    }

    /// Creates a reader with an explicit type, irrespective of extension.
    pub fn create_with_type(path: impl AsRef<Path>, file_type: FileType) -> Result<Self, Error> {
        Ok(Self {
            path: path.as_ref().into(),
            file_type,
            delimiter: b',',
            enclosure: b'"',
            encoding: encoding_rs::UTF_8,
            sheet: 1,
            sheet_name: None,
            header: true,
            header_row: 0,
            headers: None,
            trim_headers: None,
            snake_case: false,
            header_formatter: None,
            trim_values: None,
            value_formatter: None,
            formulas: false,
            format_dates: false,
            empty_rows: false,
            skip: 0,
            take: None,
        })
    }

    /// Returns the input path.
    pub fn get_path(&self) -> &Path {
        &self.path
    }
    /// Uses a zero-based nonempty row as the header; preceding rows are discarded.
    pub fn header_on_row(mut self, row: usize) -> Self {
        self.header_row = row;
        self
    }
    /// Treats the first row as data. Rows have numeric indices unless custom headers are supplied.
    pub fn no_header_row(mut self) -> Self {
        self.header = false;
        self
    }
    /// Overrides header names, without changing whether a file header is consumed.
    pub fn use_headers(mut self, headers: impl IntoIterator<Item = impl Into<String>>) -> Self {
        let headers: Vec<_> = headers.into_iter().map(Into::into).collect();
        self.headers = (!headers.is_empty()).then_some(headers);
        self
    }
    /// Sets the CSV delimiter. It must be a non-NUL ASCII byte other than CR or LF.
    pub fn use_delimiter(mut self, delimiter: u8) -> Result<Self, Error> {
        validate_csv_byte(delimiter)?;
        self.delimiter = delimiter;
        Ok(self)
    }
    /// Sets the CSV enclosure, with the same byte restrictions as the delimiter.
    pub fn use_field_enclosure(mut self, enclosure: u8) -> Result<Self, Error> {
        validate_csv_byte(enclosure)?;
        self.enclosure = enclosure;
        Ok(self)
    }
    /// Sets a WHATWG encoding label, such as UTF-8, SHIFT_JIS or Windows-1252.
    /// Decoding is streaming; malformed sequences are replaced by U+FFFD.
    pub fn use_encoding(mut self, label: &str) -> Result<Self, Error> {
        self.encoding = Encoding::for_label(label.as_bytes())
            .ok_or_else(|| Error::InvalidInput(format!("Unknown encoding: {label}")))?;
        Ok(self)
    }
    /// Trims PHP's default whitespace characters from headers.
    pub fn trim_header_row(mut self) -> Self {
        self.trim_headers = Some(TRIM.into());
        self
    }
    /// Trims a PHP-style character set, including ASCII ranges such as `a..z`.
    pub fn trim_header_row_with(mut self, characters: &str) -> Self {
        self.trim_headers = Some(expand_ranges(characters));
        self
    }
    /// Converts headers to snake_case using the PHP package's letter/digit boundaries.
    pub fn headers_to_snake_case(mut self) -> Self {
        self.snake_case = true;
        self
    }
    /// Formats headers after trimming and snake_case conversion.
    pub fn format_headers_using(
        mut self,
        callback: impl Fn(&str) -> String + Send + Sync + 'static,
    ) -> Self {
        self.header_formatter = Some(Arc::new(callback));
        self
    }
    /// Trims PHP's default whitespace characters from string values only.
    pub fn trim_values(mut self) -> Self {
        self.trim_values = Some(TRIM.into());
        self
    }
    /// Trims a PHP-style character set from string values only.
    pub fn trim_values_with(mut self, characters: &str) -> Self {
        self.trim_values = Some(expand_ranges(characters));
        self
    }
    /// Formats each value after trimming, receiving its header or numeric index as a string.
    pub fn format_values_using(
        mut self,
        callback: impl Fn(Value, &str) -> Value + Send + Sync + 'static,
    ) -> Self {
        self.value_formatter = Some(Arc::new(callback));
        self
    }
    /// Returns formula strings instead of cached results. No formula evaluator is provided.
    pub fn keep_formulas(mut self) -> Self {
        self.formulas = true;
        self
    }
    /// Returns XLSX dates and times as strings using their stored number formats.
    pub fn preserve_date_time_formatting(mut self) -> Self {
        self.format_dates = true;
        self
    }
    /// Preserves empty XLSX/ODS rows within the used range, including leading rows.
    /// CSV blank lines remain ignored, as in the PHP package.
    pub fn preserve_empty_rows(mut self) -> Self {
        self.empty_rows = true;
        self
    }
    /// Selects a sheet by its one-based number.
    pub fn from_sheet(mut self, number: usize) -> Self {
        self.sheet = number;
        self.sheet_name = None;
        self
    }
    /// Selects a sheet by its exact, case-sensitive name.
    pub fn from_sheet_name(mut self, name: impl Into<String>) -> Self {
        self.sheet_name = Some(name.into());
        self
    }
    /// Skips data rows after consuming the header.
    pub fn skip(mut self, count: usize) -> Self {
        self.skip = count;
        self
    }
    /// Limits the number of data rows returned, including zero.
    pub fn take(mut self, count: usize) -> Self {
        self.take = Some(count);
        self
    }

    /// Opens the underlying spreadsheet engine for advanced configuration.
    /// Use `get_csv_reader` for CSV. Changes to the returned engine do not alter this reader.
    pub fn get_reader(&self) -> Result<Sheets<BufReader<File>>, Error> {
        match self.file_type {
            FileType::Xlsx => Ok(Sheets::Xlsx(calamine::open_workbook(&self.path)?)),
            FileType::Ods => Ok(Sheets::Ods(
                calamine::open_workbook(&self.path).map_err(calamine::Error::Ods)?,
            )),
            FileType::Csv => Err(Error::UnsupportedOperation("Use get_csv_reader for CSV")),
        }
    }

    /// Opens the configured streaming CSV engine.
    pub fn get_csv_reader(&self) -> Result<csv::Reader<Box<dyn Read>>, Error> {
        if self.file_type != FileType::Csv {
            return Err(Error::UnsupportedOperation("This file is not CSV"));
        }
        validate_csv_pair(self.delimiter, self.enclosure)?;
        let decoder = DecodeReaderBytesBuilder::new()
            .encoding(Some(self.encoding))
            .build(File::open(&self.path)?);
        Ok(csv::ReaderBuilder::new()
            .has_headers(false)
            .flexible(true)
            .delimiter(self.delimiter)
            .quote(self.enclosure)
            .from_reader(Box::new(decoder)))
    }

    /// Returns all sheet names. CSV has a single sheet named `Sheet1`.
    pub fn get_sheet_names(&self) -> Result<Vec<String>, Error> {
        if self.file_type == FileType::Csv {
            File::open(&self.path)?;
            Ok(vec!["Sheet1".into()])
        } else {
            Ok(self.get_reader()?.sheet_names())
        }
    }

    /// Checks for an exact sheet name.
    pub fn has_sheet(&self, name: &str) -> Result<bool, Error> {
        Ok(self.get_sheet_names()?.iter().any(|n| n == name))
    }

    fn raw_rows(&self) -> Result<RawRows, Error> {
        if self.file_type == FileType::Csv {
            if self.sheet != 1 || self.sheet_name.as_deref().is_some_and(|s| s != "Sheet1") {
                return Err(Error::SheetNotFound(
                    self.sheet_name
                        .clone()
                        .unwrap_or_else(|| self.sheet.to_string()),
                ));
            }
            return Ok(RawRows::Csv(Box::new(self.get_csv_reader()?)));
        }
        let mut reader = self.get_reader()?;
        let names = reader.sheet_names();
        let name = if let Some(name) = &self.sheet_name {
            names.iter().find(|n| *n == name)
        } else {
            self.sheet.checked_sub(1).and_then(|i| names.get(i))
        }
        .ok_or_else(|| {
            Error::SheetNotFound(
                self.sheet_name
                    .clone()
                    .unwrap_or_else(|| self.sheet.to_string()),
            )
        })?;
        let data = reader.worksheet_range(name)?;
        let formulas = self
            .formulas
            .then(|| reader.worksheet_formula(name))
            .transpose()?;
        let styles = if self.format_dates {
            match &mut reader {
                Sheets::Xlsx(xlsx) => Some(xlsx.worksheet_style(name)?),
                _ => None,
            }
        } else {
            None
        };
        let start_row = if self.empty_rows {
            0
        } else {
            data.start().map_or(0, |p| p.0)
        };
        let end = data.end();
        Ok(RawRows::Sheet(Box::new(SheetRows {
            data,
            formulas,
            styles,
            next_row: start_row,
            end,
            empty_rows: self.empty_rows,
        })))
    }

    fn read_headers(&self, rows: &mut RawRows) -> Result<Option<Vec<String>>, Error> {
        if !self.header {
            return Ok(None);
        }
        for _ in 0..self.header_row {
            if rows.next().transpose()?.is_none() {
                return Ok(None);
            }
        }
        Ok(rows.next().transpose()?.map(|values| {
            values
                .into_iter()
                .map(|value| {
                    let mut header = value.to_string();
                    if let Some(chars) = &self.trim_headers {
                        header = header.trim_matches(|c| chars.contains(c)).into();
                    }
                    if self.snake_case {
                        header = snake_case(&header);
                    }
                    if let Some(callback) = &self.header_formatter {
                        header = callback(&header);
                    }
                    header
                })
                .collect()
        }))
    }

    /// Returns processed header names, or custom headers when configured.
    pub fn get_headers(&self) -> Result<Option<Vec<String>>, Error> {
        if let Some(headers) = &self.headers {
            return Ok(Some(headers.clone()));
        }
        self.get_original_headers()
    }

    /// Returns the processed file headers before custom replacements, or `None` without a header.
    pub fn get_original_headers(&self) -> Result<Option<Vec<String>>, Error> {
        self.read_headers(&mut self.raw_rows()?)
    }

    /// Opens an independent, fallible row iterator. Dropping it closes the input.
    pub fn get_rows(&self) -> Result<Rows, Error> {
        let mut raw = self.raw_rows()?;
        let original = self.read_headers(&mut raw)?;
        for _ in 0..self.skip {
            if raw.next().transpose()?.is_none() {
                break;
            }
        }
        Ok(Rows {
            raw,
            headers: self.headers.clone().or(original).filter(|h| !h.is_empty()),
            remaining: self.take,
            trim: self.trim_values.clone(),
            formatter: self.value_formatter.clone(),
            done: false,
        })
    }
}

pub(crate) fn validate_csv_byte(byte: u8) -> Result<(), Error> {
    if byte == 0 || !byte.is_ascii() || byte == b'\r' || byte == b'\n' {
        return Err(Error::InvalidInput(
            "CSV delimiter/enclosure must be a non-NUL ASCII byte other than CR or LF".into(),
        ));
    }
    Ok(())
}

pub(crate) fn validate_csv_pair(delimiter: u8, enclosure: u8) -> Result<(), Error> {
    validate_csv_byte(delimiter)?;
    validate_csv_byte(enclosure)?;
    if delimiter == enclosure {
        return Err(Error::InvalidInput(
            "CSV delimiter and enclosure must differ".into(),
        ));
    }
    Ok(())
}

fn expand_ranges(set: &str) -> String {
    let chars: Vec<char> = set.chars().collect();
    let mut result = String::new();
    let mut i = 0;
    while i < chars.len() {
        if i + 3 < chars.len()
            && chars[i + 1] == '.'
            && chars[i + 2] == '.'
            && chars[i].is_ascii()
            && chars[i + 3].is_ascii()
            && chars[i] <= chars[i + 3]
        {
            result.extend((chars[i] as u8..=chars[i + 3] as u8).map(char::from));
            i += 4;
        } else {
            result.push(chars[i]);
            i += 1;
        }
    }
    result
}

fn snake_case(header: &str) -> String {
    let mut result = String::new();
    let mut previous: Option<char> = None;
    for c in header.trim_matches(|c| TRIM.contains(c)).chars() {
        if previous.is_some_and(|p| {
            (p.is_ascii_digit() && c.is_ascii_alphabetic())
                || (p.is_ascii_alphabetic() && c.is_ascii_digit())
                || (p.is_ascii_lowercase() && c.is_ascii_uppercase())
        }) {
            result.push('_');
        }
        result.push(if c == ' ' {
            '_'
        } else {
            c.to_ascii_lowercase()
        });
        previous = Some(c);
    }
    result
}

enum RawRows {
    Csv(Box<csv::Reader<Box<dyn Read>>>),
    Sheet(Box<SheetRows>),
}

impl Iterator for RawRows {
    type Item = Result<Vec<Value>, Error>;
    fn next(&mut self) -> Option<Self::Item> {
        match self {
            Self::Csv(reader) => {
                let mut record = csv::StringRecord::new();
                match reader.read_record(&mut record) {
                    Ok(true) => Some(Ok(record.iter().map(Value::from).collect())),
                    Ok(false) => None,
                    Err(error) => Some(Err(error.into())),
                }
            }
            Self::Sheet(sheet) => sheet.next(),
        }
    }
}

struct SheetRows {
    data: Range<Data>,
    formulas: Option<Range<String>>,
    styles: Option<StyleRange>,
    next_row: u32,
    end: Option<(u32, u32)>,
    empty_rows: bool,
}

impl Iterator for SheetRows {
    type Item = Result<Vec<Value>, Error>;
    fn next(&mut self) -> Option<Self::Item> {
        let (end_row, end_col) = self.end?;
        while self.next_row <= end_row {
            let row = self.next_row;
            self.next_row += 1;
            let mut values = Vec::new();
            let mut used = false;
            for col in 0..=end_col {
                let pos = (row, col);
                if let Some(formula) = self
                    .formulas
                    .as_ref()
                    .and_then(|range| range.get_value(pos))
                    .filter(|s| !s.is_empty())
                {
                    used = true;
                    let formula = formula
                        .strip_prefix("of:")
                        .or_else(|| formula.strip_prefix("oooc:"))
                        .unwrap_or(formula);
                    values.push(Value::Formula(if formula.starts_with('=') {
                        formula.to_owned()
                    } else {
                        format!("={formula}")
                    }));
                    continue;
                }
                let data = self.data.get_value(pos).unwrap_or(&Data::Empty);
                used |= !matches!(data, Data::Empty);
                match convert(
                    data,
                    self.styles.as_ref().and_then(|styles| styles.get(pos)),
                ) {
                    Ok(value) => values.push(value),
                    Err(error) => return Some(Err(error)),
                }
            }
            // OpenSpout does not pad unkeyed rows to the widest row in the sheet.
            while values.last() == Some(&Value::Empty) {
                values.pop();
            }
            if used || self.empty_rows {
                for value in &mut values {
                    if *value == Value::Empty {
                        *value = Value::String(String::new());
                    }
                }
                return Some(Ok(values));
            }
        }
        None
    }
}

fn convert(data: &Data, style: Option<&calamine::Style>) -> Result<Value, Error> {
    Ok(match data {
        Data::Empty => Value::Empty,
        Data::String(s) => Value::String(s.clone()),
        Data::Int(n) => Value::Int(*n),
        Data::Float(n) => Value::Float(*n),
        Data::Bool(b) => Value::Bool(*b),
        Data::Error(e) => Value::Error(e.to_string()),
        Data::DateTime(d) if style.is_some() => {
            let format = style
                .and_then(|s| s.number_format.as_ref())
                .ok_or_else(|| Error::InvalidInput("Date cell has no number format".into()))?;
            let (year, month, day, ..) = d.to_ymd_hms_milli();
            let serial = if d.is_duration() {
                d.as_f64()
            } else {
                ssfmt::date_serial::date_to_serial(
                    i32::from(year),
                    u32::from(month),
                    u32::from(day),
                    ssfmt::DateSystem::Date1900,
                ) + d.as_f64().rem_euclid(1.0)
            };
            // OpenSpout deliberately differs from ECMA's built-in date displays.
            let code = match format.format_id {
                Some(14) => "m/d/yyyy",
                Some(22) => "m/d/yyyy h:mm",
                Some(47) => "mm:ss.0",
                _ => &format.format_code,
            };
            Value::String(
                ssfmt::format_default(serial, code)
                    .map_err(|e| Error::InvalidInput(format!("Date format: {e}")))?,
            )
        }
        Data::DateTime(d) if d.is_duration() => Value::Duration(
            d.as_duration()
                .ok_or_else(|| Error::InvalidInput("Invalid spreadsheet duration".into()))?,
        ),
        Data::DateTime(_) | Data::DateTimeIso(_) => Value::DateTime(
            data.as_datetime()
                .ok_or_else(|| Error::InvalidInput("Invalid spreadsheet date/time".into()))?,
        ),
        Data::DurationIso(s) => {
            // Calamine's time-of-day parser rejects valid ODS durations over 24 hours.
            let duration = s
                .parse::<jiff::SignedDuration>()
                .map_err(|_| Error::InvalidInput(format!("Invalid spreadsheet duration: {s}")))?;
            Value::Duration(
                chrono::Duration::try_seconds(duration.as_secs())
                    .and_then(|seconds| {
                        seconds.checked_add(&chrono::Duration::nanoseconds(i64::from(
                            duration.subsec_nanos(),
                        )))
                    })
                    .ok_or_else(|| Error::InvalidInput(format!("Duration out of range: {s}")))?,
            )
        }
    })
}

/// An independent fallible row iterator. Errors terminate iteration.
pub struct Rows {
    raw: RawRows,
    headers: Option<Vec<String>>,
    remaining: Option<usize>,
    trim: Option<String>,
    formatter: Option<ValueFormatter>,
    done: bool,
}

impl fmt::Debug for Rows {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Rows")
            .field("headers", &self.headers)
            .field("remaining", &self.remaining)
            .finish_non_exhaustive()
    }
}

impl Iterator for Rows {
    type Item = Result<Row, Error>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done || self.remaining == Some(0) {
            return None;
        }
        let values = match self.raw.next()? {
            Ok(values) => values,
            Err(error) => {
                self.done = true;
                return Some(Err(error));
            }
        };
        if let Some(n) = &mut self.remaining {
            *n -= 1;
        }
        let mut row = if let Some(headers) = &self.headers {
            let mut values = values;
            values.resize(headers.len(), Value::String(String::new()));
            Row::from_pairs(headers.iter().cloned().zip(values))
        } else {
            Row::new(values)
        };
        for (index, value) in row.values.iter_mut().enumerate() {
            if let (Some(chars), Value::String(s)) = (&self.trim, &mut *value) {
                *s = s.trim_matches(|c| chars.contains(c)).into();
            }
            if let Some(callback) = &self.formatter {
                let key = row
                    .headers
                    .as_ref()
                    .map_or_else(|| index.to_string(), |h| h[index].clone());
                *value = callback(std::mem::take(value), &key);
            }
        }
        Some(Ok(row))
    }
}

impl std::iter::FusedIterator for Rows {}
