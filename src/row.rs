use std::{fmt, ops::Index};

use chrono::{Duration, NaiveDate, NaiveDateTime};
use rust_xlsxwriter::Format;

/// A cell value. Text is always written literally; only `Formula` executes in Excel.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum Value {
    /// An empty cell.
    #[default]
    Empty,
    /// UTF-8 text.
    String(String),
    /// A signed integer.
    Int(i64),
    /// A floating-point number.
    Float(f64),
    /// A boolean.
    Bool(bool),
    /// A date/time without a timezone, as stored in spreadsheets.
    DateTime(NaiveDateTime),
    /// An elapsed time.
    Duration(Duration),
    /// A formula including its leading `=`. Readers return cached results by default.
    Formula(String),
    /// A spreadsheet cell error, such as `#DIV/0!`.
    Error(String),
}

impl Value {
    /// Returns text for a string, formula or error cell.
    #[must_use]
    pub fn as_str(&self) -> Option<&str> {
        match self {
            Self::String(s) | Self::Formula(s) | Self::Error(s) => Some(s),
            _ => None,
        }
    }
}

impl fmt::Display for Value {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Empty => Ok(()),
            Self::String(s) | Self::Formula(s) | Self::Error(s) => f.write_str(s),
            Self::Int(n) => n.fmt(f),
            Self::Float(n) => n.fmt(f),
            Self::Bool(b) => f.write_str(if *b { "1" } else { "0" }),
            Self::DateTime(d) => d.fmt(f),
            Self::Duration(d) => d.fmt(f),
        }
    }
}

impl From<&str> for Value {
    fn from(value: &str) -> Self {
        Self::String(value.into())
    }
}
impl From<String> for Value {
    fn from(value: String) -> Self {
        Self::String(value)
    }
}
impl From<i64> for Value {
    fn from(value: i64) -> Self {
        Self::Int(value)
    }
}
impl From<i32> for Value {
    fn from(value: i32) -> Self {
        Self::Int(i64::from(value))
    }
}
impl From<f64> for Value {
    fn from(value: f64) -> Self {
        Self::Float(value)
    }
}
impl From<bool> for Value {
    fn from(value: bool) -> Self {
        Self::Bool(value)
    }
}
impl From<NaiveDateTime> for Value {
    fn from(value: NaiveDateTime) -> Self {
        Self::DateTime(value)
    }
}
impl From<NaiveDate> for Value {
    fn from(value: NaiveDate) -> Self {
        Self::DateTime(value.and_time(chrono::NaiveTime::MIN))
    }
}
impl From<Duration> for Value {
    fn from(value: Duration) -> Self {
        Self::Duration(value)
    }
}
impl<T: Into<Value>> From<Option<T>> for Value {
    fn from(value: Option<T>) -> Self {
        value.map(Into::into).unwrap_or_default()
    }
}

/// Ordered row values, with optional header keys and XLSX cell formats.
///
/// Pair arrays automatically supply the writer's first header. `Row::new`
/// creates an unkeyed row and does not generate a header.
#[derive(Debug, Clone, Default)]
pub struct Row {
    pub(crate) headers: Option<Vec<String>>,
    pub(crate) values: Vec<Value>,
    pub(crate) formats: Vec<Option<Format>>,
}

impl Row {
    /// Creates a raw row without automatic headers.
    #[must_use]
    pub fn new<T: Into<Value>>(values: impl IntoIterator<Item = T>) -> Self {
        Self {
            values: values.into_iter().map(Into::into).collect(),
            ..Self::default()
        }
    }

    /// Creates a keyed row. Duplicate keys retain their first position and last value.
    #[must_use]
    pub fn from_pairs<K: Into<String>, V: Into<Value>>(
        pairs: impl IntoIterator<Item = (K, V)>,
    ) -> Self {
        let mut headers = Vec::new();
        let mut values = Vec::new();
        for (key, value) in pairs {
            let key = key.into();
            if let Some(index) = headers.iter().position(|h| h == &key) {
                values[index] = value.into();
            } else {
                headers.push(key);
                values.push(value.into());
            }
        }
        Self {
            headers: Some(headers),
            values,
            formats: Vec::new(),
        }
    }

    /// Returns the value at a header key, or `None` for an absent key.
    #[must_use]
    pub fn get(&self, key: &str) -> Option<&Value> {
        self.headers
            .as_ref()?
            .iter()
            .position(|h| h == key)
            .and_then(|i| self.values.get(i))
    }

    /// Returns the header keys, if this is a keyed row.
    #[must_use]
    pub fn headers(&self) -> Option<&[String]> {
        self.headers.as_deref()
    }

    /// Returns values in column order.
    #[must_use]
    pub fn values(&self) -> &[Value] {
        &self.values
    }

    /// Consumes the row, returning its values.
    #[must_use]
    pub fn into_values(self) -> Vec<Value> {
        self.values
    }

    /// Applies an XLSX format to one cell. An out-of-range column is rejected.
    pub fn set_cell_format(
        &mut self,
        column: usize,
        format: Format,
    ) -> Result<&mut Self, crate::Error> {
        if column >= self.values.len() {
            return Err(crate::Error::InvalidInput(
                "Cell format column is outside the row".into(),
            ));
        }
        self.formats.resize(self.values.len(), None);
        self.formats[column] = Some(format);
        Ok(self)
    }
}

impl<K: Into<String>, V: Into<Value>, const N: usize> From<[(K, V); N]> for Row {
    fn from(value: [(K, V); N]) -> Self {
        Self::from_pairs(value)
    }
}

impl<K: Into<String>, V: Into<Value>> From<Vec<(K, V)>> for Row {
    fn from(value: Vec<(K, V)>) -> Self {
        Self::from_pairs(value)
    }
}

impl Index<usize> for Row {
    type Output = Value;
    fn index(&self, index: usize) -> &Self::Output {
        &self.values[index]
    }
}

impl Index<&str> for Row {
    type Output = Value;
    fn index(&self, key: &str) -> &Self::Output {
        self.get(key).expect("Row key does not exist")
    }
}
