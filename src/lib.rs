#![doc = include_str!("../README.md")]

mod reader;
mod row;
mod writer;

pub use reader::{Rows, SimpleExcelReader};
pub use row::{Row, Value};
pub use rust_xlsxwriter::{self, Format};
pub use writer::{SimpleExcelWriter, WriterEngine, WriterOptions};
pub use {calamine, chrono, spreadsheet_ods};

use std::path::Path;

/// Errors opening, reading, formatting or writing a spreadsheet.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A filesystem or output stream failed.
    #[error(transparent)]
    Io(#[from] std::io::Error),
    /// CSV parsing or writing failed.
    #[error(transparent)]
    Csv(#[from] csv::Error),
    /// A spreadsheet could not be read.
    #[error(transparent)]
    Read(#[from] calamine::Error),
    /// An XLSX file could not be read.
    #[error(transparent)]
    XlsxRead(#[from] calamine::XlsxError),
    /// An XLSX file could not be written.
    #[error(transparent)]
    Xlsx(#[from] rust_xlsxwriter::XlsxError),
    /// An ODS file could not be written.
    #[error(transparent)]
    Ods(#[from] spreadsheet_ods::OdsError),
    /// A file extension is unsupported.
    #[error("Unsupported file type: {0}. Expected csv, xlsx or ods")]
    UnsupportedType(String),
    /// A sheet does not exist. Sheet numbers start at one.
    #[error("Sheet does not exist: {0}")]
    SheetNotFound(String),
    /// A configuration value or cell is invalid.
    #[error("{0}")]
    InvalidInput(String),
    /// An operation requires a different file type.
    #[error("{0}")]
    UnsupportedOperation(&'static str),
    /// The writer has already been finalized.
    #[error("The writer is closed")]
    Closed,
}

/// Supported file types. An explicit type overrides the path's extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileType {
    /// Comma-separated values.
    Csv,
    /// Excel workbook.
    Xlsx,
    /// OpenDocument spreadsheet.
    Ods,
}

impl FileType {
    /// Detects the type from a case-insensitive extension.
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self, Error> {
        match path
            .as_ref()
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_ascii_lowercase()
            .as_str()
        {
            "csv" => Ok(Self::Csv),
            "xlsx" => Ok(Self::Xlsx),
            "ods" => Ok(Self::Ods),
            extension => Err(Error::UnsupportedType(extension.into())),
        }
    }
}

#[cfg(doctest)]
mod guide {
    include!(concat!(env!("OUT_DIR"), "/documentation.rs"));
}
