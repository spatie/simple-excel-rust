# Read and write simple Excel and CSV files

[![Latest version on crates.io](https://img.shields.io/crates/v/spatie-simple-excel.svg?style=flat-square)](https://crates.io/crates/spatie-simple-excel)
[![Tests](https://img.shields.io/github/actions/workflow/status/spatie/simple-excel-rust/ci.yml?branch=main&style=flat-square)](https://github.com/spatie/simple-excel-rust/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat-square)](https://github.com/spatie/simple-excel-rust/blob/main/LICENSE.md)
[![Total downloads](https://img.shields.io/crates/d/spatie-simple-excel.svg?style=flat-square)](https://crates.io/crates/spatie-simple-excel)

This package allows you to easily read and write simple Excel and CSV files. It is a Rust port of [spatie/simple-excel](https://github.com/spatie/simple-excel).

CSV reads and writes stream row by row. XLSX writes use constant-memory worksheets backed by temporary files. Spreadsheet reads and ODS writes are not constant-memory. See the memory usage sections in the [reading guide](https://github.com/spatie/simple-excel-rust/blob/main/docs/reading-files.md#memory-usage) and [writing guide](https://github.com/spatie/simple-excel-rust/blob/main/docs/writing-files.md#memory-usage).

## Support us

[<img src="https://spatie.be/github-ad-image/simple-excel-rust.jpg" alt="Support Spatie's open source work" width="419px" />](https://spatie.be/github-ad-click/simple-excel-rust)

We invest a lot of resources into creating [best in class open source packages](https://spatie.be/open-source). You can support us by [buying one of our paid products](https://spatie.be/open-source/support-us).

We highly appreciate you sending us a postcard from your hometown, mentioning which of our package(s) you are using. You'll find our address on [our contact page](https://spatie.be/about-us). We publish all received postcards on [our virtual postcard wall](https://spatie.be/open-source/postcards).

## Installation

Requires Rust 1.89 or newer.

```sh
cargo add spatie-simple-excel
```

## Usage

Here's an example on how to read an Excel or CSV file.

```rust
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let path_to_file = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/header-and-rows.csv");
use spatie_simple_excel::SimpleExcelReader;

for row in SimpleExcelReader::create(path_to_file)?.get_rows()? {
    let row = row?;
    println!("{}", row["email"]);
#   assert!(row["email"].as_str().unwrap().contains('@'));
}
# Ok(())
# }
```

If the path ends with `.csv` a CSV file is assumed. If it ends with `.xlsx`, an Excel file is assumed. `.ods` files are supported too.

Here's how you can write a CSV file:

```rust
# fn main() -> Result<(), spatie_simple_excel::Error> {
# let directory = tempfile::tempdir()?;
# let path_to_csv = directory.path().join("export.csv");
use spatie_simple_excel::SimpleExcelWriter;

let mut writer = SimpleExcelWriter::create(&path_to_csv)?;
writer
    .add_row([("first_name", "John"), ("last_name", "Doe")])?
    .add_row([("first_name", "Jane"), ("last_name", "Doe")])?;
writer.close()?;
# assert_eq!(std::fs::read_to_string(path_to_csv)?, "\u{feff}first_name,last_name\nJohn,Doe\nJane,Doe\n");
# Ok(())
# }
```

Process rows with Rust iterators. Call `close()` to observe write errors; destruction only attempts best-effort finalization.

See the [documentation](https://spatie.be/docs/simple-excel-rust/v1/introduction), [reading guide](https://github.com/spatie/simple-excel-rust/blob/main/docs/reading-files.md) and [writing guide](https://github.com/spatie/simple-excel-rust/blob/main/docs/writing-files.md).

## Testing

```sh
cargo test --locked --all-targets --all-features
cargo test --locked --doc --all-features
```

See [all quality checks](https://github.com/spatie/simple-excel-rust/blob/main/docs/testing.md). The README and Markdown guides are tested by rustdoc, including setup on hidden `# ` lines.

## Changelog

See [CHANGELOG.md](https://github.com/spatie/simple-excel-rust/blob/main/CHANGELOG.md).

## Contributing

See [CONTRIBUTING.md](https://github.com/spatie/simple-excel-rust/blob/main/CONTRIBUTING.md) and the [Rust API Guidelines](https://rust-lang.github.io/api-guidelines/).

Using an AI coding agent? Load Apollo's [rust-best-practices skill](https://github.com/apollographql/skills/tree/main/skills/rust-best-practices) when writing or reviewing code in this package.

## Security Vulnerabilities

Report vulnerabilities privately using the instructions in [SECURITY.md](https://github.com/spatie/simple-excel-rust/blob/main/SECURITY.md).

## Credits

- [Freek Van der Herten](https://github.com/freekmurze)
- [All contributors](https://github.com/spatie/simple-excel-rust/graphs/contributors)

## License

The MIT License (MIT). See [LICENSE.md](https://github.com/spatie/simple-excel-rust/blob/main/LICENSE.md).
