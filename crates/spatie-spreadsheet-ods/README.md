# Read and write ODS files

This is Spatie's maintained copy of [spreadsheet-ods](https://github.com/thscharler/spreadsheet-ods) by Thomas Scharler. It supplies the ODS writer used by [spatie-simple-excel](https://crates.io/crates/spatie-simple-excel).

The source comes from spreadsheet-ods 1.0.4, upstream commit [ade015fe](https://github.com/thscharler/spreadsheet-ods/tree/ade015fe41d76f4a93e88e2f8263d84494f8e743). The original MIT and Apache-2.0 licenses, copyright notices, documentation and tests are retained. See [the upstream README](README.upstream.md) for the original project description.

This copy updates `quick-xml` to 0.41 and `zip` to at least 9.0.5 to use their security fixes. It uses Rust 2024 with MSRV 1.89. Source changes are limited to dependency API adjustments, toolchain lint fixes and formatting.

## Installation

```sh
cargo add spatie-spreadsheet-ods
```

The library retains the `spreadsheet_ods` import name used by the upstream API. Most applications should use `spatie-simple-excel`, which re-exports the engine as `spatie_simple_excel::spreadsheet_ods`.

## License

Copyright (c) 2014-2022 Thomas Scharler. Licensed under [MIT](LICENSE-MIT) or [Apache-2.0](LICENSE-APACHE), at your option.
