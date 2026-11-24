---
title: Installation
weight: 3
---

Requires Rust 1.89 or newer. Install Rust using [rustup](https://rustup.rs/).

Add it to your application:

```sh
cargo add spatie-simple-excel
```

No office application or external executable is required. The package uses `csv`, `calamine`, `rust_xlsxwriter` and `spatie-spreadsheet-ods` for file handling. The ODS engine is a maintained copy of [Thomas Scharler's spreadsheet-ods](https://github.com/thscharler/spreadsheet-ods), with updated XML and ZIP dependencies.

XLSX exports use temporary disk space for constant-memory worksheets. XLSX and ODS downloads need a writable temporary directory until the completed archive can be copied to your output.
