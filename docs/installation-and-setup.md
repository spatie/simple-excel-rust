---
title: Installation
weight: 3
---

Requires Rust 1.89 or newer. Install Rust using [rustup](https://rustup.rs/).

Add it to your application:

```sh
cargo add spatie-simple-excel
```

No office application, PHP runtime or external executable is required. The package uses `csv`, `calamine`, `rust_xlsxwriter` and `spreadsheet-ods` for file handling.

XLSX exports use temporary disk space for constant-memory worksheets. XLSX and ODS downloads need a writable temporary directory until the completed archive can be copied to your output.
