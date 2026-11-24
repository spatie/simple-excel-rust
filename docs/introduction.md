---
title: Introduction
weight: 1
---

This package allows you to easily read and write simple Excel and CSV files. CSV reads and writes stream row by row, and XLSX writes spool rows to temporary files. XLSX reading loads the selected sheet and shared strings; ODS reading and writing load the workbook. The [reading](reading-files.md#memory-usage) and [writing](writing-files.md#memory-usage) guides explain memory and disk usage for each format.

Here's an example on how to read an Excel or CSV.

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

If the path ends with `.csv` a CSV file is assumed. If it ends with `.xlsx`, an Excel file is assumed. `.ods` files are supported too. `create_with_type` lets you specify the format explicitly.

Continue with [installation](installation-and-setup.md), [reading files](reading-files.md) and [writing files](writing-files.md).
