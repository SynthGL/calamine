# calamine-styles

A fork of [calamine](https://github.com/tafia/calamine) that adds **cell style parsing** for xlsx files — Font, Fill, Border, Alignment, and NumberFormat.

## Relationship to calamine

This crate is API-compatible with `calamine` and re-exports everything from upstream. The key addition is the `StyleInfo` struct returned by `worksheet_range_with_style()`, which gives you:

- **Font**: bold, italic, underline, strikethrough, size, color, name
- **Fill**: pattern type, foreground/background color
- **Border**: left/right/top/bottom/diagonal style and color
- **Alignment**: horizontal, vertical, wrap text, indent, text rotation
- **NumberFormat**: format code string (e.g., `"#,##0.00"`, `"yyyy-mm-dd"`)

## Usage

```toml
[dependencies]
calamine-styles = { version = "0.1", features = ["dates"] }
```

```rust
use calamine_styles::{Reader, Xlsx, open_workbook};

let mut excel: Xlsx<_> = open_workbook("file.xlsx").unwrap();
let range = excel.worksheet_range("Sheet1").unwrap();
for row in range.rows() {
    println!("{:?}", row);
}
```

## Features

Same as calamine:

- `chrono` / `dates`: Adds support for chrono date/time types
- `picture`: Adds support for reading picture data

## License

MIT — same as upstream calamine.
