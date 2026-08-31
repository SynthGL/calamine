# calamine-styles

A maintained fork of [calamine](https://github.com/tafia/calamine) that adds
cell-style parsing for XLSX files: fonts, fills, borders, alignment, and number
formats.

## Relationship to calamine

The package is published as `calamine-styles`, while its Rust library keeps the
`calamine` crate name so existing source imports remain compatible. The current
fork baseline is calamine 0.33; promotion is gated on rebasing the style delta
onto the current upstream release and passing the full test matrix.

The main addition is [`Reader::worksheet_style()`], which returns an
RLE-compressed [`StyleRange`] alongside the existing value-only worksheet
range API:

- Font: bold, italic, underline, strikethrough, size, color, and name
- Fill: pattern type and foreground/background color
- Border: side/diagonal style and color
- Alignment: horizontal, vertical, wrapping, indent, and rotation
- Number format: format-code string such as `#,##0.00` or `yyyy-mm-dd`

## Usage

```toml
[dependencies]
calamine = { package = "calamine-styles", version = "0.1", features = ["dates"] }
```

```rust
use calamine::{open_workbook, Reader, Xlsx};

let mut excel: Xlsx<_> = open_workbook("file.xlsx").unwrap();
let values = excel.worksheet_range("Sheet1").unwrap();
let styles = excel.worksheet_style("Sheet1").unwrap();

println!("{} value rows", values.height());
for (row, column, style) in styles.cells() {
    println!("style at relative ({row}, {column}): {style:?}");
}
```

`StyleRange::get()` uses positions relative to the style range's start, just
like `Range::get()`. For bulk inspection, `StyleRange::cells()` iterates the
compressed range without cloning each `Style`.

## Features

- `chrono` / `dates`: chrono date and time types
- `picture`: raw picture data

## Maintenance contract

This fork does not claim current-upstream parity until the rebase and
conformance work is complete. Releases require the declared MSRV, stable,
beta, and nightly test lanes plus formatting, Clippy, and package validation.

## License

MIT, matching upstream calamine.
