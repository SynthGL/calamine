# Style reading in Calamine

Calamine exposes the cell formatting stored in an XLSX workbook through
`Reader::worksheet_style()`. The method returns a `StyleRange`: a compact,
run-length-encoded view of the explicit cell styles on one worksheet.

Value ranges and style ranges are separate. `worksheet_range()` returns cell
values; its cells do not contain styles. Use `worksheet_style()` when you need
formatting, and `worksheet_range()` separately when you also need values.

## Reading worksheet styles

```rust,no_run
use calamine::{open_workbook, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook("file.xlsx")?;
    let styles = workbook.worksheet_style("Sheet1")?;

    println!("distinct worksheet styles: {}", styles.unique_style_count());
    println!("RLE runs: {}", styles.run_count());

    if let Some((start_row, start_column)) = styles.start() {
        for (row_offset, column_offset, style) in styles.cells() {
            if !style.has_visible_properties() {
                continue;
            }

            // StyleRange iterator coordinates are relative to its start.
            let row = start_row + row_offset as u32;
            let column = start_column + column_offset as u32;
            println!("styled cell at ({row}, {column}): {style:?}");
        }
    }

    Ok(())
}
```

`StyleRange::start()` and `StyleRange::end()` are absolute, zero-based worksheet
coordinates. `StyleRange::get()` and the coordinates returned by
`StyleRange::cells()` are relative to `start()`. Sparse gaps inside the bounding
rectangle return the default empty `Style`; positions outside it return `None`.

The palette is compacted per worksheet. `unique_style_count()` therefore counts
the styles referenced by that sheet, excluding the synthesized empty style used
for sparse gaps.

## Inspecting a style

```rust,no_run
use calamine::{open_workbook, HorizontalAlignment, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook("file.xlsx")?;
    let styles = workbook.worksheet_style("Sheet1")?;

    for (row, column, style) in styles.cells() {
        if let Some(font) = style.get_font() {
            println!("({row}, {column}) font name: {:?}", font.name);
            println!("font size: {:?}", font.size);
            println!("bold: {}", font.is_bold());
            println!("italic: {}", font.is_italic());
            println!("underlined: {}", font.has_underline());
            println!("struck through: {}", font.has_strikethrough());

            if let Some(color) = font.color {
                println!(
                    "ARGB({}, {}, {}, {})",
                    color.alpha, color.red, color.green, color.blue
                );
            }
        }

        if let Some(fill) = style.get_fill() {
            if fill.is_visible() {
                println!("fill pattern: {:?}", fill.pattern);
                println!("fill color: {:?}", fill.get_color());
            }
        }

        if let Some(borders) = style.get_borders() {
            if borders.has_visible_borders() {
                println!("left border: {:?}", borders.left.style);
                println!("right border: {:?}", borders.right.style);
                println!("top border: {:?}", borders.top.style);
                println!("bottom border: {:?}", borders.bottom.style);
            }
        }

        if let Some(alignment) = style.get_alignment() {
            if alignment.horizontal == HorizontalAlignment::Center {
                println!("center aligned");
            }
            if alignment.wrap_text {
                println!("text wrapping enabled");
            }
            println!("vertical alignment: {:?}", alignment.vertical);
            println!("text rotation: {:?}", alignment.text_rotation);
        }

        if let Some(number_format) = style.get_number_format() {
            println!("number format ID: {:?}", number_format.format_id);
            println!("number format code: {:?}", number_format.format_code);
        }
    }

    Ok(())
}
```

Alignment fields are concrete values, not `Option`s: an omitted OOXML property
is represented by the enum or boolean default. Font name, size, color, and
family remain optional because the source record may omit them.

For locale-dependent or unknown built-in number formats, `format_code` is empty
and `format_id` preserves the workbook's numeric identifier. An empty code must
not be interpreted as `General`.

## Random access

`StyleRange::get((row, column))` takes coordinates relative to the range start:

```rust,no_run
use calamine::{open_workbook, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook("file.xlsx")?;
    let styles = workbook.worksheet_style("Sheet1")?;

    if let Some(style) = styles.get((0, 0)) {
        println!("style at the range's top-left cell: {style:?}");
    }

    Ok(())
}
```

Call `start()` first when translating an absolute worksheet coordinate into a
relative `get()` coordinate.

## Related APIs

- `worksheet_range()` reads values as `Range<Data>`.
- `worksheet_style()` reads explicit cell formatting as `StyleRange`.
- `worksheet_layout()` reads column widths, row heights, defaults, and layout
  flags as `WorksheetLayout`.
- `worksheet_cells_reader()` provides XLSX streaming access. Its styled path
  exposes cell style information; its value-only path deliberately avoids
  cloning styles.
- Rich shared and inline strings may be returned as `Data::RichText`. Call
  `RichText::plain_text()` when formatting runs are not needed.

## Supported style properties

The XLSX reader extracts:

- font name, size, family, weight, style, underline, strikethrough, and color;
- fill pattern plus foreground and background colors;
- left, right, top, bottom, and diagonal borders;
- horizontal and vertical alignment, text rotation, wrapping, indentation, and
  shrink-to-fit;
- number format ID and code; and
- cell protection flags.

Theme colors, indexed colors, and OOXML tint values are resolved using the
workbook's theme and indexed palette when those parts are present.

The `Reader` trait also exposes `worksheet_style()` for XLS, XLSB, and ODS so
generic code can compile across workbook types. Those readers currently return
an empty `StyleRange` after validating the worksheet name; populated style
extraction is currently implemented for XLSX.

## Boundaries

- Conditional-formatting rules are not evaluated into cell styles.
- Drawing and chart formatting are outside the worksheet cell-style API.
- A `StyleRange` covers explicit cell style records, not every default inherited
  from application rendering behavior.
