// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use quick_xml::{
    events::{BytesStart, Event},
    name::QName,
};
use std::{
    borrow::Borrow,
    collections::HashMap,
    io::{Read, Seek},
};

use super::{
    get_attribute, get_dimension, get_row, get_row_column, read_rich_string, style_parser,
    worksheet_column_span, Dimensions, XlReader, MAX_COLUMNS,
};
use crate::{
    datatype::DataRef,
    formats::{format_excel_f64_ref, CellFormat},
    utils::unescape_entity_to_buffer,
    Cell, Color, Data, Style, XlsxError,
};

type FormulaMap = HashMap<(u32, u32), (i64, i64)>;
type ColorPalettes<'a> = (&'a [Color; 12], &'a [Option<Color>; 64]);

/// An xlsx Cell Iterator
pub struct XlsxCellReader<'a, RS>
where
    RS: Read + Seek,
{
    xml: XlReader<'a, RS>,
    strings: &'a [Data],
    formats: &'a [CellFormat],
    styles: &'a [Style],
    theme_colors: &'a [Color; 12],
    indexed_colors: &'a [Option<Color>; 64],
    is_1904: bool,
    dimensions: Dimensions,
    row_index: u32,
    col_index: u32,
    row_style: Option<usize>,
    column_styles: Vec<Option<usize>>,
    buf: Vec<u8>,
    cell_buf: Vec<u8>,
    formulas: Vec<Option<(String, FormulaMap)>>,
}

fn row_state(
    row_element: &BytesStart<'_>,
    styles_len: usize,
) -> Result<(Option<u32>, Option<usize>), XlsxError> {
    let row_index = get_attribute(row_element.attributes(), QName(b"r"))?
        .map(get_row)
        .transpose()?;
    let custom_format = get_attribute(row_element.attributes(), QName(b"customFormat"))?
        .map(style_parser::parse_ooxml_bool)
        .transpose()?
        .unwrap_or(false);
    let row_style = get_attribute(row_element.attributes(), QName(b"s"))?
        .and_then(|style_id| atoi_simd::parse::<usize>(style_id).ok())
        .filter(|style_id| *style_id < styles_len);

    Ok((row_index, custom_format.then_some(row_style).flatten()))
}

fn resolved_cell_style_id(
    cell_element: &BytesStart<'_>,
    column: u32,
    row_style: Option<usize>,
    column_styles: &[Option<usize>],
    styles_len: usize,
) -> Result<(bool, Option<usize>), XlsxError> {
    if let Some(style_id) = get_attribute(cell_element.attributes(), QName(b"s"))? {
        return Ok((
            true,
            atoi_simd::parse::<usize>(style_id)
                .ok()
                .filter(|style_id| *style_id < styles_len),
        ));
    }

    let inherited = row_style.or_else(|| column_styles.get(column as usize).copied().flatten());
    Ok((false, inherited))
}

impl<'a, RS> XlsxCellReader<'a, RS>
where
    RS: Read + Seek,
{
    pub fn new(
        xml: XlReader<'a, RS>,
        strings: &'a [Data],
        formats: &'a [CellFormat],
        styles: &'a [Style],
        is_1904: bool,
    ) -> Result<Self, XlsxError> {
        Self::new_with_theme(
            xml,
            strings,
            formats,
            styles,
            (
                &style_parser::DEFAULT_THEME_COLORS,
                &style_parser::NO_INDEXED_COLOR_OVERRIDES,
            ),
            is_1904,
        )
    }

    pub(crate) fn new_with_theme(
        mut xml: XlReader<'a, RS>,
        strings: &'a [Data],
        formats: &'a [CellFormat],
        styles: &'a [Style],
        color_palettes: ColorPalettes<'a>,
        is_1904: bool,
    ) -> Result<Self, XlsxError> {
        let mut dimensions = Dimensions {
            start: (0, 0),
            end: (0, 0),
        };
        let mut buf = Vec::with_capacity(1024);
        let mut sheet_type: Option<String> = None;
        let mut column_styles = Vec::new();
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    match e.local_name().as_ref() {
                        b"dimension" => {
                            let attribute = get_attribute(e.attributes(), QName(b"ref"))?;
                            if let Some(range) = attribute {
                                dimensions = get_dimension(range)?;
                            }
                        }
                        b"col" => {
                            let min_column = get_attribute(e.attributes(), QName(b"min"))?
                                .and_then(|value| atoi_simd::parse::<u32>(value).ok());
                            let max_column = get_attribute(e.attributes(), QName(b"max"))?
                                .and_then(|value| atoi_simd::parse::<u32>(value).ok());
                            let style_id = get_attribute(e.attributes(), QName(b"style"))?
                                .and_then(|value| atoi_simd::parse::<usize>(value).ok())
                                .filter(|style_id| *style_id < styles.len());

                            if let (Some(min_column), Some(style_id)) = (min_column, style_id) {
                                if column_styles.is_empty() {
                                    column_styles.resize(MAX_COLUMNS as usize, None);
                                }
                                for column in worksheet_column_span(min_column, max_column)? {
                                    column_styles[column as usize] = Some(style_id);
                                }
                            }
                        }
                        b"sheetData" => {
                            return Ok(Self {
                                xml,
                                strings,
                                formats,
                                styles,
                                theme_colors: color_palettes.0,
                                indexed_colors: color_palettes.1,
                                is_1904,
                                dimensions,
                                row_index: 0,
                                col_index: 0,
                                row_style: None,
                                column_styles,
                                buf: Vec::with_capacity(1024),
                                cell_buf: Vec::with_capacity(1024),
                                formulas: Vec::with_capacity(1024),
                            });
                        }
                        typ => {
                            // Track the type of element we found (for non-worksheet detection)
                            if sheet_type.is_none() {
                                sheet_type = xml.decoder().decode(typ).ok().map(|s| s.to_string());
                            }
                        }
                    }
                }
                Ok(Event::Eof) => {
                    // If we reached EOF without finding sheetData, check if this is a non-worksheet
                    if let Some(typ) = sheet_type {
                        return Err(XlsxError::NotAWorksheet(typ));
                    } else {
                        return Err(XlsxError::XmlEof("sheetData"));
                    }
                }
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    pub fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

    pub fn next_cell(&mut self) -> Result<Option<Cell<DataRef<'a>>>, XlsxError> {
        self.next_cell_with_style(true)
    }

    /// Read the next cell without cloning its style.
    ///
    /// Worksheet value ranges discard `Cell::style`, so their internal reader
    /// path must not materialize a heap-backed `Style` for every cell. The
    /// public streaming API continues to return styled cells through
    /// [`Self::next_cell`].
    pub(crate) fn next_cell_value_only(&mut self) -> Result<Option<Cell<DataRef<'a>>>, XlsxError> {
        self.next_cell_with_style(false)
    }

    fn next_cell_with_style(
        &mut self,
        include_style: bool,
    ) -> Result<Option<Cell<DataRef<'a>>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    let (row_index, row_style) = row_state(&row_element, self.styles.len())?;
                    if let Some(row_index) = row_index {
                        self.row_index = row_index;
                    }
                    self.row_style = row_style;
                }
                Ok(Event::End(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                    self.row_style = None;
                }
                Ok(Event::Start(c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;
                    let pos = if let Some(range) = attribute {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };
                    let mut value = DataRef::Empty;
                    let mut style = None;
                    let (has_cell_override, resolved_style_id) = resolved_cell_style_id(
                        &c_element,
                        pos.1,
                        self.row_style,
                        &self.column_styles,
                        self.styles.len(),
                    )?;

                    if include_style {
                        // The public streaming API exposes the resolved style.
                        // Value-only range construction bypasses this clone.
                        let style_id = resolved_style_id.or_else(|| {
                            (!has_cell_override && !self.styles.is_empty()).then_some(0)
                        });

                        if let Some(style_id) = style_id {
                            let mut resolved_style = self.styles[style_id].clone();
                            resolved_style.style_id = Some(style_id as u32);
                            style = Some(resolved_style);
                        }
                    }

                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(e)) => {
                                value = read_value(
                                    self.strings,
                                    self.formats,
                                    (self.theme_colors, self.indexed_colors),
                                    self.is_1904,
                                    &mut self.xml,
                                    &e,
                                    (&c_element, resolved_style_id),
                                )?;
                            }
                            Ok(Event::End(e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;

                    if let Some(cell_style) = style {
                        return Ok(Some(Cell::with_style(pos, value, cell_style)));
                    } else {
                        return Ok(Some(Cell::new(pos, value)));
                    }
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    pub fn next_formula(&mut self) -> Result<Option<Cell<String>>, XlsxError> {
        self.next_formula_with_style(true)
    }

    /// Read the next formula without cloning its style for range construction.
    pub(crate) fn next_formula_value_only(&mut self) -> Result<Option<Cell<String>>, XlsxError> {
        self.next_formula_with_style(false)
    }

    fn next_formula_with_style(
        &mut self,
        include_style: bool,
    ) -> Result<Option<Cell<String>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    let (row_index, row_style) = row_state(&row_element, self.styles.len())?;
                    if let Some(row_index) = row_index {
                        self.row_index = row_index;
                    }
                    self.row_style = row_style;
                }
                Ok(Event::End(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                    self.row_style = None;
                }
                Ok(Event::Start(c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;
                    let pos = if let Some(range) = attribute {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };
                    let mut value = None;
                    let mut style = None;

                    if include_style {
                        let (has_cell_override, inherited_style_id) = resolved_cell_style_id(
                            &c_element,
                            pos.1,
                            self.row_style,
                            &self.column_styles,
                            self.styles.len(),
                        )?;
                        let style_id = inherited_style_id.or_else(|| {
                            (!has_cell_override && !self.styles.is_empty()).then_some(0)
                        });

                        if let Some(style_id) = style_id {
                            let mut resolved_style = self.styles[style_id].clone();
                            resolved_style.style_id = Some(style_id as u32);
                            style = Some(resolved_style);
                        }
                    }

                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(e)) => {
                                let formula = read_formula(&mut self.xml, &e)?;
                                if let Some(f) = formula.borrow() {
                                    value = Some(f.clone());
                                }
                                if let Ok(Some(b"shared")) =
                                    get_attribute(e.attributes(), QName(b"t"))
                                {
                                    // shared formula
                                    let mut offset_map: HashMap<(u32, u32), (i64, i64)> =
                                        HashMap::new();
                                    // shared index
                                    let shared_index =
                                        match get_attribute(e.attributes(), QName(b"si"))? {
                                            Some(res) => match atoi_simd::parse::<usize>(res) {
                                                Ok(res) => res,
                                                Err(_) => {
                                                    return Err(XlsxError::Unexpected(
                                                        "si attribute must be a number",
                                                    ));
                                                }
                                            },
                                            None => {
                                                return Err(XlsxError::Unexpected(
                                                    "si attribute is mandatory if it is shared",
                                                ));
                                            }
                                        };
                                    // shared reference
                                    match get_attribute(e.attributes(), QName(b"ref"))? {
                                        Some(res) => {
                                            // original reference formula
                                            let reference = get_dimension(res)?;

                                            for row in reference.start.0..=reference.end.0 {
                                                for column in reference.start.1..=reference.end.1 {
                                                    offset_map.insert(
                                                        (row, column),
                                                        (
                                                            row as i64 - pos.0 as i64,
                                                            column as i64 - pos.1 as i64,
                                                        ),
                                                    );
                                                }
                                            }

                                            if let Some(f) = formula.borrow() {
                                                if self.formulas.len() <= shared_index {
                                                    self.formulas.resize(shared_index + 1, None);
                                                }
                                                self.formulas[shared_index] =
                                                    Some((f.clone(), offset_map));
                                            }
                                            value = formula;
                                        }
                                        None => {
                                            // This cell uses an existing shared formula - look it up and apply offset
                                            if let Some(Some((base_formula, offset_map))) =
                                                self.formulas.get(shared_index)
                                            {
                                                if let Some(offset) = offset_map.get(&pos) {
                                                    value = Some(super::replace_cell_names(
                                                        base_formula,
                                                        *offset,
                                                    )?);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            Ok(Event::End(e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;

                    if let Some(cell_style) = style {
                        return Ok(Some(Cell::with_style(
                            pos,
                            value.unwrap_or_default(),
                            cell_style,
                        )));
                    } else {
                        return Ok(Some(Cell::new(pos, value.unwrap_or_default())));
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    pub fn next_style(&mut self) -> Result<Option<Cell<Style>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(ref row_element))
                    if row_element.local_name().as_ref() == b"row" =>
                {
                    let (row_index, row_style) = row_state(row_element, self.styles.len())?;
                    if let Some(row_index) = row_index {
                        self.row_index = row_index;
                    }
                    self.row_style = row_style;
                }
                Ok(Event::End(ref row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                    self.row_style = None;
                }
                Ok(Event::Start(ref c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;
                    let pos = if let Some(range) = attribute {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };

                    let (_, style_id) = resolved_cell_style_id(
                        c_element,
                        pos.1,
                        self.row_style,
                        &self.column_styles,
                        self.styles.len(),
                    )?;
                    let style = style_id
                        .map(|style_id| {
                            let mut style = self.styles[style_id].clone();
                            style.style_id = Some(style_id as u32);
                            style
                        })
                        .unwrap_or_default();

                    // Skip the cell content since we only care about the style
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some(Cell::new(pos, style)));
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    /// Iterate over cells, returning just position and style_id (no clone).
    ///
    /// Returns `(row, col, style_id)` where `style_id` is an index into the styles palette.
    /// This is more efficient than `next_style()` when building compressed style storage.
    pub fn next_style_id(&mut self) -> Result<Option<(u32, u32, usize)>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(ref row_element))
                    if row_element.local_name().as_ref() == b"row" =>
                {
                    let (row_index, row_style) = row_state(row_element, self.styles.len())?;
                    if let Some(row_index) = row_index {
                        self.row_index = row_index;
                    }
                    self.row_style = row_style;
                }
                Ok(Event::End(ref row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                    self.row_style = None;
                }
                Ok(Event::Start(ref c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;
                    let pos = if let Some(range) = attribute {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };

                    // An explicit cell style wins over row and column defaults.
                    // Row formatting applies only when customFormat is enabled;
                    // otherwise the applicable column style is inherited.
                    let (_, style_id) = resolved_cell_style_id(
                        c_element,
                        pos.1,
                        self.row_style,
                        &self.column_styles,
                        self.styles.len(),
                    )?;

                    // Skip the cell content since we only care about the style ID
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;

                    // Only return cells with valid explicit or inherited styles.
                    if let Some(style_id) = style_id {
                        return Ok(Some((pos.0, pos.1, style_id)));
                    }
                    // Continue to next cell if no style
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    /// Get the styles palette (reference to avoid clones)
    pub fn styles(&self) -> &[Style] {
        self.styles
    }
}

fn read_value<'s, RS>(
    strings: &'s [Data],
    formats: &[CellFormat],
    color_palettes: ColorPalettes<'_>,
    is_1904: bool,
    xml: &mut XlReader<'_, RS>,
    e: &BytesStart<'_>,
    cell: (&BytesStart<'_>, Option<usize>),
) -> Result<DataRef<'s>, XlsxError>
where
    RS: Read + Seek,
{
    Ok(match e.local_name().as_ref() {
        b"is" => {
            // inlineStr
            match read_rich_string(xml, e.name(), color_palettes.0, color_palettes.1)? {
                Some(Data::String(value)) => DataRef::String(value),
                Some(Data::RichText(value)) => DataRef::RichText(value),
                Some(_) => {
                    return Err(XlsxError::Unexpected(
                        "inline string parser returned a non-string value",
                    ))
                }
                None => DataRef::Empty,
            }
        }
        b"v" => {
            // value
            let mut v = String::new();
            let mut v_buf = Vec::new();
            loop {
                v_buf.clear();
                match xml.read_event_into(&mut v_buf)? {
                    Event::Text(t) => v.push_str(&t.xml10_content()?),
                    Event::GeneralRef(e) => unescape_entity_to_buffer(&e, &mut v)?,
                    Event::End(end) if end.name() == e.name() => break,
                    Event::Eof => return Err(XlsxError::XmlEof("v")),
                    _ => (),
                }
            }
            read_v(v, strings, formats, cell.0, cell.1, is_1904)?
        }
        b"f" => {
            xml.read_to_end_into(e.name(), &mut Vec::new())?;
            DataRef::Empty
        }
        _n => return Err(XlsxError::UnexpectedNode("v, f, or is")),
    })
}

/// read the contents of a <v> cell
fn read_v<'s>(
    v: String,
    strings: &'s [Data],
    formats: &[CellFormat],
    c_element: &BytesStart<'_>,
    style_id: Option<usize>,
    is_1904: bool,
) -> Result<DataRef<'s>, XlsxError> {
    let cell_format = style_id.and_then(|style_id| formats.get(style_id));
    match get_attribute(c_element.attributes(), QName(b"t"))? {
        Some(b"s") => {
            // Cell value is an index into the shared string table.
            let idx = atoi_simd::parse::<usize>(v.as_bytes()).unwrap_or(0);
            match strings.get(idx) {
                Some(Data::String(s)) => Ok(DataRef::SharedString(s)),
                Some(Data::RichText(rt)) => Ok(DataRef::SharedRichText(rt)),
                Some(_) => Err(XlsxError::Unexpected(
                    "Unexpected data type in shared strings table",
                )),
                None => Err(XlsxError::Unexpected(
                    "Cell string index not found in shared strings table",
                )),
            }
        }
        Some(b"b") => {
            // boolean
            Ok(DataRef::Bool(v != "0"))
        }
        Some(b"e") => {
            // error
            Ok(DataRef::Error(v.parse()?))
        }
        Some(b"d") => {
            // date
            Ok(DataRef::DateTimeIso(v))
        }
        Some(b"str") => {
            // string
            Ok(DataRef::String(v))
        }
        Some(b"n") => {
            // n - number
            if v.is_empty() {
                Ok(DataRef::Empty)
            } else {
                v.parse()
                    .map(|n| format_excel_f64_ref(n, cell_format, is_1904))
                    .map_err(XlsxError::ParseFloat)
            }
        }
        None => {
            // If type is not known, we try to parse as Float for utility, but fall back to
            // String if this fails.
            v.parse()
                .map(|n| format_excel_f64_ref(n, cell_format, is_1904))
                .or(Ok(DataRef::String(v)))
        }
        Some(b"is") => {
            // this case should be handled in outer loop over cell elements, in which
            // case read_inline_str is called instead. Case included here for completeness.
            Err(XlsxError::Unexpected(
                "called read_value on a cell of type inlineStr",
            ))
        }
        Some(t) => {
            let t = std::str::from_utf8(t).unwrap_or("<utf8 error>").to_string();
            Err(XlsxError::CellTAttribute(t))
        }
    }
}

fn read_formula<RS>(xml: &mut XlReader<RS>, e: &BytesStart) -> Result<Option<String>, XlsxError>
where
    RS: Read + Seek,
{
    match e.local_name().as_ref() {
        b"is" | b"v" => {
            xml.read_to_end_into(e.name(), &mut Vec::new())?;
            Ok(None)
        }
        b"f" => {
            let mut f_buf = Vec::with_capacity(512);
            let mut f = String::new();
            loop {
                match xml.read_event_into(&mut f_buf)? {
                    Event::Text(t) => f.push_str(&t.xml10_content()?),
                    Event::GeneralRef(e) => unescape_entity_to_buffer(&e, &mut f)?,
                    Event::End(end) if end.name() == e.name() => break,
                    Event::Eof => return Err(XlsxError::XmlEof("f")),
                    _ => (),
                }
                f_buf.clear();
            }
            Ok(Some(f))
        }
        _ => Err(XlsxError::UnexpectedNode("v, f, or is")),
    }
}
