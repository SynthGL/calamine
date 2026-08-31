// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use quick_xml::{
    events::{attributes::Attribute, BytesStart, Event},
    name::QName,
    Reader,
};
use std::io::BufRead;

use crate::style::*;
use crate::utils::unescape_entity_to_buffer;
use crate::XlsxError;

/// Default Office theme in SpreadsheetML's zero-based slot order: `dk1`,
/// `lt1`, `dk2`, `lt2`, six accents, hyperlink, and followed hyperlink.
pub(super) const DEFAULT_THEME_COLORS: [Color; 12] = [
    Color {
        alpha: 255,
        red: 0,
        green: 0,
        blue: 0,
    },
    Color {
        alpha: 255,
        red: 255,
        green: 255,
        blue: 255,
    },
    Color {
        alpha: 255,
        red: 31,
        green: 73,
        blue: 125,
    },
    Color {
        alpha: 255,
        red: 238,
        green: 236,
        blue: 225,
    },
    Color {
        alpha: 255,
        red: 79,
        green: 129,
        blue: 189,
    },
    Color {
        alpha: 255,
        red: 192,
        green: 80,
        blue: 77,
    },
    Color {
        alpha: 255,
        red: 155,
        green: 187,
        blue: 89,
    },
    Color {
        alpha: 255,
        red: 128,
        green: 100,
        blue: 162,
    },
    Color {
        alpha: 255,
        red: 75,
        green: 172,
        blue: 198,
    },
    Color {
        alpha: 255,
        red: 247,
        green: 150,
        blue: 70,
    },
    Color {
        alpha: 255,
        red: 0,
        green: 0,
        blue: 255,
    },
    Color {
        alpha: 255,
        red: 128,
        green: 0,
        blue: 128,
    },
];

pub(super) fn get_theme_color(theme: u8, theme_colors: &[Color; 12]) -> Color {
    theme_colors
        .get(usize::from(theme))
        .copied()
        .unwrap_or(Color::rgb(0, 0, 0))
}

/// Resolve an OOXML indexed color.
///
/// The `indexed` attribute is a zero-based offset into the default
/// `indexedColors` palette, not VBA's one-based `ColorIndex`.
pub(super) fn get_indexed_color(index: u8) -> Color {
    const OOXML_INDEXED_COLORS: [(u8, u8, u8); 64] = [
        (0x00, 0x00, 0x00),
        (0xFF, 0xFF, 0xFF),
        (0xFF, 0x00, 0x00),
        (0x00, 0xFF, 0x00),
        (0x00, 0x00, 0xFF),
        (0xFF, 0xFF, 0x00),
        (0xFF, 0x00, 0xFF),
        (0x00, 0xFF, 0xFF),
        (0x00, 0x00, 0x00),
        (0xFF, 0xFF, 0xFF),
        (0xFF, 0x00, 0x00),
        (0x00, 0xFF, 0x00),
        (0x00, 0x00, 0xFF),
        (0xFF, 0xFF, 0x00),
        (0xFF, 0x00, 0xFF),
        (0x00, 0xFF, 0xFF),
        (0x80, 0x00, 0x00),
        (0x00, 0x80, 0x00),
        (0x00, 0x00, 0x80),
        (0x80, 0x80, 0x00),
        (0x80, 0x00, 0x80),
        (0x00, 0x80, 0x80),
        (0xC0, 0xC0, 0xC0),
        (0x80, 0x80, 0x80),
        (0x99, 0x99, 0xFF),
        (0x99, 0x33, 0x66),
        (0xFF, 0xFF, 0xCC),
        (0xCC, 0xFF, 0xFF),
        (0x66, 0x00, 0x66),
        (0xFF, 0x80, 0x80),
        (0x00, 0x66, 0xCC),
        (0xCC, 0xCC, 0xFF),
        (0x00, 0x00, 0x80),
        (0xFF, 0x00, 0xFF),
        (0xFF, 0xFF, 0x00),
        (0x00, 0xFF, 0xFF),
        (0x80, 0x00, 0x80),
        (0x80, 0x00, 0x00),
        (0x00, 0x80, 0x80),
        (0x00, 0x00, 0xFF),
        (0x00, 0xCC, 0xFF),
        (0xCC, 0xFF, 0xFF),
        (0xCC, 0xFF, 0xCC),
        (0xFF, 0xFF, 0x99),
        (0x99, 0xCC, 0xFF),
        (0xFF, 0x99, 0xCC),
        (0xCC, 0x99, 0xFF),
        (0xFF, 0xCC, 0x99),
        (0x33, 0x66, 0xFF),
        (0x33, 0xCC, 0xCC),
        (0x99, 0xCC, 0x00),
        (0xFF, 0xCC, 0x00),
        (0xFF, 0x99, 0x00),
        (0xFF, 0x66, 0x00),
        (0x66, 0x66, 0x99),
        (0x96, 0x96, 0x96),
        (0x00, 0x33, 0x66),
        (0x33, 0x99, 0x66),
        (0x00, 0x33, 0x00),
        (0x33, 0x33, 0x00),
        (0x99, 0x33, 0x00),
        (0x99, 0x33, 0x66),
        (0x33, 0x33, 0x99),
        (0x33, 0x33, 0x33),
    ];

    let (red, green, blue) = OOXML_INDEXED_COLORS
        .get(usize::from(index))
        .copied()
        .unwrap_or((0, 0, 0));
    Color::rgb(red, green, blue)
}

/// Parse color from XML attributes
fn parse_color(
    attributes: &[Attribute],
    theme_colors: &[Color; 12],
) -> Result<Option<Color>, XlsxError> {
    for attr in attributes {
        match attr.key.as_ref() {
            b"rgb" => {
                let rgb_str = attr.value.as_ref();
                if rgb_str.len() == 6 {
                    // RGB format (6 characters)
                    let r = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[0..2]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid red color value"))?;
                    let g = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[2..4]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid green color value"))?;
                    let b = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[4..6]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid blue color value"))?;
                    return Ok(Some(Color::rgb(r, g, b)));
                } else if rgb_str.len() == 8 {
                    // ARGB format (8 characters)
                    let a = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[0..2]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid alpha color value"))?;
                    let r = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[2..4]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid red color value"))?;
                    let g = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[4..6]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid green color value"))?;
                    let b = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[6..8]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid blue color value"))?;
                    return Ok(Some(Color::new(a, r, g, b)));
                }
            }
            b"theme" => {
                let theme_str = String::from_utf8_lossy(&attr.value);
                if let Ok(theme_value) = theme_str.parse::<u8>() {
                    return Ok(Some(get_theme_color(theme_value, theme_colors)));
                }
            }
            b"indexed" => {
                let indexed_str = String::from_utf8_lossy(&attr.value);
                if let Ok(indexed_value) = indexed_str.parse::<u8>() {
                    return Ok(Some(get_indexed_color(indexed_value)));
                }
            }
            _ => {}
        }
    }
    Ok(None)
}

/// Parse font weight from string
fn parse_font_weight(s: &str) -> FontWeight {
    match s {
        "bold" | "700" => FontWeight::Bold,
        "normal" | "400" => FontWeight::Normal,
        _ => {
            // Try to parse as numeric weight
            if let Ok(weight) = s.parse::<u16>() {
                if weight >= 600 {
                    FontWeight::Bold
                } else {
                    FontWeight::Normal
                }
            } else {
                FontWeight::Normal
            }
        }
    }
}

/// Parse font style from string
fn parse_font_style(s: &str) -> FontStyle {
    match s {
        "italic" | "oblique" => FontStyle::Italic,
        "normal" => FontStyle::Normal,
        _ => FontStyle::Normal,
    }
}

/// Parse underline style from string
fn parse_underline_style(s: &str) -> UnderlineStyle {
    match s {
        "single" => UnderlineStyle::Single,
        "double" => UnderlineStyle::Double,
        "singleAccounting" => UnderlineStyle::SingleAccounting,
        "doubleAccounting" => UnderlineStyle::DoubleAccounting,
        _ => UnderlineStyle::None,
    }
}

/// Parse horizontal alignment from string
fn parse_horizontal_alignment(s: &str) -> HorizontalAlignment {
    match s {
        "left" => HorizontalAlignment::Left,
        "center" => HorizontalAlignment::Center,
        "right" => HorizontalAlignment::Right,
        "justify" => HorizontalAlignment::Justify,
        "distributed" => HorizontalAlignment::Distributed,
        "fill" => HorizontalAlignment::Fill,
        _ => HorizontalAlignment::General,
    }
}

/// Parse vertical alignment from string
fn parse_vertical_alignment(s: &str) -> VerticalAlignment {
    match s {
        "top" => VerticalAlignment::Top,
        "center" => VerticalAlignment::Center,
        "bottom" => VerticalAlignment::Bottom,
        "justify" => VerticalAlignment::Justify,
        "distributed" => VerticalAlignment::Distributed,
        _ => VerticalAlignment::Bottom,
    }
}

/// Parse fill pattern from string
fn parse_fill_pattern(s: &str) -> FillPattern {
    match s {
        "solid" => FillPattern::Solid,
        "darkGray" => FillPattern::DarkGray,
        "mediumGray" => FillPattern::MediumGray,
        "lightGray" => FillPattern::LightGray,
        "gray125" => FillPattern::Gray125,
        "gray0625" => FillPattern::Gray0625,
        "darkHorizontal" => FillPattern::DarkHorizontal,
        "darkVertical" => FillPattern::DarkVertical,
        "darkDown" => FillPattern::DarkDown,
        "darkUp" => FillPattern::DarkUp,
        "darkGrid" => FillPattern::DarkGrid,
        "darkTrellis" => FillPattern::DarkTrellis,
        "lightHorizontal" => FillPattern::LightHorizontal,
        "lightVertical" => FillPattern::LightVertical,
        "lightDown" => FillPattern::LightDown,
        "lightUp" => FillPattern::LightUp,
        "lightGrid" => FillPattern::LightGrid,
        "lightTrellis" => FillPattern::LightTrellis,
        _ => FillPattern::None,
    }
}

/// Parse border style from string
fn parse_border_style(s: &str) -> BorderStyle {
    match s {
        "thin" => BorderStyle::Thin,
        "medium" => BorderStyle::Medium,
        "thick" => BorderStyle::Thick,
        "double" => BorderStyle::Double,
        "hair" => BorderStyle::Hair,
        "dashed" => BorderStyle::Dashed,
        "dotted" => BorderStyle::Dotted,
        "mediumDashed" => BorderStyle::MediumDashed,
        "dashDot" => BorderStyle::DashDot,
        "dashDotDot" => BorderStyle::DashDotDot,
        "slantDashDot" => BorderStyle::SlantDashDot,
        _ => BorderStyle::None,
    }
}

/// Parse font element
pub fn parse_font<RS: BufRead>(
    xml: &mut Reader<RS>,
    _start_elem: &BytesStart,
    theme_colors: &[Color; 12],
) -> Result<Font, XlsxError> {
    let mut font = Font::new();

    // Font elements can have attributes like outline, shadow, etc.
    // TODO(ddimaria): Add specific font attribute parsing here if needed

    // // Parse attributes from the opening font element
    // for attr in start_elem.attributes() {
    //     let attr = attr?;
    //     match attr.key.as_ref() {
    //         _ => {}
    //     }
    // }

    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"name" => {
                    // Check if name is in val attribute first
                    let mut name = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"val" {
                            name = Some(String::from_utf8_lossy(&attr.value).to_string());
                            break;
                        }
                    }
                    // If not in attribute, try reading as text content
                    if name.is_none() {
                        name = read_string(xml, QName(b"name"))?;
                    } else {
                        // Skip to end of element
                        xml.read_to_end_into(e.name(), &mut Vec::new())?;
                    }
                    if let Some(n) = name {
                        font = font.with_name(n);
                    }
                }
                b"sz" => {
                    // Check if size is in val attribute first
                    let mut size_str = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"val" {
                            size_str = Some(String::from_utf8_lossy(&attr.value).to_string());
                            break;
                        }
                    }
                    // If not in attribute, try reading as text content
                    if size_str.is_none() {
                        size_str = read_string(xml, QName(b"sz"))?;
                    } else {
                        // Skip to end of element
                        xml.read_to_end_into(e.name(), &mut Vec::new())?;
                    }
                    if let Some(s) = size_str {
                        if let Ok(size) = s.parse::<f64>() {
                            font = font.with_size(size);
                        }
                    }
                }
                b"b" => {
                    // Check if the element has a 'val' attribute
                    let mut weight = FontWeight::Bold; // Default to bold
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"val" {
                            let val_str = String::from_utf8_lossy(&attr.value);
                            weight = parse_font_weight(&val_str);
                            break;
                        }
                    }
                    font = font.with_weight(weight);
                }
                b"i" => {
                    // Check if the element has a 'val' attribute
                    let mut style = FontStyle::Italic; // Default to italic
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"val" {
                            let val_str = String::from_utf8_lossy(&attr.value);
                            style = parse_font_style(&val_str);
                            break;
                        }
                    }
                    font = font.with_style(style);
                }
                b"u" => {
                    // Check if the element has a 'val' attribute
                    let mut underline_style = UnderlineStyle::Single; // Default to single underline
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"val" {
                            let val_str = String::from_utf8_lossy(&attr.value);
                            underline_style = parse_underline_style(&val_str);
                            break;
                        }
                    }
                    font = font.with_underline(underline_style);
                }
                b"strike" => {
                    font = font.with_strikethrough(true);
                }
                b"color" => {
                    if let Some(color) = parse_color(
                        &e.attributes().collect::<Result<Vec<_>, _>>()?,
                        theme_colors,
                    )? {
                        font = font.with_color(color);
                    }
                }
                b"family" => {
                    let mut family = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"val" {
                            family = Some(String::from_utf8_lossy(&attr.value).to_string());
                            break;
                        }
                    }
                    if family.is_none() {
                        family = read_string(xml, QName(b"family"))?;
                    } else {
                        xml.read_to_end_into(e.name(), &mut Vec::new())?;
                    }
                    if let Some(family) = family {
                        font = font.with_family(family);
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"font" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("font")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(font)
}

/// Parse fill element
pub fn parse_fill<RS: BufRead>(
    xml: &mut Reader<RS>,
    _start_elem: &BytesStart,
    theme_colors: &[Color; 12],
) -> Result<Fill, XlsxError> {
    let mut fill = Fill::new();

    // Fill elements can have attributes like type, etc.
    // TODO(ddimaria): Add specific fill attribute parsing here if needed

    // // Parse attributes from the opening fill element
    // for attr in start_elem.attributes() {
    //     let attr = attr?;
    //     match attr.key.as_ref() {
    //         //
    //         _ => {}
    //     }
    // }

    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"patternFill" => {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"patternType" {
                            let pattern_str = String::from_utf8_lossy(&attr.value);
                            fill = fill.with_pattern(parse_fill_pattern(&pattern_str));
                        }
                    }
                }
                b"fgColor" => {
                    if let Some(color) = parse_color(
                        &e.attributes().collect::<Result<Vec<_>, _>>()?,
                        theme_colors,
                    )? {
                        fill = fill.with_foreground_color(color);
                    }
                }
                b"bgColor" => {
                    if let Some(color) = parse_color(
                        &e.attributes().collect::<Result<Vec<_>, _>>()?,
                        theme_colors,
                    )? {
                        fill = fill.with_background_color(color);
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fill" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("fill")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(fill)
}

pub(super) fn parse_ooxml_bool(value: &[u8]) -> Result<bool, XlsxError> {
    match value {
        b"1" | b"true" => Ok(true),
        b"0" | b"false" => Ok(false),
        _ => Err(XlsxError::Unexpected("invalid OOXML boolean")),
    }
}

fn border_from_element(
    element: &BytesStart<'_>,
    theme_colors: &[Color; 12],
) -> Result<Border, XlsxError> {
    let attributes = element.attributes().collect::<Result<Vec<_>, _>>()?;
    let mut style = BorderStyle::None;
    for attribute in &attributes {
        if attribute.key.as_ref() == b"style" {
            style = parse_border_style(&String::from_utf8_lossy(&attribute.value));
        }
    }

    Ok(match parse_color(&attributes, theme_colors)? {
        Some(color) => Border::with_color(style, color),
        None => Border::new(style),
    })
}

fn apply_border(
    borders: &mut Borders,
    side: &[u8],
    border: Border,
    diagonal_down: bool,
    diagonal_up: bool,
) {
    match side {
        b"left" => borders.left = border,
        b"right" => borders.right = border,
        b"top" => borders.top = border,
        b"bottom" => borders.bottom = border,
        b"diagonal" => {
            if diagonal_down {
                borders.diagonal_down = border.clone();
            }
            if diagonal_up {
                borders.diagonal_up = border;
            }
        }
        _ => {}
    }
}

fn is_border_side(name: &[u8]) -> bool {
    matches!(name, b"left" | b"right" | b"top" | b"bottom" | b"diagonal")
}

/// Parse border element
pub fn parse_border<RS: BufRead>(
    xml: &mut Reader<RS>,
    start_elem: &BytesStart<'_>,
    theme_colors: &[Color; 12],
) -> Result<Borders, XlsxError> {
    let mut borders = Borders::new();
    let mut diagonal_down = false;
    let mut diagonal_up = false;

    // OOXML stores diagonal direction on the parent <border>, not on the
    // <diagonal> child that carries the line style and color.
    for attribute in start_elem.attributes() {
        let attribute = attribute?;
        match attribute.key.as_ref() {
            b"diagonalDown" => diagonal_down = parse_ooxml_bool(&attribute.value)?,
            b"diagonalUp" => diagonal_up = parse_ooxml_bool(&attribute.value)?,
            _ => {}
        }
    }

    let mut buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref element)) if is_border_side(element.local_name().as_ref()) => {
                let side = element.local_name().as_ref().to_vec();
                let closing = element.name();
                let mut border = border_from_element(element, theme_colors)?;
                let mut inner_buf = Vec::new();

                loop {
                    inner_buf.clear();
                    match xml.read_event_into(&mut inner_buf) {
                        Ok(Event::Start(ref inner)) if inner.local_name().as_ref() == b"color" => {
                            let attributes = inner.attributes().collect::<Result<Vec<_>, _>>()?;
                            border.color = parse_color(&attributes, theme_colors)?;
                            xml.read_to_end_into(inner.name(), &mut Vec::new())?;
                        }
                        Ok(Event::Empty(ref inner)) if inner.local_name().as_ref() == b"color" => {
                            let attributes = inner.attributes().collect::<Result<Vec<_>, _>>()?;
                            border.color = parse_color(&attributes, theme_colors)?;
                        }
                        Ok(Event::Start(ref inner)) => {
                            xml.read_to_end_into(inner.name(), &mut Vec::new())?;
                        }
                        Ok(Event::End(ref inner)) if inner.name() == closing => break,
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("border side")),
                        Err(error) => return Err(XlsxError::Xml(error)),
                        _ => {}
                    }
                }

                apply_border(&mut borders, &side, border, diagonal_down, diagonal_up);
            }
            Ok(Event::Empty(ref element)) if is_border_side(element.local_name().as_ref()) => {
                let border = border_from_element(element, theme_colors)?;
                apply_border(
                    &mut borders,
                    element.local_name().as_ref(),
                    border,
                    diagonal_down,
                    diagonal_up,
                );
            }
            Ok(Event::End(ref element)) if element.local_name().as_ref() == b"border" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("border")),
            Err(error) => return Err(XlsxError::Xml(error)),
            _ => {}
        }
    }

    Ok(borders)
}

/// Parse alignment element
pub fn parse_alignment<RS: BufRead>(
    _xml: &mut Reader<RS>,
    start_elem: &BytesStart,
) -> Result<Alignment, XlsxError> {
    let mut alignment = Alignment::new();

    for attr in start_elem.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"horizontal" => {
                let horizontal_str = String::from_utf8_lossy(&attr.value);
                alignment = alignment.with_horizontal(parse_horizontal_alignment(&horizontal_str));
            }
            b"vertical" => {
                let vertical_str = String::from_utf8_lossy(&attr.value);
                alignment = alignment.with_vertical(parse_vertical_alignment(&vertical_str));
            }
            b"wrapText" => {
                let wrap_str = String::from_utf8_lossy(&attr.value);
                if wrap_str == "1" || wrap_str == "true" {
                    alignment = alignment.with_wrap_text(true);
                }
            }
            b"textRotation" => {
                if let Ok(rotation) = String::from_utf8_lossy(&attr.value).parse::<u16>() {
                    alignment = alignment.with_text_rotation(TextRotation::Degrees(rotation));
                }
            }
            b"indent" => {
                if let Ok(indent) = String::from_utf8_lossy(&attr.value).parse::<u8>() {
                    alignment = alignment.with_indent(indent);
                }
            }
            b"shrinkToFit" => {
                let shrink_str = String::from_utf8_lossy(&attr.value);
                if shrink_str == "1" || shrink_str == "true" {
                    alignment = alignment.with_shrink_to_fit(true);
                }
            }
            _ => {}
        }
    }

    Ok(alignment)
}

/// Parse protection element
pub fn parse_protection<RS: BufRead>(
    _xml: &mut Reader<RS>,
    start_elem: &BytesStart,
) -> Result<Protection, XlsxError> {
    let mut protection = Protection::new();

    for attr in start_elem.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"locked" => {
                let locked_str = String::from_utf8_lossy(&attr.value);
                if locked_str == "1" || locked_str == "true" {
                    protection = protection.with_locked(true);
                }
            }
            b"hidden" => {
                let hidden_str = String::from_utf8_lossy(&attr.value);
                if hidden_str == "1" || hidden_str == "true" {
                    protection = protection.with_hidden(true);
                }
            }
            _ => {}
        }
    }

    Ok(protection)
}
/// Read string content from XML element
fn read_string<RS: BufRead>(
    xml: &mut Reader<RS>,
    closing: QName,
) -> Result<Option<String>, XlsxError> {
    let mut buf = Vec::new();
    let mut content = String::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Text(e)) => {
                content.push_str(&e.xml10_content()?);
            }
            Ok(Event::GeneralRef(e)) => {
                unescape_entity_to_buffer(&e, &mut content)?;
            }
            Ok(Event::End(ref e)) if e.local_name() == closing.into() => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("string")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    if content.is_empty() {
        Ok(None)
    } else {
        Ok(Some(content))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    fn parse_font_xml(source: &[u8]) -> Result<Font, XlsxError> {
        let mut xml = Reader::from_reader(Cursor::new(source));
        xml.config_mut().expand_empty_elements = true;
        let mut buf = Vec::new();
        let start = match xml.read_event_into(&mut buf).unwrap() {
            Event::Start(element) => element.into_owned(),
            event => panic!("expected font start, got {event:?}"),
        };
        parse_font(&mut xml, &start, &DEFAULT_THEME_COLORS)
    }

    fn parse_border_xml(source: &[u8]) -> Result<Borders, XlsxError> {
        let mut xml = Reader::from_reader(Cursor::new(source));
        let mut buf = Vec::new();
        let start = match xml.read_event_into(&mut buf).unwrap() {
            Event::Start(element) => element.into_owned(),
            event => panic!("expected border start, got {event:?}"),
        };
        parse_border(&mut xml, &start, &DEFAULT_THEME_COLORS)
    }

    #[test]
    fn indexed_colors_use_zero_based_ooxml_offsets() {
        assert_eq!(get_indexed_color(0), Color::rgb(0, 0, 0));
        assert_eq!(get_indexed_color(1), Color::rgb(255, 255, 255));
        assert_eq!(get_indexed_color(2), Color::rgb(255, 0, 0));
        assert_eq!(get_indexed_color(7), Color::rgb(0, 255, 255));
        assert_eq!(get_indexed_color(63), Color::rgb(51, 51, 51));
    }

    #[test]
    fn theme_colors_use_ooxml_slot_order() {
        assert_eq!(
            get_theme_color(0, &DEFAULT_THEME_COLORS),
            Color::rgb(0, 0, 0)
        );
        assert_eq!(
            get_theme_color(1, &DEFAULT_THEME_COLORS),
            Color::rgb(255, 255, 255)
        );
        assert_eq!(
            get_theme_color(2, &DEFAULT_THEME_COLORS),
            Color::rgb(31, 73, 125)
        );
        assert_eq!(
            get_theme_color(3, &DEFAULT_THEME_COLORS),
            Color::rgb(238, 236, 225)
        );
        assert_eq!(
            get_theme_color(10, &DEFAULT_THEME_COLORS),
            Color::rgb(0, 0, 255)
        );
    }

    #[test]
    fn font_family_uses_val_attribute() {
        let font = parse_font_xml(br#"<font><family val="2"/></font>"#).unwrap();
        assert_eq!(font.family.as_deref(), Some("2"));
    }

    #[test]
    fn diagonal_direction_comes_from_parent_border_attributes() {
        let borders = parse_border_xml(
            br#"<border diagonalUp="1" diagonalDown="true"><diagonal style="thin"><color indexed="2"/></diagonal></border>"#,
        )
        .unwrap();

        assert_eq!(borders.diagonal_up.style, BorderStyle::Thin);
        assert_eq!(borders.diagonal_down.style, BorderStyle::Thin);
        assert_eq!(borders.diagonal_up.color, Some(Color::rgb(255, 0, 0)));
        assert_eq!(borders.diagonal_down.color, Some(Color::rgb(255, 0, 0)));
    }

    #[test]
    fn self_closing_diagonal_side_retains_parent_direction() {
        let borders =
            parse_border_xml(br#"<border diagonalUp="1"><diagonal style="thin"/></border>"#)
                .unwrap();

        assert_eq!(borders.diagonal_up.style, BorderStyle::Thin);
        assert_eq!(borders.diagonal_down.style, BorderStyle::None);
    }

    #[test]
    fn malformed_diagonal_direction_is_rejected() {
        assert!(parse_border_xml(
            br#"<border diagonalUp="sometimes"><diagonal style="thin"/></border>"#,
        )
        .is_err());
    }
}
