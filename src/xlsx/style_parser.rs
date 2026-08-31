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

/// Default Office theme in SpreadsheetML's effective zero-based index order:
/// `lt1`, `dk1`, `lt2`, `dk2`, six accents, hyperlink, and followed hyperlink.
pub(super) const DEFAULT_THEME_COLORS: [Color; 12] = [
    Color {
        alpha: 255,
        red: 255,
        green: 255,
        blue: 255,
    },
    Color {
        alpha: 255,
        red: 0,
        green: 0,
        blue: 0,
    },
    Color {
        alpha: 255,
        red: 238,
        green: 236,
        blue: 225,
    },
    Color {
        alpha: 255,
        red: 31,
        green: 73,
        blue: 125,
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

pub(super) const NO_INDEXED_COLOR_OVERRIDES: [Option<Color>; 64] = [None; 64];

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
pub(super) fn get_indexed_color(index: u8, indexed_colors: &[Option<Color>; 64]) -> Color {
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

    if let Some(color) = indexed_colors.get(usize::from(index)).copied().flatten() {
        return color;
    }
    let (red, green, blue) = OOXML_INDEXED_COLORS
        .get(usize::from(index))
        .copied()
        .unwrap_or((0, 0, 0));
    Color::rgb(red, green, blue)
}

fn parse_hex_component(value: &[u8], component: &'static str) -> Result<u8, XlsxError> {
    let value =
        std::str::from_utf8(value).map_err(|_| XlsxError::Unexpected("Invalid color value"))?;
    u8::from_str_radix(value, 16).map_err(|_| XlsxError::Unexpected(component))
}

fn parse_rgb_color(value: &[u8]) -> Result<Color, XlsxError> {
    match value.len() {
        6 => Ok(Color::rgb(
            parse_hex_component(&value[0..2], "Invalid red color value")?,
            parse_hex_component(&value[2..4], "Invalid green color value")?,
            parse_hex_component(&value[4..6], "Invalid blue color value")?,
        )),
        8 => Ok(Color::new(
            parse_hex_component(&value[0..2], "Invalid alpha color value")?,
            parse_hex_component(&value[2..4], "Invalid red color value")?,
            parse_hex_component(&value[4..6], "Invalid green color value")?,
            parse_hex_component(&value[6..8], "Invalid blue color value")?,
        )),
        _ => Err(XlsxError::Unexpected("Invalid RGB color length")),
    }
}

fn hue_to_rgb(p: f64, q: f64, mut hue: f64) -> f64 {
    if hue < 0.0 {
        hue += 1.0;
    }
    if hue > 1.0 {
        hue -= 1.0;
    }
    if hue < 1.0 / 6.0 {
        p + (q - p) * 6.0 * hue
    } else if hue < 0.5 {
        q
    } else if hue < 2.0 / 3.0 {
        p + (q - p) * (2.0 / 3.0 - hue) * 6.0
    } else {
        p
    }
}

fn apply_tint(color: Color, tint: f64) -> Color {
    if tint == 0.0 {
        return color;
    }
    let red = f64::from(color.red) / 255.0;
    let green = f64::from(color.green) / 255.0;
    let blue = f64::from(color.blue) / 255.0;
    let max = red.max(green).max(blue);
    let min = red.min(green).min(blue);
    let mut hue = 0.0;
    let mut saturation = 0.0;
    let luminance = (max + min) / 2.0;

    if max != min {
        let delta = max - min;
        saturation = if luminance > 0.5 {
            delta / (2.0 - max - min)
        } else {
            delta / (max + min)
        };
        hue = if max == red {
            (green - blue) / delta + if green < blue { 6.0 } else { 0.0 }
        } else if max == green {
            (blue - red) / delta + 2.0
        } else {
            (red - green) / delta + 4.0
        } / 6.0;
    }

    let tinted_luminance = if tint < 0.0 {
        luminance * (1.0 + tint)
    } else {
        luminance * (1.0 - tint) + tint
    };
    let (red, green, blue) = if saturation == 0.0 {
        (tinted_luminance, tinted_luminance, tinted_luminance)
    } else {
        let q = if tinted_luminance < 0.5 {
            tinted_luminance * (1.0 + saturation)
        } else {
            tinted_luminance + saturation - tinted_luminance * saturation
        };
        let p = 2.0 * tinted_luminance - q;
        (
            hue_to_rgb(p, q, hue + 1.0 / 3.0),
            hue_to_rgb(p, q, hue),
            hue_to_rgb(p, q, hue - 1.0 / 3.0),
        )
    };
    // Convert the transformed HSL channels back to the nearest 8-bit value.
    let to_byte = |component: f64| (component * 255.0).round().clamp(0.0, 255.0) as u8;
    Color::new(color.alpha, to_byte(red), to_byte(green), to_byte(blue))
}

/// Parse and resolve color attributes, including workbook palettes and tint.
pub(super) fn parse_color(
    attributes: &[Attribute],
    theme_colors: &[Color; 12],
    indexed_colors: &[Option<Color>; 64],
) -> Result<Option<Color>, XlsxError> {
    let mut color = None;
    let mut tint = None;
    for attr in attributes {
        match attr.key.as_ref() {
            b"rgb" => {
                color = Some(parse_rgb_color(attr.value.as_ref())?);
            }
            b"theme" => {
                let theme_value = std::str::from_utf8(attr.value.as_ref())
                    .map_err(|_| XlsxError::Unexpected("Invalid theme color index"))?
                    .parse::<u8>()
                    .map_err(|_| XlsxError::Unexpected("Invalid theme color index"))?;
                color = Some(get_theme_color(theme_value, theme_colors));
            }
            b"indexed" => {
                let indexed_value = std::str::from_utf8(attr.value.as_ref())
                    .map_err(|_| XlsxError::Unexpected("Invalid indexed color"))?
                    .parse::<u8>()
                    .map_err(|_| XlsxError::Unexpected("Invalid indexed color"))?;
                color = Some(get_indexed_color(indexed_value, indexed_colors));
            }
            b"tint" => {
                let value = std::str::from_utf8(attr.value.as_ref())
                    .map_err(|_| XlsxError::Unexpected("Invalid color tint"))?
                    .parse::<f64>()
                    .map_err(|_| XlsxError::Unexpected("Invalid color tint"))?;
                if !value.is_finite() || !(-1.0..=1.0).contains(&value) {
                    return Err(XlsxError::Unexpected("Invalid color tint"));
                }
                tint = Some(value);
            }
            _ => {}
        }
    }
    Ok(color.map(|color| tint.map_or(color, |tint| apply_tint(color, tint))))
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
        "centerContinuous" => HorizontalAlignment::CenterContinuous,
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
        "mediumDashDot" => BorderStyle::MediumDashDot,
        "mediumDashDotDot" => BorderStyle::MediumDashDotDot,
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
    indexed_colors: &[Option<Color>; 64],
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
                    let mut bold = true;
                    for attribute in e.attributes() {
                        let attribute = attribute?;
                        if attribute.key.as_ref() == b"val" {
                            bold = parse_ooxml_bool(&attribute.value)?;
                            break;
                        }
                    }
                    font = font.with_weight(if bold {
                        FontWeight::Bold
                    } else {
                        FontWeight::Normal
                    });
                }
                b"i" => {
                    let mut italic = true;
                    for attribute in e.attributes() {
                        let attribute = attribute?;
                        if attribute.key.as_ref() == b"val" {
                            italic = parse_ooxml_bool(&attribute.value)?;
                            break;
                        }
                    }
                    font = font.with_style(if italic {
                        FontStyle::Italic
                    } else {
                        FontStyle::Normal
                    });
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
                    let mut strikethrough = true;
                    for attribute in e.attributes() {
                        let attribute = attribute?;
                        if attribute.key.as_ref() == b"val" {
                            strikethrough = parse_ooxml_bool(&attribute.value)?;
                            break;
                        }
                    }
                    font = font.with_strikethrough(strikethrough);
                }
                b"color" => {
                    if let Some(color) = parse_color(
                        &e.attributes().collect::<Result<Vec<_>, _>>()?,
                        theme_colors,
                        indexed_colors,
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
    indexed_colors: &[Option<Color>; 64],
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
                        indexed_colors,
                    )? {
                        fill = fill.with_foreground_color(color);
                    }
                }
                b"bgColor" => {
                    if let Some(color) = parse_color(
                        &e.attributes().collect::<Result<Vec<_>, _>>()?,
                        theme_colors,
                        indexed_colors,
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
    indexed_colors: &[Option<Color>; 64],
) -> Result<Border, XlsxError> {
    let attributes = element.attributes().collect::<Result<Vec<_>, _>>()?;
    let mut style = BorderStyle::None;
    for attribute in &attributes {
        if attribute.key.as_ref() == b"style" {
            style = parse_border_style(&String::from_utf8_lossy(&attribute.value));
        }
    }

    Ok(
        match parse_color(&attributes, theme_colors, indexed_colors)? {
            Some(color) => Border::with_color(style, color),
            None => Border::new(style),
        },
    )
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
    indexed_colors: &[Option<Color>; 64],
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
                let mut border = border_from_element(element, theme_colors, indexed_colors)?;
                let mut inner_buf = Vec::new();

                loop {
                    inner_buf.clear();
                    match xml.read_event_into(&mut inner_buf) {
                        Ok(Event::Start(ref inner)) if inner.local_name().as_ref() == b"color" => {
                            let attributes = inner.attributes().collect::<Result<Vec<_>, _>>()?;
                            border.color = parse_color(&attributes, theme_colors, indexed_colors)?;
                            xml.read_to_end_into(inner.name(), &mut Vec::new())?;
                        }
                        Ok(Event::Empty(ref inner)) if inner.local_name().as_ref() == b"color" => {
                            let attributes = inner.attributes().collect::<Result<Vec<_>, _>>()?;
                            border.color = parse_color(&attributes, theme_colors, indexed_colors)?;
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
                let border = border_from_element(element, theme_colors, indexed_colors)?;
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
                let rotation = String::from_utf8_lossy(&attr.value).parse::<u16>()?;
                let rotation = match rotation {
                    0..=180 => TextRotation::Degrees(rotation),
                    255 => TextRotation::Stacked,
                    _ => return Err(XlsxError::Unexpected("invalid text rotation")),
                };
                alignment = alignment.with_text_rotation(rotation);
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
    // Excel interprets omitted protection attributes as `locked=true` and
    // `hidden=false`, even though CT_CellProtection declares no schema default.
    let mut protection = Protection::new().with_locked(true);

    for attr in start_elem.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"locked" => {
                protection = protection.with_locked(parse_ooxml_bool(&attr.value)?);
            }
            b"hidden" => {
                protection = protection.with_hidden(parse_ooxml_bool(&attr.value)?);
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
        parse_font(
            &mut xml,
            &start,
            &DEFAULT_THEME_COLORS,
            &NO_INDEXED_COLOR_OVERRIDES,
        )
    }

    fn parse_border_xml(source: &[u8]) -> Result<Borders, XlsxError> {
        let mut xml = Reader::from_reader(Cursor::new(source));
        let mut buf = Vec::new();
        let start = match xml.read_event_into(&mut buf).unwrap() {
            Event::Start(element) => element.into_owned(),
            event => panic!("expected border start, got {event:?}"),
        };
        parse_border(
            &mut xml,
            &start,
            &DEFAULT_THEME_COLORS,
            &NO_INDEXED_COLOR_OVERRIDES,
        )
    }

    fn parse_color_xml(source: &[u8]) -> Result<Option<Color>, XlsxError> {
        let mut xml = Reader::from_reader(Cursor::new(source));
        let mut buf = Vec::new();
        let element = match xml.read_event_into(&mut buf).unwrap() {
            Event::Empty(element) => element.into_owned(),
            event => panic!("expected color element, got {event:?}"),
        };
        let attributes = element.attributes().collect::<Result<Vec<_>, _>>()?;
        parse_color(
            &attributes,
            &DEFAULT_THEME_COLORS,
            &NO_INDEXED_COLOR_OVERRIDES,
        )
    }

    fn parse_alignment_xml(source: &[u8]) -> Result<Alignment, XlsxError> {
        let mut xml = Reader::from_reader(Cursor::new(source));
        let mut buf = Vec::new();
        let element = match xml.read_event_into(&mut buf).unwrap() {
            Event::Start(element) | Event::Empty(element) => element.into_owned(),
            event => panic!("expected alignment element, got {event:?}"),
        };
        parse_alignment(&mut xml, &element)
    }

    fn parse_protection_xml(source: &[u8]) -> Result<Protection, XlsxError> {
        let mut xml = Reader::from_reader(Cursor::new(source));
        let mut buf = Vec::new();
        let element = match xml.read_event_into(&mut buf).unwrap() {
            Event::Start(element) | Event::Empty(element) => element.into_owned(),
            event => panic!("expected protection element, got {event:?}"),
        };
        parse_protection(&mut xml, &element)
    }

    #[test]
    fn indexed_colors_use_zero_based_ooxml_offsets() {
        let indexed = &NO_INDEXED_COLOR_OVERRIDES;
        assert_eq!(get_indexed_color(0, indexed), Color::rgb(0, 0, 0));
        assert_eq!(get_indexed_color(1, indexed), Color::rgb(255, 255, 255));
        assert_eq!(get_indexed_color(2, indexed), Color::rgb(255, 0, 0));
        assert_eq!(get_indexed_color(7, indexed), Color::rgb(0, 255, 255));
        assert_eq!(get_indexed_color(63, indexed), Color::rgb(51, 51, 51));
    }

    #[test]
    fn theme_colors_use_ooxml_slot_order() {
        assert_eq!(
            get_theme_color(0, &DEFAULT_THEME_COLORS),
            Color::rgb(255, 255, 255)
        );
        assert_eq!(
            get_theme_color(1, &DEFAULT_THEME_COLORS),
            Color::rgb(0, 0, 0)
        );
        assert_eq!(
            get_theme_color(2, &DEFAULT_THEME_COLORS),
            Color::rgb(238, 236, 225)
        );
        assert_eq!(
            get_theme_color(3, &DEFAULT_THEME_COLORS),
            Color::rgb(31, 73, 125)
        );
        assert_eq!(
            get_theme_color(10, &DEFAULT_THEME_COLORS),
            Color::rgb(0, 0, 255)
        );
    }

    #[test]
    fn theme_tint_uses_ooxml_hsl_luminance_transform() {
        assert_eq!(
            parse_color_xml(br#"<color tint="0.4" theme="4"/>"#).unwrap(),
            Some(Color::rgb(0x95, 0xB3, 0xD7))
        );
        assert_eq!(
            parse_color_xml(br#"<color theme="4" tint="-1"/>"#).unwrap(),
            Some(Color::rgb(0, 0, 0))
        );
    }

    #[test]
    fn malformed_or_out_of_range_tint_is_rejected() {
        for source in [
            br#"<color theme="4" tint="invalid"/>"#.as_slice(),
            br#"<color theme="4" tint="NaN"/>"#.as_slice(),
            br#"<color theme="4" tint="1.01"/>"#.as_slice(),
            br#"<color theme="4" tint="-1.01"/>"#.as_slice(),
        ] {
            assert!(parse_color_xml(source).is_err());
        }
    }

    #[test]
    fn font_family_uses_val_attribute() {
        let font = parse_font_xml(br#"<font><family val="2"/></font>"#).unwrap();
        assert_eq!(font.family.as_deref(), Some("2"));
    }

    #[test]
    fn font_emphasis_honors_ooxml_boolean_values() {
        let font = parse_font_xml(br#"<font><b/><i/><strike/></font>"#).unwrap();
        assert!(font.is_bold());
        assert!(font.is_italic());
        assert!(font.has_strikethrough());

        for value in ["0", "false"] {
            let source = format!(
                r#"<font><b val="{value}"/><i val="{value}"/><strike val="{value}"/></font>"#
            );
            let font = parse_font_xml(source.as_bytes()).unwrap();
            assert!(!font.is_bold());
            assert!(!font.is_italic());
            assert!(!font.has_strikethrough());
        }

        let font =
            parse_font_xml(br#"<font><b val="true"/><i val="1"/><strike val="true"/></font>"#)
                .unwrap();
        assert!(font.is_bold());
        assert!(font.is_italic());
        assert!(font.has_strikethrough());

        assert!(parse_font_xml(br#"<font><b val="sometimes"/></font>"#).is_err());
        assert!(parse_font_xml(br#"<font><i val="sometimes"/></font>"#).is_err());
        assert!(parse_font_xml(br#"<font><strike val="sometimes"/></font>"#).is_err());
    }

    #[test]
    fn text_rotation_255_maps_to_stacked_text() {
        assert_eq!(
            parse_alignment_xml(br#"<alignment textRotation="255"/>"#)
                .unwrap()
                .text_rotation,
            TextRotation::Stacked
        );
        assert_eq!(
            parse_alignment_xml(br#"<alignment textRotation="180"/>"#)
                .unwrap()
                .text_rotation,
            TextRotation::Degrees(180)
        );
        assert!(parse_alignment_xml(br#"<alignment textRotation="181"/>"#).is_err());
        assert!(parse_alignment_xml(br#"<alignment textRotation="invalid"/>"#).is_err());
    }

    #[test]
    fn alignment_preserves_every_ooxml_horizontal_and_vertical_token() {
        for (token, expected) in [
            ("general", HorizontalAlignment::General),
            ("left", HorizontalAlignment::Left),
            ("center", HorizontalAlignment::Center),
            ("right", HorizontalAlignment::Right),
            ("fill", HorizontalAlignment::Fill),
            ("justify", HorizontalAlignment::Justify),
            ("centerContinuous", HorizontalAlignment::CenterContinuous),
            ("distributed", HorizontalAlignment::Distributed),
        ] {
            assert_eq!(parse_horizontal_alignment(token), expected);
        }
        for (token, expected) in [
            ("top", VerticalAlignment::Top),
            ("center", VerticalAlignment::Center),
            ("bottom", VerticalAlignment::Bottom),
            ("justify", VerticalAlignment::Justify),
            ("distributed", VerticalAlignment::Distributed),
        ] {
            assert_eq!(parse_vertical_alignment(token), expected);
        }

        let center_across =
            parse_alignment_xml(br#"<alignment horizontal="centerContinuous"/>"#).unwrap();
        assert_eq!(
            center_across.horizontal,
            HorizontalAlignment::CenterContinuous
        );
        assert!(Style::new()
            .with_alignment(center_across)
            .has_visible_properties());
    }

    #[test]
    fn protection_uses_excel_locked_default_and_ooxml_boolean_values() {
        let protection = parse_protection_xml(br#"<protection hidden="1"/>"#).unwrap();
        assert!(protection.locked);
        assert!(protection.hidden);

        let protection =
            parse_protection_xml(br#"<protection locked="false" hidden="0"/>"#).unwrap();
        assert!(!protection.locked);
        assert!(!protection.hidden);

        assert!(parse_protection_xml(br#"<protection locked="sometimes"/>"#).is_err());
    }

    #[test]
    fn medium_dash_dot_border_variants_are_preserved() {
        let borders = parse_border_xml(
            br#"<border><left style="mediumDashDot"/><right style="mediumDashDotDot"/></border>"#,
        )
        .unwrap();

        assert_eq!(borders.left.style, BorderStyle::MediumDashDot);
        assert_eq!(borders.right.style, BorderStyle::MediumDashDotDot);
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
