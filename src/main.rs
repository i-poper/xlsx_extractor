use clap::{error::ErrorKind, ArgAction, CommandFactory, Parser};
use clap_derive::{Parser, ValueEnum};
use csv::{QuoteStyle, Writer, WriterBuilder};
use std::collections::HashMap;
use std::{error::Error, io};
use umya_spreadsheet::helper::coordinate::CellCoordinates;
use umya_spreadsheet::*;
use unescape::unescape;

#[derive(Debug, Default, serde::Deserialize)]
#[serde(deny_unknown_fields)]
struct Config {
    #[serde(default)]
    builtin_formats: FormatConfig,
}

#[derive(Debug, Default, Clone, serde::Deserialize)]
#[serde(transparent)]
struct FormatConfig {
    formats: HashMap<String, String>,
}

const BUILTIN_FORMATS: &[(u32, &str)] = &[
    (14, "short_date"),
    (15, "date_abbr_month"),
    (16, "day_abbr_month"),
    (17, "abbr_month_year"),
    (18, "time_ampm"),
    (19, "time_seconds_ampm"),
    (20, "short_time"),
    (21, "long_time"),
    (22, "short_date_time"),
];

impl FormatConfig {
    fn get(&self, num_fmt_id: u32) -> Option<&str> {
        let name = BUILTIN_FORMATS
            .iter()
            .find_map(|(id, name)| (*id == num_fmt_id).then_some(*name))?;
        self.formats.get(name).map(String::as_str)
    }

    fn set(&mut self, name: &str, format: String) -> Result<(), String> {
        if !is_builtin_format_name(name) {
            return Err(format!("unknown built-in format name `{name}`"));
        }
        self.formats.insert(name.to_string(), format);
        Ok(())
    }

    fn validate(&self) -> Result<(), String> {
        for name in self.formats.keys() {
            if !is_builtin_format_name(name) {
                return Err(format!("unknown built-in format name `{name}`"));
            }
        }
        Ok(())
    }
}

fn is_builtin_format_name(name: &str) -> bool {
    BUILTIN_FORMATS
        .iter()
        .any(|(_id, format_name)| *format_name == name)
}

#[derive(Parser, Debug)]
#[command(version, about="Tool to extract data from xlsx(xlsm) by specifying headers.", long_about = None)]
struct Args {
    /// Excel file(.xlsx or .xlsm)
    #[arg(short = 'f', long = "file")]
    xlsx: String,
    /// Output delimiter
    #[arg(short, long, value_parser = escaped_u8, default_value="\t")]
    delimiter: u8,
    /// Sheet name
    #[arg(short, long)]
    sheet: Option<String>,
    /// Suppress header output
    #[arg(short='H', long, action=ArgAction::SetFalse, default_value_t=true)]
    header: bool,
    /// Quote
    #[arg(short, long, value_parser = escaped_u8, default_value="\"")]
    quote: u8,
    /// Quote Style
    #[arg(short='t', long, value_enum, default_value_t=Style::Necessary)]
    style: Style,
    /// Place the output into <FILE>
    #[arg(short = 'o', long = "output")]
    file: Option<String>,
    /// Header names
    #[arg(value_parser = escaped_string)]
    headers: Vec<String>,
    /// Config file
    #[arg(short = 'c', long = "config")]
    config_file: Option<String>,
    /// Set the date and time formats
    /// Key words:
    /// short_date, short_date_time, short_time, long_time
    ///
    /// Example:
    /// -X 'short_date_time=yyyy/m/d h:mm'
    #[arg(
        short = 'X', long = "format", value_name = "NAME=FORMAT",
        value_parser = parse_format_override, verbatim_doc_comment
    )]
    format_overrides: Vec<(String, String)>,
}

///
fn parse_format_override(s: &str) -> Result<(String, String), String> {
    let (name, format) = s
        .split_once('=')
        .ok_or_else(|| "format override must be NAME=FORMAT".to_string())?;

    if name.is_empty() {
        return Err("format override name must not be empty".to_string());
    }

    if format.is_empty() {
        return Err("format override value must not be empty".to_string());
    }

    Ok((name.to_string(), format.to_string()))
}

#[derive(Debug, Copy, Clone, PartialEq, Eq, PartialOrd, Ord, ValueEnum)]
enum Style {
    Always,
    Necessary,
    NonNumeric,
    Never,
}

impl From<Style> for QuoteStyle {
    fn from(value: Style) -> Self {
        match value {
            Style::Always => QuoteStyle::Always,
            Style::Necessary => QuoteStyle::Necessary,
            Style::NonNumeric => QuoteStyle::NonNumeric,
            Style::Never => QuoteStyle::Never,
        }
    }
}

/// header information
struct HeaderInfo {
    /// Header row
    row: u32,
    /// Header column positions
    header_column: Vec<u32>,
}

//================================================================================
/// Extraction context for a single worksheet.
///
/// It finds rows by header names and converts cell values to output text using
/// the configured formatting options.
struct WorksheetExtractor<'a> {
    sheet: &'a Worksheet,
    options: FormatConfig,
}

//================================================================================
impl<'a> WorksheetExtractor<'a> {
    fn new(sheet: &'a Worksheet, options: FormatConfig) -> Self {
        WorksheetExtractor { sheet, options }
    }

    //--------------------------------------------------------------------------------
    /// Iterates over non-empty data rows matched by the specified headers.
    ///
    /// The header row is detected from `headers`. Each yielded row contains cell
    /// text values in the same order as `headers`.
    fn get_iterator<'b>(&'b self, headers: &[String]) -> impl Iterator<Item = Vec<String>> + 'b {
        // Find a Header
        let header = self.find_header(headers).unwrap_or_else(|| {
            invalid_value("`[HEADERS]...` not found.");
        });

        let start_row = header.row + 1;
        let end_row = self.sheet.highest_row();
        let columns = header.header_column;

        (start_row..=end_row)
            .map(move |row| {
                columns
                    .iter()
                    .map(|x| self.text((*x, row)))
                    .collect::<Vec<String>>()
            })
            .filter(|x| !x.iter().all(|y| y.is_empty()))
    }

    //--------------------------------------------------------------------------------
    /// Find the rows in the sheet that contain all the characters specified as headers.
    fn find_header(&self, headers: &[String]) -> Option<HeaderInfo> {
        for row in 1..=self.sheet.highest_row() {
            if let Some(header_column) = self.find_header_in_row(row, headers) {
                return Some(HeaderInfo { row, header_column });
            }
        }
        None
    }

    //--------------------------------------------------------------------------------
    /// Determines whether a header exists in a row.
    fn find_header_in_row(&self, row: u32, headers: &[String]) -> Option<Vec<u32>> {
        let mut indexes: Vec<(&String, Option<u32>)> =
            headers.iter().map(|x| (x, None as Option<u32>)).collect();

        for col in 1..=self.sheet.highest_column() {
            let text = self.text((col, row));
            if text.is_empty() || !headers.iter().any(|x| x == &text) {
                continue;
            }

            if let Some(h) = indexes.iter_mut().find(|(h, c)| *h == &text && c.is_none()) {
                h.1 = Some(col);
            }
            if indexes.iter().all(|(_, x)| x.is_some()) {
                return Some(indexes.into_iter().map(|(_, c)| c.unwrap()).collect());
            }
        }
        None
    }

    //--------------------------------------------------------------------------------
    /// get Cell text
    fn text<T>(&self, coordinate: T) -> String
    where
        T: Into<CellCoordinates>,
    {
        let coordinate: CellCoordinates = coordinate.into();
        match self.sheet.cell_value(coordinate.clone()).raw_value() {
            CellRawValue::Numeric(value) => {
                if let Some(format) = self.sheet.style(coordinate.clone()).number_format() {
                    let format_id = format.number_format_id();
                    let format_code = if let Some(format) = self.options.get(format_id) {
                        format
                    } else {
                        format.format_code()
                    };
                    if format_code != NumberingFormat::FORMAT_GENERAL {
                        if let Ok(text) = ssfmt::format_default(*value, format_code) {
                            return text.trim_end().to_string();
                        }
                    }
                }
                self.sheet.formatted_value(coordinate)
            }
            _ => self.sheet.formatted_value(coordinate),
        }
    }
}

//--------------------------------------------------------------------------------
/// Parsing of escape sequence char
fn escaped_u8(s: &str) -> Result<u8, String> {
    let d = unescape(s).ok_or(format!("`{s}` is not a valid escape string."))?;
    let d = d.as_bytes();
    if d.len() != 1 {
        return Err("Specified by ASCII characters.".to_string());
    }
    Ok(d[0])
}

//--------------------------------------------------------------------------------
/// Parsing of escape sequence strings
fn escaped_string(s: &str) -> Result<String, String> {
    unescape(s).ok_or(format!("`{s}` is not a valid escape string."))
}

//--------------------------------------------------------------------------------
/// Loads configuration from a TOML file.
///
/// Returns the default configuration when `path` is `None`.
fn load_config_file(path: Option<&str>) -> Config {
    let config = if let Some(file) = path {
        toml::from_str::<Config>(
            &std::fs::read_to_string(&file)
                .unwrap_or_else(|e| invalid_value(&format!("Can't read file `{file}`: {e}"))),
        )
        .unwrap_or_else(|e| invalid_value(&format!("Invalid config: {e}")))
    } else {
        Config::default()
    };
    if let Err(e) = config.builtin_formats.validate() {
        invalid_value(&format!("Invalid config: {e}"));
    }
    config
}

//--------------------------------------------------------------------------------
/// Applies command-line format overrides to the configuration.
///
/// Overrides take precedence over values loaded from the config file.
fn apply_format_overrides(config: &mut Config, overrides: &[(String, String)]) {
    for (k, v) in overrides {
        config
            .builtin_formats
            .set(k, v.clone())
            .unwrap_or_else(|e| invalid_value(&e))
    }
}

//================================================================================
/// main
fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    let path = std::path::Path::new(&args.xlsx);
    let mut book = reader::xlsx::lazy_read(path)
        .unwrap_or_else(|e| invalid_value(&format!("Can't read `{}`: {e}", path.display())));

    let mut config = load_config_file(args.config_file.as_deref());
    apply_format_overrides(&mut config, &args.format_overrides);

    // Check sheet name
    let sheet = if let Some(ref sheet_name) = args.sheet {
        let index = book
            .sheet_collection_no_check()
            .iter()
            .position(|x| x.name() == sheet_name)
            .unwrap_or_else(|| invalid_value(&format!("Sheet not found:{}", sheet_name)));
        book.read_sheet(index).sheet(index).unwrap()
    } else {
        book.read_sheet(0_usize)
            .sheet(0_usize)
            .unwrap_or_else(|_| invalid_value("There is no sheet."))
    };

    let mut binding = WriterBuilder::new();
    let builder = binding
        .delimiter(args.delimiter)
        .quote(args.quote)
        .quote_style(args.style.into());

    let extractor = WorksheetExtractor::new(&sheet, config.builtin_formats);
    let mut data_iter = extractor.get_iterator(&args.headers);
    let show_headers = args.header.then_some(args.headers);
    // Output data based on headers
    if let Some(output) = args.file {
        let mut writer = builder
            .from_path(&output)
            .unwrap_or_else(|x| invalid_value(&format!("Can't create {output}: {x}")));
        output_table_data(show_headers, &mut data_iter, &mut writer)
    } else {
        let mut writer = builder.from_writer(io::stdout());
        output_table_data(show_headers, &mut data_iter, &mut writer)
    }
}

//--------------------------------------------------------------------------------
/// Output the data of the columns recognized as headers.
fn output_table_data<W: io::Write>(
    show_header: Option<Vec<String>>,
    data_iter: &mut impl Iterator<Item = Vec<String>>,
    writer: &mut Writer<W>,
) -> Result<(), Box<dyn Error>> {
    // Output headers
    if let Some(header) = show_header {
        writer.write_record(header)?;
    }
    // Output Datas
    data_iter.try_for_each(|x| writer.write_record(x))?;
    Ok(writer.flush()?)
}

//--------------------------------------------------------------------------------
/// Termination process when an illegal value is detected
fn invalid_value(msg: &str) -> ! {
    let mut cmd = Args::command();
    cmd.error(ErrorKind::InvalidValue, msg).exit();
}
