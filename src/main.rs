use anyhow::{bail, Result};
use calamine::{open_workbook, DataType, Reader, Xlsx};
use clap::Parser;
use itertools::Itertools;
use std::io::{stdout, BufWriter, Write};
use std::path::{Path, PathBuf};

fn get_target_sheet<P: AsRef<Path>, S: AsRef<str>>(
    path: P,
    sheet_name: Option<S>,
) -> Result<calamine::Range<DataType>> {
    let mut book: Xlsx<_> = open_workbook(path)?;
    let sheet_name = match sheet_name {
        Some(name) => name.as_ref().to_string(),
        None => {
            let names = book.sheet_names();
            names[0].clone()
        }
    };

    let range = book.worksheet_range(&sheet_name);
    match range {
        Some(range) => Ok(range?),
        None => bail!("{sheet_name} not found"),
    }
}

fn write<W: Write>(
    mut writer: W,
    sheet: &calamine::Range<DataType>,
    sep: &str,
    eol: &str,
    print_empty_row: bool,
) -> Result<()> {
    let mut line = Vec::with_capacity(sheet.width());
    for row in sheet.rows() {
        line.clear();
        for cell in row.iter() {
            line.push(cell);
        }

        if !print_empty_row && line.iter().all(|x| x.is_empty()) {
            continue;
        }

        writer.write_all(line.iter().join(sep).as_bytes())?;
        writer.write_all(eol.as_bytes())?;
    }

    Ok(())
}

#[derive(Parser, Debug)]
#[command(author, version)]
struct Cli {
    #[clap(value_parser)]
    file: PathBuf,

    sheet: Option<String>,

    #[clap(short, long, value_parser, default_value_t = String::from("\t"))]
    separator: String,

    #[clap(long, action)]
    print_empty_row: bool,

    #[cfg(not(target_os = "windows"))]
    #[clap(long, value_parser, default_value_t = String::from("\n"))]
    eol: String,

    #[cfg(target_os = "windows")]
    #[clap(long, value_parser, default_value_t = String::from("\r\n"))]
    eol: String,
}

fn main() -> Result<()> {
    let cli = Cli::parse();
    let sheet = get_target_sheet(&cli.file, cli.sheet.as_ref())?;

    let out = stdout();
    let writer = BufWriter::new(out.lock());
    write(
        writer,
        &sheet,
        &cli.separator,
        &cli.eol,
        cli.print_empty_row,
    )?;
    Ok(())
}
