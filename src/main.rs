use calamine::{open_workbook, Reader, Xlsx};

use clap::{App, Arg};

use std::fs::File;
use std::io::{self, Write};

fn main() {
    let matches = App::new(env!("CARGO_PKG_NAME"))
        .version(env!("CARGO_PKG_VERSION"))
        .author(env!("CARGO_PKG_AUTHORS"))
        .about("Convert excel to csv for git diffing")
        .arg(
            Arg::with_name("INPUT")
                .help("Select input file")
                .required(true)
                .index(1),
        )
        .arg(
            Arg::with_name("output")
                .short("o")
                .long("output")
                .takes_value(true)
                .help("Write output to file instead"),
        )
        .get_matches();

    let input_file = matches.value_of("INPUT").unwrap();

    let mut output: Box<dyn Write> = if let Some(output_file) = matches.value_of("output") {
        Box::new(File::create(output_file).expect("Failed to create output file"))
    } else {
        Box::new(io::stdout())
    };

    let mut excel: Xlsx<_> = open_workbook(input_file).unwrap();
    let sheet_names: Vec<_> = excel.sheet_names().into();
    for sheet_name in sheet_names {
        writeln!(output, "\n\nSheet '{}'", sheet_name).expect("Failed to write to file");
        if let Some(Ok(r)) = excel.worksheet_range(&sheet_name) {
            for row in r.rows() {
                writeln!(output, "{:?}", row).unwrap();
            }
        }
    }
}
