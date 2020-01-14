use calamine::{Reader, Xlsx, open_workbook};

use clap::{Arg, App};

use std::fs::File;
use std::io::Write;

fn main() {
    let matches = App::new(env!("CARGO_PKG_NAME"))
        .version(env!("CARGO_PKG_VERSION"))
        .author(env!("CARGO_PKG_AUTHORS"))
        .about("Convert excel to csv for git diffing")
        .arg(Arg::with_name("INPUT")
            .help("Select input file")
            .required(true)
            .index(1))
        .arg(Arg::with_name("output")
            .short("o")
            .long("output")
            .takes_value(true)
            .help("Write output to file instead"))
        .get_matches();

    let input_file = matches.value_of("INPUT").unwrap();

    if let Some(output_file) = matches.value_of("output") {
        let mut excel: Xlsx<_> = open_workbook(input_file).unwrap();
        let sheet_names: Vec<_> = excel.sheet_names().into();
        let mut file = File::create(output_file).expect("Failed to create file");
        for sheet_name in sheet_names {
            writeln!(file, "\n\nSheet '{}'", sheet_name).expect("Failed to write to file");
            if let Some(Ok(r)) = excel.worksheet_range(&sheet_name) {
                for row in r.rows() {
                    writeln!(file, "{:?}", row).unwrap();
                }
            }
        }
    } else {
        let mut excel: Xlsx<_> = open_workbook(input_file).unwrap();
        let sheet_names: Vec<_> = excel.sheet_names().into();
        for sheet_name in sheet_names {
            println!("\n\nSheet '{}'", sheet_name);
            if let Some(Ok(r)) = excel.worksheet_range(&sheet_name) {
                for row in r.rows() {
                    println!("{:?}", row);
                }
            }
        }
    }
}



