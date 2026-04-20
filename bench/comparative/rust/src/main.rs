use serde::{Deserialize, Serialize};
use std::path::Path;
use std::time::Instant;
use umya_spreadsheet::{reader, writer, Workbook};

#[derive(Deserialize)]
struct Spec {
    template: Template,
    manipulation: Manipulation,
}

#[derive(Deserialize)]
struct Template {
    summary_sheets: Vec<SummarySheet>,
    data_sheets: Vec<SheetShape>,
    reference_sheets: Vec<SheetShape>,
}

#[derive(Deserialize)]
struct SummarySheet {
    name: String,
    #[serde(default)]
    merges: Vec<String>,
    #[serde(default)]
    formulas: Vec<(String, String)>,
}

#[derive(Deserialize)]
struct SheetShape {
    name: String,
    rows: u32,
    cols: u32,
}

#[derive(Deserialize)]
struct Manipulation {
    rows_per_data_sheet: u32,
    cols_per_data_sheet: u32,
}

#[derive(Serialize)]
struct Result {
    lang: &'static str,
    mode: String,
    wall_ms: f64,
    cells_written: u64,
    cells_cleared: u64,
}

fn gen_cell(book: &mut Workbook, sheet: &str, row: u32, col: u32) {
    let ws = book.get_sheet_by_name_mut(sheet).expect("sheet exists");
    let cell = ws.get_cell_mut((col, row));
    match col {
        1 => {
            cell.set_value_string(format!("SYM{:04}", row));
        }
        2 => {
            cell.set_value_number(row as f64 * 1.5 + 0.5);
        }
        3 => {
            cell.set_value_number(row as f64 * 0.1 + 0.03);
        }
        4 => {
            cell.set_value_number((row * 4) as f64);
        }
        5 => {
            cell.set_value_string(format!("LBL-{:04}", row));
        }
        6 => {
            cell.set_value_bool(row % 2 == 0);
        }
        7 => {
            cell.set_value_string("B".to_string());
        }
        _ => {
            cell.set_value_number((row + col) as f64);
        }
    }
}

fn write_sheet(book: &mut Workbook, name: &str, rows: u32, cols: u32) -> u64 {
    let mut count = 0u64;
    for row in 1..=rows {
        for col in 1..=cols {
            gen_cell(book, name, row, col);
            count += 1;
        }
    }
    count
}

fn clear_sheet(book: &mut Workbook, name: &str, rows: u32, cols: u32) -> u64 {
    let mut count = 0u64;
    let ws = book.get_sheet_by_name_mut(name).expect("sheet exists");
    for row in 1..=rows {
        for col in 1..=cols {
            let _ = ws.remove_cell((col, row));
            count += 1;
        }
    }
    count
}

fn create(spec: &Spec, output: &Path) -> (u64, u64) {
    let mut book = umya_spreadsheet::new_file();
    // umya's new_file() seeds one sheet ("Sheet1"); rename + add to match spec
    book.remove_sheet_by_name("Sheet1").ok();

    let mut written = 0u64;

    for s in &spec.template.summary_sheets {
        book.new_sheet(&s.name).expect("new sheet");
    }

    for s in &spec.template.data_sheets {
        book.new_sheet(&s.name).expect("new sheet");
        written += write_sheet(&mut book, &s.name, s.rows, s.cols);
    }

    for s in &spec.template.reference_sheets {
        book.new_sheet(&s.name).expect("new sheet");
        written += write_sheet(&mut book, &s.name, s.rows, s.cols);
    }

    for s in &spec.template.summary_sheets {
        let ws = book
            .get_sheet_by_name_mut(&s.name)
            .expect("summary sheet exists");
        for rng in &s.merges {
            ws.add_merge_cells(rng);
        }
        for (cell, formula) in &s.formulas {
            let c = ws.get_cell_mut(cell.as_str());
            c.set_formula(formula.trim_start_matches('='));
        }
    }

    writer::xlsx::write(&book, output).expect("write");
    (written, 0)
}

fn edit(spec: &Spec, input: &Path, output: &Path) -> (u64, u64) {
    let mut book = reader::xlsx::read(input).expect("read");
    let rows = spec.manipulation.rows_per_data_sheet;
    let cols = spec.manipulation.cols_per_data_sheet;

    let data_names: Vec<String> = spec
        .template
        .data_sheets
        .iter()
        .map(|s| s.name.clone())
        .collect();

    let sheet_names: Vec<String> = book
        .get_sheet_collection()
        .iter()
        .map(|ws| ws.name().to_string())
        .collect();

    let present: Vec<String> = data_names
        .into_iter()
        .filter(|n| sheet_names.contains(n))
        .collect();

    let mut cleared = 0u64;
    let mut written = 0u64;

    for name in &present {
        cleared += clear_sheet(&mut book, name, rows, cols);
    }
    for name in &present {
        written += write_sheet(&mut book, name, rows, cols);
    }

    writer::xlsx::write(&book, output).expect("write");
    (written, cleared)
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 5 {
        eprintln!("usage: bench <spec.json> <create|edit> <input|-> <output>");
        std::process::exit(2);
    }
    let spec_path = &args[1];
    let mode = &args[2];
    let input = &args[3];
    let output = &args[4];

    let spec_bytes = std::fs::read(spec_path).expect("read spec");
    let spec: Spec = serde_json::from_slice(&spec_bytes).expect("parse spec");

    let start = Instant::now();
    let (written, cleared) = match mode.as_str() {
        "create" => create(&spec, Path::new(output)),
        "edit" => edit(&spec, Path::new(input), Path::new(output)),
        other => {
            eprintln!("unknown mode: {}", other);
            std::process::exit(2);
        }
    };
    let wall_ms = start.elapsed().as_secs_f64() * 1000.0;

    let result = Result {
        lang: "rust",
        mode: mode.clone(),
        wall_ms: (wall_ms * 1000.0).round() / 1000.0,
        cells_written: written,
        cells_cleared: cleared,
    };
    println!("{}", serde_json::to_string(&result).unwrap());
}
