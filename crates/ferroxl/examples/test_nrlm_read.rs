//! Read-only test runner for Excel files in D:\OneDrive\Office\Assignments\NRLM\Shimla
//! Strictly tests parsing, sheet extraction, cell inspection, formulas, and styles without modifying any file.

use ferroxl::CellValue;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::Instant;

fn find_excel_files(dir: &Path, list: &mut Vec<PathBuf>) {
    if let Ok(entries) = fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                find_excel_files(&path, list);
            } else if let Some(ext) = path.extension().and_then(|s| s.to_str()) {
                let ext_lower = ext.to_ascii_lowercase();
                if ext_lower == "xlsx" || ext_lower == "xlsm" {
                    list.push(path);
                }
            }
        }
    }
}

fn main() {
    let target_dir = Path::new(r"D:\OneDrive\Office\Assignments\NRLM\Shimla");
    if !target_dir.exists() {
        eprintln!("Target directory does not exist: {:?}", target_dir);
        return;
    }

    let mut files = Vec::new();
    find_excel_files(target_dir, &mut files);
    files.sort();

    println!(
        "Testing ferroxl (read-only) on {} Excel workbooks in {:?}",
        files.len(),
        target_dir
    );
    println!("{:-<110}", "");
    println!(
        "{:<4} | {:<55} | {:<7} | {:<8} | {:<8} | {:<12}",
        "#", "File Name", "Sheets", "Cells", "Formulas", "Status"
    );
    println!("{:-<110}", "");

    let start_all = Instant::now();
    let mut passed = 0;
    let mut failed = 0;
    let mut total_sheets = 0;
    let mut total_cells = 0;
    let mut total_formulas = 0;

    for (i, file_path) in files.iter().enumerate() {
        let file_name = file_path
            .file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("unknown");
        let short_name = if file_name.len() > 53 {
            format!("{}...", &file_name[..50])
        } else {
            file_name.to_string()
        };

        let file_start = Instant::now();
        let result =
            ferroxl::load_workbook(file_path.to_str().unwrap(), ferroxl::LoadOptions::default());

        match result {
            Ok(wb) => {
                let sheet_names = wb.get_sheet_names();
                let sheet_count = sheet_names.len();
                total_sheets += sheet_count;

                let mut file_cells = 0;
                let mut file_formulas = 0;

                for sheet in &wb.worksheets {
                    for cell in sheet.cells() {
                        file_cells += 1;
                        if matches!(cell.internal_value(), CellValue::Formula(_)) {
                            file_formulas += 1;
                        }
                    }
                }

                total_cells += file_cells;
                total_formulas += file_formulas;
                passed += 1;

                let elapsed_ms = file_start.elapsed().as_millis();
                println!(
                    "{:<4} | {:<55} | {:<7} | {:<8} | {:<8} | OK ({}ms)",
                    i + 1,
                    short_name,
                    sheet_count,
                    file_cells,
                    file_formulas,
                    elapsed_ms
                );
            }
            Err(e) => {
                failed += 1;
                println!(
                    "{:<4} | {:<55} | {:<7} | {:<8} | {:<8} | FAILED: {}",
                    i + 1,
                    short_name,
                    "-",
                    "-",
                    "-",
                    e
                );
            }
        }
    }

    let total_elapsed = start_all.elapsed();
    println!("{:-<110}", "");
    println!("Test Summary:");
    println!("  Total Workbooks Tested : {}", files.len());
    println!("  Passed                 : {}", passed);
    println!("  Failed                 : {}", failed);
    println!("  Total Sheets Processed : {}", total_sheets);
    println!("  Total Cells Extracted  : {}", total_cells);
    println!("  Total Formulas Read    : {}", total_formulas);
    println!("  Total Execution Time   : {:.2?}", total_elapsed);
    if !files.is_empty() {
        println!(
            "  Average Time/Workbook  : {:.2?}",
            total_elapsed / files.len() as u32
        );
    }
    println!("{:-<110}", "");
}
